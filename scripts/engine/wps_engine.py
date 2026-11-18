import os
import zipfile
import xml.etree.ElementTree as ET
from typing import Dict, List, Set, Tuple
from .safe_io import ensure_cellimage_relationship, ensure_cellimage_content_type, atomic_replace_file

WPS_NS = {
    'etc': 'http://www.wps.cn/officeDocument/2017/etCustomData',
    'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'rel': 'http://schemas.openxmlformats.org/package/2006/relationships'
}

for prefix, uri in WPS_NS.items():
    if prefix not in ('rel',):
        ET.register_namespace(prefix, uri)
ET.register_namespace('', WPS_NS['rel'])


class WpsImageEngine:
    """Handles parsing, deduplication, and lossless injection of WPS DISPIMG cell images."""

    @staticmethod
    def extract_image_manifest(xlsx_path: str) -> Dict[str, Tuple[str, str]]:
        """
        Extracts image metadata: img_name ("ID_...") -> (rId, media_target_path)
        """
        manifest = {}
        with zipfile.ZipFile(xlsx_path, 'r') as z:
            if 'xl/cellimages.xml' not in z.namelist():
                return manifest
            
            rels = {}
            if 'xl/_rels/cellimages.xml.rels' in z.namelist():
                tree_rels = ET.fromstring(z.read('xl/_rels/cellimages.xml.rels'))
                for r in tree_rels.findall(f"{{{WPS_NS['rel']}}}Relationship"):
                    rels[r.attrib.get('Id')] = r.attrib.get('Target')

            ci_tree = ET.fromstring(z.read('xl/cellimages.xml'))
            for c in ci_tree.findall('.//etc:cellImage', WPS_NS):
                cNvPr = c.find('.//xdr:cNvPr', WPS_NS)
                blip = c.find('.//a:blip', WPS_NS)
                if cNvPr is not None and blip is not None:
                    name = cNvPr.attrib.get('name')
                    embed = None
                    for k, v in blip.attrib.items():
                        if k.endswith('embed'):
                            embed = v
                            break
                    if name and embed and embed in rels:
                        manifest[name] = (embed, rels[embed])
        return manifest

    @staticmethod
    def supplement_missing_images(target_path: str, source_paths: List[str], output_path: str) -> int:
        """
        Merges missing images from source_paths into target_path.
        Returns total number of newly added images.
        """
        with zipfile.ZipFile(target_path, 'r') as z_tgt:
            tgt_files = {name: z_tgt.read(name) for name in z_tgt.namelist()}

        if 'xl/cellimages.xml' in tgt_files:
            ci_root = ET.fromstring(tgt_files['xl/cellimages.xml'])
        else:
            ci_root = ET.Element(f"{{{WPS_NS['etc']}}}cellImages")

        if 'xl/_rels/cellimages.xml.rels' in tgt_files:
            rels_root = ET.fromstring(tgt_files['xl/_rels/cellimages.xml.rels'])
        else:
            rels_root = ET.Element(f"{{{WPS_NS['rel']}}}Relationships")

        existing_names: Set[str] = set()
        max_pic_id = 0
        for cNvPr in ci_root.findall('.//xdr:cNvPr', WPS_NS):
            name = cNvPr.attrib.get('name')
            if name:
                existing_names.add(name)
            pid = cNvPr.attrib.get('id')
            if pid and pid.isdigit():
                max_pic_id = max(max_pic_id, int(pid))

        max_rel_id = 0
        for rel in rels_root.findall(f"{{{WPS_NS['rel']}}}Relationship"):
            rid = rel.attrib.get('Id', '')
            if rid.startswith('rId') and rid[3:].isdigit():
                max_rel_id = max(max_rel_id, int(rid[3:]))

        media_counter = 1
        total_added = 0

        for src_path in source_paths:
            if not os.path.exists(src_path):
                continue
            with zipfile.ZipFile(src_path, 'r') as z_src:
                src_namelist = z_src.namelist()
                if 'xl/cellimages.xml' not in src_namelist or 'xl/_rels/cellimages.xml.rels' not in src_namelist:
                    continue

                src_ci_root = ET.fromstring(z_src.read('xl/cellimages.xml'))
                src_rels_root = ET.fromstring(z_src.read('xl/_rels/cellimages.xml.rels'))

                src_rels = {}
                for rel in src_rels_root.findall(f"{{{WPS_NS['rel']}}}Relationship"):
                    src_rels[rel.attrib.get('Id')] = {
                        'Target': rel.attrib.get('Target'),
                        'Type': rel.attrib.get('Type')
                    }

                for cell_img in src_ci_root.findall('.//etc:cellImage', WPS_NS):
                    cNvPr = cell_img.find('.//xdr:cNvPr', WPS_NS)
                    blip = cell_img.find('.//a:blip', WPS_NS)
                    if cNvPr is None or blip is None:
                        continue

                    img_name = cNvPr.attrib.get('name')
                    if not img_name or img_name in existing_names:
                        continue

                    src_rid = [v for k, v in blip.attrib.items() if k.endswith('embed')]
                    if not src_rid or src_rid[0] not in src_rels:
                        continue

                    rel_info = src_rels[src_rid[0]]
                    src_media = rel_info['Target']
                    src_full = src_media if src_media.startswith('xl/') else f"xl/{src_media}"
                    if src_full not in src_namelist:
                        continue

                    img_bytes = z_src.read(src_full)
                    ext = os.path.splitext(src_media)[1] or '.png'

                    while True:
                        new_media_rel = f"media/supp_{media_counter}{ext}"
                        new_media_full = f"xl/{new_media_rel}"
                        if new_media_full not in tgt_files:
                            break
                        media_counter += 1

                    max_rel_id += 1
                    new_rid = f"rId{max_rel_id}"

                    max_pic_id += 1
                    new_pic_id = str(max_pic_id)

                    tgt_files[new_media_full] = img_bytes

                    ET.SubElement(rels_root, f"{{{WPS_NS['rel']}}}Relationship", {
                        'Id': new_rid,
                        'Type': rel_info['Type'],
                        'Target': new_media_rel
                    })

                    copy_elem = ET.fromstring(ET.tostring(cell_img, encoding='utf-8'))
                    copy_elem.find('.//xdr:cNvPr', WPS_NS).attrib['id'] = new_pic_id
                    blip_copy = copy_elem.find('.//a:blip', WPS_NS)
                    for k in list(blip_copy.attrib.keys()):
                        if k.endswith('embed'):
                            blip_copy.attrib[k] = new_rid

                    ci_root.append(copy_elem)
                    existing_names.add(img_name)
                    total_added += 1

        tgt_files['xl/cellimages.xml'] = ET.tostring(ci_root, encoding='utf-8', xml_declaration=True)
        tgt_files['xl/_rels/cellimages.xml.rels'] = ET.tostring(rels_root, encoding='utf-8', xml_declaration=True)
        tgt_files['xl/_rels/workbook.xml.rels'] = ensure_cellimage_relationship(tgt_files['xl/_rels/workbook.xml.rels'])
        tgt_files['[Content_Types].xml'] = ensure_cellimage_content_type(tgt_files['[Content_Types].xml'])

        temp_out = f"{output_path}.tmp_wps"
        with zipfile.ZipFile(temp_out, 'w', compression=zipfile.ZIP_DEFLATED) as z_out:
            for fname, fbytes in tgt_files.items():
                z_out.writestr(fname, fbytes)

        atomic_replace_file(temp_out, output_path)
        return total_added

    @staticmethod
    def restore_images_to_temp_workbook(orig_xlsx: str, temp_xlsx: str, output_path: str):
        """
        Takes openpyxl-saved temp_xlsx (which lacks cellimages) and restores all original WPS cellimages
        and media files from orig_xlsx into output_path.
        """
        with zipfile.ZipFile(orig_xlsx, 'r') as z_orig:
            orig_media = {name: z_orig.read(name) for name in z_orig.namelist() if name.startswith('xl/media/')}
            orig_ci = z_orig.read('xl/cellimages.xml') if 'xl/cellimages.xml' in z_orig.namelist() else None
            orig_ci_rels = z_orig.read('xl/_rels/cellimages.xml.rels') if 'xl/_rels/cellimages.xml.rels' in z_orig.namelist() else None

        with zipfile.ZipFile(temp_xlsx, 'r') as z_temp:
            temp_files = {name: z_temp.read(name) for name in z_temp.namelist()}

        for m_name, m_bytes in orig_media.items():
            temp_files[m_name] = m_bytes
        if orig_ci:
            temp_files['xl/cellimages.xml'] = orig_ci
        if orig_ci_rels:
            temp_files['xl/_rels/cellimages.xml.rels'] = orig_ci_rels

        temp_files['xl/_rels/workbook.xml.rels'] = ensure_cellimage_relationship(temp_files['xl/_rels/workbook.xml.rels'])
        temp_files['[Content_Types].xml'] = ensure_cellimage_content_type(temp_files['[Content_Types].xml'])

        temp_out = f"{output_path}.atomic_restore"
        with zipfile.ZipFile(temp_out, 'w', compression=zipfile.ZIP_DEFLATED) as z_out:
            for fname, fbytes in temp_files.items():
                z_out.writestr(fname, fbytes)

        atomic_replace_file(temp_out, output_path)
