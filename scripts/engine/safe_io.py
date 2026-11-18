import os
import shutil
import zipfile
import xml.etree.ElementTree as ET
from typing import Dict, Optional

def is_file_locked(path: str) -> bool:
    """Checks if a file is exclusively locked by another process (e.g. WPS / Excel)."""
    if not os.path.exists(path):
        return False
    try:
        with open(path, 'a+b') as f:
            pass
        return False
    except (PermissionError, IOError):
        return True

def create_backup(path: str, suffix: str = "_备份") -> str:
    """Creates a backup file in the same directory if it does not already exist."""
    if not os.path.exists(path):
        return ""
    base, ext = os.path.splitext(path)
    # Strip any trailing space before ext
    backup_path = f"{base.rstrip()}{suffix}{ext}"
    if not os.path.exists(backup_path):
        shutil.copy2(path, backup_path)
    return backup_path

def atomic_replace_file(src_temp: str, dest_target: str, fallback_suffix: str = "_已同步") -> str:
    """
    Safely moves src_temp to dest_target.
    If dest_target is locked by WPS/Excel, finds an available safe filename instead of crashing.
    """
    final_dest = dest_target
    counter = 1
    while is_file_locked(final_dest):
        base, ext = os.path.splitext(dest_target)
        suffix = fallback_suffix if counter == 1 else f"{fallback_suffix}_{counter}"
        final_dest = f"{base.rstrip()}{suffix}{ext}"
        print(f"[提示] 文件处于锁定状态，安全另存为: {final_dest}")
        counter += 1

    if os.path.exists(final_dest):
        try:
            os.remove(final_dest)
        except Exception:
            pass
    shutil.move(src_temp, final_dest)
    return final_dest

def ensure_cellimage_relationship(workbook_rels_bytes: bytes) -> bytes:
    """Ensures xl/_rels/workbook.xml.rels includes the WPS cellImage relationship."""
    rel_ns = 'http://schemas.openxmlformats.org/package/2006/relationships'
    ET.register_namespace('', rel_ns)
    root = ET.fromstring(workbook_rels_bytes)

    has_ci = False
    max_rid = 0
    for rel in root.findall(f"{{{rel_ns}}}Relationship"):
        rid = rel.attrib.get('Id', '')
        if rid.startswith('rId') and rid[3:].isdigit():
            max_rid = max(max_rid, int(rid[3:]))
        if 'cellImage' in rel.attrib.get('Type', ''):
            has_ci = True

    if not has_ci:
        new_rid = f"rId{max_rid + 1}"
        ET.SubElement(root, f"{{{rel_ns}}}Relationship", {
            'Id': new_rid,
            'Type': 'http://www.wps.cn/officeDocument/2020/cellImage',
            'Target': 'cellimages.xml'
        })
    return ET.tostring(root, encoding='utf-8', xml_declaration=True)

def ensure_cellimage_content_type(content_types_bytes: bytes) -> bytes:
    """Ensures [Content_Types].xml includes the WPS cellimage override and image default extensions."""
    ct_ns = 'http://schemas.openxmlformats.org/package/2006/content-types'
    ET.register_namespace('', ct_ns)
    root = ET.fromstring(content_types_bytes)

    has_ci = any(ov.attrib.get('PartName') == '/xl/cellimages.xml' 
                 for ov in root.findall(f"{{{ct_ns}}}Override"))
    if not has_ci:
        ET.SubElement(root, f"{{{ct_ns}}}Override", {
            'PartName': '/xl/cellimages.xml',
            'ContentType': 'application/vnd.wps-officedocument.cellimage+xml'
        })
    return ET.tostring(root, encoding='utf-8', xml_declaration=True)
