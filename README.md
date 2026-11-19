# excel-Agent (Universal WPS & Office Excel Toolkit)

> **通用型 WPS / 微软 Office 电商表格自动化操作工具集 · AI Agent 增强套件**
> 专为解决电商/企业级复杂表格中：**WPS `DISPIMG` 浮动/嵌入图片丢失**、**微软 `RichData` (XLRICHVALUE) 跨表补图**、**无损行过滤**、**公式智能平移**与**跨表状态同步**而设计。

---

## 核心特性

- 🖼️ **双引擎图片无损恢复 (`fill-images`)**：
  - 完美支持 **WPS `DISPIMG` 嵌入单元格图片** 与 **Microsoft Office / 腾讯文档 `RichData` (XLRICHVALUE)**。
  - 多规格级联探测：支持 `(链接 + 规格)` -> `(链接 + 名称)` -> `(69码 + 规格)` -> `(编码 + 规格)` 级联匹配，防止多规格混淆。
- ⚡ **物理级零损条件行过滤 (`filter`)**：
  - 完整保留 100% 格式与 6,000+ 嵌入图片。
  - 100MB+ 大底表秒级完成无损筛选（支持表达式如 `毛利率 >= 0.15`）。
- 🔄 **跨表字段/状态智能平移 (`transfer`)**：
  - 同步上下架状态、检测备注、实际售价等字段。
  - **右侧公式自适应偏移**：插入新列时自动修正受影响的右侧公式（如 `(L2-N2)/L2` 自动平移为 `(M2-O2)/M2`）。
- 📊 **表格差异对比与核对 (`diff`)**：
  - 比对新旧两期产品目录，输出价格变更、下架明细差异报告。
- 📌 **表结构感知与合计行保护 (`summary_detector`)**：
  - 智能识别表尾/表中的“合计、小计、Total、Average”等汇总行。
  - 过滤与清洗时自动置于表尾保护，彻底杜绝明细过滤时合计行被误删、或重复计入导致统计翻倍（+161% 业务事故）。
- 📐 **自适应列宽排版计算 (`layout_utils`)**：
  - 精确计算中英文混排视觉宽度（全角 2.0，半角 1.0），预留过滤图标与边框间距，根治 Excel 渲染出的 `###` 及表头文字截断。
- 🔒 **公式注入与提示词越狱拦截 (`security`)**：
  - 拦截包含 `WEBSERVICE`、`DDE`、`EXEC` 等高危外部调用函数，避免数据外带泄漏。
  - 对以 `=`, `+`, `-`, `@` 开头的非数字文本进行安全脱敏（CWE-1236 防护），并能探测单元格中的恶意 Prompt 越狱语句。
- 📑 **多级表头解析与扁平化 (`header_parser`)**：
  - 启发式探测 2~3 层复合多行表头，跨列合并单元格向前填充，自动生成层级扁平化列名（如 `基本信息_商品编码`、`一季度_1月`），准确定位明细起始行。
- 🧹 **脏数据规范化与分列 (`data_cleaner`)**：
  - 假数字清洗：自动去除千分位逗号、货币符号（`¥`、`$`、`元`）、括号负数 `(150.0)`、百分号转小数，将 `"—"`、`"N/A"` 等占位符置为空值。
  - 混杂日期规范化：统一将 Excel 日期序列号（如 `45615`）、中文年月日（`2026年11月19日`）、点号斜杠日期转为标准 ISO `YYYY-MM-DD`。
  - 文本精准分列：支持逗号、分号等多分隔符的中英文自适应文本分列清洗。
- ⚡ **原生联动公式构建器 (`formula_builder`)**：
  - 自动构建带绝对/相对引用锁定的原生 `=XLOOKUP(...)`、全版本通用的 `=INDEX(MATCH(...))`、多条件 `=SUMIFS(...)` 与 `=IFERROR(...)` 容错包裹，杜绝填入硬编码死数据。
- 🔍 **表格类型智能嗅探 (`sniff`)**：
  - 自动识别工作簿采用的是 WPS DISPIMG、微软 RichData、标准 Drawings 还是纯数据。
- 🛡️ **安全 I/O 与冲突保护**：
  - 自动感知 WPS / Excel 文件锁定占用，被占用时安全输出至 `_已同步.xlsx`，防止崩溃与文件损坏。
  - 操作前后自动生成备份文件。

---

## 项目结构

```text
excel-Agent/
├── scripts/
│   ├── cli.py                     # 统一命令行入口
│   ├── engine/                    # 核心底层引擎
│   │   ├── formula_parser.py      # 公式引用平移与列标计算
│   │   ├── matcher.py             # 多规格级联复合主键匹配
│   │   ├── msoffice_engine.py     # 微软 Office RichData 处理引擎
│   │   ├── wps_engine.py          # WPS DISPIMG 处理引擎
│   │   ├── xml_patcher.py         # 流式 XML 补丁器
│   │   ├── safe_io.py             # 原子性文件 I/O 与防锁备份
│   │   ├── sniffer.py             # Excel 表格类型探测
│   │   ├── summary_detector.py    # 汇总行/小计行特征识别引擎
│   │   ├── layout_utils.py        # 排版显示宽度与自适应列宽
│   │   ├── security.py            # 公式注入与 Prompt 注入安全防护
│   │   ├── header_parser.py       # 多级复合表头解析与扁平化
│   │   ├── data_cleaner.py        # 假数字/混杂日期规范化与分列
│   │   └── formula_builder.py     # 原生 XLOOKUP/SUMIFS 联动公式构建器
│   └── ops/                       # 核心业务算子
│       ├── fill_images_op.py      # 补图算子
│       ├── filter_op.py           # 零损过滤算子 (含小计行保护)
│       ├── transfer_op.py         # 状态与字段迁移算子
│       └── diff_op.py             # 双表比对算子
├── tests/
│   └── test_toolkit.py            # 单元测试与端到端回归用例
├── SKILL.md                       # AI Agent Skill 规则规范定义
├── requirements.txt               # Python 依赖
├── .gitignore                     # Git 忽略配置
└── README.md                      # 项目说明文档
```

---

## 快速上手

### 1. 环境安装
```bash
git clone https://github.com/BoBo-v/excel-Agent.git
cd excel-Agent
pip install -r requirements.txt
```

### 2. 命令行调用 (CLI)

#### ① 物理级无损过滤 (`filter`)
按毛利率 >= 15% 过滤并保留空值行，零损保留所有图片与公式：
```bash
python scripts/cli.py filter \
    --file "path/to/table.xlsx" \
    --expr "毛利率 >= 0.15" \
    --keep-empty \
    --output "path/to/output.xlsx"
```

#### ② 跨表回填与公式平移 (`transfer`)
将处理后的来源表字段同步回大底表，自动平移受影响的右侧公式：
```bash
python scripts/cli.py transfer \
    --source "path/to/source.xlsx" \
    --target "path/to/target.xlsx" \
    --sheets auto \
    --cols auto \
    --keys auto \
    --insert-after "唯一识别码"
```

#### ③ 跨表补充商品图片 (`fill-images`)
自动识别 WPS / Office 双引擎，多规格级联回填商品图片：
```bash
python scripts/cli.py fill-images \
    --target "path/to/target.xlsx" \
    --sources "path/to/source1.xlsx" "path/to/source2.xlsx" \
    --image-col "商品图片"
```

#### ④ 双表差异核对 (`diff`)
```bash
python scripts/cli.py diff \
    --base "path/to/old.xlsx" \
    --compare "path/to/new.xlsx" \
    --keys "商品链接" "产品69码" \
    --watch-cols "售价" "含税进价" "是否已下架" \
    --out-report "path/to/diff_report.xlsx"
```

#### ⑤ 表格类型嗅探 (`sniff`)
```bash
python scripts/cli.py sniff --file "path/to/table.xlsx"
```

---

## 作为 AI Agent Skill 使用

本项目原生遵循 Agent Skill 规范（包含 `SKILL.md`），可直接作为扩展技能接入各种大模型 Agent 客户端：

- **Antigravity CLI**：放置在 `~/.gemini/config/skills/wps-excel-toolkit`
- **Claude Code**：放置在 `~/.claude/skills/wps-excel-toolkit`
- **Cursor / Codex**：放置在项目 `.agents/skills/`

接入后，Agent 即可在接收到电商业务指令时自动调用本工具完成操作。

---

## 运行测试

```bash
python tests/test_toolkit.py
```

---

## 开源协议

MIT License
