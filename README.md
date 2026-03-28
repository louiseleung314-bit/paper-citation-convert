# paper-citation-convert

用于处理论文参考文献的脚本，支持从 Word 文档提取文献、按规则规范化、去重、分类，并输出新的 `.docx` 结果文档。
现在支持整套格式的 `profile` 配置（不仅是标题和字体）。

## 目录结构（输入-处理-输出）

```text
paper-citation-convert/
├── input/                # 把待处理的 .docx 放这里
├── output/               # 处理结果默认输出到这里
├── config/
│   └── profiles/
│       ├── gb2015.json          # 默认 profile
│       ├── custom_template.json # 自定义模板示例
│       └── mystyle.json         # 你可直接改的完整模板
├── process_references.py # 处理脚本
└── README.md
```

## 快速开始

1. 把源文件放到 `input/` 目录（支持多个 `.docx`）。
2. 运行：

```bash
python3 process_references.py
```

3. 到 `output/` 查看结果文件：`<原文件名>_processed_references.docx`。

## 单文件模式

```bash
python3 process_references.py /path/to/your.docx
```

指定输出路径：

```bash
python3 process_references.py /path/to/your.docx -o /path/to/result.docx
```

## Profile（整套格式）怎么改

默认 profile 是 `gb2015`，对应 `config/profiles/gb2015.json`。

直接切换 profile：

```bash
python3 process_references.py --profile gb2015
python3 process_references.py --profile-file config/profiles/custom_template.json
python3 process_references.py --profile-file config/profiles/mystyle.json
```

你要改“整套引用格式”时，建议直接改 `config/profiles/mystyle.json`。

profile 结构：

- `category_titles`
  - 控制分类标题文本（例如 `"普通图书："`）。
- `docx_style`
  - 控制输出字体和版式。
  - `font_cn`：中文字体
  - `font_en`：英文字体
  - `font_size_half_points`：字号（half-point，`24` = 小四 12pt）
  - `line_spacing_twips`：行距（`360` = 1.5 倍行距）
- `rules`
  - `strip_page_for_non_excerpt_types`：哪些类型（如 `M`、`C`）在非析出文献时移除年份后页码。
  - `page_merge_exclude_categories`：哪些分类不参与“同文献页码范围合并”。
- `category_detection`
  - 控制文献类型标识如何映射到分类（如 `J -> journal`）。
  - 控制析出文献识别（如 `marker: //`）。
- `reference_span_patterns`
  - 提取“有效文献主干”的正则列表。
- `normalization_rules`
  - 这是核心：每条包含 `regex + template`，用于把原文献重写成目标格式。
  - 你换成非 GB2015 时，主要改这里。

## 当前默认规则

- 从 `.docx` 提取文献：优先脚注/尾注，缺失时回退到正文段落。
- `[M]`、`[C]` 非析出文献自动去掉年份后的页码。
- 去重忽略大小写、空白和标点差异。
- 除 `book`、`monograph_excerpt` 外，若同一文献仅页码不同，则合并为最小-最大页码范围。

## 常用参数

```bash
python3 process_references.py --help
```

- `--input-dir`：批量输入目录（默认 `input`）
- `--output-dir`：批量输出目录（默认 `output`）
- `--profile`：profile 名称（默认 `gb2015`，即 `config/profiles/gb2015.json`）
- `--profile-file`：直接指定 profile 文件路径
- `--config`：兼容参数，等同 `--profile-file`
- `--keep-original-number`：保留原始 `[n]` 编号
