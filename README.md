# paper-citation-convert

从 Word 文档中**一键提取、规范化、去重、分类**论文参考文献，输出格式统一的新 `.docx` 文件。

零第三方依赖，仅需 Python 标准库。默认遵循 GB/T 7714-2015 标准，通过 JSON Profile 可切换为任意引用格式。

## 它解决什么问题？

写论文时，脚注里的参考文献经常格式混乱、重复引用、标点不一致。手动整理几十上百条文献费时费力。

这个脚本帮你：把 `.docx` 里的文献**捞出来 → 洗干净 → 按标准重写 → 去掉重复 → 分类排好 → 输出新文档**。

## 处理流水线

```
┌─────────┐    ┌─────────┐    ┌──────────┐    ┌──────────┐    ┌─────────┐
│  提 取   │ →  │  清 洗   │ →  │  规范化   │ →  │ 去重合并  │ →  │ 分类输出 │
│ .docx   │    │ 标点/空格 │    │ 正则+模板 │    │ 页码范围  │    │ 新.docx │
└─────────┘    └─────────┘    └──────────┘    └──────────┘    └─────────┘
```

### 1. 提取 — 从 Word 文档捞文献

`.docx` 本质是 zip 压缩包，内含 XML 文件。脚本直接解析 XML，按优先级提取：

**脚注 → 尾注 → 正文段落**

识别依据：文本包含 `[M]`、`[J]`、`[D]` 等文献类型标记，且含四位数年份。

### 2. 清洗 — 统一基础格式

- 全角标点 → 半角（`。`→`.`，`［`→`[`，`，`→`,`）
- 多余空格、换行合并
- 类型标记后补 `. `
- 截掉文献主体之后的多余文字

### 3. 规范化 — 用正则 + 模板重写

每种文献类型对应一条规则：

| 类型 | 说明 | 示例输出 |
|------|------|----------|
| `[M]` | 普通图书 | `作者. 书名[M]. 出版地: 出版社, 年份.` |
| `[J]` | 期刊 | `作者. 篇名[J]. 刊名, 年份, 卷(期): 页码.` |
| `[D]` | 学位论文 | `作者. 题名[D]. 地点: 学校, 年份.` |
| `[N]` | 报纸 | `作者. 篇名[N]. 报纸名, 日期.` |
| `[C]` | 论文集 | `作者. 题名[C]. 出版地: 出版社, 年份.` |
| `[A]` | 档案 | `作者. 题名[A]. 地点: 机构, 年份.` |
| `[EB/OL]` | 电子资源 | 保留 URL 原样 |

图书 `[M]` 和论文集 `[C]` 的非析出文献会**自动去掉年份后的页码**。

### 4. 去重 + 页码合并

- **去重**：将文本 Unicode 归一化，去掉大小写/空格/标点后比较
- **页码合并**：同一文献不同页码引用（如 `45-60` 和 `80-90`）合并为 `45-90`
- 图书和专著析出文献不参与页码合并

### 5. 分类输出

按 9 大类分组输出，每类加标题，统一重新编号，生成新的 `.docx`：

```
普通图书：
[1]张三. 论文标题[M]. 北京: 人民出版社, 2020.

期刊：
[2]李四. 期刊文章[J]. 某期刊, 2021, 12(3): 100-105.

学位论文：
[3]王五. 学位论文题目[D]. 上海: 复旦大学, 2019.
```

## 快速开始

**无需安装任何依赖**，只需要 Python 3.6+。

### 批量处理

```bash
# 1. 把 .docx 文件放入 input/ 目录
# 2. 运行
python3 process_references.py
# 3. 到 output/ 查看结果
```

### 单文件处理

```bash
python3 process_references.py /path/to/your.docx
python3 process_references.py /path/to/your.docx -o /path/to/result.docx
```

## 目录结构

```
paper-citation-convert/
├── input/                 # 待处理的 .docx 放这里
├── output/                # 处理结果输出到这里
├── config/profiles/
│   ├── gb2015.json        # 默认 profile（GB/T 7714-2015）
│   ├── custom_template.json
│   └── mystyle.json       # 可直接修改的完整模板
├── process_references.py  # 主脚本（单文件，零依赖）
└── README.md
```

## Profile 配置系统

核心设计思想：**规则与代码分离**。处理逻辑固定在 Python 代码中，所有格式规则都在 JSON Profile 里。换引用标准只改 JSON，不改代码。

### 切换 Profile

```bash
python3 process_references.py --profile gb2015
python3 process_references.py --profile-file config/profiles/mystyle.json
```

### Profile 结构说明

```jsonc
{
  // 分类标题文本
  "category_titles": {
    "book": "普通图书：",
    "journal": "期刊：",
    // ...
  },
  // 输出文档样式
  "docx_style": {
    "font_cn": "宋体",
    "font_en": "Times New Roman",
    "font_size_half_points": 24,  // 24 = 小四 12pt
    "line_spacing_twips": 360     // 360 = 1.5 倍行距
  },
  // 处理规则
  "rules": {
    "strip_page_for_non_excerpt_types": ["M", "C"],
    "page_merge_exclude_categories": ["book", "monograph_excerpt"]
  },
  // 文献类型 → 分类映射
  "category_detection": { /* ... */ },
  // 提取文献主干的正则
  "reference_span_patterns": [ /* ... */ ],
  // 核心：正则 + 模板规范化规则（换格式主要改这里）
  "normalization_rules": [
    {
      "name": "journal",
      "regex": "^(?P<author>.+?)\\.\\s*(?P<title>.+?)\\[J\\]...",
      "template": "{author}. {title}[J]. {journal}, {year}, {volume_issue}: {pages}."
    }
  ]
}
```

## 全部参数

| 参数 | 说明 | 默认值 |
|------|------|--------|
| `input_docx` | 单个输入文件路径 | 无（批量处理 `input/`） |
| `--input-dir` | 批量输入目录 | `input` |
| `--output-dir` | 批量输出目录 | `output` |
| `-o` / `--output` | 单文件输出路径（`.docx` 或 `.txt`） | 自动生成 |
| `--profile` | Profile 名称 | `gb2015` |
| `--profile-file` | 直接指定 Profile 文件路径 | 无 |
| `--keep-original-number` | 保留原始 `[n]` 编号 | 否（重新编号） |

## 工程设计思路

如果你想做类似的"文本规范化处理工具"，这套代码的架构可以复用：

```
固定的流水线代码（引擎）
        ↕
可替换的 JSON 配置（规则集）
```

1. **直接解析文件格式** — `.docx` = zip + XML，用标准库读取，避免引入重量级依赖
2. **正则 + 模板的规则引擎** — 每种模式一条 regex 拆字段，一条 template 拼输出，全部外置到 JSON
3. **归一化去重** — Unicode NFKC + 去标点/空格作为比较 key，附加页码范围合并逻辑
4. **手搓最小 DOCX** — 直接拼 Word Open XML 写入 zip，零依赖生成可用的 Word 文档
5. **优先级回退** — 脚注 → 尾注 → 正文，确保各种文档格式都能处理
