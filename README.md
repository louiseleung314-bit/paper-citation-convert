# paper-citation-convert

用于处理论文参考文献的脚本，支持从 Word 文档中提取文献、按规则规范化、去重、分类，并输出新的 `.docx` 结果文档。

## 功能

- 从 `.docx` 提取参考文献（优先脚注/尾注，缺失时回退到正文段落）
- 参考 `GB/T 7714` 示例格式进行基础校验与修正
- `[M]`、`[C]` 非析出文献自动去掉年份后的页码
- 去重时忽略大小写、空白和标点差异
- 除著作类外，若同一文献仅页码不同，会合并为最小-最大页码范围
- 按类型分组输出，并保留出现顺序
- 输出独立新文档，源文档不覆盖

## 环境

- Python 3.9+
- 无第三方依赖

## 用法

```bash
python3 process_references.py input.docx
```

默认输出：

```text
input_processed_references.docx
```

指定输出路径：

```bash
python3 process_references.py input.docx -o output.docx
```

保留原始编号：

```bash
python3 process_references.py input.docx --keep-original-number
```

## 输出样式（DOCX）

- 中文：宋体
- 英文：Times New Roman
- 字号：小四（12pt）
- 行距：1.5 倍

