#!/usr/bin/env python3
"""Process DOCX footnote references with GB/T 7714-like normalization rules.

Usage:
  python3 process_references.py input.docx -o output.docx
"""

from __future__ import annotations

import argparse
import re
import sys
import unicodedata
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple
from xml.sax.saxutils import escape


NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

LEADING_INDEX_RE = re.compile(r"^\s*[\[［](\d+)[\]］]\s*")
TYPE_MARK_RE = re.compile(r"\[(EB/OL|[A-Z](?:/[A-Z]+)?)\]")

BOOK_OR_PROC_RE = re.compile(
    r"^(?P<author>.+?)\.\s*(?P<title>.+?)\[(?P<type>M|C)\]\.\s*"
    r"(?P<place>.+?)\s*[:：]\s*(?P<publisher>.+?)\s*[,，]\s*"
    r"(?P<year>\d{4})(?:\s*[:：]\s*(?P<pages>[^.\[]+))?\.\s*$"
)
THESIS_RE = re.compile(
    r"^(?P<author>.+?)\.\s*(?P<title>.+?)\[D\]\.\s*"
    r"(?P<place>.+?)\s*[:：]\s*(?P<publisher>.+?)\s*[,，]\s*(?P<year>\d{4})"
    r"(?:\s*[:：]\s*(?P<pages>[^.\[]+))?\.\s*$"
)
ARCHIVE_RE = re.compile(
    r"^(?P<author>.+?)\.\s*(?P<title>.+?)\[A\]\.\s*"
    r"(?P<place>.+?)\s*[:：]\s*(?P<publisher>.+?)\s*[,，]\s*(?P<year>\d{4})"
    r"(?:\s*[:：]\s*(?P<pages>[^.\[]+))?\.\s*$"
)
JOURNAL_RE = re.compile(
    r"^(?P<author>.+?)\.\s*(?P<title>.+?)\[J\]\.\s*(?P<journal>.+?)\s*[,，]\s*"
    r"(?P<year>\d{4})\s*[,，]\s*(?P<volume_issue>[^:：]+)\s*[:：]\s*(?P<pages>[^.]+)\.\s*$"
)
NEWSPAPER_RE = re.compile(
    r"^(?P<author>.+?)\.\s*(?P<title>.+?)\[N\]\.\s*(?P<newspaper>.+?)\s*[,，]\s*"
    r"(?P<date>\d{4}-\d{2}-\d{2})(?P<issue>\([^)]+\))?\.\s*$"
)
PAGE_TAIL_COLON_RE = re.compile(
    r"^(?P<prefix>.+?)\s*[:：]\s*(?P<pages>\d[\d+\-, ]*)\.\s*$"
)
PAGE_TAIL_PAREN_RE = re.compile(
    r"^(?P<prefix>.+\d{4}-\d{2}-\d{2})\((?P<pages>\d[\d+\-, ]*)\)\.\s*$"
)

REFERENCE_SPAN_PATTERNS = [
    re.compile(r"^(.+?\[M\](?://.+?)?\.\s*.+?[,，]\s*\d{4}(?:\s*[:：]\s*[^.\[]+)?\.)"),
    re.compile(r"^(.+?\[C\](?://.+?)?\.\s*.+?[,，]\s*\d{4}(?:\s*[:：]\s*[^.\[]+)?\.)"),
    re.compile(r"^(.+?\[D\]\.\s*.+?[,，]\s*\d{4}(?:\s*[:：]\s*[^.\[]+)?\.)"),
    re.compile(r"^(.+?\[A\]\.\s*.+?[,，]\s*\d{4}(?:\s*[:：]\s*[^.\[]+)?\.)"),
    re.compile(r"^(.+?\[J\]\.\s*.+?[,，]\s*\d{4}[^.]*\.)"),
    re.compile(
        r"^(.+?\[N\]\.\s*.+?\d{4}-\d{2}-\d{2}"
        r"(?:\([^)]+\))?(?:\s*[:：]\s*[^.\s]+)?\.)"
    ),
    re.compile(r"^(.*?\[EB/OL\].*?(?:https?://\S+))"),
]

CATEGORY_ORDER: List[Tuple[str, str]] = [
    ("book", "普通图书："),
    ("proceedings", "论文集："),
    ("thesis", "学位论文："),
    ("monograph_excerpt", "专著中析出文献："),
    ("newspaper", "报纸中析出文献："),
    ("archive", "档案资源："),
    ("journal", "期刊："),
    ("electronic", "电子资源："),
    ("other", "其他类型："),
]


@dataclass
class RefItem:
    original_index: Optional[int]
    raw_text: str
    normalized_text: str
    category: str
    changed: bool


@dataclass
class PageMergeInfo:
    prefix: str
    bounds: Tuple[int, int]
    style: str  # "colon" or "paren"


def extract_text_from_node(node: ET.Element) -> str:
    chunks: List[str] = []
    for elem in node.iter():
        tag = elem.tag.rsplit("}", 1)[-1]
        if tag == "t":
            chunks.append(elem.text or "")
        elif tag in {"tab"}:
            chunks.append(" ")
        elif tag in {"br", "cr"}:
            chunks.append(" ")
    return "".join(chunks)


def build_note_map(note_root: ET.Element, note_tag: str) -> Dict[str, str]:
    note_map: Dict[str, str] = {}
    for note in note_root.findall(f".//w:{note_tag}", NS):
        nid = note.get(f"{{{NS['w']}}}id")
        if nid is None or nid.startswith("-"):
            continue
        text = extract_text_from_node(note)
        text = re.sub(r"\s+", " ", text).strip()
        if text:
            note_map[nid] = text
    return note_map


def trim_to_reference_candidate(text: str) -> Optional[str]:
    text = re.sub(r"\s+", " ", text).strip()
    text = text.replace("［", "[").replace("］", "]")
    marker = TYPE_MARK_RE.search(text)
    if not marker:
        return None

    prefix = text[: marker.start()]
    start = 0

    for token in ("参见",):
        pos = prefix.rfind(token)
        if pos != -1:
            start = max(start, pos + len(token))
    for ch in ("。", "；"):
        pos = prefix.rfind(ch)
        if pos != -1:
            start = max(start, pos + 1)

    candidate = text[start:].strip(" ：:，,;；。")
    candidate = re.sub(r"^\[[^\]]{1,8}\]\s*", "", candidate)

    if not TYPE_MARK_RE.search(candidate):
        return None
    if not re.search(r"\d{4}", candidate):
        return None
    if len(candidate) < 12:
        return None
    return candidate


def extract_references_in_order(docx_path: Path) -> Tuple[List[str], str]:
    with zipfile.ZipFile(docx_path) as zf:
        try:
            document_xml = zf.read("word/document.xml")
            footnotes_xml = zf.read("word/footnotes.xml")
        except KeyError as exc:
            raise RuntimeError(f"DOCX 缺少必要 XML 文件: {exc}") from exc

        endnotes_xml: Optional[bytes]
        try:
            endnotes_xml = zf.read("word/endnotes.xml")
        except KeyError:
            endnotes_xml = None

    document_root = ET.fromstring(document_xml)
    footnotes_root = ET.fromstring(footnotes_xml)

    footnote_map = build_note_map(footnotes_root, "footnote")

    foot_refs: List[str] = []
    for ref in document_root.findall(".//w:footnoteReference", NS):
        fid = ref.get(f"{{{NS['w']}}}id")
        if not fid:
            continue
        text = footnote_map.get(fid, "").strip()
        if text:
            foot_refs.append(text)
    if foot_refs:
        return foot_refs, "footnotes"

    if endnotes_xml is not None:
        endnotes_root = ET.fromstring(endnotes_xml)
        endnote_map = build_note_map(endnotes_root, "endnote")
        end_refs: List[str] = []
        for ref in document_root.findall(".//w:endnoteReference", NS):
            eid = ref.get(f"{{{NS['w']}}}id")
            if not eid:
                continue
            text = endnote_map.get(eid, "").strip()
            if text:
                end_refs.append(text)
        if end_refs:
            return end_refs, "endnotes"

    # Fallback: many source docs keep one reference per paragraph in body text.
    body_refs: List[str] = []
    for para in document_root.findall(".//w:body/w:p", NS):
        text = extract_text_from_node(para)
        candidate = trim_to_reference_candidate(text)
        if candidate:
            body_refs.append(candidate)
    return body_refs, "body paragraphs"


def basic_cleanup(text: str) -> str:
    text = text.strip().replace("\u00a0", " ")
    text = text.replace("［", "[").replace("］", "]")
    text = text.replace("。", ".")
    text = re.sub(
        r"(\[(?:EB/OL|M|C|D|A|J|N)\])(?![./])\s+",
        r"\1. ",
        text,
    )
    text = re.sub(r"\s+", " ", text).strip()
    text = re.sub(r"\s*//\s*", "//", text)
    text = re.sub(r"\s+\.", ".", text)
    text = re.sub(r"\s+,", ",", text)
    text = re.sub(r"\s+:", ":", text)
    return text


def trim_after_reference_tail(text: str) -> str:
    for pattern in REFERENCE_SPAN_PATTERNS:
        m = pattern.match(text)
        if m:
            return m.group(1).strip()
    return text


def split_leading_index(text: str) -> Tuple[Optional[int], str]:
    m = LEADING_INDEX_RE.match(text)
    if not m:
        return None, text.strip()
    idx = int(m.group(1))
    return idx, text[m.end() :].strip()


def strip_trailing_page_for_m_c_non_excerpt(text: str) -> str:
    if "//" in text:
        return text
    if "[M]" not in text and "[C]" not in text:
        return text

    # Remove any trailing page segment after year for [M]/[C] non-excerpt refs.
    # Examples:
    #   ..., 1913:4,12-15. -> ..., 1913.
    #   ..., 2002: 20.     -> ..., 2002.
    updated = re.sub(
        r"([,，]\s*\d{4})\s*[:：]\s*[^.\[]+(?=\.\s*$)",
        r"\1",
        text,
    )
    updated = re.sub(
        r"([,，]\s*\d{4})\s*[:：]\s*[^.\[]+\s*$",
        r"\1",
        updated,
    )
    updated = re.sub(r"\.\.+$", ".", updated)
    if not updated.endswith("."):
        updated = re.sub(r"([,，]\s*\d{4})\s*$", r"\1.", updated)
    return updated


def enforce_terminal_period(text: str) -> str:
    if re.search(r"[.]$", text):
        return text
    return text + "."


def fix_book_or_proceedings(text: str) -> str:
    m = BOOK_OR_PROC_RE.match(text)
    if not m:
        return text
    return (
        f"{m.group('author').strip()}. "
        f"{m.group('title').strip()}[{m.group('type')}]. "
        f"{m.group('place').strip()}: {m.group('publisher').strip()}, {m.group('year').strip()}."
    )


def fix_thesis(text: str) -> str:
    m = THESIS_RE.match(text)
    if not m:
        return text
    return (
        f"{m.group('author').strip()}. "
        f"{m.group('title').strip()}[D]. "
        f"{m.group('place').strip()}: {m.group('publisher').strip()}, {m.group('year').strip()}."
    )


def fix_archive(text: str) -> str:
    m = ARCHIVE_RE.match(text)
    if not m:
        return text
    return (
        f"{m.group('author').strip()}. "
        f"{m.group('title').strip()}[A]. "
        f"{m.group('place').strip()}: {m.group('publisher').strip()}, {m.group('year').strip()}."
    )


def fix_journal(text: str) -> str:
    m = JOURNAL_RE.match(text)
    if not m:
        return text
    return (
        f"{m.group('author').strip()}. "
        f"{m.group('title').strip()}[J]. "
        f"{m.group('journal').strip()}, {m.group('year').strip()}, "
        f"{m.group('volume_issue').strip()}: {m.group('pages').strip()}."
    )


def fix_newspaper(text: str) -> str:
    m = NEWSPAPER_RE.match(text)
    if not m:
        return text
    issue = m.group("issue") or ""
    return (
        f"{m.group('author').strip()}. "
        f"{m.group('title').strip()}[N]. "
        f"{m.group('newspaper').strip()}, {m.group('date').strip()}{issue}."
    )


def detect_category(text: str) -> str:
    mark = TYPE_MARK_RE.search(text)
    if "/OL" in text or (mark and "/OL" in mark.group(1)):
        return "electronic"
    if "[M]" in text and "//" in text:
        return "monograph_excerpt"
    if "[M]" in text:
        return "book"
    if "[C]" in text:
        return "proceedings"
    if "[D]" in text:
        return "thesis"
    if "[N]" in text:
        return "newspaper"
    if "[A]" in text:
        return "archive"
    if "[J]" in text:
        return "journal"
    if "[EB/OL]" in text:
        return "electronic"
    return "other"


def normalize_reference(text: str) -> str:
    text = basic_cleanup(text)
    text = trim_after_reference_tail(text)
    text = strip_trailing_page_for_m_c_non_excerpt(text)
    text = fix_book_or_proceedings(text)
    text = fix_thesis(text)
    text = fix_archive(text)
    text = fix_newspaper(text)
    text = fix_journal(text)
    text = enforce_terminal_period(text)
    return text


def dedupe_key(normalized_text: str) -> str:
    text = unicodedata.normalize("NFKC", normalized_text).casefold()
    # Keep only letters/numbers so case, spaces, and punctuation style won't affect dedupe.
    return "".join(ch for ch in text if unicodedata.category(ch)[0] in {"L", "N"})


def parse_page_bounds(page_text: str) -> Optional[Tuple[int, int]]:
    nums = [int(x) for x in re.findall(r"\d+", page_text)]
    if not nums:
        return None
    return min(nums), max(nums)


def is_page_merge_eligible(category: str) -> bool:
    # "著作类型"默认指 [M]（含专著中析出文献 [M]//）,
    # 其余类型若仅页码不同则合并页码范围。
    return category not in {"book", "monograph_excerpt"}


def split_for_page_merge(
    normalized_text: str,
    category: str,
) -> Tuple[str, Optional[PageMergeInfo]]:
    if not is_page_merge_eligible(category):
        return normalized_text, None

    match = PAGE_TAIL_PAREN_RE.match(normalized_text)
    if match:
        prefix = match.group("prefix").strip()
        bounds = parse_page_bounds(match.group("pages"))
        if bounds is not None:
            base_text = f"{prefix}."
            return base_text, PageMergeInfo(prefix=prefix, bounds=bounds, style="paren")

    match = PAGE_TAIL_COLON_RE.match(normalized_text)
    if match:
        prefix = match.group("prefix").strip()
        bounds = parse_page_bounds(match.group("pages"))
        if bounds is not None:
            base_text = f"{prefix}."
            return base_text, PageMergeInfo(prefix=prefix, bounds=bounds, style="colon")

    return normalized_text, None


def build_text_with_bounds(info: PageMergeInfo) -> str:
    prefix = info.prefix
    bounds = info.bounds
    start, end = bounds
    page_text = str(start) if start == end else f"{start}-{end}"
    if info.style == "paren":
        return f"{prefix}({page_text})."
    return f"{prefix}: {page_text}."


def process_references(footnotes: List[str]) -> List[RefItem]:
    items: List[RefItem] = []
    seen: Dict[str, int] = {}
    page_merge_by_item: Dict[int, PageMergeInfo] = {}

    for ref in footnotes:
        raw = basic_cleanup(ref)
        original_idx, body = split_leading_index(raw)
        normalized = normalize_reference(body)
        category = detect_category(normalized)
        base_text, page_merge = split_for_page_merge(
            normalized_text=normalized,
            category=category,
        )
        key_text = base_text if page_merge else normalized
        key = dedupe_key(key_text)

        if key in seen:
            existing_idx = seen[key]
            if page_merge:
                if existing_idx in page_merge_by_item:
                    old = page_merge_by_item[existing_idx]
                    merged_info = PageMergeInfo(
                        prefix=old.prefix,
                        style=old.style,
                        bounds=(
                            min(old.bounds[0], page_merge.bounds[0]),
                            max(old.bounds[1], page_merge.bounds[1]),
                        ),
                    )
                    page_merge_by_item[existing_idx] = merged_info
                    merged_text = build_text_with_bounds(merged_info)
                    if items[existing_idx].normalized_text != merged_text:
                        items[existing_idx].normalized_text = merged_text
                        items[existing_idx].changed = True
                else:
                    page_merge_by_item[existing_idx] = page_merge
                    merged_text = build_text_with_bounds(page_merge)
                    if items[existing_idx].normalized_text != merged_text:
                        items[existing_idx].normalized_text = merged_text
                        items[existing_idx].changed = True
            continue

        output_text = normalized
        if page_merge:
            output_text = build_text_with_bounds(page_merge)

        changed = output_text != body
        items.append(
            RefItem(
                original_index=original_idx,
                raw_text=ref,
                normalized_text=output_text,
                category=category,
                changed=changed,
            )
        )
        item_idx = len(items) - 1
        seen[key] = item_idx
        if page_merge:
            page_merge_by_item[item_idx] = page_merge
    return items


def render_output_lines(
    items: List[RefItem],
    keep_original_number: bool = False,
) -> List[str]:
    grouped: Dict[str, List[RefItem]] = {k: [] for k, _ in CATEGORY_ORDER}
    for item in items:
        grouped.setdefault(item.category, []).append(item)

    lines: List[str] = []
    next_index = 1
    for key, title in CATEGORY_ORDER:
        refs = grouped.get(key, [])
        if not refs:
            continue
        lines.append(title)
        for item in refs:
            if keep_original_number and item.original_index is not None:
                idx = item.original_index
            else:
                idx = next_index
                next_index += 1
            lines.append(f"[{idx}]{item.normalized_text}")
        lines.append("")
    return lines


def write_docx_output(lines: List[str], output_path: Path) -> None:
    def para_xml(line: str) -> str:
        ppr_xml = '<w:pPr><w:spacing w:line="360" w:lineRule="auto"/></w:pPr>'
        if not line:
            return f"<w:p>{ppr_xml}</w:p>"
        text = escape(line)
        return (
            f"<w:p>{ppr_xml}<w:r><w:rPr>"
            '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" '
            'w:cs="Times New Roman" w:eastAsia="宋体"/>'
            '<w:sz w:val="24"/><w:szCs w:val="24"/>'
            "</w:rPr>"
            f'<w:t xml:space="preserve">{text}</w:t>'
            "</w:r></w:p>"
        )

    body_xml = "".join(para_xml(line) for line in lines)
    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f"<w:body>{body_xml}"
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" '
        'w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>'
        "</w:body></w:document>"
    )
    content_types_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" '
        'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        "</Types>"
    )
    rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
        'Target="word/document.xml"/>'
        "</Relationships>"
    )

    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml)
        zf.writestr("_rels/.rels", rels_xml)
        zf.writestr("word/document.xml", document_xml)


def write_output(
    items: List[RefItem],
    output_path: Path,
    keep_original_number: bool = False,
) -> None:
    lines = render_output_lines(
        items=items,
        keep_original_number=keep_original_number,
    )
    ext = output_path.suffix.lower()
    if ext == ".docx":
        write_docx_output(lines=lines, output_path=output_path)
        return
    output_path.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")


def parse_args(argv: List[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="从 DOCX 脚注提取参考文献，并按规则校验、修正、去重和分类输出。"
    )
    parser.add_argument("input_docx", type=Path, help="包含脚注文献的 Word 文件（.docx）")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        help="输出路径，支持 .docx 或 .txt（默认: 输入名_processed_references.docx）",
    )
    parser.add_argument(
        "--keep-original-number",
        action="store_true",
        help="输出时保留原始 [n] 编号；默认按分类后的顺序重新编号",
    )
    return parser.parse_args(argv)


def main(argv: List[str]) -> int:
    args = parse_args(argv)

    if not args.input_docx.exists():
        print(f"输入文件不存在: {args.input_docx}", file=sys.stderr)
        return 1
    if args.input_docx.suffix.lower() != ".docx":
        print("仅支持 .docx 文件。", file=sys.stderr)
        return 1

    output_path = args.output
    if output_path is None:
        output_path = args.input_docx.with_name(
            f"{args.input_docx.stem}_processed_references.docx"
        )
    if output_path.suffix.lower() not in {".docx", ".txt"}:
        print("输出文件仅支持 .docx 或 .txt 后缀。", file=sys.stderr)
        return 1
    if output_path.resolve() == args.input_docx.resolve():
        print("输出文件不能与输入文件同名，请更换输出路径。", file=sys.stderr)
        return 1

    try:
        refs, ref_source = extract_references_in_order(args.input_docx)
    except Exception as exc:  # pylint: disable=broad-except
        print(f"提取文献失败: {exc}", file=sys.stderr)
        return 1

    if not refs:
        print("未在文档中检测到可处理的参考文献内容。", file=sys.stderr)
        return 1

    items = process_references(refs)
    write_output(
        items=items,
        output_path=output_path,
        keep_original_number=args.keep_original_number,
    )

    changed_count = sum(1 for i in items if i.changed)
    print(f"文献来源: {ref_source}")
    print(f"提取文献: {len(refs)} 条")
    print(f"去重后文献: {len(items)} 条")
    print(f"被修正条目: {changed_count} 条")
    print(f"输出文件: {output_path}")
    print(f"源文件保持不变: {args.input_docx}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
