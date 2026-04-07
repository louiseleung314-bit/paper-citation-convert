#!/usr/bin/env python3
"""Process DOCX references with profile-driven normalization rules.

Usage:
  python3 process_references.py input.docx
  python3 process_references.py
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import unicodedata
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Pattern, Tuple
from xml.sax.saxutils import escape


NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
SCRIPT_DIR = Path(__file__).resolve().parent

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

CATEGORY_KEYS: List[str] = [
    "book",
    "proceedings",
    "thesis",
    "monograph_excerpt",
    "newspaper",
    "archive",
    "journal",
    "electronic",
    "other",
]

DEFAULT_CATEGORY_TITLES: Dict[str, str] = {
    "book": "普通图书：",
    "proceedings": "论文集：",
    "thesis": "学位论文：",
    "monograph_excerpt": "专著中析出文献：",
    "newspaper": "报纸中析出文献：",
    "archive": "档案资源：",
    "journal": "期刊：",
    "electronic": "电子资源：",
    "other": "其他类型：",
}

DEFAULT_DOCX_STYLE: Dict[str, Any] = {
    "font_cn": "宋体",
    "font_en": "Times New Roman",
    "font_size_half_points": 24,
    "line_spacing_twips": 360,
}

DEFAULT_RULES: Dict[str, Any] = {
    "strip_page_for_non_excerpt_types": ["M", "C"],
    "page_merge_exclude_categories": ["book", "monograph_excerpt"],
    "loose_dedupe_categories": ["book"],
}

DEFAULT_CATEGORY_DETECTION: Dict[str, Any] = {
    "type_to_category": {
        "M": "book",
        "C": "proceedings",
        "D": "thesis",
        "N": "newspaper",
        "A": "archive",
        "J": "journal",
        "EB/OL": "electronic",
    },
    "excerpt": {
        "type": "M",
        "marker": "//",
        "category": "monograph_excerpt",
    },
    "electronic_keywords": ["/OL"],
    "fallback": "other",
}

DEFAULT_REFERENCE_SPAN_PATTERNS: List[str] = [
    r"^(.+?\[M\](?://.+?)?\.\s*.+?[,，]\s*\d{4}(?:\s*[:：]\s*[^.\[]+)?\.)",
    r"^(.+?\[C\](?://.+?)?\.\s*.+?[,，]\s*\d{4}(?:\s*[:：]\s*[^.\[]+)?\.)",
    r"^(.+?\[D\]\.\s*.+?[,，]\s*\d{4}(?:\s*[:：]\s*[^.\[]+)?\.)",
    r"^(.+?\[A\]\.\s*.+?[,，]\s*\d{4}(?:\s*[:：]\s*[^.\[]+)?\.)",
    r"^(.+?\[J\]\.\s*.+?[,，]\s*\d{4}[^.]*\.)",
    r"^(.+?\[N\]\.\s*.+?\d{4}-\d{2}-\d{2}(?:\([^)]+\))?(?:\s*[:：]\s*[^.\s]+)?\.)",
    r"^(.*?\[EB/OL\].*?(?:https?://\S+))",
]

DEFAULT_NORMALIZATION_RULES: List[Dict[str, str]] = [
    {
        "name": "book_or_proceedings",
        "regex": BOOK_OR_PROC_RE.pattern,
        "template": "{author}. {title}[{type}]. {place}: {publisher}, {year}.",
    },
    {
        "name": "thesis",
        "regex": THESIS_RE.pattern,
        "template": "{author}. {title}[D]. {place}: {publisher}, {year}.",
    },
    {
        "name": "archive",
        "regex": ARCHIVE_RE.pattern,
        "template": "{author}. {title}[A]. {place}: {publisher}, {year}.",
    },
    {
        "name": "newspaper",
        "regex": NEWSPAPER_RE.pattern,
        "template": "{author}. {title}[N]. {newspaper}, {date}{issue}.",
    },
    {
        "name": "journal",
        "regex": JOURNAL_RE.pattern,
        "template": "{author}. {title}[J]. {journal}, {year}, {volume_issue}: {pages}.",
    },
]

DEFAULT_PROFILE_NAME = "gb2015"
DEFAULT_PROFILE_DIR = SCRIPT_DIR / "config/profiles"


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


@dataclass
class NormalizationRule:
    name: str
    pattern: Pattern[str]
    template: str


@dataclass
class RuntimeConfig:
    profile_name: str
    profile_path: Path
    category_order: List[Tuple[str, str]]
    docx_style: Dict[str, Any]
    rules: Dict[str, Any]
    category_detection: Dict[str, Any]
    reference_span_patterns: List[Pattern[str]]
    normalization_rules: List[NormalizationRule]


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


def as_str_list(value: Any, fallback: List[str]) -> List[str]:
    if not isinstance(value, list):
        return list(fallback)
    output: List[str] = []
    for item in value:
        if isinstance(item, str) and item.strip():
            output.append(item.strip())
    return output or list(fallback)


def as_dict(value: Any, fallback: Dict[str, Any]) -> Dict[str, Any]:
    if isinstance(value, dict):
        return value
    return dict(fallback)


def choose_profile_path(profile_name: str, profile_file: Optional[Path]) -> Path:
    if profile_file is not None:
        return profile_file
    return DEFAULT_PROFILE_DIR / f"{profile_name}.json"


def compile_regex_list(patterns: List[str], field_name: str) -> List[Pattern[str]]:
    compiled: List[Pattern[str]] = []
    for idx, pattern in enumerate(patterns, 1):
        try:
            compiled.append(re.compile(pattern))
        except re.error as exc:
            raise RuntimeError(
                f"{field_name} 第 {idx} 条正则无效: {pattern} ({exc})"
            ) from exc
    return compiled


def compile_normalization_rules(raw_rules: List[Dict[str, str]]) -> List[NormalizationRule]:
    compiled: List[NormalizationRule] = []
    for idx, rule in enumerate(raw_rules, 1):
        name = str(rule.get("name") or f"rule_{idx}")
        pattern_str = rule.get("regex")
        template = rule.get("template")
        if not isinstance(pattern_str, str) or not pattern_str.strip():
            raise RuntimeError(f"normalization_rules 第 {idx} 条缺少 regex。")
        if not isinstance(template, str) or not template.strip():
            raise RuntimeError(f"normalization_rules 第 {idx} 条缺少 template。")
        try:
            pattern = re.compile(pattern_str)
        except re.error as exc:
            raise RuntimeError(
                f"normalization_rules 第 {idx} 条 regex 无效: {pattern_str} ({exc})"
            ) from exc
        compiled.append(
            NormalizationRule(
                name=name,
                pattern=pattern,
                template=template,
            )
        )
    return compiled


def load_runtime_config(profile_name: str, profile_file: Optional[Path]) -> RuntimeConfig:
    profile_path = choose_profile_path(profile_name=profile_name, profile_file=profile_file)

    category_titles = dict(DEFAULT_CATEGORY_TITLES)
    docx_style = dict(DEFAULT_DOCX_STYLE)
    rules = dict(DEFAULT_RULES)
    category_detection = dict(DEFAULT_CATEGORY_DETECTION)
    reference_span_patterns_raw = list(DEFAULT_REFERENCE_SPAN_PATTERNS)
    normalization_rules_raw: List[Dict[str, str]] = list(DEFAULT_NORMALIZATION_RULES)

    if profile_path.exists():
        try:
            raw = json.loads(profile_path.read_text(encoding="utf-8"))
        except json.JSONDecodeError as exc:
            raise RuntimeError(f"配置文件 JSON 解析失败: {profile_path} ({exc})") from exc
        if not isinstance(raw, dict):
            raise RuntimeError(f"配置文件格式无效: {profile_path}")

        raw_titles = raw.get("category_titles")
        if isinstance(raw_titles, dict):
            for key in CATEGORY_KEYS:
                value = raw_titles.get(key)
                if isinstance(value, str) and value.strip():
                    category_titles[key] = value.strip()

        raw_style = raw.get("docx_style")
        if isinstance(raw_style, dict):
            if isinstance(raw_style.get("font_cn"), str) and raw_style["font_cn"].strip():
                docx_style["font_cn"] = raw_style["font_cn"].strip()
            if isinstance(raw_style.get("font_en"), str) and raw_style["font_en"].strip():
                docx_style["font_en"] = raw_style["font_en"].strip()
            if isinstance(raw_style.get("font_size_half_points"), int):
                docx_style["font_size_half_points"] = raw_style["font_size_half_points"]
            if isinstance(raw_style.get("line_spacing_twips"), int):
                docx_style["line_spacing_twips"] = raw_style["line_spacing_twips"]

        raw_rules = raw.get("rules")
        if isinstance(raw_rules, dict):
            rules["strip_page_for_non_excerpt_types"] = as_str_list(
                raw_rules.get("strip_page_for_non_excerpt_types"),
                DEFAULT_RULES["strip_page_for_non_excerpt_types"],
            )
            rules["page_merge_exclude_categories"] = as_str_list(
                raw_rules.get("page_merge_exclude_categories"),
                DEFAULT_RULES["page_merge_exclude_categories"],
            )
            rules["loose_dedupe_categories"] = as_str_list(
                raw_rules.get("loose_dedupe_categories"),
                DEFAULT_RULES["loose_dedupe_categories"],
            )

        raw_category_detection = raw.get("category_detection")
        if isinstance(raw_category_detection, dict):
            category_detection = dict(DEFAULT_CATEGORY_DETECTION)
            type_to_category = as_dict(
                raw_category_detection.get("type_to_category"),
                DEFAULT_CATEGORY_DETECTION["type_to_category"],
            )
            category_detection["type_to_category"] = {
                str(k): str(v) for k, v in type_to_category.items()
            }
            excerpt = as_dict(
                raw_category_detection.get("excerpt"),
                DEFAULT_CATEGORY_DETECTION["excerpt"],
            )
            category_detection["excerpt"] = {
                "type": str(excerpt.get("type", DEFAULT_CATEGORY_DETECTION["excerpt"]["type"])),
                "marker": str(
                    excerpt.get("marker", DEFAULT_CATEGORY_DETECTION["excerpt"]["marker"])
                ),
                "category": str(
                    excerpt.get("category", DEFAULT_CATEGORY_DETECTION["excerpt"]["category"])
                ),
            }
            category_detection["electronic_keywords"] = as_str_list(
                raw_category_detection.get("electronic_keywords"),
                DEFAULT_CATEGORY_DETECTION["electronic_keywords"],
            )
            category_detection["fallback"] = str(
                raw_category_detection.get("fallback", DEFAULT_CATEGORY_DETECTION["fallback"])
            )

        raw_ref_patterns = raw.get("reference_span_patterns")
        if isinstance(raw_ref_patterns, list):
            reference_span_patterns_raw = as_str_list(
                raw_ref_patterns,
                DEFAULT_REFERENCE_SPAN_PATTERNS,
            )

        raw_norm_rules = raw.get("normalization_rules")
        if isinstance(raw_norm_rules, list) and raw_norm_rules:
            normalization_rules_raw = []
            for rule in raw_norm_rules:
                if isinstance(rule, dict):
                    normalization_rules_raw.append(rule)

    else:
        raise RuntimeError(
            f"未找到 profile 文件: {profile_path}。请使用 --profile-file 指定，或在 config/profiles/ 下创建。"
        )

    category_order = [(key, category_titles[key]) for key in CATEGORY_KEYS]
    reference_span_patterns = compile_regex_list(
        reference_span_patterns_raw,
        field_name="reference_span_patterns",
    )
    normalization_rules = compile_normalization_rules(normalization_rules_raw)

    return RuntimeConfig(
        profile_name=profile_name,
        profile_path=profile_path,
        category_order=category_order,
        docx_style=docx_style,
        rules=rules,
        category_detection=category_detection,
        reference_span_patterns=reference_span_patterns,
        normalization_rules=normalization_rules,
    )


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


def trim_after_reference_tail(text: str, reference_span_patterns: List[Pattern[str]]) -> str:
    for pattern in reference_span_patterns:
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


def strip_trailing_page_for_non_excerpt_types(
    text: str,
    strip_types: List[str],
) -> str:
    if "//" in text:
        return text
    if not any(f"[{doc_type}]" in text for doc_type in strip_types):
        return text

    # Remove trailing page segment after year for configured non-excerpt types.
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


def detect_category(text: str, category_detection: Dict[str, Any]) -> str:
    type_to_category = as_dict(
        category_detection.get("type_to_category"),
        DEFAULT_CATEGORY_DETECTION["type_to_category"],
    )
    fallback = str(category_detection.get("fallback", DEFAULT_CATEGORY_DETECTION["fallback"]))

    electronic_keywords = as_str_list(
        category_detection.get("electronic_keywords"),
        DEFAULT_CATEGORY_DETECTION["electronic_keywords"],
    )
    for marker in electronic_keywords:
        if marker and marker in text:
            return str(type_to_category.get("EB/OL", "electronic"))

    excerpt = as_dict(category_detection.get("excerpt"), DEFAULT_CATEGORY_DETECTION["excerpt"])
    excerpt_type = str(excerpt.get("type", DEFAULT_CATEGORY_DETECTION["excerpt"]["type"]))
    excerpt_marker = str(excerpt.get("marker", DEFAULT_CATEGORY_DETECTION["excerpt"]["marker"]))
    excerpt_category = str(
        excerpt.get("category", DEFAULT_CATEGORY_DETECTION["excerpt"]["category"])
    )
    if excerpt_marker in text and f"[{excerpt_type}]" in text:
        return excerpt_category

    mark = TYPE_MARK_RE.search(text)
    if mark:
        detected_type = mark.group(1)
        mapped = type_to_category.get(detected_type)
        if isinstance(mapped, str) and mapped:
            return mapped
    return fallback


def apply_normalization_rules(
    text: str,
    normalization_rules: List[NormalizationRule],
) -> str:
    for rule in normalization_rules:
        match = rule.pattern.match(text)
        if not match:
            continue
        group_data: Dict[str, str] = {}
        for key, value in match.groupdict().items():
            if value is None:
                group_data[key] = ""
            elif isinstance(value, str):
                group_data[key] = value.strip()
            else:
                group_data[key] = str(value)
        try:
            rendered = rule.template.format(**group_data)
        except KeyError as exc:
            raise RuntimeError(
                f"normalization rule '{rule.name}' 模板字段缺失: {exc}"
            ) from exc
        rendered = re.sub(r"\s+", " ", rendered).strip()
        rendered = re.sub(r"\s+([,.:])", r"\1", rendered)
        return rendered
    return text


def normalize_reference(
    text: str,
    rules: Dict[str, Any],
    reference_span_patterns: List[Pattern[str]],
    normalization_rules: List[NormalizationRule],
) -> str:
    text = basic_cleanup(text)
    text = trim_after_reference_tail(text, reference_span_patterns=reference_span_patterns)
    strip_types = as_str_list(
        rules.get("strip_page_for_non_excerpt_types"),
        DEFAULT_RULES["strip_page_for_non_excerpt_types"],
    )
    text = strip_trailing_page_for_non_excerpt_types(
        text,
        strip_types=strip_types,
    )
    text = apply_normalization_rules(
        text,
        normalization_rules=normalization_rules,
    )
    text = enforce_terminal_period(text)
    return text


def dedupe_key(normalized_text: str) -> str:
    text = unicodedata.normalize("NFKD", normalized_text).casefold()
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    # Keep only letters/numbers so case, spaces, punctuation and accents won't affect dedupe.
    return "".join(ch for ch in text if unicodedata.category(ch)[0] in {"L", "N"})


def parse_page_bounds(page_text: str) -> Optional[Tuple[int, int]]:
    nums = [int(x) for x in re.findall(r"\d+", page_text)]
    if not nums:
        return None
    return min(nums), max(nums)


def is_page_merge_eligible(category: str, rules: Dict[str, Any]) -> bool:
    excluded = set(
        as_str_list(
            rules.get("page_merge_exclude_categories"),
            DEFAULT_RULES["page_merge_exclude_categories"],
        )
    )
    return category not in excluded


def split_for_page_merge(
    normalized_text: str,
    category: str,
    rules: Dict[str, Any],
) -> Tuple[str, Optional[PageMergeInfo]]:
    if not is_page_merge_eligible(category, rules):
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


def build_loose_dedupe_key(
    normalized_text: str,
    category: str,
    rules: Dict[str, Any],
) -> Optional[str]:
    loose_categories = set(
        as_str_list(
            rules.get("loose_dedupe_categories"),
            DEFAULT_RULES["loose_dedupe_categories"],
        )
    )
    if category not in loose_categories:
        return None

    match = BOOK_OR_PROC_RE.match(normalized_text)
    if not match:
        return None

    base = "|".join(
        [
            match.group("author").strip(),
            match.group("title").strip(),
            match.group("type").strip(),
            match.group("place").strip(),
            match.group("year").strip(),
        ]
    )
    return dedupe_key(base)


def process_references(footnotes: List[str], runtime_cfg: RuntimeConfig) -> List[RefItem]:
    items: List[RefItem] = []
    seen_strict: Dict[str, int] = {}
    seen_loose: Dict[str, int] = {}
    page_merge_by_item: Dict[int, PageMergeInfo] = {}

    for ref in footnotes:
        raw = basic_cleanup(ref)
        original_idx, body = split_leading_index(raw)
        normalized = normalize_reference(
            body,
            rules=runtime_cfg.rules,
            reference_span_patterns=runtime_cfg.reference_span_patterns,
            normalization_rules=runtime_cfg.normalization_rules,
        )
        category = detect_category(
            normalized,
            category_detection=runtime_cfg.category_detection,
        )
        base_text, page_merge = split_for_page_merge(
            normalized_text=normalized,
            category=category,
            rules=runtime_cfg.rules,
        )
        key_text = base_text if page_merge else normalized
        strict_key = dedupe_key(key_text)
        loose_key = build_loose_dedupe_key(
            normalized_text=normalized,
            category=category,
            rules=runtime_cfg.rules,
        )

        existing_idx = seen_strict.get(strict_key)
        if existing_idx is None and loose_key:
            existing_idx = seen_loose.get(loose_key)

        if existing_idx is not None:
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
        seen_strict[strict_key] = item_idx
        if loose_key:
            seen_loose.setdefault(loose_key, item_idx)
        if page_merge:
            page_merge_by_item[item_idx] = page_merge
    return items


def render_output_lines(
    items: List[RefItem],
    category_order: List[Tuple[str, str]],
    keep_original_number: bool = False,
) -> List[str]:
    grouped: Dict[str, List[RefItem]] = {k: [] for k, _ in category_order}
    for item in items:
        grouped.setdefault(item.category, []).append(item)

    lines: List[str] = []
    next_index = 1
    for key, title in category_order:
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


def write_docx_output(
    lines: List[str],
    output_path: Path,
    docx_style: Dict[str, Any],
) -> None:
    font_cn = str(docx_style.get("font_cn", DEFAULT_DOCX_STYLE["font_cn"]))
    font_en = str(docx_style.get("font_en", DEFAULT_DOCX_STYLE["font_en"]))
    font_size_half_points = int(
        docx_style.get("font_size_half_points", DEFAULT_DOCX_STYLE["font_size_half_points"])
    )
    line_spacing_twips = int(
        docx_style.get("line_spacing_twips", DEFAULT_DOCX_STYLE["line_spacing_twips"])
    )

    def para_xml(line: str) -> str:
        ppr_xml = (
            f'<w:pPr><w:spacing w:line="{line_spacing_twips}" '
            'w:lineRule="auto"/></w:pPr>'
        )
        if not line:
            return f"<w:p>{ppr_xml}</w:p>"
        text = escape(line)
        return (
            f"<w:p>{ppr_xml}<w:r><w:rPr>"
            f'<w:rFonts w:ascii="{escape(font_en)}" w:hAnsi="{escape(font_en)}" '
            f'w:cs="{escape(font_en)}" w:eastAsia="{escape(font_cn)}"/>'
            f'<w:sz w:val="{font_size_half_points}"/>'
            f'<w:szCs w:val="{font_size_half_points}"/>'
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
    category_order: List[Tuple[str, str]],
    docx_style: Dict[str, Any],
    keep_original_number: bool = False,
) -> None:
    lines = render_output_lines(
        items=items,
        category_order=category_order,
        keep_original_number=keep_original_number,
    )
    ext = output_path.suffix.lower()
    if ext == ".docx":
        write_docx_output(
            lines=lines,
            output_path=output_path,
            docx_style=docx_style,
        )
        return
    output_path.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")


def parse_args(argv: List[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="从 DOCX 脚注提取参考文献，并按规则校验、修正、去重和分类输出。"
    )
    parser.add_argument(
        "input_docx",
        nargs="?",
        type=Path,
        help="单个输入 Word 文件（.docx）；不提供时会批量处理 input/ 目录",
    )
    parser.add_argument(
        "--input-dir",
        type=Path,
        default=Path("input"),
        help="批量模式输入目录（默认: input）",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("output"),
        help="默认输出目录（默认: output）",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=None,
        help="单文件模式输出路径，支持 .docx 或 .txt",
    )
    parser.add_argument(
        "--profile",
        type=str,
        default=DEFAULT_PROFILE_NAME,
        help="引用格式 profile 名称（默认: gb2015，对应 config/profiles/gb2015.json）",
    )
    parser.add_argument(
        "--profile-file",
        type=Path,
        default=None,
        help="直接指定 profile JSON 文件路径（优先于 --profile）",
    )
    parser.add_argument(
        "--config",
        type=Path,
        default=None,
        help="兼容参数，等同 --profile-file",
    )
    parser.add_argument(
        "--keep-original-number",
        action="store_true",
        help="输出时保留原始 [n] 编号；默认按分类后的顺序重新编号",
    )
    return parser.parse_args(argv)


def main(argv: List[str]) -> int:
    args = parse_args(argv)

    profile_file = args.profile_file or args.config
    try:
        runtime_cfg = load_runtime_config(
            profile_name=args.profile,
            profile_file=profile_file,
        )
    except Exception as exc:  # pylint: disable=broad-except
        print(f"加载配置失败: {exc}", file=sys.stderr)
        return 1

    if args.input_docx:
        input_paths = [args.input_docx]
    else:
        if not args.input_dir.exists():
            print(f"输入目录不存在: {args.input_dir}", file=sys.stderr)
            print("请将待处理 .docx 放入 input/ 目录，或传入单个 input.docx 路径。", file=sys.stderr)
            return 1
        input_paths = sorted(
            p
            for p in args.input_dir.glob("*.docx")
            if not p.name.startswith("~$") and not p.name.startswith(".~")
        )
        if not input_paths:
            print(f"未在目录中找到可处理的 .docx: {args.input_dir}", file=sys.stderr)
            return 1

    if args.output and len(input_paths) > 1:
        print("批量处理模式下不能使用 --output，请使用 --output-dir。", file=sys.stderr)
        return 1

    args.output_dir.mkdir(parents=True, exist_ok=True)
    processed_files = 0

    for input_docx in input_paths:
        if not input_docx.exists():
            print(f"输入文件不存在: {input_docx}", file=sys.stderr)
            return 1
        if input_docx.suffix.lower() != ".docx":
            print(f"仅支持 .docx 文件: {input_docx}", file=sys.stderr)
            return 1

        if args.output:
            output_path = args.output
        else:
            output_path = args.output_dir / f"{input_docx.stem}_processed_references.docx"

        if output_path.suffix.lower() not in {".docx", ".txt"}:
            print(f"输出文件仅支持 .docx 或 .txt 后缀: {output_path}", file=sys.stderr)
            return 1
        if output_path.resolve() == input_docx.resolve():
            print(f"输出文件不能与输入文件同名: {output_path}", file=sys.stderr)
            return 1

        try:
            refs, ref_source = extract_references_in_order(input_docx)
        except Exception as exc:  # pylint: disable=broad-except
            print(f"提取文献失败 ({input_docx}): {exc}", file=sys.stderr)
            return 1

        if not refs:
            print(f"未在文档中检测到可处理的参考文献内容: {input_docx}", file=sys.stderr)
            return 1

        items = process_references(refs, runtime_cfg=runtime_cfg)
        write_output(
            items=items,
            output_path=output_path,
            category_order=runtime_cfg.category_order,
            docx_style=runtime_cfg.docx_style,
            keep_original_number=args.keep_original_number,
        )

        changed_count = sum(1 for i in items if i.changed)
        print(f"输入文件: {input_docx}")
        print(f"使用 profile: {runtime_cfg.profile_path}")
        print(f"文献来源: {ref_source}")
        print(f"提取文献: {len(refs)} 条")
        print(f"去重后文献: {len(items)} 条")
        print(f"被修正条目: {changed_count} 条")
        print(f"输出文件: {output_path}")
        print(f"源文件保持不变: {input_docx}")
        print("")
        processed_files += 1

    print(f"处理完成，共 {processed_files} 个文件。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
