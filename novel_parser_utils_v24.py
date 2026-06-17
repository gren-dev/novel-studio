from __future__ import annotations
import re


def chinese_num_to_int(raw: str, fallback: int) -> int:
    try:
        return int(raw)
    except Exception:
        pass
    digits = {"零": 0, "一": 1, "二": 2, "两": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}
    if raw == "十":
        return 10
    if "十" in raw:
        left, _, right = raw.partition("十")
        tens = digits.get(left, 1) if left else 1
        ones = digits.get(right, 0) if right else 0
        return tens * 10 + ones
    return digits.get(raw, fallback)


def _clean_chapter_title(title: str) -> str:
    title = re.sub(r"^[#>*\-\s]+", "", str(title)).strip()
    title = re.sub(r"\*\*|__|`", "", title).strip()
    title = re.sub(r"\s+", " ", title).strip(" ：:-—–")
    return title or "未命名章节"


def _append_chapter(chapters: list[dict], seen_nums: set[int], num: int, title: str, summary: str = "") -> None:
    if num in seen_nums:
        return
    title = _clean_chapter_title(title)
    summary = str(summary or "").strip().lstrip("—–-:： ").strip()
    chapters.append({"num": num, "title": title, "summary": summary})
    seen_nums.add(num)


def parse_chapters(outline_text: str) -> list[dict]:
    chapters: list[dict] = []
    seen_nums: set[int] = set()
    lines = outline_text.splitlines()
    heading_patterns = [
        r"^Chapter\s+([0-9]+)\s*[:：\-.、]?\s*(.+?)(?:\s*[—–-]\s*(.+))?$",
        r"^Ch\.?\s*([0-9]+)\s*[:：\-.、]?\s*(.+?)(?:\s*[—–-]\s*(.+))?$",
        r"^第\s*([0-9一二两三四五六七八九十百]+)\s*[章节章回卷部]+\s*[:：\-.、]?\s*(.+?)(?:\s*[—–-]\s*(.+))?$",
        r"^([0-9]{1,3})\s*[\.、)）]\s*(.+?)(?:\s*[—–-]\s*(.+))?$",
        r"^([一二两三四五六七八九十百]+)\s*[、\.．)]\s*(.+?)(?:\s*[—–-]\s*(.+))?$",
    ]
    skip_foreshadowing_section = False
    for raw_line in lines:
        raw_clean = raw_line.strip()
        is_markdown_heading = bool(re.match(r"^\s{0,4}#{1,6}\s+", raw_clean))
        clean = raw_clean
        if not clean:
            continue
        clean = re.sub(r"^\s{0,4}(#{1,6})\s*", "", clean).strip()
        heading_plain = re.sub(r"\*\*|__|`", "", clean).strip()
        if is_markdown_heading and re.match(r"^(Foreshadowing|Foreshadowing\s+Map|伏笔|线索|线索回收)", heading_plain, re.I):
            skip_foreshadowing_section = True
            continue
        if is_markdown_heading and re.match(r"^(Chapter\s+Outline|章节大纲)", heading_plain, re.I):
            skip_foreshadowing_section = False
            continue
        if skip_foreshadowing_section:
            continue
        clean = re.sub(r"^[-*+]\s+", "", clean).strip()
        clean = re.sub(r"\*\*|__|`", "", clean).strip()
        if not clean:
            continue
        if re.match(r"^(Chapter\s+Outline|章节大纲|大纲|Story\s+Arc|Foreshadowing|Characters)\b", clean, re.I):
            continue
        for pat in heading_patterns:
            m = re.match(pat, clean, re.IGNORECASE)
            if not m:
                continue
            raw_num, title = m.group(1), m.group(2)
            summary = m.group(3) if len(m.groups()) >= 3 else ""
            num = chinese_num_to_int(raw_num, len(chapters) + 1)
            if len(title) > 160 and not summary:
                title, summary = title[:60], title[60:]
            _append_chapter(chapters, seen_nums, num, title, summary)
            break
    if not chapters:
        for raw_line in lines:
            clean = raw_line.strip().lstrip("#-*+ >").strip()
            clean = re.sub(r"\*\*|__|`", "", clean).strip()
            if re.search(r"chapter\s+\d+|第\s*[0-9一二两三四五六七八九十百]+\s*[章节章回]", clean, re.IGNORECASE):
                _append_chapter(chapters, seen_nums, len(chapters) + 1, clean, "")
    chapters.sort(key=lambda c: c.get("num", 0))
    return chapters


def infer_chapters_from_story_bible_timeline(story_bible: str) -> list[dict]:
    if not story_bible or not story_bible.strip():
        return []
    lines = story_bible.splitlines()
    in_timeline = False
    timeline_lines: list[str] = []
    for raw in lines:
        clean = raw.strip()
        heading = re.sub(r"^#+\s*", "", clean).strip()
        heading_plain = re.sub(r"\*\*|__|`", "", heading).strip()
        if re.search(r"时间线|Timeline", heading_plain, re.IGNORECASE):
            in_timeline = True
            continue
        if in_timeline and clean.startswith("##") and not re.search(r"时间线|Timeline", heading_plain, re.IGNORECASE):
            break
        if in_timeline and clean:
            timeline_lines.append(clean)
    if not timeline_lines:
        return []
    chapters: list[dict] = []
    seen_nums: set[int] = set()
    for raw in timeline_lines:
        clean = raw.strip().lstrip("-*+ >").strip()
        clean = re.sub(r"\*\*|__|`", "", clean).strip()
        if not clean:
            continue
        m = re.match(r"^([0-9一二两三四五六七八九十百]+)\s*[\.、)）]?\s*(.+)$", clean)
        if not m:
            continue
        raw_num = m.group(1)
        body = m.group(2).strip()
        num = chinese_num_to_int(raw_num, len(chapters) + 1)
        title = body
        summary = ""
        body = re.sub(r"^第?[一二两三四五六七八九十百0-9]+阶段[：:]?", "", body).strip()
        if "——" in body:
            title, summary = body.split("——", 1)
        elif "—" in body:
            title, summary = body.split("—", 1)
        elif " - " in body:
            title, summary = body.split(" - ", 1)
        elif "：" in body:
            maybe_title, maybe_summary = body.split("：", 1)
            if len(maybe_title) <= 30:
                title, summary = maybe_title, maybe_summary
        title = re.sub(r"^[（(].*?[）)]", "", title).strip(" ：:-—")
        title = title or f"阶段 {num}"
        summary = summary.strip() or clean
        _append_chapter(chapters, seen_nums, num, title, summary)
    chapters.sort(key=lambda c: c.get("num", 0))
    return chapters
