from novel_parser_utils_v24 import chinese_num_to_int, parse_chapters, infer_chapters_from_story_bible_timeline


def test_chinese_num_to_int_common_cases():
    cases = {
        "一": 1,
        "二": 2,
        "两": 2,
        "十": 10,
        "十一": 11,
        "二十": 20,
        "二十一": 21,
        "三十六": 36,
        "9": 9,
    }
    for raw, expected in cases.items():
        assert chinese_num_to_int(raw, 99) == expected


def test_parse_chapters_chinese_standard_format():
    outline = """
## Chapter Outline
第1章：夜入顾府 — 沈清和第一次进入顾府，发现旧案线索。
第2章：灯下问话 — 顾廷川试探她，她避开关键问题。
第3章：雨声停后 — 两人找到一封被藏起的旧信。
"""
    chapters = parse_chapters(outline)
    assert [c["num"] for c in chapters] == [1, 2, 3]
    assert chapters[0]["title"] == "夜入顾府"
    assert "旧案线索" in chapters[0]["summary"]


def test_parse_chapters_common_ai_variants():
    outline = """
# 暗涌
**第一章：夜入顾府**
## 第 2 章 灯下问话
Chapter 3: Rain Stops - The hidden letter appears.
Ch. 4 - Second Gate
5. 庭前旧影 — 她看见旧日证人。
六、归档之前 — 她把证据放回匣中。
"""
    chapters = parse_chapters(outline)
    assert [c["num"] for c in chapters] == [1, 2, 3, 4, 5, 6]
    assert chapters[1]["title"] == "灯下问话"
    assert chapters[2]["title"] == "Rain Stops"
    assert "hidden letter" in chapters[2]["summary"]


def test_parse_chapters_ignores_section_headings_and_duplicates():
    outline = """
## Characters
沈清和：女主。
## Story Arc
她回到顾府。
第1章：夜入顾府 — 摘要一。
第1章：重复章节 — 不应加入。
## Foreshadowing Map
第7章回收旧信。
"""
    chapters = parse_chapters(outline)
    assert len(chapters) == 1
    assert chapters[0]["title"] == "夜入顾府"


def test_parse_chapters_without_summary_still_detects_chapters():
    outline = """
第1章 夜入顾府
第2章 灯下问话
第3章 雨声停后
"""
    chapters = parse_chapters(outline)
    assert len(chapters) == 3
    assert chapters[0]["summary"] == ""


def test_infer_chapters_from_story_bible_timeline():
    bible = """
## 核心前提
沈清和重回顾府。

## 时间线
1. 第一阶段：夜入顾府 —— 她进入顾府，避开巡夜人。
2. 第二阶段：灯下问话 —— 顾廷川逼问旧案。
3. 第三阶段：雨声停后 —— 旧信被找到。

## 主要角色
这里不应继续解析。
4. 伪章节 —— 不应加入。
"""
    chapters = infer_chapters_from_story_bible_timeline(bible)
    assert [c["num"] for c in chapters] == [1, 2, 3]
    assert chapters[0]["title"] == "夜入顾府"
    assert "避开巡夜人" in chapters[0]["summary"]


def test_infer_chapters_empty_timeline_returns_empty():
    assert infer_chapters_from_story_bible_timeline("## 主要角色\n无") == []
