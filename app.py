from __future__ import annotations

import io
import json
import os
import re
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime
from typing import Optional

import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# =============================================================================
# Page config
# =============================================================================
st.set_page_config(
    page_title="AI Novel Studio",
    page_icon="📖",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
<style>
    .stage-header { font-size: 1.35rem; font-weight: 700; margin-bottom: 0.25rem; }
    .stage-desc { color: #777; font-size: 0.92rem; margin-bottom: 1.2rem; }
    .small-note { color: #777; font-size: 0.86rem; }
    .project-card { border: 1px solid rgba(128,128,128,.25); border-radius: 12px; padding: 14px; margin: 8px 0; }
    .accent-line { height: 3px; background: linear-gradient(90deg, #1D9E75, transparent); border-radius: 3px; margin: 16px 0; }
    .version-box { border-left: 4px solid #1D9E75; padding-left: 12px; margin: 8px 0; }
    .stProgress > div > div { background-color: #1D9E75; }
    div[data-testid="stSidebarContent"] { padding-top: 1.2rem; }
</style>
""",
    unsafe_allow_html=True,
)

# =============================================================================
# AI provider config: Gemini first, BigModel/Zhipu second
# =============================================================================
DEFAULT_GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-2.5-flash")
GEMINI_FALLBACK_MODELS = [
    DEFAULT_GEMINI_MODEL,
    "gemini-2.5-flash-lite",
    "gemini-2.0-flash",
    "gemini-1.5-flash",
]
DEFAULT_GLM_MODEL = os.getenv("BIGMODEL_MODEL", "glm-4-flash-250414")


def _read_secret_or_env(secret_name: str, env_name: str | None = None) -> Optional[str]:
    env_name = env_name or secret_name
    try:
        value = st.secrets.get(secret_name, None)
        if value:
            return str(value).strip()
    except Exception:
        pass
    value = os.getenv(env_name)
    return value.strip() if value else None


def get_gemini_api_key() -> Optional[str]:
    return _read_secret_or_env("GEMINI_API_KEY")


def get_bigmodel_api_key() -> Optional[str]:
    return (
        _read_secret_or_env("BIGMODEL_API_KEY")
        or _read_secret_or_env("ZHIPU_API_KEY")
        or _read_secret_or_env("ZAI_API_KEY")
    )


def _extract_gemini_text(result: dict) -> str:
    candidates = result.get("candidates", [])
    if not candidates:
        raise RuntimeError(f"No candidates returned from Gemini: {json.dumps(result, ensure_ascii=False)[:800]}")
    parts = candidates[0].get("content", {}).get("parts", [])
    texts = [p.get("text", "") for p in parts if p.get("text")]
    if texts:
        return "\n".join(texts).strip()
    raise RuntimeError(f"No text returned from Gemini: {json.dumps(result, ensure_ascii=False)[:800]}")


def _call_one_gemini_model(api_key: str, prompt: str, model_name: str, max_tokens: int) -> str:
    clean_model = model_name.replace("models/", "").strip()
    model = urllib.parse.quote(clean_model, safe="")
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={urllib.parse.quote(api_key)}"
    payload = {
        "contents": [{"role": "user", "parts": [{"text": prompt}]}],
        "generationConfig": {
            "temperature": 0.82,
            "topP": 0.95,
            "maxOutputTokens": max_tokens,
        },
    }
    request = urllib.request.Request(
        url,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=90) as response:
            body = response.read().decode("utf-8")
    except urllib.error.HTTPError as exc:
        error_body = exc.read().decode("utf-8", errors="ignore")
        raise RuntimeError(f"HTTP {exc.code} for {clean_model}: {error_body[:800]}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Network error calling Gemini API: {exc.reason}") from exc
    return _extract_gemini_text(json.loads(body))


def list_generate_content_models(api_key: str) -> list[str]:
    url = f"https://generativelanguage.googleapis.com/v1beta/models?key={urllib.parse.quote(api_key)}"
    request = urllib.request.Request(url, method="GET")
    try:
        with urllib.request.urlopen(request, timeout=30) as response:
            body = response.read().decode("utf-8")
        payload = json.loads(body)
    except Exception:
        return []
    models = []
    for model in payload.get("models", []):
        methods = model.get("supportedGenerationMethods", [])
        name = model.get("name", "").replace("models/", "")
        if name and "generateContent" in methods:
            models.append(name)
    return models


def call_gemini_rest(api_key: str, prompt: str, max_tokens: int = 3000) -> str:
    candidate_models: list[str] = []
    for m in GEMINI_FALLBACK_MODELS:
        clean = m.replace("models/", "").strip()
        if clean and clean not in candidate_models:
            candidate_models.append(clean)
    errors = []
    for model in candidate_models:
        try:
            return _call_one_gemini_model(api_key, prompt, model, max_tokens)
        except Exception as exc:
            errors.append(str(exc))
    available = list_generate_content_models(api_key)
    for model in available:
        if model not in candidate_models:
            try:
                return _call_one_gemini_model(api_key, prompt, model, max_tokens)
            except Exception as exc:
                errors.append(str(exc))
    raise RuntimeError("Could not call Gemini. Last errors: " + " | ".join(errors[-3:]))


def _extract_openai_style_text(result: dict, provider: str) -> str:
    choices = result.get("choices", [])
    if not choices:
        raise RuntimeError(f"No choices returned from {provider}: {json.dumps(result, ensure_ascii=False)[:800]}")
    content = choices[0].get("message", {}).get("content")
    if isinstance(content, str) and content.strip():
        return content.strip()
    raise RuntimeError(f"No text returned from {provider}: {json.dumps(result, ensure_ascii=False)[:800]}")


def call_bigmodel_glm(api_key: str, prompt: str, max_tokens: int = 3000) -> str:
    url = "https://open.bigmodel.cn/api/paas/v4/chat/completions"
    payload = {
        "model": DEFAULT_GLM_MODEL,
        "messages": [
            {"role": "system", "content": "You are a professional novelist, story architect, and manuscript editor."},
            {"role": "user", "content": prompt},
        ],
        "temperature": 0.82,
        "max_tokens": max_tokens,
        "stream": False,
    }
    request = urllib.request.Request(
        url,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json", "Authorization": f"Bearer {api_key}"},
        method="POST",
    )
    try:
        with urllib.request.urlopen(request, timeout=90) as response:
            body = response.read().decode("utf-8")
    except urllib.error.HTTPError as exc:
        error_body = exc.read().decode("utf-8", errors="ignore")
        raise RuntimeError(f"HTTP {exc.code} for {DEFAULT_GLM_MODEL}: {error_body[:800]}") from exc
    except urllib.error.URLError as exc:
        raise RuntimeError(f"Network error calling BigModel GLM API: {exc.reason}") from exc
    return _extract_openai_style_text(json.loads(body), "BigModel GLM")


def call_ai(system: str, user: str, max_tokens: int = 3000) -> str:
    prompt = f"{system.strip()}\n\nUSER REQUEST:\n{user.strip()}"
    gemini_key = get_gemini_api_key()
    bigmodel_key = get_bigmodel_api_key()
    errors = []
    if gemini_key:
        try:
            return call_gemini_rest(gemini_key, prompt, max_tokens=max_tokens)
        except Exception as exc:
            errors.append(f"Gemini failed: {exc}")
    if bigmodel_key:
        try:
            return call_bigmodel_glm(bigmodel_key, prompt, max_tokens=max_tokens)
        except Exception as exc:
            errors.append(f"BigModel/Zhipu failed: {exc}")
    if not gemini_key and not bigmodel_key:
        raise RuntimeError("No AI key is configured. Add GEMINI_API_KEY or BIGMODEL_API_KEY in Streamlit secrets.")
    raise RuntimeError("AI providers failed. " + " | ".join(errors[-3:]))


def generate_text(system: str, user: str, placeholder, max_tokens: int = 3000) -> str:
    try:
        text = call_ai(system, user, max_tokens=max_tokens)
    except Exception as exc:
        st.error(f"AI call failed: {exc}")
        return ""
    shown = ""
    for chunk in re.split(r"(\n\n+)", text):
        shown += chunk
        placeholder.markdown(shown + "▌")
        time.sleep(0.005)
    placeholder.markdown(text)
    return text

# =============================================================================
# State and project helpers
# =============================================================================
PROJECT_FIELDS = [
    "stage", "project_name", "novel_title", "language", "chinese_style", "genre", "target_length",
    "story_idea", "style_note", "global_instruction", "banned_phrases", "story_bible", "characters",
    "worldbuilding", "timeline", "glossary", "outline", "chapters", "chapter_beats", "chapter_drafts",
    "chapter_versions", "pending_revisions", "full_draft", "logic_report", "format_report", "polished_draft",
    "quality_reports", "revision_notes",
]

DEFAULTS = {
    "stage": 1,
    "project_name": "my_novel_project",
    "novel_title": "My Novel",
    "language": "Chinese 中文",
    "chinese_style": "现代网文",
    "genre": "Xianxia / 仙侠",
    "target_length": "Novel (~18 chapters)",
    "story_idea": "",
    "style_note": "",
    "global_instruction": "",
    "banned_phrases": "她深吸一口气\n空气仿佛凝固了\n命运的齿轮开始转动\n这一切才刚刚开始\n眼中闪过一丝复杂",
    "story_bible": "",
    "characters": "",
    "worldbuilding": "",
    "timeline": "",
    "glossary": "",
    "outline": "",
    "chapters": [],
    "chapter_beats": {},
    "chapter_drafts": {},
    "chapter_versions": {},
    "pending_revisions": {},
    "full_draft": "",
    "logic_report": "",
    "format_report": "",
    "polished_draft": "",
    "quality_reports": {},
    "revision_notes": {},
}

for key, value in DEFAULTS.items():
    if key not in st.session_state:
        st.session_state[key] = value


def normalize_loaded_project(data: dict) -> dict:
    normalized = dict(DEFAULTS)
    normalized.update({k: v for k, v in data.items() if k in DEFAULTS})
    for dict_key in ["chapter_beats", "chapter_drafts", "chapter_versions", "pending_revisions", "quality_reports", "revision_notes"]:
        value = normalized.get(dict_key, {})
        if isinstance(value, dict):
            normalized[dict_key] = {int(k) if str(k).isdigit() else k: v for k, v in value.items()}
    return normalized


def project_data() -> dict:
    data = {field: st.session_state.get(field, DEFAULTS.get(field)) for field in PROJECT_FIELDS}
    data["saved_at"] = datetime.now().isoformat(timespec="seconds")
    data["app_version"] = "novel-studio-project-v2"
    return data


def load_project(data: dict) -> None:
    normalized = normalize_loaded_project(data)
    for key, value in normalized.items():
        st.session_state[key] = value
    rebuild_full_draft()


def safe_filename(name: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9_\-\u4e00-\u9fff]+", "_", name).strip("_")
    return cleaned or "novel_project"


def rebuild_full_draft() -> None:
    chapters = st.session_state.chapters or []
    drafts = st.session_state.chapter_drafts or {}
    parts = []
    for i, ch in enumerate(chapters):
        if i in drafts and drafts[i]:
            heading = chapter_heading(ch)
            parts.append(f"## {heading}\n\n{drafts[i]}")
    st.session_state.full_draft = "\n\n---\n\n".join(parts)


def chapter_heading(ch: dict) -> str:
    num = ch.get("num", "")
    title = ch.get("title", "")
    if "Chinese" in st.session_state.language or "Bilingual" in st.session_state.language:
        return f"第{num}章 {title}"
    return f"Chapter {num}: {title}"


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


def parse_chapters(outline_text: str) -> list[dict]:
    chapters = []
    for line in outline_text.splitlines():
        clean = line.strip().lstrip("#-* ").strip()
        patterns = [
            r"Chapter\s+(\d+)[:：]\s*(.+?)\s*[—–-]\s*(.+)",
            r"第\s*([0-9一二两三四五六七八九十百]+)\s*[章节章回]+[:：]?\s*(.+?)\s*[—–-]\s*(.+)",
            r"第\s*([0-9一二两三四五六七八九十百]+)\s*[章节章回]+\s+(.+?)\s*[—–-]\s*(.+)",
        ]
        for pat in patterns:
            m = re.match(pat, clean, re.IGNORECASE)
            if m:
                raw = m.group(1)
                num = chinese_num_to_int(raw, len(chapters) + 1)
                chapters.append({"num": num, "title": m.group(2).strip(), "summary": m.group(3).strip()})
                break
    if not chapters:
        for line in outline_text.splitlines():
            clean = line.strip().lstrip("#-* ").strip()
            if re.search(r"chapter\s+\d+|第.+[章节章回]", clean, re.IGNORECASE):
                chapters.append({"num": len(chapters) + 1, "title": clean, "summary": ""})
    return chapters


def compact_story_bible(limit: int = 3500) -> str:
    parts = [
        ("Story Bible", st.session_state.story_bible),
        ("Characters", st.session_state.characters),
        ("Worldbuilding", st.session_state.worldbuilding),
        ("Timeline", st.session_state.timeline),
        ("Glossary / Terms", st.session_state.glossary),
        ("Banned phrases", st.session_state.banned_phrases),
    ]
    text = "\n\n".join(f"## {label}\n{value}" for label, value in parts if str(value).strip())
    return text[:limit]


def writing_context() -> str:
    rules = [
        f"Project title: {st.session_state.novel_title}",
        f"Writing language: {st.session_state.language}",
        f"Genre: {st.session_state.genre}",
        f"Chinese style: {st.session_state.chinese_style}",
    ]
    if "Chinese" in st.session_state.language or "Bilingual" in st.session_state.language:
        rules.extend([
            "Write in natural Chinese novel prose unless the user explicitly asks otherwise.",
            "Avoid translation-like sentence structure and stiff exposition.",
            "Use scene action, dialogue, and subtext instead of direct explanation whenever possible.",
            "For web-novel styles, keep conflict clear, pacing strong, and chapter endings hooky.",
        ])
    else:
        rules.append("Write in polished, natural English fiction prose.")
    if st.session_state.style_note.strip():
        rules.append(f"Style note: {st.session_state.style_note.strip()}")
    if st.session_state.global_instruction.strip():
        rules.append(f"Global instruction: {st.session_state.global_instruction.strip()}")
    if st.session_state.banned_phrases.strip():
        rules.append("Avoid these banned or overused phrases when possible:\n" + st.session_state.banned_phrases.strip())
    if compact_story_bible().strip():
        rules.append("Follow this project Story Bible and do not contradict it:\n" + compact_story_bible())
    return "\n".join(f"- {rule}" for rule in rules)


def push_version(chapter_idx: int, text: str, label: str) -> None:
    versions = st.session_state.chapter_versions.setdefault(chapter_idx, [])
    versions.append({"time": datetime.now().isoformat(timespec="seconds"), "label": label, "text": text})


def revision_prompt_for_mode(mode: str, custom: str) -> str:
    presets = {
        "Make it more suspenseful": "Increase suspense, reduce early explanation, plant questions, and make the scene ending stronger.",
        "Make dialogue more natural": "Rewrite dialogue to sound more natural, character-specific, and less formal. Add subtext.",
        "Add Chinese web-novel pacing": "Make pacing clearer and more addictive: stronger conflict, faster scene movement, and a hook ending.",
        "Make prose less AI-like": "Remove generic AI phrasing, repetitive rhythm, over-explanation, and cliché emotional descriptions.",
        "Increase emotional tension": "Increase emotional stakes through action, silence, subtext, and character choices rather than direct explanation.",
        "Add sensory detail": "Add concrete sensory details tied to character perspective without slowing the scene too much.",
        "Reduce exposition": "Cut exposition and replace it with action, dialogue, conflict, and implication.",
        "Add conflict": "Add sharper conflict, resistance, obstacles, and consequences within the scene.",
        "Make ending stronger": "Rewrite the ending to create curiosity, danger, emotional impact, or an unanswered question.",
        "Rewrite in 古风 tone": "Rewrite with a more elegant 古风 tone, but avoid over-decorated language and keep readability.",
    }
    base = presets.get(mode, "")
    if custom.strip():
        return base + "\nAdditional user instruction: " + custom.strip()
    return base

# =============================================================================
# Sidebar
# =============================================================================
STAGES = [
    ("🧭", "Project Setup"),
    ("📚", "Story Bible"),
    ("📝", "Outline"),
    ("🧩", "Chapter Planner"),
    ("✍️", "Draft Writer"),
    ("🔁", "Revision Studio"),
    ("🔍", "Consistency Editor"),
    ("✨", "Polish"),
    ("📥", "Export"),
]

with st.sidebar:
    st.markdown("## 📖 AI Novel Studio")
    st.caption("Gemini first → BigModel/Zhipu fallback")
    gemini_key = get_gemini_api_key()
    bigmodel_key = get_bigmodel_api_key()
    if gemini_key and bigmodel_key:
        st.success("AI ready: Gemini + BigModel")
    elif gemini_key:
        st.success("AI ready: Gemini only")
    elif bigmodel_key:
        st.success("AI ready: BigModel only")
    else:
        st.warning("Add API keys in Streamlit secrets")
    st.markdown("---")
    for i, (icon, label) in enumerate(STAGES, 1):
        btn_type = "primary" if st.session_state.stage == i else "secondary"
        if st.button(f"{icon} {i}. {label}", use_container_width=True, type=btn_type):
            st.session_state.stage = i
            st.rerun()
    st.markdown("---")
    st.caption(f"Project: {st.session_state.project_name}")

# =============================================================================
# Stage 1: Project Setup
# =============================================================================
if st.session_state.stage == 1:
    st.markdown('<div class="stage-header">🧭 Project Setup / 项目设置</div>', unsafe_allow_html=True)
    st.markdown('<div class="stage-desc">Create or load a reusable novel project. This app saves project data as JSON so you can continue later.</div>', unsafe_allow_html=True)

    tab_new, tab_load, tab_backup = st.tabs(["Create / Edit Project", "Load Project JSON", "Backup"])

    with tab_new:
        col1, col2 = st.columns(2)
        with col1:
            st.session_state.project_name = st.text_input("Project file name", value=st.session_state.project_name)
            st.session_state.novel_title = st.text_input("Novel title / 书名", value=st.session_state.novel_title)
            languages = ["Chinese 中文", "English", "Bilingual 中英双语"]
            st.session_state.language = st.selectbox("Writing language", languages, index=languages.index(st.session_state.language))
            styles = ["现代网文", "古风/仙侠", "女频仙侠", "玄幻升级流", "悬疑推理", "都市情感", "轻小说", "宫斗权谋", "中式恐怖", "文学感", "自定义"]
            st.session_state.chinese_style = st.selectbox(
                "Chinese style preset", styles, index=styles.index(st.session_state.chinese_style) if st.session_state.chinese_style in styles else 0
            )
        with col2:
            genres = ["Xianxia / 仙侠", "Wuxia / 武侠", "Fantasy", "Science Fiction", "Romance", "Thriller", "Mystery", "Historical Fiction", "Urban Fantasy / 都市异能", "Horror", "Literary Fiction", "Other"]
            st.session_state.genre = st.selectbox("Genre", genres, index=genres.index(st.session_state.genre) if st.session_state.genre in genres else 0)
            lengths = ["Short story (~5 chapters)", "Novella (~10 chapters)", "Novel (~18 chapters)", "Long web novel (~30 chapters)"]
            st.session_state.target_length = st.selectbox(
                "Target length", lengths, index=lengths.index(st.session_state.target_length) if st.session_state.target_length in lengths else 2
            )
            st.session_state.style_note = st.text_input("Style / tone note", value=st.session_state.style_note, placeholder="例如：古风但不堆辞藻，节奏快，对白自然")

        st.session_state.story_idea = st.text_area(
            "Story idea / 故事想法",
            value=st.session_state.story_idea,
            height=130,
            placeholder="例如：一个被逐出宗门的少女，在废弃灵矿中发现会说话的古剑……",
        )
        st.session_state.global_instruction = st.text_area(
            "Global writing instruction / 全局写作要求",
            value=st.session_state.global_instruction,
            height=90,
            placeholder="例如：用中文写作；女主冷静克制；少解释，多用动作和对白推进剧情；避免翻译腔。",
        )
        st.session_state.banned_phrases = st.text_area(
            "Banned / overused phrases to avoid / 避免使用的套话",
            value=st.session_state.banned_phrases,
            height=110,
        )
        st.info("Recommended workflow: Project Setup → Story Bible → Outline → Chapter Planner → Draft Writer → Revision Studio → Consistency Editor → Polish → Export.")
        if st.button("Save settings and go to Story Bible →", type="primary"):
            st.session_state.stage = 2
            st.rerun()

    with tab_load:
        uploaded = st.file_uploader("Upload project JSON", type=["json"])
        if uploaded is not None:
            try:
                data = json.loads(uploaded.read().decode("utf-8"))
                load_project(data)
                st.success("Project loaded successfully.")
                st.rerun()
            except Exception as exc:
                st.error(f"Could not load project: {exc}")

    with tab_backup:
        backup = json.dumps(project_data(), ensure_ascii=False, indent=2)
        st.download_button(
            "Download project backup JSON",
            data=backup,
            file_name=f"{safe_filename(st.session_state.project_name)}.json",
            mime="application/json",
            use_container_width=True,
        )

# =============================================================================
# Stage 2: Story Bible
# =============================================================================
elif st.session_state.stage == 2:
    st.markdown('<div class="stage-header">📚 Story Bible / 小说设定集</div>', unsafe_allow_html=True)
    st.markdown('<div class="stage-desc">Keep names, relationships, rules, timeline, locations, clues, and style constraints consistent across chapters.</div>', unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        st.session_state.characters = st.text_area("Characters / 人物设定", value=st.session_state.characters, height=220)
        st.session_state.timeline = st.text_area("Timeline / 时间线", value=st.session_state.timeline, height=180)
    with col2:
        st.session_state.worldbuilding = st.text_area("Worldbuilding, rules, factions / 世界观、规则、势力", value=st.session_state.worldbuilding, height=220)
        st.session_state.glossary = st.text_area("Glossary: names, terms, objects / 术语表", value=st.session_state.glossary, height=180)

    st.session_state.story_bible = st.text_area(
        "Full Story Bible / 总设定集",
        value=st.session_state.story_bible,
        height=220,
        placeholder="核心设定、人物关系、伏笔、禁忌、世界规则、写作风格规则……",
    )

    col_a, col_b = st.columns(2)
    with col_a:
        if st.button("Generate / improve Story Bible ✨", type="primary", use_container_width=True):
            with st.spinner("Building Story Bible..."):
                placeholder = st.empty()
                result = generate_text(
                    system="You are a senior story architect. Build a practical Story Bible for long-form fiction.",
                    user=f"""Writing context:
{writing_context()}

Story idea:
{st.session_state.story_idea}

Existing character notes:
{st.session_state.characters}

Existing worldbuilding notes:
{st.session_state.worldbuilding}

Existing timeline notes:
{st.session_state.timeline}

Existing glossary:
{st.session_state.glossary}

Create or improve a Story Bible with these sections:
## Core premise
## Main characters
## Relationships
## World rules / power system
## Key locations and factions
## Timeline
## Important objects / clues / foreshadowing
## Style rules
## Must avoid
""",
                    placeholder=placeholder,
                    max_tokens=4000,
                )
            if result:
                st.session_state.story_bible = result
                st.success("Story Bible updated.")
    with col_b:
        custom = st.text_area("Story Bible revision instruction", height=90, placeholder="例如：增加修炼等级体系；把男主身份隐藏更久；加入三条伏笔。")
        if st.button("Revise Story Bible 🔁", use_container_width=True):
            if not st.session_state.story_bible:
                st.error("Generate or enter a Story Bible first.")
            elif not custom.strip():
                st.error("Enter a revision instruction.")
            else:
                with st.spinner("Revising Story Bible..."):
                    placeholder = st.empty()
                    result = generate_text(
                        system="You are a senior story architect. Revise the Story Bible while preserving useful details.",
                        user=f"""Writing context:
{writing_context()}

Current Story Bible:
{st.session_state.story_bible}

Revision instruction:
{custom}

Return the revised Story Bible.
""",
                        placeholder=placeholder,
                        max_tokens=4000,
                    )
                if result:
                    st.session_state.story_bible = result
                    st.success("Story Bible revised.")

    if st.button("Next: Outline →", type="primary"):
        st.session_state.stage = 3
        st.rerun()

# =============================================================================
# Stage 3: Outline
# =============================================================================
elif st.session_state.stage == 3:
    st.markdown('<div class="stage-header">📝 Outline / 大纲</div>', unsafe_allow_html=True)
    st.markdown('<div class="stage-desc">Generate a chapter-by-chapter outline grounded in the Story Bible.</div>', unsafe_allow_html=True)

    chapter_count = re.search(r"~(\d+)", st.session_state.target_length)
    count = chapter_count.group(1) if chapter_count else "18"

    if st.button("Generate outline ✨", type="primary"):
        if not st.session_state.story_idea.strip():
            st.error("Please enter a story idea in Project Setup first.")
        else:
            with st.spinner("Writing outline..."):
                placeholder = st.empty()
                result = generate_text(
                    system="You are a professional novel editor and story architect. Create detailed, compelling outlines.",
                    user=f"""Writing context:
{writing_context()}

Create a detailed chapter-by-chapter outline for about {count} chapters.

Story idea:
{st.session_state.story_idea}

Required format:
# [Title]

## Characters
[2-6 main characters with brief descriptions]

## Story Arc
[Main conflict, escalation, midpoint, climax, ending direction]

## Chapter Outline
Chapter 1: [Title] — [2-3 sentence summary]
Chapter 2: [Title] — [2-3 sentence summary]
...continue for all chapters

## Foreshadowing Map
[List major clues and when they should pay off]
""",
                    placeholder=placeholder,
                    max_tokens=5000,
                )
            if result:
                st.session_state.outline = result
                st.session_state.chapters = parse_chapters(result)
                st.success(f"Outline complete. {len(st.session_state.chapters)} chapters detected.")

    if st.session_state.outline:
        st.markdown(st.session_state.outline)
        st.divider()
        revision = st.text_area("Revise outline instruction", height=90, placeholder="例如：增加反派动机；减少爱情线；每三章有一个小高潮。")
        if st.button("Revise outline 🔁"):
            if not revision.strip():
                st.error("Enter a revision instruction.")
            else:
                with st.spinner("Revising outline..."):
                    placeholder = st.empty()
                    result = generate_text(
                        system="You are a story architect. Revise the outline while preserving good structure.",
                        user=f"""Writing context:
{writing_context()}

Current outline:
{st.session_state.outline}

Revision instruction:
{revision}

Return the revised outline in the same chapter format.
""",
                        placeholder=placeholder,
                        max_tokens=5000,
                    )
                if result:
                    st.session_state.outline = result
                    st.session_state.chapters = parse_chapters(result)
                    st.success(f"Outline revised. {len(st.session_state.chapters)} chapters detected.")
                    st.rerun()
        if st.button("Next: Chapter Planner →", type="primary"):
            st.session_state.stage = 4
            st.rerun()

# =============================================================================
# Stage 4: Chapter Planner
# =============================================================================
elif st.session_state.stage == 4:
    st.markdown('<div class="stage-header">🧩 Chapter Planner / 章节规划</div>', unsafe_allow_html=True)
    st.markdown('<div class="stage-desc">Create a beat sheet before drafting. This improves chapter quality and consistency.</div>', unsafe_allow_html=True)

    if not st.session_state.chapters:
        st.warning("Generate an outline first.")
        st.stop()

    labels = [f"{chapter_heading(ch)} {'✓' if i in st.session_state.chapter_beats else ''}" for i, ch in enumerate(st.session_state.chapters)]
    idx = st.selectbox("Select chapter", range(len(labels)), format_func=lambda i: labels[i])
    ch = st.session_state.chapters[idx]
    st.markdown(f"**Outline summary:** {ch.get('summary', '')}")

    current = st.session_state.chapter_beats.get(idx, "")
    st.session_state.chapter_beats[idx] = st.text_area("Chapter beat sheet", value=current, height=260, placeholder="POV, location, goal, conflict, reveal, emotional turn, ending hook, must include, must avoid...")

    col1, col2 = st.columns(2)
    with col1:
        if st.button("Generate beat sheet ✨", type="primary", use_container_width=True):
            with st.spinner("Planning chapter..."):
                placeholder = st.empty()
                result = generate_text(
                    system="You are a chapter-level story planner. Create practical beat sheets that help draft strong scenes.",
                    user=f"""Writing context:
{writing_context()}

Overall outline:
{st.session_state.outline[:5000]}

Chapter to plan:
{chapter_heading(ch)}
Summary: {ch.get('summary', '')}

Create a chapter beat sheet with:
## POV character
## Scene location
## Chapter goal
## External conflict
## Internal/emotional conflict
## New information revealed
## Character change
## Must include
## Must avoid
## Ending hook
## Drafting notes
""",
                    placeholder=placeholder,
                    max_tokens=2500,
                )
            if result:
                st.session_state.chapter_beats[idx] = result
                st.success("Beat sheet generated.")
                st.rerun()
    with col2:
        instruction = st.text_area("Beat revision instruction", height=90, placeholder="例如：结尾更强；增加女主隐藏实力；减少男主出场。")
        if st.button("Revise beat sheet 🔁", use_container_width=True):
            if not st.session_state.chapter_beats.get(idx):
                st.error("Generate or enter a beat sheet first.")
            elif not instruction.strip():
                st.error("Enter a revision instruction.")
            else:
                with st.spinner("Revising beat sheet..."):
                    placeholder = st.empty()
                    result = generate_text(
                        system="You are a chapter planner. Revise the beat sheet clearly and practically.",
                        user=f"""Writing context:
{writing_context()}

Current beat sheet:
{st.session_state.chapter_beats[idx]}

Revision instruction:
{instruction}

Return the revised beat sheet.
""",
                        placeholder=placeholder,
                        max_tokens=2500,
                    )
                if result:
                    st.session_state.chapter_beats[idx] = result
                    st.success("Beat sheet revised.")
                    st.rerun()

    if st.button("Next: Draft Writer →", type="primary"):
        st.session_state.stage = 5
        st.rerun()

# =============================================================================
# Stage 5: Draft Writer
# =============================================================================
elif st.session_state.stage == 5:
    st.markdown('<div class="stage-header">✍️ Draft Writer / 章节写作</div>', unsafe_allow_html=True)
    st.markdown('<div class="stage-desc">Draft chapters from outline + Story Bible + beat sheet. Revisions are proposed side-by-side before acceptance.</div>', unsafe_allow_html=True)

    if not st.session_state.chapters:
        st.warning("Generate an outline first.")
        st.stop()

    labels = [f"{chapter_heading(ch)} {'✓' if i in st.session_state.chapter_drafts else ''}" for i, ch in enumerate(st.session_state.chapters)]
    idx = st.selectbox("Select chapter", range(len(labels)), format_func=lambda i: labels[i])
    ch = st.session_state.chapters[idx]
    beat = st.session_state.chapter_beats.get(idx, "")

    st.markdown(f"**Outline summary:** {ch.get('summary', '')}")
    if beat:
        with st.expander("View beat sheet"):
            st.markdown(beat)
    else:
        st.warning("This chapter has no beat sheet yet. You can still draft, but quality may be lower.")

    length_target = st.selectbox("Length target", ["Short scene", "Normal chapter", "Long chapter"], index=1)
    length_rule = {
        "Short scene": "700-1000 English words or 1000-1800 Chinese characters",
        "Normal chapter": "1000-1600 English words or 1800-3000 Chinese characters",
        "Long chapter": "1600-2400 English words or 3000-4500 Chinese characters",
    }[length_target]

    if st.button(f"Draft {chapter_heading(ch)} ✨", type="primary"):
        with st.spinner("Drafting chapter..."):
            placeholder = st.empty()
            result = generate_text(
                system="You are a skilled fiction writer. Draft vivid, natural, scene-driven prose.",
                user=f"""Writing context:
{writing_context()}

Overall outline:
{st.session_state.outline[:5000]}

Chapter:
{chapter_heading(ch)}
Summary: {ch.get('summary', '')}

Beat sheet:
{beat or 'No beat sheet provided. Use the outline summary.'}

Length target: {length_rule}

Draft the chapter as polished prose.
Use natural dialogue, scene action, conflict, emotional movement, and a strong ending hook.
Do not include planning notes. Return only the chapter text.
""",
                placeholder=placeholder,
                max_tokens=5000,
            )
        if result:
            if st.session_state.chapter_drafts.get(idx):
                push_version(idx, st.session_state.chapter_drafts[idx], "Before new draft")
            st.session_state.chapter_drafts[idx] = result
            rebuild_full_draft()
            st.success("Chapter drafted.")
            st.rerun()

    if st.session_state.chapter_drafts.get(idx):
        st.subheader("Current draft")
        st.text_area("Draft text", value=st.session_state.chapter_drafts[idx], height=360, key=f"draft_view_{idx}")

        st.subheader("Revise this chapter")
        mode = st.selectbox("Quick revision mode", [
            "Make it more suspenseful", "Make dialogue more natural", "Add Chinese web-novel pacing",
            "Make prose less AI-like", "Increase emotional tension", "Add sensory detail", "Reduce exposition",
            "Add conflict", "Make ending stronger", "Rewrite in 古风 tone",
        ])
        custom = st.text_area("Additional revision instruction", height=90, placeholder="例如：让女主少说话，用动作表现她的防备。")
        if st.button("Create revision proposal 🔁"):
            instruction = revision_prompt_for_mode(mode, custom)
            with st.spinner("Creating revision proposal..."):
                placeholder = st.empty()
                result = generate_text(
                    system="You are a professional fiction editor. Create a revised version without losing useful details.",
                    user=f"""Writing context:
{writing_context()}

Chapter:
{chapter_heading(ch)}

Current draft:
{st.session_state.chapter_drafts[idx]}

Revision instruction:
{instruction}

Return the revised chapter text only.
""",
                    placeholder=placeholder,
                    max_tokens=5000,
                )
            if result:
                st.session_state.pending_revisions[idx] = result
                st.success("Revision proposal created.")
                st.rerun()

        if st.session_state.pending_revisions.get(idx):
            st.subheader("Compare before accepting")
            left, right = st.columns(2)
            with left:
                st.markdown("**Original**")
                st.text_area("Original", value=st.session_state.chapter_drafts[idx], height=360, key=f"orig_{idx}")
            with right:
                st.markdown("**Proposed revision**")
                st.text_area("Proposed", value=st.session_state.pending_revisions[idx], height=360, key=f"prop_{idx}")
            a, b, c = st.columns(3)
            with a:
                if st.button("Accept revision", type="primary"):
                    push_version(idx, st.session_state.chapter_drafts[idx], "Before accepted revision")
                    st.session_state.chapter_drafts[idx] = st.session_state.pending_revisions.pop(idx)
                    rebuild_full_draft()
                    st.success("Revision accepted.")
                    st.rerun()
            with b:
                if st.button("Save proposal as version"):
                    push_version(idx, st.session_state.pending_revisions[idx], "Saved proposal")
                    st.success("Proposal saved in version history.")
            with c:
                if st.button("Discard proposal"):
                    st.session_state.pending_revisions.pop(idx, None)
                    st.rerun()

    if st.button("Next: Revision Studio →", type="primary"):
        st.session_state.stage = 6
        st.rerun()

# =============================================================================
# Stage 6: Revision Studio
# =============================================================================
elif st.session_state.stage == 6:
    st.markdown('<div class="stage-header">🔁 Revision Studio / 修改工作台</div>', unsafe_allow_html=True)
    st.markdown('<div class="stage-desc">Continue writing, rewrite passages, score chapter quality, and manage version history.</div>', unsafe_allow_html=True)

    if not st.session_state.chapter_drafts:
        st.warning("Draft at least one chapter first.")
        st.stop()

    drafted_indices = sorted(st.session_state.chapter_drafts.keys())
    idx = st.selectbox("Select drafted chapter", drafted_indices, format_func=lambda i: chapter_heading(st.session_state.chapters[i]) if i < len(st.session_state.chapters) else f"Chapter {i+1}")
    draft = st.session_state.chapter_drafts[idx]
    st.text_area("Current chapter draft", value=draft, height=320, key=f"revision_current_{idx}")

    tab_continue, tab_quality, tab_versions = st.tabs(["Continue / Rewrite", "Quality Score", "Version History"])
    with tab_continue:
        continue_from = st.text_area("Paste last paragraph or selected passage", height=120, placeholder="Paste the end of the scene or a passage to continue/rewrite...")
        action = st.selectbox("Action", ["Continue 800 Chinese characters / 500 English words", "Continue until scene ending", "Rewrite selected passage", "Rewrite with more dialogue", "Rewrite with more action", "Rewrite with more emotional conflict"])
        extra = st.text_area("Extra instruction", height=90)
        if st.button("Run continuation / rewrite ✨", type="primary"):
            source = continue_from.strip() or draft[-1200:]
            with st.spinner("Writing..."):
                placeholder = st.empty()
                result = generate_text(
                    system="You are a fiction writer and line editor. Continue or rewrite prose seamlessly.",
                    user=f"""Writing context:
{writing_context()}

Current chapter context:
{draft[-2500:]}

Selected passage or continuation point:
{source}

Action: {action}
Extra instruction: {extra or 'None'}

Return only the new or rewritten prose.
""",
                    placeholder=placeholder,
                    max_tokens=3500,
                )
            if result:
                st.session_state.pending_revisions[idx] = result
                st.success("Output created. Copy it into the chapter or accept it as a replacement in Draft Writer.")
                st.markdown(result)

    with tab_quality:
        if st.button("Score chapter quality 📊"):
            with st.spinner("Scoring chapter..."):
                placeholder = st.empty()
                result = generate_text(
                    system="You are a strict but helpful fiction editor. Score chapter craft and give actionable fixes.",
                    user=f"""Writing context:
{writing_context()}

Chapter draft:
{draft[:6000]}

Score with 1-10 ratings:
- Hook strength
- Conflict
- Character voice
- Dialogue naturalness
- Pacing
- Emotional tension
- Continuity risk
- AI-sounding risk

Then give the top 5 concrete fixes.
""",
                    placeholder=placeholder,
                    max_tokens=2500,
                )
            if result:
                st.session_state.quality_reports[idx] = result
        if st.session_state.quality_reports.get(idx):
            st.markdown(st.session_state.quality_reports[idx])

    with tab_versions:
        versions = st.session_state.chapter_versions.get(idx, [])
        if not versions:
            st.info("No saved versions yet. Versions are created when you accept revisions or save proposals.")
        for v_i, version in enumerate(reversed(versions), 1):
            with st.expander(f"Version {len(versions)-v_i+1}: {version.get('label')} — {version.get('time')}"):
                st.text_area("Version text", value=version.get("text", ""), height=220, key=f"version_{idx}_{v_i}")

    if st.button("Next: Consistency Editor →", type="primary"):
        st.session_state.stage = 7
        st.rerun()

# =============================================================================
# Stage 7: Consistency Editor
# =============================================================================
elif st.session_state.stage == 7:
    st.markdown('<div class="stage-header">🔍 Consistency Editor / 连续性与设定检查</div>', unsafe_allow_html=True)
    st.markdown('<div class="stage-desc">Check manuscript against Story Bible: character facts, timeline, rules, clues, tone, and Chinese name consistency.</div>', unsafe_allow_html=True)

    rebuild_full_draft()
    draft = st.session_state.full_draft or st.session_state.outline
    if not draft:
        st.warning("Draft chapters first.")
        st.stop()

    focus = st.multiselect(
        "Check types",
        ["Character consistency", "Timeline consistency", "Worldbuilding/rule consistency", "Relationship consistency", "Foreshadowing/clue consistency", "Chinese name consistency", "Tone consistency", "Actionable rewrite plan"],
        default=["Character consistency", "Timeline consistency", "Worldbuilding/rule consistency", "Foreshadowing/clue consistency", "Actionable rewrite plan"],
    )
    extra = st.text_area("Extra check instruction", height=80, placeholder="例如：重点检查修仙等级是否前后矛盾，伏笔是否太早泄露。")

    if st.button("Run consistency check 🔍", type="primary"):
        with st.spinner("Checking consistency..."):
            placeholder = st.empty()
            result = generate_text(
                system="You are a professional continuity editor for long-form fiction.",
                user=f"""Writing context:
{writing_context()}

Check types:
{', '.join(focus)}

Extra instruction:
{extra or 'None'}

Story Bible:
{compact_story_bible(5000)}

Outline:
{st.session_state.outline[:4000]}

Manuscript excerpt / current draft:
{draft[:9000]}

Return:
## High-priority continuity problems
## Possible contradictions with Story Bible
## Timeline / relationship / rule issues
## Foreshadowing and unresolved clues
## Specific chapter-level fixes
## Safe rewrite plan
""",
                placeholder=placeholder,
                max_tokens=4500,
            )
        if result:
            st.session_state.logic_report = result
            st.success("Consistency check complete.")

    if st.session_state.logic_report:
        st.markdown(st.session_state.logic_report)
        revision = st.text_area("Revise consistency report instruction", height=80, placeholder="例如：改成表格；只保留最重要的10个问题；给出逐章修复清单。")
        if st.button("Revise report 🔁"):
            if revision.strip():
                with st.spinner("Revising report..."):
                    placeholder = st.empty()
                    result = generate_text(
                        system="You are a continuity editor. Revise the report to be more useful.",
                        user=f"Current report:\n{st.session_state.logic_report}\n\nRevision instruction:\n{revision}",
                        placeholder=placeholder,
                        max_tokens=3500,
                    )
                if result:
                    st.session_state.logic_report = result
                    st.rerun()

    if st.button("Next: Polish →", type="primary"):
        st.session_state.stage = 8
        st.rerun()

# =============================================================================
# Stage 8: Polish
# =============================================================================
elif st.session_state.stage == 8:
    st.markdown('<div class="stage-header">✨ Polish / 润色去AI味</div>', unsafe_allow_html=True)
    st.markdown('<div class="stage-desc">Polish selected chapters for natural prose, Chinese rhythm, dialogue, pacing, and style.</div>', unsafe_allow_html=True)

    rebuild_full_draft()
    if not st.session_state.full_draft:
        st.warning("Draft chapters first.")
        st.stop()

    choices = ["Full draft excerpt"] + [chapter_heading(ch) for i, ch in enumerate(st.session_state.chapters) if i in st.session_state.chapter_drafts]
    section = st.selectbox("Section to polish", choices)
    intensity = st.selectbox("Polish intensity", ["Light", "Medium", "Deep"], index=1)
    polish_targets = st.multiselect(
        "Polish targets",
        ["Remove translation tone", "Reduce repetitive sentence patterns", "Avoid generic AI phrases", "Make dialogue less formal", "Add subtext", "Use natural Chinese rhythm", "Avoid too much exposition", "Strengthen hook", "Improve sensory detail"],
        default=["Avoid generic AI phrases", "Use natural Chinese rhythm", "Make dialogue less formal", "Avoid too much exposition"],
    )
    extra = st.text_area("Extra polish instruction", height=90, placeholder="例如：更像女频仙侠，情绪克制但张力更强。")

    if section == "Full draft excerpt":
        text_to_polish = st.session_state.full_draft[:5000]
    else:
        idx = next((i for i, ch in enumerate(st.session_state.chapters) if chapter_heading(ch) == section), 0)
        text_to_polish = st.session_state.chapter_drafts.get(idx, st.session_state.full_draft[:5000])

    if st.button("Polish ✨", type="primary"):
        with st.spinner("Polishing..."):
            placeholder = st.empty()
            result = generate_text(
                system="You are a master literary editor specializing in natural Chinese and English fiction prose.",
                user=f"""Writing context:
{writing_context()}

Section: {section}
Intensity: {intensity}
Targets: {', '.join(polish_targets)}
Extra instruction: {extra or 'None'}

Original:
{text_to_polish}

Return:
## What changed
[brief notes]

## Polished version
[rewritten text]
""",
                placeholder=placeholder,
                max_tokens=5000,
            )
        if result:
            st.session_state.polished_draft = result
            st.success("Polish complete.")

    if st.session_state.polished_draft:
        st.markdown(st.session_state.polished_draft)
        revision = st.text_area("Revise polished version instruction", height=80)
        if st.button("Revise polished version 🔁"):
            if revision.strip():
                with st.spinner("Revising polished version..."):
                    placeholder = st.empty()
                    result = generate_text(
                        system="You are a literary editor. Revise the polished version based on the user's instruction.",
                        user=f"Writing context:\n{writing_context()}\n\nCurrent polished result:\n{st.session_state.polished_draft}\n\nRevision instruction:\n{revision}",
                        placeholder=placeholder,
                        max_tokens=5000,
                    )
                if result:
                    st.session_state.polished_draft = result
                    st.rerun()

    if st.button("Next: Export →", type="primary"):
        st.session_state.stage = 9
        st.rerun()

# =============================================================================
# Stage 9: Export
# =============================================================================
elif st.session_state.stage == 9:
    st.markdown('<div class="stage-header">📥 Export / 导出</div>', unsafe_allow_html=True)
    st.markdown('<div class="stage-desc">Export manuscript, Story Bible, notes, DOCX, TXT, Markdown, and JSON project backup.</div>', unsafe_allow_html=True)

    rebuild_full_draft()
    final = st.session_state.polished_draft or st.session_state.full_draft or st.session_state.outline
    word_count = len(final.split())
    char_count = len(final)
    ch_count = len(st.session_state.chapter_drafts)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Drafted chapters", ch_count)
    c2.metric("Approx. words", f"{word_count:,}")
    c3.metric("Characters", f"{char_count:,}")
    c4.metric("Language", st.session_state.language)

    def build_markdown(include_notes: bool = True) -> str:
        parts = [
            f"# {st.session_state.novel_title}",
            f"- Language: {st.session_state.language}",
            f"- Genre: {st.session_state.genre}",
            f"- Style: {st.session_state.chinese_style}",
            "\n---\n",
            "# Story Bible",
            compact_story_bible(20000),
            "\n---\n",
            "# Outline",
            st.session_state.outline,
            "\n---\n",
            "# Full Draft",
            st.session_state.full_draft,
        ]
        if include_notes and st.session_state.logic_report:
            parts += ["\n---\n", "# Consistency Report", st.session_state.logic_report]
        if include_notes and st.session_state.polished_draft:
            parts += ["\n---\n", "# Polished Sections", st.session_state.polished_draft]
        return "\n\n".join(p for p in parts if p)

    def build_docx() -> bytes:
        doc = Document()
        title_para = doc.add_paragraph()
        title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = title_para.add_run(st.session_state.novel_title.upper())
        run.bold = True
        run.font.size = Pt(24)
        sub = doc.add_paragraph()
        sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
        sub.add_run(f"Language: {st.session_state.language} | Genre: {st.session_state.genre}")
        doc.add_page_break()
        for heading, text in [
            ("Story Bible", compact_story_bible(20000)),
            ("Outline", st.session_state.outline),
            ("Full Draft", st.session_state.full_draft),
            ("Consistency Report", st.session_state.logic_report),
            ("Polished Sections", st.session_state.polished_draft),
        ]:
            if text:
                doc.add_heading(heading, level=1)
                for line in text.splitlines():
                    if line.startswith("# "):
                        doc.add_heading(line[2:], level=1)
                    elif line.startswith("## "):
                        doc.add_heading(line[3:], level=2)
                    elif line.strip():
                        doc.add_paragraph(line.strip())
                doc.add_page_break()
        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        return buf.getvalue()

    filename = safe_filename(st.session_state.project_name or st.session_state.novel_title)
    md = build_markdown()
    txt = md.replace("#", "")
    json_backup = json.dumps(project_data(), ensure_ascii=False, indent=2)

    col_a, col_b, col_c, col_d = st.columns(4)
    with col_a:
        st.download_button("Download TXT", txt, f"{filename}.txt", "text/plain", use_container_width=True)
    with col_b:
        st.download_button("Download Markdown", md, f"{filename}.md", "text/markdown", use_container_width=True)
    with col_c:
        st.download_button("Download DOCX", build_docx(), f"{filename}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
    with col_d:
        st.download_button("Download Project JSON", json_backup, f"{filename}.json", "application/json", use_container_width=True)

    st.divider()
    st.markdown("**Preview**")
    st.text(final[:2500] + ("..." if len(final) > 2500 else ""))
