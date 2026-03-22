"""
学研·工科科研助手 v4.0 (Xueyan OS)
双引擎架构：论文内容写作 + 确定性格式对齐
"""

import os
import re
import time
from collections import Counter
from copy import deepcopy
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Dict, List, Tuple

import anthropic
import fitz  # pymupdf
import streamlit as st
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
from tenacity import retry, retry_if_exception_type, stop_after_attempt, wait_exponential

# ── Configuration ────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="学研·工科科研助手 v4.0",
    layout="wide",
    initial_sidebar_state="expanded",
)

CLAUDE_API_KEY = os.getenv("ANTHROPIC_AUTH_TOKEN", "").strip()
CLAUDE_BASE_URL = os.getenv("ANTHROPIC_BASE_URL", "https://aiapi.aixia.tech").rstrip("/")
CLAUDE_MODEL = os.getenv("ANTHROPIC_MODEL", "claude-opus-4-6")

WRITING_PAGE = "✍️ 论文内容写作"
FORMATTING_PAGE = "📏 确定性格式对齐"
SECTION_NAMES = ["摘要", "引言", "方法", "结果", "讨论", "结论"]
MODIFICATION_FUNCTIONS = {"📋 Redlining修订", "✨ 表达润色", "🤖 去AI味 (Humanizer)", "🎯 精修模式"}
SHADOW_FUNCTIONS = {"✍️ 逐段起草", "💡 研究想法构思", "📄 节节头脑风暴", "✍️ 影子写作"}

DEFAULT_STATES = {
    "engine_mode": WRITING_PAGE,
    "writing_active_section": "摘要",
    "writing_history": [],
    "writing_reference_docs": {},
    "writing_import_targets": {},
    "format_guideline_text": "",
    "format_guideline_summary": "",
    "format_audit_report": {},
    "format_fix_options": [],
    "format_output_docx_bytes": b"",
    "format_last_filename": "",
}

for key, default in DEFAULT_STATES.items():
    if key not in st.session_state:
        st.session_state[key] = deepcopy(default)

# ── 常量定义 ─────────────────────────────────────────────────────────────────
FUNCTION_MATRIX = {
    "📝 中转英翻译": {
        "description": "执行中英双向精准学术翻译，严格跨语种输出",
        "rules": ["✅ 必须跨语种输出", "❌ 禁止同语输出", "✅ 工科论文表达"]
    },
    "✨ 表达润色": {
        "description": "提升学术地道性，同语言优化",
        "rules": ["❌ 禁止翻译", "❌ 禁止改变原意", "✅ 仅限同语言"]
    },
    "🔍 逻辑检查": {
        "description": "检查因果链条和衔接",
        "rules": ["❌ 仅查逻辑", "❌ 不修改文本", "✅ 指出问题"]
    },
    "🤖 去AI味 (Humanizer)": {
        "description": "消除AI痕迹，模仿真人语序",
        "rules": ["❌ 禁止翻译", "❌ 禁止增加论点", "✅ CN优化CN/EN优化EN"]
    },
    "👨‍⚖️ Reviewer视角": {
        "description": "模拟审稿人视角审视",
        "rules": ["✅ 提供改进建议", "✅ 指出薄弱环节"]
    },
    "💡 研究想法构思": {
        "description": "从零构思研究方向",
        "rules": ["✅ 基于研究领域", "✅ 提供创新点"]
    },
    "🧠 ML论文写作": {
        "description": "NeurIPS/ICML级别写作",
        "rules": ["✅ 引用格式验证", "✅ 图表描述规范"]
    },
    "📊 概念图设计": {
        "description": "生成论文概念图设计",
        "rules": ["✅ 设计哲学", "✅ 绘图提示词"]
    },
    "📄 节节头脑风暴": {
        "description": "按章节进行头脑风暴",
        "rules": ["✅ 针对当前板块", "✅ 提供思路"]
    },
    "✍️ 逐段起草": {
        "description": "将大纲扩充为正式段落",
        "rules": ["✅ 模仿标杆文献", "✅ 保持学术规范"]
    },
    "✍️ 影子写作": {
        "description": "优先模仿标杆文献叙事节奏进行起草",
        "rules": ["✅ 优先模仿标杆文献", "✅ 保持当前语种", "✅ 对齐当前板块"]
    },
    "🎯 精修模式": {
        "description": "深度精修，达到顶刊水准",
        "rules": ["✅ 模仿标杆文献", "✅ 顶刊标准"]
    },
    "📝 引用验证": {
        "description": "检查引用格式和完整性",
        "rules": ["✅ 格式验证", "✅ 完整性检查"]
    },
    "🎨 图表规范检查": {
        "description": "检查图表描述规范性",
        "rules": ["✅ 色盲友好", "✅ 规范验证"]
    },
    "📋 Redlining修订": {
        "description": "带修订痕迹的修改建议",
        "rules": ["✅ 显示修改痕迹", "✅ 对比视图"]
    },
    "🔄 版本对比": {
        "description": "对比不同版本差异",
        "rules": ["✅ 高亮差异", "✅ 修改说明"]
    },
}

SECTIONS = {
    "摘要": {"focus": "开门见山，数据支撑", "max_words": 250},
    "引言": {"focus": "背景转折，研究空白", "max_words": 800},
    "方法": {"focus": "流程清晰，参数精确", "max_words": 1500},
    "结果": {"focus": "数据客观，图表引用", "max_words": 1200},
    "讨论": {"focus": "深度解读，文献对比", "max_words": 1000},
    "结论": {"focus": "总结贡献，展望未来", "max_words": 300},
}

DOMAINS = {
    "🔋 能源电池": {
        "hard_lock": [r"\$Li\^\+\$", r"\$Na\^\+\$", "NCM523", "NCM622", "XRD", "SEM", "capacity retention", "intercalation", "晶格", "应力应变"],
        "focus": "电化学性能"
    },
    "🏗️ 固废/土木": {
        "hard_lock": ["GGBS", "水化动力学", "compressive strength", "C-S-H", "pozzolanic"],
        "focus": "材料性能"
    },
    "🔩 机械/材料": {
        "hard_lock": ["tensile strength", "grain boundary", "dislocation", "microstructure", "XRD", "SEM", "晶格", "应力应变"],
        "focus": "力学性能"
    },
    "🧪 催化/化工": {
        "hard_lock": ["TOF", "turnover frequency", "BET", "heterogeneous catalysis"],
        "focus": "催化性能"
    },
    "🌊 环境工程": {
        "hard_lock": ["COD", "BOD", "MBR", "activated sludge", "removal efficiency"],
        "focus": "处理效果"
    },
    "📱 电子半导体": {
        "hard_lock": ["MOSFET", "bandgap", "GaN", "SiC", "carrier mobility"],
        "focus": "电学性能"
    },
    "🧬 生物材料": {
        "hard_lock": ["hydrogel", "cell adhesion", "biocompatibility", "MTT assay"],
        "focus": "生物相容性"
    },
    "🧠 机器学习": {
        "hard_lock": ["CNN", "Transformer", "overfitting", "gradient descent", "F1-score"],
        "focus": "模型性能"
    },
    "💻 软件工程": {
        "hard_lock": ["CI/CD", "Agile", "DevOps", "API", "microservices"],
        "focus": "工程实践"
    },
}

@st.cache_resource
def load_local_skills() -> Dict[str, str]:
    skills: Dict[str, str] = {}
    base_path = Path("./awesome-ai-research-writing")
    if not base_path.exists():
        return skills
    for md_file in base_path.rglob("*.md"):
        try:
            skills[md_file.stem] = md_file.read_text(encoding="utf-8")
        except Exception:
            continue
    return skills

LOCAL_SKILLS = load_local_skills()
SKILL_GROUPS = {
    "humanizer": HUMANIZER_RULES if "HUMANIZER_RULES" in globals() else "",
}

HUMANIZER_RULES = """
## Humanizer 核心规则（零篡位）
- 清理 AI 套话：综上所述、总而言之、in conclusion、it is worth noting that。
- 清理 AI 高危词：delve, landscape, leverage, underscore, utilize。
- 保留原语种，禁止翻译。
- 用长短句变化制造呼吸感，但不新增论点。
- 不夸大贡献，不制造不存在的创新性语气。
"""

TRANSLATION_SKILL_RULES = """
## Translation Mastery（精准跨语种）
- 必须执行跨语种转换，严禁同语输出。
- 中译英：使用工科学术表达，补全冠词、时态、被动结构。
- 英译中：去翻译腔，输出自然、凝练、可直接入文的中文。
- 保留 LaTeX、化学式、缩写、材料名、术语格式。
"""

RESULTS_SKILL_RULES = """
## Results 专项写作协议
- 结果板块优先写趋势、对比、异常点解释、图表引用。
- 遇到实验数据时，优先明确 increase/decrease、higher/lower、plateau、fluctuation 等关系。
- 自动保留 LaTeX、化学式、材料名和单位。
- 对工科实验结果保持客观，避免宣传腔。
"""

DOCX_SKILL_RULES = """
## DOCX / Redlining 协议
- Word 导出需保留板块、功能、领域和时间元信息。
- Redlining 仅用于展示修订痕迹，不改变功能职责边界。
"""

ML_CHECKLIST = """
## ML Paper Writing Checklist
- 引用格式统一
- 所有方法引用有明确出处
- 基线模型正确引用
- 图表配色、图例、误差线清晰
- 超参数表与实验设置可复现
- 统计显著性标注完整
"""

# ── 工具函数 ─────────────────────────────────────────────────────────────────
def writing_input_key(section: str) -> str:
    return f"writing_input_{section}"


def writing_output_key(section: str) -> str:
    return f"writing_output_{section}"


def writing_note_key(section: str) -> str:
    return f"writing_note_{section}"


def init_writing_state() -> None:
    for section in SECTION_NAMES:
        st.session_state.setdefault(writing_input_key(section), "")
        st.session_state.setdefault(writing_output_key(section), "")
        st.session_state.setdefault(writing_note_key(section), "")


def append_skill_preview() -> Dict[str, str]:
    readme_text = LOCAL_SKILLS.get("README", "")
    return {
        "humanizer": "humanizer" if "humanizer" in readme_text.lower() else "内置规则",
        "ml-paper": "20-ml-paper-writing" if "20-ml-paper-writing" in readme_text.lower() else "内置规则",
        "docx": "docx" if "docx" in readme_text.lower() else "内置规则",
        "thesis-formatting": "thesis-skills 设计参考",
    }


def protect_hard_terms(text: str, domain: str) -> Tuple[str, Dict[str, str]]:
    protected = text
    mapping: Dict[str, str] = {}
    counter = 0

    patterns = [
        r"\$[^$]+?\$",
        r"\$\$[^$]+?\$\$|\\\[.*?\\\]",
        r"\\[a-zA-Z]+(?:\{[^}]*\}|\[[^\]]*\]|[_^]\{[^}]*\})?",
        r"\b[A-Z][a-z]?\d*(?:_[\d-]+|[+\-]?\d*)?\b",
        r"\d+\s*(?:°C|K|MPa|GPa|kPa|Pa|wt%|vol%|at%|mol%|nm|μm|mm|cm|m|km|Hz|kHz|MHz|GHz)",
    ]

    for pattern in patterns:
        for match in re.finditer(pattern, text, re.DOTALL):
            token = match.group()
            if token in protected:
                placeholder = f"__LOCK_{counter}__"
                mapping[placeholder] = token
                protected = protected.replace(token, placeholder, 1)
                counter += 1

    for term in DOMAINS.get(domain, {}).get("hard_lock", []):
        if isinstance(term, str) and term in protected:
            placeholder = f"__LOCK_{counter}__"
            mapping[placeholder] = term
            protected = protected.replace(term, placeholder, 1)
            counter += 1

    return protected, mapping


def restore_hard_terms(text: str, mapping: Dict[str, str]) -> str:
    restored = text
    for placeholder, original in mapping.items():
        restored = restored.replace(placeholder, original)
    return restored


def detect_language(text: str) -> str:
    total_chars = len(text.strip())
    if total_chars == 0:
        return "unknown"
    chinese_chars = len(re.findall(r"[\u4e00-\u9fff]", text))
    return "zh" if chinese_chars / total_chars > 0.3 else "en"


def extract_language_genes(text: str) -> Dict:
    genes = {
        "sentence_length_pattern": [],
        "connecting_words": Counter(),
        "academic_phrases": Counter(),
        "voice_pattern": {"passive": 0, "active": 0},
        "citation_style": Counter(),
        "complexity_markers": Counter(),
    }
    sentences = [s.strip() for s in re.split(r"[.!?。！？]", text) if s.strip()]
    connectors = ["however", "therefore", "furthermore", "moreover", "thus", "hence"]
    phrases = ["previous studies", "recent work", "suggest that", "indicate that", "demonstrate that"]
    complexity = ["although", "while", "despite", "whereas", "whether or"]

    for sent in sentences[:100]:
        genes["sentence_length_pattern"].append(len(sent.split()))
        lower = sent.lower()
        for conn in connectors:
            if re.search(rf"\b{re.escape(conn)}\b", lower):
                genes["connecting_words"][conn] += 1
        for phrase in phrases:
            if phrase in lower:
                genes["academic_phrases"][phrase] += 1
        if re.search(r"\b(?:was|were|is|are|been|be)\s+\w+ed\b", lower):
            genes["voice_pattern"]["passive"] += 1
        else:
            genes["voice_pattern"]["active"] += 1
        for cite in re.findall(r"\[\d+\]|\([^)]+\d+[^)]*\)|\w+\s+et\s+al\.", sent):
            genes["citation_style"][cite[:20]] += 1
        for marker in complexity:
            if marker in lower:
                genes["complexity_markers"][marker] += 1

    if genes["sentence_length_pattern"]:
        avg = sum(genes["sentence_length_pattern"]) / len(genes["sentence_length_pattern"])
        std = (
            sum((x - avg) ** 2 for x in genes["sentence_length_pattern"]) / len(genes["sentence_length_pattern"])
        ) ** 0.5
        genes["avg_sentence_length"] = avg
        genes["sentence_length_std"] = std
    else:
        genes["avg_sentence_length"] = 0
        genes["sentence_length_std"] = 0
    return genes


def extract_text(file_bytes: bytes, filename: str) -> str:
    try:
        lower_name = filename.lower()
        if lower_name.endswith(".pdf"):
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            return "\n\n".join(page.get_text() for page in doc)
        if lower_name.endswith(".docx"):
            doc = Document(BytesIO(file_bytes))
            return "\n\n".join(p.text for p in doc.paragraphs if p.text.strip())
        if lower_name.endswith((".txt", ".md")):
            return file_bytes.decode("utf-8", errors="ignore")
    except Exception as e:
        return f"解析失败: {e}"
    return ""


def analyze_reference_paper(file_bytes: bytes, filename: str) -> str:
    try:
        if filename.lower().endswith(".docx"):
            doc = Document(BytesIO(file_bytes))
            paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
            text = "\n".join(paragraphs[:80])
        else:
            text = extract_text(file_bytes, filename)
            if not text or text.startswith("解析"):
                return f"分析失败: 无法解析 {filename}"
            paragraphs = [p.strip() for p in text.split("\n\n") if p.strip()]
            text = "\n".join(paragraphs[:80])

        lang = detect_language(text)
        genes = extract_language_genes(text)
        return f"""
## 🧬 标杆文献语言基因提取: {filename}
- **段落数**: {len(paragraphs)}
- **语言**: {"中文" if lang == "zh" else "英文"}
- **平均句长**: {genes.get('avg_sentence_length', 0):.1f}
- **句长波动**: {genes.get('sentence_length_std', 0):.1f}
- **连接词偏好**: {', '.join([f'{k}({v})' for k, v in genes['connecting_words'].most_common(5)]) if genes['connecting_words'] else '无明显偏好'}
- **学术短语**: {', '.join([f'{k}({v})' for k, v in genes['academic_phrases'].most_common(5)]) if genes['academic_phrases'] else '无常用短语'}
- **模仿策略**: 优先模仿句长、连接词、语态和复杂度节奏。
""".strip()
    except Exception as e:
        return f"分析失败: {e}"


def create_strict_prompt(function: str, input_text: str, section: str, domain: str, reference_styles: List[str], lang: str) -> str:
    skill_groups = append_skill_preview()
    lang_lock = {
        "zh": "所有输出默认保持中文；若功能为翻译，则必须输出英文。",
        "en": "All output should remain English by default; if translation mode is selected, output must be Chinese.",
        "unknown": "保持输入语言语义一致。",
    }
    base = f"""
# Xueyan OS v4.0
- 功能模式: {function}
- 论文板块: {section}
- 研究领域: {domain}
- 输入语言: {"中文" if lang == "zh" else "英文" if lang == "en" else "未知"}
- 本地规则包: humanizer={skill_groups['humanizer']}, ml-paper={skill_groups['ml-paper']}, docx={skill_groups['docx']}, thesis-formatting={skill_groups['thesis-formatting']}
- 语言约束: {lang_lock.get(lang, '')}

{HUMANIZER_RULES}
"""

    if function == "📝 中转英翻译":
        base += TRANSLATION_SKILL_RULES + "\n"
    if function in MODIFICATION_FUNCTIONS:
        base += DOCX_SKILL_RULES + "\n"
    if section == "结果":
        base += RESULTS_SKILL_RULES + "\n"
    if section == "结果" and domain == "🧠 机器学习":
        base += ML_CHECKLIST + "\n"
    if reference_styles and function in SHADOW_FUNCTIONS:
        base += "\n## 标杆文献风格参考\n" + "\n".join(reference_styles[:3]) + "\n"

    base += f"""
## 功能规则
{chr(10).join(f'- {rule}' for rule in FUNCTION_MATRIX.get(function, {}).get('rules', []))}

## 当前输入
{input_text}
"""

    if function == "📝 中转英翻译":
        return base + """
请严格执行跨语种精准学术翻译：
1. 输入中文时输出英文；输入英文时输出中文。
2. 严禁同语种复述。
3. 工科术语、LaTeX、化学式、材料名必须原样保留。
4. 不新增事实，不夸大贡献。
5. 只输出翻译后的正文，不附加解释。
"""

    if function == "🤖 去AI味 (Humanizer)":
        return base + """
请只做同语种去AI化重构：
1. 保持原语种。
2. 清理模板腔和机械连接词。
3. 用长短句变化增强真人节奏。
4. 不翻译，不新增论点。
5. 先给修改后文本，再给3-5条修改说明。
"""

    if function == "📋 Redlining修订":
        return base + """
请输出“修改后文本”与“修改说明”。保持原语种，不翻译。
"""

    return base + "\n请严格按功能职责输出结果，避免越界到其他功能。"


def get_client():
    if not CLAUDE_API_KEY:
        st.error("❌ 未检测到 ANTHROPIC_AUTH_TOKEN")
        st.stop()
    return anthropic.Anthropic(
        api_key=CLAUDE_API_KEY,
        base_url=CLAUDE_BASE_URL if CLAUDE_BASE_URL != "https://api.anthropic.com" else None,
    )


@retry(
    stop=stop_after_attempt(3),
    wait=wait_exponential(multiplier=1, min=4, max=60),
    retry=retry_if_exception_type((anthropic.APITimeoutError, anthropic.InternalServerError)),
)
def call_api(prompt: str, timeout: int = 300) -> str:
    client = get_client()
    message = client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=8192,
        temperature=0.2,
        timeout=timeout,
        messages=[{"role": "user", "content": prompt}],
    )
    return "".join(block.text for block in message.content if getattr(block, "type", "") == "text")


def create_docx_with_redlines(original: str, revised: str, metadata: dict) -> bytes:
    doc = Document()
    doc.add_heading(f"学研·修订模式 - {metadata['function']}", 0)
    info = doc.add_paragraph()
    info.add_run(f"板块: {metadata.get('section', '未指定')}\n")
    info.add_run(f"领域: {metadata.get('domain', '未指定')}\n")
    info.add_run(f"时间: {metadata.get('timestamp', '')}\n")
    doc.add_heading("原文", 1)
    doc.add_paragraph(original)
    doc.add_heading("修改后", 1)
    doc.add_paragraph(revised)
    doc.add_heading("修改说明", 2)
    original_words = set(original.lower().split())
    revised_words = set(revised.lower().split())
    added = revised_words - original_words
    removed = original_words - revised_words
    if added:
        p = doc.add_paragraph()
        run = p.add_run("新增词汇: ")
        run.font.color.rgb = RGBColor(0, 128, 0)
        p.add_run(", ".join(list(added)[:20]))
    if removed:
        p = doc.add_paragraph()
        run = p.add_run("删除词汇: ")
        run.font.color.rgb = RGBColor(255, 0, 0)
        p.add_run(", ".join(list(removed)[:20]))
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.read()


def create_docx(content: str, metadata: dict) -> bytes:
    doc = Document()
    doc.add_heading(f"学研·工科科研助手 - {metadata.get('function', '导出')}", 0)
    info = doc.add_paragraph()
    for key in ("section", "domain", "timestamp"):
        if metadata.get(key):
            info.add_run(f"{key}: {metadata[key]}\n")
    doc.add_heading("处理结果", 1)
    doc.add_paragraph(content)
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.read()


def parse_format_guidelines(guideline_text: str, guideline_file_bytes: bytes | None = None, filename: str = "") -> Dict:
    text = guideline_text.strip()
    if guideline_file_bytes and filename:
        extracted = extract_text(guideline_file_bytes, filename)
        if extracted and not extracted.startswith("解析"):
            text = f"{text}\n{extracted}".strip()

    rules = {
        "raw_text": text,
        "title_font": "黑体" if "黑体" in text else "",
        "title_size": "三号" if "三号" in text else "",
        "body_font": "宋体" if "宋体" in text else "",
        "body_size": "小四" if "小四" in text else "",
        "line_spacing": 20 if "20磅" in text else None,
        "header_required": any(token in text for token in ["页眉", "校名"]),
        "reference_gbt": "GB/T 7714" in text,
    }
    summary = []
    if rules["title_font"] or rules["title_size"]:
        summary.append(f"标题规则：{rules['title_font'] or '未指定'} {rules['title_size'] or ''}".strip())
    if rules["body_font"] or rules["body_size"]:
        summary.append(f"正文规则：{rules['body_font'] or '未指定'} {rules['body_size'] or ''}".strip())
    if rules["line_spacing"]:
        summary.append(f"行距：{rules['line_spacing']} 磅")
    if rules["header_required"]:
        summary.append("需要检查页眉页脚")
    if rules["reference_gbt"]:
        summary.append("需要检查参考文献格式/悬挂缩进")
    rules["summary"] = "；".join(summary) if summary else "已读取格式指南，但未识别出明确格式条目。"
    return rules


def _paragraph_text(paragraph) -> str:
    return paragraph.text.strip()


def _is_heading(paragraph) -> bool:
    style_name = paragraph.style.name if paragraph.style else ""
    text = _paragraph_text(paragraph)
    return style_name.startswith("Heading") or bool(re.match(r"^(第[一二三四五六七八九十]+[章节]|[一二三四五六七八九十]+、|\d+\.)", text))


def audit_docx_format(docx_bytes: bytes, guideline_rules: Dict) -> Dict:
    doc = Document(BytesIO(docx_bytes))
    issues = []
    stats = {
        "sections": len(doc.sections),
        "paragraphs": len(doc.paragraphs),
        "headings": 0,
        "body_paragraphs": 0,
    }

    for idx, paragraph in enumerate(doc.paragraphs):
        text = _paragraph_text(paragraph)
        if not text:
            continue
        if _is_heading(paragraph):
            stats["headings"] += 1
            run = paragraph.runs[0] if paragraph.runs else None
            if guideline_rules.get("title_font") and run and run.font.name and guideline_rules["title_font"] not in run.font.name:
                issues.append({
                    "id": f"heading_font_{idx}",
                    "type": "heading_font",
                    "label": "对齐各级标题字体",
                    "detail": f"第 {idx + 1} 段标题字体疑似不是 {guideline_rules['title_font']}。",
                })
        else:
            stats["body_paragraphs"] += 1
            fmt = paragraph.paragraph_format
            if guideline_rules.get("line_spacing") and fmt.line_spacing and abs(float(fmt.line_spacing.pt) - guideline_rules["line_spacing"]) > 0.5:
                issues.append({
                    "id": f"body_spacing_{idx}",
                    "type": "body_spacing",
                    "label": "统一正文行距",
                    "detail": f"第 {idx + 1} 段行距不是 {guideline_rules['line_spacing']} 磅。",
                })

    if guideline_rules.get("header_required"):
        for sec_index, section in enumerate(doc.sections):
            header_text = "".join(p.text for p in section.header.paragraphs).strip()
            if not header_text:
                issues.append({
                    "id": f"header_{sec_index}",
                    "type": "header",
                    "label": "修正全局页眉页脚",
                    "detail": f"第 {sec_index + 1} 个 section 缺少页眉内容。",
                })

    if guideline_rules.get("reference_gbt"):
        ref_hits = [p for p in doc.paragraphs if any(token in p.text for token in ["参考文献", "References"])]
        if not ref_hits:
            issues.append({
                "id": "refs_missing",
                "type": "references",
                "label": "检查参考文献格式",
                "detail": "未检测到明确的参考文献标题，无法确认 GB/T 7714 对齐情况。",
            })
        else:
            issues.append({
                "id": "refs_hanging",
                "type": "references",
                "label": "统一参考文献悬挂缩进",
                "detail": "检测到参考文献区，建议统一悬挂缩进与段间距。",
            })

    unique_options = []
    seen = set()
    for issue in issues:
        if issue["label"] not in seen:
            unique_options.append(issue)
            seen.add(issue["label"])

    return {
        "issues": issues,
        "fix_options": unique_options,
        "stats": stats,
    }


def _set_run_font(run, font_name: str | None = None, size_pt: float | None = None, bold: bool | None = None):
    if font_name:
        run.font.name = font_name
        rpr = run._element.get_or_add_rPr()
        rfonts = rpr.rFonts or OxmlElement("w:rFonts")
        rfonts.set(qn("w:eastAsia"), font_name)
        rfonts.set(qn("w:ascii"), font_name)
        rfonts.set(qn("w:hAnsi"), font_name)
        if rpr.rFonts is None:
            rpr.append(rfonts)
    if size_pt:
        run.font.size = Pt(size_pt)
    if bold is not None:
        run.font.bold = bold


def apply_docx_fixes(docx_bytes: bytes, guideline_rules: Dict, selected_fixes: List[str]) -> bytes:
    doc = Document(BytesIO(docx_bytes))

    if "对齐各级标题字体" in selected_fixes:
        for paragraph in doc.paragraphs:
            if _is_heading(paragraph):
                for run in paragraph.runs:
                    _set_run_font(run, guideline_rules.get("title_font") or "黑体", 16, True)
                paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

    if "统一正文行距" in selected_fixes:
        for paragraph in doc.paragraphs:
            if _paragraph_text(paragraph) and not _is_heading(paragraph):
                fmt = paragraph.paragraph_format
                fmt.line_spacing_rule = WD_LINE_SPACING.EXACTLY
                fmt.line_spacing = Pt(guideline_rules.get("line_spacing") or 20)
                fmt.first_line_indent = Cm(0.74)
                for run in paragraph.runs:
                    _set_run_font(run, guideline_rules.get("body_font") or "宋体", 12)

    if "修正全局页眉页脚" in selected_fixes:
        for section in doc.sections:
            if section.start_type in (WD_SECTION.NEW_PAGE, WD_SECTION.CONTINUOUS):
                header = section.header
                if not header.paragraphs:
                    header.add_paragraph()
                header.paragraphs[0].text = "学研 · 确定性格式对齐"
                header.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in header.paragraphs[0].runs:
                    _set_run_font(run, guideline_rules.get("body_font") or "宋体", 10)

    if "统一参考文献悬挂缩进" in selected_fixes:
        in_refs = False
        for paragraph in doc.paragraphs:
            text = _paragraph_text(paragraph)
            if any(token in text for token in ["参考文献", "References"]):
                in_refs = True
                continue
            if in_refs and text:
                fmt = paragraph.paragraph_format
                fmt.left_indent = Cm(0.74)
                fmt.first_line_indent = Cm(-0.74)
                fmt.line_spacing_rule = WD_LINE_SPACING.EXACTLY
                fmt.line_spacing = Pt(guideline_rules.get("line_spacing") or 20)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.read()


def create_format_audit_summary(audit_report: Dict) -> str:
    stats = audit_report.get("stats", {})
    issues = audit_report.get("issues", [])
    return (
        f"共检测 {stats.get('sections', 0)} 个 section，"
        f"{stats.get('paragraphs', 0)} 个段落，"
        f"发现 {len(issues)} 条格式问题。"
    )


# ── UI 样式 ──────────────────────────────────────────────────────────────────
st.markdown(
    """
<style>
    .stApp { background: linear-gradient(135deg, #f8fafc 0%, #e2e8f0 100%); }
    .main-header {
        text-align: center; padding: 2rem 0;
        background: linear-gradient(135deg, #1e3a8a 0%, #3b82f6 100%);
        color: white; border-radius: 12px; margin-bottom: 1.5rem;
        box-shadow: 0 10px 25px rgba(30, 58, 138, 0.2);
    }
    .comparison-box, .engine-box {
        background: white;
        padding: 1rem 1rem 1.25rem 1rem;
        border-radius: 12px;
        border: 1px solid #dbe4f0;
        min-height: 520px;
        box-shadow: 0 6px 18px rgba(15, 23, 42, 0.05);
    }
    .result-toolbar {
        margin: 0.75rem 0 0.25rem 0;
        padding: 0.75rem;
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
    }
</style>
""",
    unsafe_allow_html=True,
)


def render_header() -> None:
    st.markdown(
        """
<div class="main-header">
    <h1>🧪 学研·工科科研助手 v4.0</h1>
    <p style="margin: 0.5rem 0 0 0; opacity: 0.9;">双引擎架构 · 15+功能矩阵 · 影子写作 · 确定性格式对齐</p>
</div>
""",
        unsafe_allow_html=True,
    )
    col1, col2, col3 = st.columns(3)
    with col1:
        st.success("✅ API 就绪" if CLAUDE_API_KEY else "❌ API 未配置")
    with col2:
        st.info(f"📚 {len(append_skill_preview())} 组规则包")
    with col3:
        st.success("⏱️ 300s 超时 | 3次重试")


def render_result_actions(section_name: str, current_input: str, previous_output: str, function: str, domain: str) -> None:
    st.markdown("<div class='result-toolbar'><strong>快捷操作</strong></div>", unsafe_allow_html=True)
    col1, col2 = st.columns([1, 1])
    with col1:
        st.download_button(
            "📋 下载结果.txt",
            previous_output.encode("utf-8"),
            file_name=f"xueyan_{section_name}_result.txt",
            mime="text/plain",
            key=f"download_preview_{section_name}_{function}",
        )
    with col2:
        if function in MODIFICATION_FUNCTIONS:
            redline = create_docx_with_redlines(current_input, previous_output, {
                "function": function,
                "section": section_name,
                "domain": domain,
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            })
            st.download_button(
                "📋 Redlining 修订",
                redline,
                file_name=f"xueyan_redline_{section_name}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key=f"redline_preview_{section_name}_{function}",
            )
        else:
            st.caption("💡 修改类功能执行后可导出 Redlining")


def handle_writing_process(section_name: str, function: str, domain: str, reference_styles: List[str]) -> None:
    input_key = writing_input_key(section_name)
    output_key = writing_output_key(section_name)
    note_key = writing_note_key(section_name)
    current_input = st.session_state.get(input_key, "")
    if not current_input.strip():
        st.warning("请先输入文本。")
        return

    input_lang = detect_language(current_input)
    with st.spinner(f"🔄 处理中 [{function}] | 板块: {section_name} | 领域: {domain}..."):
        protected_input, term_mapping = protect_hard_terms(current_input, domain)
        full_prompt = create_strict_prompt(function, protected_input, section_name, domain, reference_styles, input_lang)
        start_time = time.time()
        response = call_api(full_prompt, timeout=300)
        elapsed = time.time() - start_time

        if function == "📝 中转英翻译":
            output_lang = detect_language(response)
            if output_lang == input_lang:
                st.error("❌ 翻译校验失败：输出语言与输入相同，请重试")
                return

        final_output = restore_hard_terms(response, term_mapping)
        st.session_state[output_key] = final_output
        if function == "🤖 去AI味 (Humanizer)":
            note_text = f"✅ 已针对【{section_name}】完成【{function}】，保持原语种，仅重构语序与节奏。"
        elif function == "📝 中转英翻译":
            note_text = f"✅ 已针对【{section_name}】完成【{function}】，已执行跨语种精准学术翻译并保留术语。"
        else:
            note_text = f"✅ 已针对【{section_name}】完成【{function}】，严格按单一职责执行。"
        st.session_state[note_key] = note_text
        st.session_state.writing_history.insert(0, {
            "id": len(st.session_state.writing_history) + 1,
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "function": function,
            "section": section_name,
            "domain": domain,
            "input_lang": input_lang,
            "input": current_input,
            "output": final_output,
            "elapsed": f"{elapsed:.1f}s",
        })
        st.session_state.writing_history = st.session_state.writing_history[:100]


def render_writing_section(section_name: str, function: str, domain: str, reference_styles: List[str]) -> None:
    input_key = writing_input_key(section_name)
    output_key = writing_output_key(section_name)
    note_key = writing_note_key(section_name)
    current_input = st.session_state.get(input_key, "")
    previous_output = st.session_state.get(output_key, "")
    previous_note = st.session_state.get(note_key, "")

    col_left, col_right = st.columns([1, 1], gap="large")
    with col_left:
        st.markdown("<div class='comparison-box'>", unsafe_allow_html=True)
        st.markdown(f"### 📝 原文输入 [{section_name}]")
        st.text_area(
            "原文输入",
            value=current_input,
            key=input_key,
            height=400,
            label_visibility="collapsed",
        )
        current_input = st.session_state.get(input_key, "")
        st.caption(f"📊 {len(current_input)} 字符 | 语言: {'中文' if detect_language(current_input) == 'zh' else '英文'}")
        btn1, btn2, btn3 = st.columns([3, 1, 1])
        with btn1:
            if st.button(f"✨ 执行 {function}", key=f"process_{section_name}", type="primary", use_container_width=True):
                handle_writing_process(section_name, function, domain, reference_styles)
                st.rerun()
        with btn2:
            if st.button("🗑️", key=f"clear_{section_name}"):
                st.session_state[input_key] = ""
                st.session_state[output_key] = ""
                st.session_state[note_key] = ""
                st.rerun()
        with btn3:
            if st.button("🔄", key=f"rerun_{section_name}"):
                st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

    with col_right:
        st.markdown("<div class='comparison-box'>", unsafe_allow_html=True)
        st.markdown(f"### 👁️ 处理结果 [{section_name}]")
        if previous_output:
            render_result_actions(section_name, st.session_state.get(input_key, ""), previous_output, function, domain)
            st.text_area(
                "结果文本",
                value=previous_output,
                key=f"result_preview_{section_name}",
                height=340,
                disabled=True,
                label_visibility="collapsed",
            )
            if previous_note:
                st.caption(previous_note)
            meta = {
                "function": function,
                "section": section_name,
                "domain": domain,
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
            export1, export2 = st.columns([1, 1])
            with export1:
                st.download_button(
                    "📥 导出 Word",
                    create_docx(previous_output, meta),
                    file_name=f"xueyan_{section_name}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key=f"export_docx_{section_name}",
                )
            with export2:
                if function in MODIFICATION_FUNCTIONS:
                    st.download_button(
                        "📋 导出 Redlining",
                        create_docx_with_redlines(st.session_state.get(input_key, ""), previous_output, meta),
                        file_name=f"xueyan_redline_{section_name}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"export_redline_{section_name}",
                    )
        else:
            st.info("处理结果将显示在这里。")
        st.markdown("</div>", unsafe_allow_html=True)


def render_writing_engine() -> None:
    init_writing_state()
    with st.sidebar:
        st.markdown("---")
        st.subheader("🎯 写作引擎")
        function = st.selectbox("选择功能", list(FUNCTION_MATRIX.keys()), key="writing_function")
        func_info = FUNCTION_MATRIX.get(function, {})
        st.caption(func_info.get("description", ""))
        for rule in func_info.get("rules", []):
            st.caption(rule)

        domain = st.selectbox("研究领域", list(DOMAINS.keys()), key="writing_domain")

        st.markdown("---")
        st.subheader("📚 影子合著者：标杆文献")
        reference_files = st.file_uploader(
            "上传标杆文献 (1-3篇)",
            type=["pdf", "docx"],
            accept_multiple_files=True,
            key="writing_reference_files",
        )
        reference_styles: List[str] = []
        if reference_files:
            for ref_file in reference_files:
                if ref_file.name not in st.session_state.writing_reference_docs:
                    ref_bytes = ref_file.read()
                    text = extract_text(ref_bytes, ref_file.name)
                    if text and not text.startswith("解析"):
                        st.session_state.writing_reference_docs[ref_file.name] = analyze_reference_paper(ref_bytes, ref_file.name)
                if ref_file.name in st.session_state.writing_reference_docs:
                    reference_styles.append(st.session_state.writing_reference_docs[ref_file.name])
        if st.session_state.writing_reference_docs:
            st.success(f"✅ 已加载 {len(st.session_state.writing_reference_docs)} 篇标杆文献")

        st.markdown("---")
        st.subheader("📁 普通文件上传")
        uploaded_file = st.file_uploader("上传待处理文件", type=["pdf", "docx"], key="writing_uploaded_file")
        if uploaded_file:
            uploaded_bytes = uploaded_file.read()
            extracted = extract_text(uploaded_bytes, uploaded_file.name)
            if extracted and not extracted.startswith("解析"):
                st.success(f"✅ 提取 {len(extracted)} 字符")
                if st.button("📥 填入当前板块", use_container_width=True, key="fill_current_section"):
                    active = st.session_state.writing_active_section
                    st.session_state[writing_input_key(active)] = extracted[:10000]
                    st.rerun()

        st.markdown("---")
        preview = append_skill_preview()
        st.caption(f"规则包：{', '.join(f'{k}={v}' for k, v in preview.items())}")

    st.markdown("---")
    tabs = st.tabs(SECTION_NAMES)
    for section_name, tab in zip(SECTION_NAMES, tabs):
        with tab:
            st.session_state.writing_active_section = section_name
            render_writing_section(section_name, function, domain, reference_styles)

    st.markdown("---")
    st.subheader("⏰ 版本时光机")
    filter_section = st.selectbox("筛选板块", ["全部"] + list(SECTIONS.keys()), key="history_section_filter")
    filter_function = st.selectbox("筛选功能", ["全部"] + list(FUNCTION_MATRIX.keys()), key="history_function_filter")
    filtered_history = st.session_state.writing_history
    if filter_section != "全部":
        filtered_history = [h for h in filtered_history if h["section"] == filter_section]
    if filter_function != "全部":
        filtered_history = [h for h in filtered_history if h["function"] == filter_function]

    if filtered_history:
        for entry in filtered_history[:20]:
            with st.expander(f"#{entry['id']} · {entry['function']} · {entry['section']} · {entry['timestamp']}"):
                c1, c2, c3 = st.columns([3, 1, 1])
                with c1:
                    st.caption(f"领域: {entry['domain']} | 语言: {'中文' if entry['input_lang'] == 'zh' else '英文'} | 耗时: {entry['elapsed']}")
                with c2:
                    if st.button("📥 恢复", key=f"restore_{entry['id']}"):
                        st.session_state[writing_input_key(entry['section'])] = entry['input']
                        st.session_state[writing_output_key(entry['section'])] = entry['output']
                        st.session_state[writing_note_key(entry['section'])] = f"已恢复 {entry['timestamp']} 的结果"
                        st.rerun()
                with c3:
                    if st.button("🗑️", key=f"delete_{entry['id']}"):
                        st.session_state.writing_history = [h for h in st.session_state.writing_history if h['id'] != entry['id']]
                        st.rerun()
                left, right = st.columns(2)
                with left:
                    st.markdown("**原始输入**")
                    st.text_area("", entry["input"], height=150, key=f"orig_{entry['id']}", disabled=True)
                with right:
                    st.markdown("**处理输出**")
                    st.text_area("", entry["output"], height=150, key=f"out_{entry['id']}", disabled=True)
    else:
        st.info("📭 暂无符合条件的记录")


def render_formatting_engine() -> None:
    with st.sidebar:
        st.markdown("---")
        st.subheader("📏 排版引擎")
        st.caption("先审查，后修改。仅做样式级修复，不改正文逻辑与公式内容。")
        st.caption("设计参考：thesis-skills 工作流思想 + python-docx 确定性修复。")

    st.markdown("---")
    left, middle, right = st.columns([1, 1, 1.2], gap="large")

    with left:
        st.markdown("<div class='engine-box'>", unsafe_allow_html=True)
        st.markdown("### 📐 格式指南要求")
        guideline_text = st.text_area(
            "格式指南",
            value=st.session_state.format_guideline_text,
            height=360,
            key="format_guideline_text",
            placeholder="例如：一级标题黑体三号；正文宋体小四；行距20磅；页眉含校名；参考文献符合 GB/T 7714",
        )
        guideline_file = st.file_uploader("上传格式指南", type=["txt", "md", "docx"], key="guideline_file")
        if st.button("🔍 解析指南", use_container_width=True, key="parse_guideline"):
            file_bytes = guideline_file.read() if guideline_file else None
            filename = guideline_file.name if guideline_file else ""
            rules = parse_format_guidelines(guideline_text, file_bytes, filename)
            st.session_state.format_guideline_summary = rules["summary"]
            st.session_state.format_guideline_rules = rules
            st.rerun()
        if st.session_state.get("format_guideline_summary"):
            st.success(st.session_state.format_guideline_summary)
        st.markdown("</div>", unsafe_allow_html=True)

    with middle:
        st.markdown("<div class='engine-box'>", unsafe_allow_html=True)
        st.markdown("### 📄 待改 Word 上传")
        target_doc = st.file_uploader("上传待改 Word", type=["docx"], key="target_doc")
        if target_doc:
            st.info(f"已载入：{target_doc.name}")
            extracted = extract_text(target_doc.read(), target_doc.name)
            if extracted and not extracted.startswith("解析"):
                st.text_area("正文预览", extracted[:3000], height=360, disabled=True, key="target_doc_preview")
        st.caption("支持 .docx；当前版本优先修复标题、正文、页眉页脚、参考文献缩进等确定性样式问题。")
        st.markdown("</div>", unsafe_allow_html=True)

    with right:
        st.markdown("<div class='engine-box'>", unsafe_allow_html=True)
        st.markdown("### 🔍 审查与修复区")
        target_doc_for_audit = st.session_state.get("target_doc")
        if st.button("🩺 开始审查", use_container_width=True, key="start_audit"):
            if not guideline_text.strip() and not st.session_state.get("format_guideline_summary"):
                st.warning("请先填写或解析格式指南。")
            elif not target_doc:
                st.warning("请先上传待改 Word。")
            else:
                target_bytes = target_doc.getvalue()
                rules = st.session_state.get("format_guideline_rules") or parse_format_guidelines(guideline_text)
                report = audit_docx_format(target_bytes, rules)
                st.session_state.format_audit_report = report
                st.session_state.format_fix_options = report.get("fix_options", [])
                st.session_state.format_last_filename = target_doc.name
                st.rerun()

        audit_report = st.session_state.get("format_audit_report", {})
        if audit_report:
            st.success(create_format_audit_summary(audit_report))
            for issue in audit_report.get("issues", []):
                st.markdown(f"- {issue['detail']}")
            labels = [item["label"] for item in audit_report.get("fix_options", [])]
            selected = []
            if labels:
                st.markdown("---")
                st.markdown("**可选修复项**")
                for label in labels:
                    if st.checkbox(label, value=True, key=f"fix_{label}"):
                        selected.append(label)
                if st.button("🛠️ 应用所选修复", use_container_width=True, key="apply_fixes"):
                    if not target_doc:
                        st.warning("请重新上传待改 Word。")
                    else:
                        rules = st.session_state.get("format_guideline_rules") or parse_format_guidelines(guideline_text)
                        output_bytes = apply_docx_fixes(target_doc.getvalue(), rules, selected)
                        st.session_state.format_output_docx_bytes = output_bytes
                        st.rerun()

        if st.session_state.get("format_output_docx_bytes"):
            st.markdown("---")
            st.download_button(
                "📥 下载修复后 Word",
                st.session_state.format_output_docx_bytes,
                file_name=f"fixed_{st.session_state.get('format_last_filename') or 'document.docx'}",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="download_fixed_docx",
            )
            st.caption("仅对勾选项做样式级修复；正文逻辑、翻译与公式语义不会被改写。")
        else:
            st.info("先解析指南，再审查文档，最后勾选修复项导出。")
        st.markdown("</div>", unsafe_allow_html=True)


def main() -> None:
    render_header()
    with st.sidebar:
        st.markdown("---")
        st.session_state.engine_mode = st.radio(
            "核心引擎切换",
            [WRITING_PAGE, FORMATTING_PAGE],
            index=0 if st.session_state.engine_mode == WRITING_PAGE else 1,
            key="engine_mode",
        )

    if st.session_state.engine_mode == WRITING_PAGE:
        render_writing_engine()
    else:
        render_formatting_engine()

    st.markdown(
        """
---
<div style="text-align: center; color: #64748b; font-size: 0.8rem; padding: 2rem 0;">
    <p><strong>🧪 学研·工科科研助手 v4.0</strong></p>
    <p>双引擎架构 · 写作协作 OS · 确定性格式对齐 · Hugging Face 就绪</p>
</div>
""",
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
