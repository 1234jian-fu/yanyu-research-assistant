"""
学研·工科科研助手 v4.0 (Xueyan OS)
双引擎架构：论文内容写作 + 确定性格式对齐
"""

import os
import re
import time
import json
from collections import Counter
from copy import deepcopy
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Dict, List, Tuple

import anthropic
import fitz  # pymupdf
import streamlit as st
import streamlit.components.v1 as components
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

WRITING_PAGE = "✍️ 论文写作"
FORMATTING_PAGE = "📐 格式对齐"

CONFIG = {
    "sections": {
        "摘要": {"focus": "开门见山，数据支撑", "max_words": 250},
        "引言": {"focus": "背景转折，研究空白", "max_words": 800},
        "方法": {"focus": "流程清晰，参数精确", "max_words": 1500},
        "结果": {"focus": "数据客观，图表引用", "max_words": 1200},
        "讨论": {"focus": "深度解读，文献对比", "max_words": 1000},
        "结论": {"focus": "总结贡献，展望未来", "max_words": 300},
    },
    "functions": {
        "📝 中转英翻译": {
            "description": "执行中英双向精准学术翻译，严格跨语种输出",
            "rules": ["✅ 必须跨语种输出", "❌ 禁止同语输出", "✅ 工科论文表达"],
        },
        "✨ 表达润色": {
            "description": "提升学术地道性，同语言优化",
            "rules": ["❌ 禁止翻译", "❌ 禁止改变原意", "✅ 仅限同语言"],
        },
        "🔍 逻辑检查": {
            "description": "检查因果链条和衔接",
            "rules": ["❌ 仅查逻辑", "❌ 不修改文本", "✅ 指出问题"],
        },
        "🤖 去AI味 (Humanizer)": {
            "description": "消除AI痕迹，模仿真人语序",
            "rules": ["❌ 禁止翻译", "❌ 禁止增加论点", "✅ CN优化CN/EN优化EN"],
        },
        "👨‍⚖️ Reviewer视角": {
            "description": "模拟审稿人视角审视",
            "rules": ["✅ 提供改进建议", "✅ 指出薄弱环节"],
        },
        "💡 研究想法构思": {
            "description": "从零构思研究方向",
            "rules": ["✅ 基于研究领域", "✅ 提供创新点"],
        },
        "🧠 ML论文写作": {
            "description": "NeurIPS/ICML级别写作",
            "rules": ["✅ 引用格式验证", "✅ 图表描述规范"],
        },
        "📊 概念图设计": {
            "description": "生成论文概念图设计",
            "rules": ["✅ 设计哲学", "✅ 绘图提示词"],
        },
        "📄 节节头脑风暴": {
            "description": "按章节进行头脑风暴",
            "rules": ["✅ 针对当前板块", "✅ 提供思路"],
        },
        "✍️ 逐段起草": {
            "description": "将大纲扩充为正式段落",
            "rules": ["✅ 模仿标杆文献", "✅ 保持学术规范"],
        },
        "✍️ 影子写作": {
            "description": "优先模仿标杆文献叙事节奏进行起草",
            "rules": ["✅ 优先模仿标杆文献", "✅ 保持当前语种", "✅ 对齐当前板块"],
        },
        "🎯 精修模式": {
            "description": "深度精修，达到顶刊水准",
            "rules": ["✅ 模仿标杆文献", "✅ 顶刊标准"],
        },
        "📝 引用验证": {
            "description": "检查引用格式和完整性",
            "rules": ["✅ 格式验证", "✅ 完整性检查"],
        },
        "🎨 图表规范检查": {
            "description": "检查图表描述规范性",
            "rules": ["✅ 色盲友好", "✅ 规范验证"],
        },
        "📋 Redlining修订": {
            "description": "带修订痕迹的修改建议",
            "rules": ["✅ 显示修改痕迹", "✅ 对比视图"],
        },
        "🔄 版本对比": {
            "description": "对比不同版本差异",
            "rules": ["✅ 高亮差异", "✅ 修改说明"],
        },
    },
    "function_groups": {
        "✨ 核心润色": ["✨ 表达润色", "🤖 去AI味 (Humanizer)", "🎯 精修模式", "📋 Redlining修订", "🔍 逻辑检查", "👨‍⚖️ Reviewer视角"],
        "🔄 翻译转换": ["📝 中转英翻译", "📝 引用验证", "🎨 图表规范检查", "🔄 版本对比"],
        "📝 影子写作": ["✍️ 逐段起草", "✍️ 影子写作", "📄 节节头脑风暴", "💡 研究想法构思", "🧠 ML论文写作", "📊 概念图设计"],
    },
    "function_nav": [
        {"label": "✨ 表达润色 (核心)", "value": "✨ 表达润色"},
        {"label": "🤖 去AI味 (Humanizer)", "value": "🤖 去AI味 (Humanizer)"},
        {"label": "🔍 逻辑检查", "value": "🔍 逻辑检查"},
        {"label": "📋 Redlining修订", "value": "📋 Redlining修订"},
        {"label": "👨‍⚖️ Reviewer视角", "value": "👨‍⚖️ Reviewer视角"},
        {"label": "🎯 精修模式", "value": "🎯 精修模式"},
        {"label": "📝 影子写作", "value": "✍️ 影子写作"},
        {"label": "✍️ 逐段起草", "value": "✍️ 逐段起草"},
        {"label": "📄 节节头脑风暴", "value": "📄 节节头脑风暴"},
        {"label": "💡 研究想法构思", "value": "💡 研究想法构思"},
        {"label": "🌍 中转英学术翻译", "value": "📝 中转英翻译"},
        {"label": "📊 图表规范检查", "value": "🎨 图表规范检查"},
        {"label": "📝 引用规范检查", "value": "📝 引用验证"},
        {"label": "🔄 版本对比", "value": "🔄 版本对比"},
        {"label": "🧠 ML论文写作", "value": "🧠 ML论文写作"},
        {"label": "📊 概念图设计", "value": "📊 概念图设计"},
    ],
    "domains": {
        "🔋 能源电池 (Battery)": {
            "focus": "电化学储能与离子传输",
            "brain_prompt": "聚焦电化学机理、离子迁移、倍率性能、循环稳定性与界面演化，保持电池论文语气克制、数据导向。",
            "hard_lock_terms": ["NCM523", "NCM622", "XRD", "SEM", "capacity retention", "intercalation", "solid electrolyte interphase", "SEI"],
            "hard_lock_regex": [r"\$Li\^\+\$", r"\$Na\^\+\$", r"lithium[- ]ion", r"sodium[- ]ion", r"coulombic efficiency", r"diffusion coefficient"],
        },
        "🔬 先进材料 (Materials)": {
            "focus": "微结构演化与晶体缺陷调控",
            "brain_prompt": "聚焦晶格畸变、晶界、位错、相变、微结构表征与材料性能之间的因果链，避免空泛描述。",
            "hard_lock_terms": ["grain boundary", "dislocation", "microstructure", "XRD", "SEM", "TEM", "EBSD"],
            "hard_lock_regex": [r"晶格畸变", r"应力应变", r"phase transformation", r"lattice distortion", r"fracture toughness", r"yield strength"],
        },
        "🏗️ 工业循环 (Ecology)": {
            "focus": "固废资源化与工业副产物循环利用",
            "brain_prompt": "聚焦红泥、粉煤灰、矿渣等工业固废的资源化路径、反应机理、强度/活性/去除效率表现，强调过程闭环。",
            "hard_lock_terms": ["red mud", "fly ash", "GGBS", "C-S-H", "compressive strength", "pozzolanic", "removal efficiency"],
            "hard_lock_regex": [r"红泥", r"粉煤灰", r"工业固废", r"资源化", r"水化动力学", r"协同处置"],
        },
        "⚙️ 机械制造": {
            "focus": "制造过程、载荷响应与寿命评价",
            "brain_prompt": "聚焦加工参数、应力分布、疲劳寿命、断裂行为和制造质量之间的关系，保持工程化表达。",
            "hard_lock_terms": ["tensile strength", "fatigue life", "fracture toughness", "surface roughness", "residual stress"],
            "hard_lock_regex": [r"应力应变", r"stress-strain", r"finite element", r"machining", r"wear resistance", r"fracture morphology"],
        },
        "📊 信息智能": {
            "focus": "模型性能、系统架构与智能算法",
            "brain_prompt": "聚焦模型结构、训练稳定性、性能指标、系统架构与工程可复现性，避免营销式措辞。",
            "hard_lock_terms": ["CNN", "Transformer", "F1-score", "API", "microservices", "CI/CD", "DevOps"],
            "hard_lock_regex": [r"precision", r"recall", r"gradient descent", r"overfitting", r"retrieval[- ]augmented generation", r"latency"],
        },
    },
    "global_hard_lock_regex": [
        r"\$\$[\s\S]+?\$\$",
        r"\\\[[\s\S]+?\\\]",
        r"\$[^$\n]+?\$",
        r"\\[a-zA-Z]+(?:\{[^}]*\}|\[[^\]]*\]|[_^]\{[^}]*\})*",
        r"\b[A-Z][a-z]?\d*(?:_[\d-]+|[+\-]?\d*)?\b",
        r"\d+\s*(?:°C|K|MPa|GPa|kPa|Pa|wt%|vol%|at%|mol%|nm|μm|mm|cm|m|km|Hz|kHz|MHz|GHz)",
    ],
    "prompt_rules": {
        "humanizer": """
## Humanizer 核心规则（零篡位）
- 清理 AI 套话：综上所述、总而言之、in conclusion、it is worth noting that。
- 清理 AI 高危词：delve, landscape, leverage, underscore, utilize。
- 保留原语种，禁止翻译。
- 用长短句变化制造呼吸感，但不新增论点。
- 不夸大贡献，不制造不存在的创新性语气。
""",
        "translation": """
## Translation Mastery（精准跨语种）
- 必须执行跨语种转换，严禁同语输出。
- 中译英：使用工科学术表达，补全冠词、时态、被动结构。
- 英译中：去翻译腔，输出自然、凝练、可直接入文的中文。
- 保留 LaTeX、化学式、缩写、材料名、术语格式。
""",
        "results": """
## Results 专项写作协议
- 结果板块优先写趋势、对比、异常点解释、图表引用。
- 遇到实验数据时，优先明确 increase/decrease、higher/lower、plateau、fluctuation 等关系。
- 自动保留 LaTeX、化学式、材料名和单位。
- 对工科实验结果保持客观，避免宣传腔。
""",
        "docx": """
## DOCX / Redlining 协议
- Word 导出需保留板块、功能、领域和时间元信息。
- Redlining 仅用于展示修订痕迹，不改变功能职责边界。
""",
        "ml_checklist": """
## ML Paper Writing Checklist
- 引用格式统一
- 所有方法引用有明确出处
- 基线模型正确引用
- 图表配色、图例、误差线清晰
- 超参数表与实验设置可复现
- 统计显著性标注完整
""",
    },
    "builtin_local_skills": {
        "README": {
            "label": "README",
            "handler": "meta",
            "description": "Xueyan GOOD local skill pack for engineering research writing and deterministic formatting.",
        },
        "humanizer": {"label": "🤖 去AI味 (Humanizer)", "handler": "same_language_rewrite", "description": "同语种去AI味重写，保留原意与术语。"},
        "translation": {"label": "📝 中转英翻译", "handler": "cross_language_translation", "description": "中英双向学术翻译，严格跨语种输出。"},
        "reviewer": {"label": "👨‍⚖️ Reviewer视角", "handler": "diagnostic_review", "description": "Reviewer 视角诊断论证薄弱点与可改进项。"},
        "logic-check": {"label": "🔍 逻辑检查", "handler": "diagnostic_review", "description": "检查因果链、衔接、论证闭环。"},
        "brainstorming": {"label": "📄 节节头脑风暴", "handler": "ideation", "description": "围绕选题、章节与实验设计进行头脑风暴。"},
        "shadow-writing": {"label": "✍️ 影子写作", "handler": "reference_guided_drafting", "description": "模仿标杆文献叙事节奏进行影子写作。"},
        "paragraph-drafting": {"label": "✍️ 逐段起草", "handler": "reference_guided_drafting", "description": "根据提纲或要点扩写为正式学术段落。"},
        "precision-polish": {"label": "🎯 精修模式", "handler": "same_language_rewrite", "description": "顶刊风格精修，保持克制表达。"},
        "citation-check": {"label": "📝 引用验证", "handler": "citation_audit", "description": "检查引用格式、一致性与完整性。"},
        "figure-check": {"label": "🎨 图表规范检查", "handler": "figure_audit", "description": "检查图表标题、配色、图例与描述规范。"},
        "redlining": {"label": "📋 Redlining修订", "handler": "redline_export", "description": "输出带修订痕迹的对照修改建议。"},
        "version-compare": {"label": "🔄 版本对比", "handler": "diff_compare", "description": "比较不同版本文本差异并总结改动。"},
        "ml-paper-writing": {"label": "🧠 ML论文写作", "handler": "ml_writing", "description": "面向 ML 论文的章节写作与清单校验。"},
        "concept-figure": {"label": "📊 概念图设计", "handler": "concept_design", "description": "生成概念图设计说明与绘图提示。"},
        "thesis-formatting": {"label": "📐 格式对齐", "handler": "formatting_audit", "description": "学位论文格式规则抽取、审计与修复建议。"},
        "docx": {"label": "📄 Word导出", "handler": "docx_export", "description": "Word 文档导出、修订痕迹与格式修复支持。"},
    },
    "ui": {
        "format_export_button_label": "📥 导出 Word",
        "format_redline_button_label": "📋 导出 Redlining",
        "primary_button_css": """
<link rel=\"preconnect\" href=\"https://fonts.googleapis.com\">
<link rel=\"preconnect\" href=\"https://fonts.gstatic.com\" crossorigin>
<link href=\"https://fonts.googleapis.com/css2?family=Poppins:wght@600;700&display=swap\" rel=\"stylesheet\">
<style>
    .stApp {
        background:
            radial-gradient(circle at top left, rgba(59, 130, 246, 0.10), transparent 28%),
            linear-gradient(180deg, #f3f6fb 0%, #eef3f9 100%);
    }
    header[data-testid=\"stHeader\"] {
        background: transparent !important;
        border: none !important;
        box-shadow: none !important;
    }
    .stAppHeader {
        background: transparent !important;
        box-shadow: none !important;
    }
    .block-container {
        padding-top: 0.35rem;
        padding-bottom: 1.2rem;
    }
    section[data-testid=\"stSidebar\"] {
        min-width: 20rem;
    }
    section[data-testid=\"stSidebar\"] .block-container {
        padding-top: 0.75rem;
    }
    .xueyan-badge {
        display: flex;
        align-items: center;
        justify-content: center;
        padding: 0.75rem 0.9rem;
        margin: 0 0 0.8rem 0;
        border-radius: 12px;
        background: #E0F2FE;
        color: #1E3A8A;
        font-family: 'Poppins', sans-serif;
        font-weight: 700;
        font-size: 1.05rem;
        letter-spacing: 0.12em;
        text-transform: uppercase;
        box-shadow: 0 6px 18px rgba(30, 58, 138, 0.10);
        border: 1px solid rgba(30, 58, 138, 0.08);
    }
    .sidebar-panel {
        padding: 0.7rem 0.8rem;
        margin: 0.3rem 0 0.75rem 0;
        border-radius: 12px;
        background: rgba(255,255,255,0.72);
        border: 1px solid #dbe7f5;
        box-shadow: 0 6px 18px rgba(15, 23, 42, 0.05);
    }
    section[data-testid=\"stSidebar\"] [data-testid=\"stExpander\"] {
        border: 1px solid #dbe7f5;
        border-radius: 12px;
        background: rgba(255,255,255,0.88);
        overflow: hidden;
        margin-bottom: 0.55rem;
    }
    section[data-testid=\"stSidebar\"] [data-testid=\"stExpander\"] details summary {
        padding: 0.35rem 0.6rem;
    }
    section[data-testid=\"stSidebar\"] .stRadio > div {
        gap: 0.35rem;
    }
    .sidebar-panel-title {
        margin: 0 0 0.2rem 0;
        color: #0f172a;
        font-size: 0.92rem;
        font-weight: 700;
    }
    .sidebar-panel-note {
        margin: 0;
        color: #64748b;
        font-size: 0.78rem;
        line-height: 1.45;
    }
    .stTabs [data-baseweb=\"tab-list\"] {
        gap: 0.35rem;
        margin-bottom: 0.15rem;
    }
    .stTabs [data-baseweb=\"tab\"] {
        height: 2.45rem;
        padding: 0 0.9rem;
    }
    .stTabs [data-baseweb=\"tab-panel\"] {
        padding-top: 0.1rem;
    }
    .main-header {
        text-align: left;
        padding: 0.9rem 1.1rem;
        background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 72%, #2563eb 100%);
        color: white;
        border-radius: 16px;
        margin-bottom: 0.6rem;
        box-shadow: 0 16px 38px rgba(15, 23, 42, 0.16);
        border: 1px solid rgba(255, 255, 255, 0.12);
        position: relative;
        overflow: hidden;
    }
    .main-header::after {
        content: \"\";
        position: absolute;
        inset: 0;
        background: linear-gradient(120deg, rgba(255,255,255,0.10), transparent 42%, transparent 58%, rgba(255,255,255,0.08));
        pointer-events: none;
    }
    .main-header h1 {
        margin: 0;
        font-size: 1.7rem;
        letter-spacing: 0.01em;
    }
    .main-header p {
        margin: 0.3rem 0 0 0;
        opacity: 0.88;
        font-size: 0.92rem;
    }
    .comparison-box, .engine-box, .workbench-card, .dashboard-stat {
        background: rgba(255, 255, 255, 0.94);
        border-radius: 14px;
        border: 1px solid #d8e1ee;
        box-shadow: 0 8px 24px rgba(15, 23, 42, 0.05);
    }
    .comparison-box, .engine-box {
        padding: 0.95rem 1rem 1rem 1rem;
        min-height: 560px;
    }
    .comparison-box h3, .engine-box h3 {
        margin-top: 0;
        margin-bottom: 0.55rem;
    }
    .workbench-card {
        padding: 0.85rem 1rem;
        margin-bottom: 0.75rem;
    }
    .workbench-card.compact {
        padding: 0.75rem 0.9rem;
        margin-bottom: 0.55rem;
    }
    .workbench-title {
        display: flex;
        justify-content: space-between;
        align-items: flex-start;
        gap: 1rem;
        margin-bottom: 0.5rem;
    }
    .workbench-title h3 {
        margin: 0;
        color: #0f172a;
        font-size: 1.05rem;
    }
    .workbench-title p {
        margin: 0.2rem 0 0 0;
        color: #475569;
        font-size: 0.86rem;
    }
    .workbench-chip {
        display: inline-flex;
        align-items: center;
        gap: 0.35rem;
        padding: 0.22rem 0.58rem;
        border-radius: 999px;
        background: #dbeafe;
        color: #1d4ed8;
        font-size: 0.76rem;
        font-weight: 600;
        white-space: nowrap;
    }
    .dashboard-stat {
        padding: 0.78rem 0.9rem;
        min-height: 104px;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
    }
    .dashboard-stat-label {
        font-size: 0.74rem;
        letter-spacing: 0.08em;
        text-transform: uppercase;
        color: #64748b;
        margin-bottom: 0.4rem;
    }
    .dashboard-stat-value {
        font-size: 1rem;
        font-weight: 700;
        color: #0f172a;
        line-height: 1.3;
    }
    .dashboard-stat-meta {
        font-size: 0.8rem;
        color: #475569;
        margin-top: 0.3rem;
    }
    .subtle-kpi-row {
        padding: 0.2rem 0 0.55rem 0;
        color: #475569;
        font-size: 0.84rem;
    }
    .subtle-kpi-row strong {
        color: #0f172a;
    }
    .result-toolbar {
        margin: 0.55rem 0 0.15rem 0;
        padding: 0.68rem 0.8rem;
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
    }
    .sidebar-section-note {
        font-size: 0.8rem;
        color: #64748b;
        margin-bottom: 0.2rem;
    }
    div.stButton > button, div.stDownloadButton > button {
        margin: 0.12rem 0.28rem 0.45rem 0;
        border-radius: 10px;
    }
    div.stDownloadButton > button[kind=\"secondary\"] {
        background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%) !important;
        color: #ffffff !important;
        border: 1px solid #1d4ed8 !important;
        box-shadow: 0 6px 16px rgba(37, 99, 235, 0.22);
        font-weight: 600;
    }
    div.stDownloadButton > button[kind=\"secondary\"]:hover {
        border-color: #1e40af !important;
        background: linear-gradient(135deg, #1d4ed8 0%, #1e40af 100%) !important;
        color: #ffffff !important;
    }
    div[data-testid=\"stTextArea\"] {
        margin-top: 0 !important;
        padding-top: 0 !important;
    }
    div[data-testid=\"stTextArea\"] > label {
        margin-top: 0 !important;
        padding-top: 0 !important;
        min-height: 0 !important;
    }
    .stTextArea textarea {
        font-family: \"Iosevka Term\", \"Sarasa Mono SC\", \"Noto Serif CJK SC\", \"Source Han Serif SC\", serif;
        line-height: 1.72;
        font-size: 0.96rem;
    }
    .stTextArea [data-baseweb=\"textarea\"] {
        border-radius: 12px;
        background: #fcfdff;
        border: 1px solid #dbe7f5;
    }
    .stTextArea [data-baseweb=\"textarea\"]:focus-within {
        border-color: #93c5fd;
        box-shadow: 0 0 0 1px #93c5fd;
    }
    .compact-caption {
        color: #64748b;
        font-size: 0.8rem;
        margin-top: -0.15rem;
        margin-bottom: 0.45rem;
    }
    .format-preview-box {
        min-height: 340px;
    }
    .format-chat-anchor {
        position: sticky;
        bottom: 0.75rem;
        z-index: 20;
        padding-top: 0.75rem;
        background: linear-gradient(180deg, rgba(243,246,251,0) 0%, rgba(243,246,251,0.92) 35%, rgba(243,246,251,1) 100%);
    }
</style>
""",
    },
}

SECTION_NAMES = list(CONFIG["sections"].keys())
FUNCTION_MATRIX = CONFIG["functions"]
FUNCTION_GROUPS = CONFIG["function_groups"]
FUNCTION_NAV = CONFIG["function_nav"]
DOMAIN_PROFILES = CONFIG["domains"]
DOMAIN_ORDER = list(DOMAIN_PROFILES.keys())
SECTIONS = CONFIG["sections"]
GLOBAL_HARD_LOCK_REGEX = CONFIG["global_hard_lock_regex"]
BUILTIN_LOCAL_SKILLS = CONFIG["builtin_local_skills"]
UI_CONFIG = CONFIG["ui"]
MODIFICATION_FUNCTIONS = {"📋 Redlining修订", "✨ 表达润色", "🤖 去AI味 (Humanizer)", "🎯 精修模式"}
SHADOW_FUNCTIONS = {"✍️ 逐段起草", "💡 研究想法构思", "📄 节节头脑风暴", "✍️ 影子写作"}

DEFAULT_STATES = {
    "engine_mode": WRITING_PAGE,
    "writing_active_section": SECTION_NAMES[0],
    "writing_domain": DOMAIN_ORDER[0],
    "writing_function": "✨ 表达润色",
    "writing_history": [],
    "writing_reference_docs": {},
    "formatting_domain": DOMAIN_ORDER[0],
    "format_ruleset": "tsinghua-thesis",
    "format_guideline_text": "",
    "format_guideline_summary": "",
    "format_guideline_rules": {},
    "format_audit_report": {},
    "format_fix_options": [],
    "format_output_docx_bytes": b"",
    "format_last_filename": "",
    "format_micro_tune_request": "",
    "format_target_doc_bytes": b"",
    "format_target_doc_name": "",
    "format_pipeline_summary": "",
    "format_pipeline_steps": [],
    "format_audit_rows": [],
    "format_audited_docx_bytes": b"",
    "format_audited_docx_name": "",
}

for key, default in DEFAULT_STATES.items():
    if key not in st.session_state:
        st.session_state[key] = deepcopy(default)

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


BUILTIN_LOCAL_SKILLS = CONFIG["builtin_local_skills"]
LOCAL_SKILLS: Dict[str, Dict[str, str]] = deepcopy(BUILTIN_LOCAL_SKILLS)
for skill_name, skill_content in load_local_skills().items():
    LOCAL_SKILLS.setdefault(
        skill_name,
        {
            "label": skill_name,
            "handler": "external_reference",
            "description": skill_content[:120].strip() or "外部规则包",
        },
    )


def trim_session_state() -> None:
    """保守清理，避免误删 Streamlit 组件状态导致界面闪烁或重复刷新。"""
    for key, default in DEFAULT_STATES.items():
        st.session_state.setdefault(key, deepcopy(default))


# ── 工具函数 ─────────────────────────────────────────────────────────────────
def writing_input_key(section: str) -> str:
    return f"writing_input_{section}"


def writing_output_key(section: str) -> str:
    return f"writing_output_{section}"


def writing_note_key(section: str) -> str:
    return f"writing_note_{section}"


def writing_term_count_key(section: str) -> str:
    return f"writing_term_count_{section}"


def init_writing_state() -> None:
    for section in SECTION_NAMES:
        st.session_state.setdefault(writing_input_key(section), "")
        st.session_state.setdefault(writing_output_key(section), "")
        st.session_state.setdefault(writing_note_key(section), "")
        st.session_state.setdefault(writing_term_count_key(section), 0)


FORMAT_RULESET_LIBRARY = {
    "tsinghua-thesis": {
        "label": "清华论文规则包",
        "summary": "按 thesis-skills 的 tsinghua-thesis 工作流组织，强调中文学位论文标题、正文行距、页眉与参考文献一致性。",
        "guidance": [
            "标题层级优先检查黑体/字号与层级结构。",
            "正文优先检查宋体小四、20 磅行距、中文标点与中英文混排间距。",
            "参考文献区优先检查 [n] 引用连续性与悬挂缩进。",
        ],
    },
    "university-generic": {
        "label": "通用高校规则包",
        "summary": "适合先做确定性预对齐，再根据校方模板微调。重点覆盖标题、页眉页脚、图表题注与参考文献区。",
        "guidance": [
            "优先保证标题层级与正文样式统一。",
            "统一中文标点、空格、数字单位间距。",
            "补齐图表题注与正文引用的基本连贯性。",
        ],
    },
    "journal-generic": {
        "label": "通用期刊规则包",
        "summary": "更偏向投稿前清稿，强调英文/中英混排、图表编号、参考文献与题注一致性。",
        "guidance": [
            "优先压缩样式漂移，减少投稿前人工回查。",
            "统一 caption 编号与正文 Figure/Table 引用。",
            "保持正文内容不改写，仅做确定性样式修复。",
        ],
    },
}


def get_format_ruleset_profile(ruleset: str) -> Dict:
    return FORMAT_RULESET_LIBRARY.get(ruleset, FORMAT_RULESET_LIBRARY["tsinghua-thesis"])


def append_skill_preview() -> Dict[str, str]:
    readme_text = LOCAL_SKILLS.get("README", {}).get("description", "")
    return {
        "humanizer": "humanizer" if "humanizer" in readme_text.lower() else "内置规则",
        "ml-paper": "20-ml-paper-writing" if "20-ml-paper-writing" in readme_text.lower() else "内置规则",
        "docx": "docx" if "docx" in readme_text.lower() else "内置规则",
        "thesis-formatting": "thesis-skills 设计参考",
    }


def lock_matches(protected: str, mapping: Dict[str, str], counter: int, pattern: str) -> Tuple[str, int]:
    compiled = re.compile(pattern, re.DOTALL)
    while True:
        match = compiled.search(protected)
        if not match:
            break
        token = match.group(0)
        placeholder = f"__LOCK_{counter}__"
        mapping[placeholder] = token
        protected = f"{protected[:match.start()]}{placeholder}{protected[match.end():]}"
        counter += 1
    return protected, counter


def protect_hard_terms(text: str, domain: str) -> Tuple[str, Dict[str, str]]:
    protected = text
    mapping: Dict[str, str] = {}
    counter = 0
    domain_profile = DOMAIN_PROFILES.get(domain, {})

    for pattern in GLOBAL_HARD_LOCK_REGEX:
        protected, counter = lock_matches(protected, mapping, counter, pattern)

    for pattern in domain_profile.get("hard_lock_regex", []):
        protected, counter = lock_matches(protected, mapping, counter, pattern)

    for term in sorted(domain_profile.get("hard_lock_terms", []), key=len, reverse=True):
        while term and term in protected:
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
    domain_profile = DOMAIN_PROFILES.get(domain, {})
    prompt_rules = CONFIG["prompt_rules"]
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
- 领域焦点: {domain_profile.get('focus', '科研表达')}
- 学科大脑: {domain_profile.get('brain_prompt', '保持工科学术表达与术语稳定性。')}
- 输入语言: {"中文" if lang == "zh" else "英文" if lang == "en" else "未知"}
- 本地规则包: humanizer={skill_groups['humanizer']}, ml-paper={skill_groups['ml-paper']}, docx={skill_groups['docx']}, thesis-formatting={skill_groups['thesis-formatting']}
- 语言约束: {lang_lock.get(lang, '')}

{prompt_rules['humanizer']}
"""

    if function == "📝 中转英翻译":
        base += prompt_rules["translation"] + "\n"
    if function in MODIFICATION_FUNCTIONS:
        base += prompt_rules["docx"] + "\n"
        base += """
## 核心锁约束
- 禁止改写任何已锁定 placeholder 对应内容。
- 禁止修改任何 $...$ / $$...$$ / \\[...\\] LaTeX 公式。
- 禁止拆分、翻译或润色化学式、材料缩写、单位和核心术语。
"""
    if section == "结果":
        base += prompt_rules["results"] + "\n"
    if section == "结果" and domain == "📊 信息智能":
        base += prompt_rules["ml_checklist"] + "\n"
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


def build_export_filename(prefix: str, name_hint: str, suffix: str = ".docx") -> str:
    safe_name = re.sub(r"[^\w\u4e00-\u9fff.-]+", "_", (name_hint or "document").strip())
    safe_name = safe_name.strip("._") or "document"
    if not safe_name.lower().endswith(suffix):
        safe_name = f"{safe_name}{suffix}"
    return f"{prefix}_{safe_name}"


def ensure_guideline_rules(guideline_text: str, guideline_file=None) -> Dict:
    file_bytes = guideline_file.getvalue() if guideline_file else None
    filename = guideline_file.name if guideline_file else ""
    rules = parse_format_guidelines(guideline_text, file_bytes, filename)
    st.session_state.format_guideline_summary = rules["summary"]
    st.session_state.format_guideline_rules = rules
    return rules


def build_format_preview_export(target_doc_bytes: bytes, target_doc_name: str, guideline_summary: str, domain: str) -> bytes:
    extracted = extract_text(target_doc_bytes, target_doc_name)
    content = extracted[:6000] if extracted and not extracted.startswith("解析") else "文档预览不可用"
    return create_docx(
        content,
        {
            "function": "确定性格式对齐",
            "section": guideline_summary,
            "domain": domain,
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        },
    )


def parse_format_guidelines(guideline_text: str, guideline_file_bytes: bytes | None = None, filename: str = "") -> Dict:
    text = guideline_text.strip()
    if guideline_file_bytes and filename:
        extracted = extract_text(guideline_file_bytes, filename)
        if extracted and not extracted.startswith("解析"):
            text = f"{text}\n{extracted}".strip()

    text = text or "默认规则：一级标题黑体三号；正文宋体小四；行距20磅；页眉含校名；参考文献符合 GB/T 7714；中文标点统一；数字与单位之间保留半角空格。"
    rules = {
        "raw_text": text,
        "title_font": "黑体" if "黑体" in text else "",
        "title_size": "三号" if "三号" in text else "",
        "body_font": "宋体" if "宋体" in text else "",
        "body_size": "小四" if "小四" in text else "",
        "line_spacing": 20 if "20磅" in text else None,
        "header_required": any(token in text for token in ["页眉", "校名"]),
        "reference_gbt": "GB/T 7714" in text,
        "enforce_cn_punctuation": any(token in text for token in ["中文标点", "全角", "标点统一"]),
        "enforce_unit_spacing": any(token in text for token in ["单位", "半角空格", "数字与单位"]),
        "check_captions": any(token in text for token in ["图表", "题注", "caption", "Figure", "Table"]),
        "check_citation_order": any(token in text for token in ["引用", "[n]", "参考文献"]),
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
    if rules["enforce_cn_punctuation"]:
        summary.append("开启中文标点统一")
    if rules["enforce_unit_spacing"]:
        summary.append("开启数字与单位间距修正")
    if rules["check_captions"]:
        summary.append("检查图表题注与正文引用")
    if rules["check_citation_order"]:
        summary.append("检查 [n] 引用顺序")
    rules["summary"] = "；".join(summary) if summary else "已读取格式指南，但未识别出明确格式条目。"
    return rules


def _iter_paragraph_runs_text(paragraph) -> str:
    return "".join(run.text for run in paragraph.runs) if paragraph.runs else paragraph.text


def _replace_text_with_style(paragraph, new_text: str, font_name: str | None = None, size_pt: float | None = None, bold: bool | None = None) -> None:
    if paragraph.runs:
        first_run = paragraph.runs[0]
        first_run.text = new_text
        _set_run_font(first_run, font_name, size_pt, bold)
        for extra_run in paragraph.runs[1:]:
            extra_run.text = ""
    else:
        run = paragraph.add_run(new_text)
        _set_run_font(run, font_name, size_pt, bold)


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
        "citations": 0,
        "references": 0,
        "figures": 0,
        "tables": 0,
    }
    citation_numbers = []
    reference_numbers = []
    seen_reference_section = False
    figure_captions = []
    table_captions = []
    figure_refs = set()
    table_refs = set()

    for idx, paragraph in enumerate(doc.paragraphs):
        text = _paragraph_text(paragraph)
        if not text:
            continue
        citations = [int(num) for num in re.findall(r"\[(\d+)\]", text)]
        citation_numbers.extend(citations)
        stats["citations"] += len(citations)

        if re.search(r"(?:图|Figure)\s*\d+", text):
            figure_refs.update(int(num) for num in re.findall(r"(?:图|Figure)\s*(\d+)", text))
        if re.search(r"(?:表|Table)\s*\d+", text):
            table_refs.update(int(num) for num in re.findall(r"(?:表|Table)\s*(\d+)", text))

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
            continue

        stats["body_paragraphs"] += 1
        fmt = paragraph.paragraph_format
        if guideline_rules.get("line_spacing") and fmt.line_spacing and abs(float(fmt.line_spacing.pt) - guideline_rules["line_spacing"]) > 0.5:
            issues.append({
                "id": f"body_spacing_{idx}",
                "type": "body_spacing",
                "label": "统一正文行距",
                "detail": f"第 {idx + 1} 段行距不是 {guideline_rules['line_spacing']} 磅。",
            })

        if guideline_rules.get("enforce_cn_punctuation") and re.search(r"[,;:!?]", text) and re.search(r"[\u4e00-\u9fff]", text):
            issues.append({
                "id": f"cn_punc_{idx}",
                "type": "language",
                "label": "统一中文标点与空格",
                "detail": f"第 {idx + 1} 段存在中文语境下的半角标点或异常空格。",
            })

        if guideline_rules.get("enforce_unit_spacing") and re.search(r"\d+(?:\.\d+)?(?:%|℃|°C|K|kg|g|mg|μm|nm|mm|cm|mL|L|MPa|GPa|Pa|h|min|s)\b", text):
            issues.append({
                "id": f"unit_spacing_{idx}",
                "type": "language",
                "label": "修正数字与单位间距",
                "detail": f"第 {idx + 1} 段检测到数字与单位直接相连。",
            })

        if re.match(r"^(参考文献|References)\s*$", text, re.IGNORECASE):
            seen_reference_section = True
            continue

        if seen_reference_section and re.match(r"^\[(\d+)\]", text):
            ref_num = int(re.match(r"^\[(\d+)\]", text).group(1))
            reference_numbers.append(ref_num)
            stats["references"] += 1

        if re.match(r"^(图|Figure)\s*\d+", text):
            numbers = [int(num) for num in re.findall(r"(?:图|Figure)\s*(\d+)", text)]
            figure_captions.extend(numbers)
            stats["figures"] += len(numbers)
        if re.match(r"^(表|Table)\s*\d+", text):
            numbers = [int(num) for num in re.findall(r"(?:表|Table)\s*(\d+)", text)]
            table_captions.extend(numbers)
            stats["tables"] += len(numbers)

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
        if not reference_numbers:
            issues.append({
                "id": "refs_missing",
                "type": "references",
                "label": "检查参考文献格式",
                "detail": "未检测到明确的 [n] 参考文献条目，无法确认 GB/T 7714 对齐情况。",
            })
        else:
            issues.append({
                "id": "refs_hanging",
                "type": "references",
                "label": "统一参考文献悬挂缩进",
                "detail": "检测到参考文献区，建议统一悬挂缩进与段间距。",
            })

    if guideline_rules.get("check_citation_order") and citation_numbers:
        expected = list(range(1, max(citation_numbers) + 1))
        if sorted(set(citation_numbers)) != expected:
            issues.append({
                "id": "citation_order",
                "type": "references",
                "label": "校验正文引用编号顺序",
                "detail": "正文 [n] 编号存在跳号或乱序。",
            })
        if reference_numbers and set(citation_numbers) - set(reference_numbers):
            issues.append({
                "id": "citation_reference_mismatch",
                "type": "references",
                "label": "对齐正文引用与参考文献",
                "detail": "正文引用与末尾参考文献 [n] 编号未完全对齐。",
            })

    if guideline_rules.get("check_captions"):
        missing_fig_refs = sorted(set(figure_captions) - figure_refs)
        missing_table_refs = sorted(set(table_captions) - table_refs)
        if missing_fig_refs:
            issues.append({
                "id": "figure_caption_refs",
                "type": "caption",
                "label": "核对图题注与正文引用",
                "detail": f"检测到图题注编号 {missing_fig_refs} 未在正文中引用。",
            })
        if missing_table_refs:
            issues.append({
                "id": "table_caption_refs",
                "type": "caption",
                "label": "核对表题注与正文引用",
                "detail": f"检测到表题注编号 {missing_table_refs} 未在正文中引用。",
            })

    unique_options = []
    seen = set()
    for issue in issues:
        if issue["label"] not in seen:
            unique_options.append(issue)
            seen.add(issue["label"])

    pipeline_steps = [
        {"label": "解析规则包", "detail": guideline_rules.get("summary", "未提供规则摘要")},
        {"label": "引用完整性检查", "detail": f"正文检测到 {stats['citations']} 处引用，参考文献区检测到 {stats['references']} 条。"},
        {"label": "语言规范检查", "detail": "检查中文标点、空格、数字与单位间距。"},
        {"label": "结构格式检查", "detail": f"标题 {stats['headings']} 个，图题注 {stats['figures']} 个，表题注 {stats['tables']} 个。"},
    ]

    return {
        "issues": issues,
        "fix_options": unique_options,
        "stats": stats,
        "pipeline_steps": pipeline_steps,
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


def apply_docx_fixes(docx_bytes: bytes, guideline_rules: Dict, selected_fixes: List[str], micro_tune_request: str = "") -> bytes:
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

    if "统一中文标点与空格" in selected_fixes:
        punctuation_map = {
            ",": "，",
            ";": "；",
            ":": "：",
            "!": "！",
            "?": "？",
        }
        for paragraph in doc.paragraphs:
            text = _paragraph_text(paragraph)
            if not text or not re.search(r"[\u4e00-\u9fff]", text):
                continue
            fixed = text
            for src, dst in punctuation_map.items():
                fixed = re.sub(rf"(?<=[\u4e00-\u9fff])\{src}", dst, fixed)
                fixed = re.sub(rf"\{src}(?=[\u4e00-\u9fff])", dst, fixed)
            fixed = re.sub(r"([\u4e00-\u9fff])\s+([，。；：！？])", r"\1\2", fixed)
            fixed = re.sub(r"([，。；：！？])\s+([\u4e00-\u9fff])", r"\1\2", fixed)
            if fixed != text:
                _replace_text_with_style(paragraph, fixed, guideline_rules.get("body_font") or "宋体", 12)

    if "修正数字与单位间距" in selected_fixes:
        for paragraph in doc.paragraphs:
            text = _paragraph_text(paragraph)
            if not text:
                continue
            fixed = re.sub(r"(\d+(?:\.\d+)?)(%|℃|°C|K|kg|g|mg|μm|nm|mm|cm|mL|L|MPa|GPa|Pa|h|min|s)\b", r"\1 \2", text)
            if fixed != text:
                _replace_text_with_style(paragraph, fixed, guideline_rules.get("body_font") or "宋体", 12)

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

    if "校验正文引用编号顺序" in selected_fixes or "对齐正文引用与参考文献" in selected_fixes:
        seen_reference_section = False
        next_ref_number = 1
        next_body_number = 1
        for paragraph in doc.paragraphs:
            text = _paragraph_text(paragraph)
            if not text:
                continue
            if re.match(r"^(参考文献|References)\s*$", text, re.IGNORECASE):
                seen_reference_section = True
                continue
            if seen_reference_section and re.match(r"^\[(\d+)\]", text):
                fixed = re.sub(r"^\[(\d+)\]", f"[{next_ref_number}]", text, count=1)
                next_ref_number += 1
                if fixed != text:
                    _replace_text_with_style(paragraph, fixed, guideline_rules.get("body_font") or "宋体", 12)
            elif not seen_reference_section and re.search(r"\[(\d+)\]", text):
                def repl(_match):
                    nonlocal next_body_number
                    value = f"[{next_body_number}]"
                    next_body_number += 1
                    return value
                fixed = re.sub(r"\[(\d+)\]", repl, text)
                if fixed != text:
                    _replace_text_with_style(paragraph, fixed, guideline_rules.get("body_font") or "宋体", 12)

    if "核对图题注与正文引用" in selected_fixes or "核对表题注与正文引用" in selected_fixes:
        in_body = True
        figure_seen = set()
        table_seen = set()
        for paragraph in doc.paragraphs:
            text = _paragraph_text(paragraph)
            if not text:
                continue
            if re.match(r"^(参考文献|References)\s*$", text, re.IGNORECASE):
                in_body = False
            if in_body:
                figure_seen.update(int(num) for num in re.findall(r"(?:图|Figure)\s*(\d+)", text))
                table_seen.update(int(num) for num in re.findall(r"(?:表|Table)\s*(\d+)", text))
                continue
            if re.match(r"^(图|Figure)\s*(\d+)", text):
                num = int(re.match(r"^(?:图|Figure)\s*(\d+)", text).group(1))
                if num not in figure_seen:
                    fixed = f"图 {num}（正文暂未引用，已标记补查）" + re.sub(r"^(?:图|Figure)\s*\d+", "", text, count=1)
                    _replace_text_with_style(paragraph, fixed, guideline_rules.get("body_font") or "宋体", 12)
            if re.match(r"^(表|Table)\s*(\d+)", text):
                num = int(re.match(r"^(?:表|Table)\s*(\d+)", text).group(1))
                if num not in table_seen:
                    fixed = f"表 {num}（正文暂未引用，已标记补查）" + re.sub(r"^(?:表|Table)\s*\d+", "", text, count=1)
                    _replace_text_with_style(paragraph, fixed, guideline_rules.get("body_font") or "宋体", 12)

    request_lower = micro_tune_request.lower().strip()
    if request_lower:
        if any(token in request_lower for token in ["页脚", "footer"]):
            footer_size = 11 if any(token in request_lower for token in ["大", "bigger", "larger"]) else 9 if any(token in request_lower for token in ["小", "smaller"]) else 10
            for section in doc.sections:
                footer = section.footer
                if not footer.paragraphs:
                    footer.add_paragraph()
                if not footer.paragraphs[0].text.strip():
                    footer.paragraphs[0].text = "学研 · 格式微调"
                footer.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in footer.paragraphs[0].runs:
                    _set_run_font(run, guideline_rules.get("body_font") or "宋体", footer_size)
        if any(token in request_lower for token in ["页眉", "header"]):
            header_size = 11 if any(token in request_lower for token in ["大", "bigger", "larger"]) else 9 if any(token in request_lower for token in ["小", "smaller"]) else 10
            for section in doc.sections:
                header = section.header
                if not header.paragraphs:
                    header.add_paragraph()
                if not header.paragraphs[0].text.strip():
                    header.paragraphs[0].text = "学研 · 确定性格式对齐"
                header.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in header.paragraphs[0].runs:
                    _set_run_font(run, guideline_rules.get("body_font") or "宋体", header_size)
        if any(token in request_lower for token in ["行距", "line spacing"]):
            spacing_pt = 24 if any(token in request_lower for token in ["大", "更大", "larger"]) else 18 if any(token in request_lower for token in ["小", "更小", "smaller"]) else (guideline_rules.get("line_spacing") or 20)
            for paragraph in doc.paragraphs:
                if _paragraph_text(paragraph) and not _is_heading(paragraph):
                    paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
                    paragraph.paragraph_format.line_spacing = Pt(spacing_pt)

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
        f"{stats.get('citations', 0)} 处正文引用，"
        f"{stats.get('references', 0)} 条参考文献，"
        f"发现 {len(issues)} 条待处理问题。"
    )


def build_format_audit_table(audit_report: Dict) -> str:
    issues = audit_report.get("issues", [])
    if not issues:
        return "| 检查项 | 状态 | 改进建议 |\n| --- | --- | --- |\n| 格式审计 | 🟢 已达标 | 未发现需要修复的样式问题。 |"

    rows = ["| 检查项 | 状态 | 改进建议 |", "| --- | --- | --- |"]
    severe_types = {"header", "references", "caption"}
    for issue in issues:
        status = "🔴 需优先处理" if issue.get("type") in severe_types else "🟡 建议修正"
        rows.append(
            f"| {issue['label']} | {status} | {issue['detail']} |"
        )
    return "\n".join(rows)


def build_format_audit_rows(audit_report: Dict) -> List[Dict[str, str]]:
    rows: List[Dict[str, str]] = []
    severe_types = {"header", "references", "caption"}
    for issue in audit_report.get("issues", []):
        rows.append(
            {
                "检查项": issue.get("label", "未命名问题"),
                "优先级": "高" if issue.get("type") in severe_types else "中",
                "问题类型": issue.get("type", "unknown"),
                "审计结果": issue.get("detail", ""),
            }
        )
    return rows


def run_formatting_pipeline(docx_bytes: bytes, guideline_rules: Dict, micro_tune_request: str = "") -> Tuple[Dict, bytes]:
    report = audit_docx_format(docx_bytes, guideline_rules)
    selected_fixes = [item["label"] for item in report.get("fix_options", [])]
    output_bytes = apply_docx_fixes(docx_bytes, guideline_rules, selected_fixes, micro_tune_request)
    return report, output_bytes


def render_copy_text(text: str, key: str) -> None:
    payload = json.dumps(text)
    components.html(
        f"""
        <script>
        const text = {payload};
        const btn = window.parent.document.getElementById('{key}');
        if (btn && !btn.dataset.copyBound) {{
            btn.dataset.copyBound = '1';
            btn.addEventListener('click', async () => {{
                try {{
                    await navigator.clipboard.writeText(text);
                }} catch (e) {{}}
            }});
        }}
        </script>
        """,
        height=0,
    )


def inject_custom_css() -> None:
    st.markdown(
        UI_CONFIG["primary_button_css"],
        unsafe_allow_html=True,
    )


def render_sidebar_brand() -> None:
    st.sidebar.markdown("<div class='xueyan-badge'>Xueyan GOOD</div>", unsafe_allow_html=True)


def render_sidebar_panel(title: str, note: str) -> None:
    st.sidebar.markdown(
        f"<div class='sidebar-panel'><div class='sidebar-panel-title'>{title}</div><p class='sidebar-panel-note'>{note}</p></div>",
        unsafe_allow_html=True,
    )


def render_header() -> None:
    st.markdown(
        """
<div class="main-header">
    <h1>学研 · Xueyan Workstation</h1>
    <p>双引擎科研工作台 · 写作协作与确定性排版在同一界面内完成</p>
</div>
""",
        unsafe_allow_html=True,
    )
    col1, col2, col3 = st.columns(3, gap="small")
    with col1:
        st.success("✅ API 就绪" if CLAUDE_API_KEY else "❌ API 未配置")
    with col2:
        st.info(f"📚 {len(append_skill_preview())} 组规则包")
    with col3:
        st.success("⏱️ 300s 超时 | 3次重试")


def render_result_actions(section_name: str, current_input: str, previous_output: str, function: str, domain: str) -> None:
    st.markdown("<div class='result-toolbar'><strong>快捷操作</strong></div>", unsafe_allow_html=True)
    copy_button_id = f"copy_result_btn_{section_name}_{function}".replace(" ", "_")
    col1, col2 = st.columns([1, 1])
    with col1:
        render_copy_text(previous_output, copy_button_id)
        if st.button("📋 复制结果", key=copy_button_id, use_container_width=True):
            st.toast("结果已复制到系统剪贴板")
        st.download_button(
            "📄 下载结果.txt",
            previous_output.encode("utf-8"),
            file_name=f"xueyan_{section_name}_result.txt",
            mime="text/plain",
            key=f"download_preview_{section_name}_{function}",
            use_container_width=True,
        )
    with col2:
        if function in MODIFICATION_FUNCTIONS:
            redline = create_docx_with_redlines(current_input, previous_output, {
                "function": function,
                "section": section_name,
                "domain": domain,
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            })
            if st.download_button(
                "📋 Redlining 修订",
                redline,
                file_name=f"xueyan_redline_{section_name}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key=f"redline_preview_{section_name}_{function}",
                use_container_width=True,
            ):
                st.toast("Redlining 文档已准备下载")
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
    with st.status("处理中...", expanded=True) as status:
        status.write(f"已锁定板块：{section_name}")
        status.write(f"正在应用功能：{function}")
        status.write(f"正在保护术语与 LaTeX：{domain}")
        protected_input, term_mapping = protect_hard_terms(current_input, domain)
        full_prompt = create_strict_prompt(function, protected_input, section_name, domain, reference_styles, input_lang)
        status.write("正在调用写作引擎")
        start_time = time.time()
        response = call_api(full_prompt, timeout=300)
        elapsed = time.time() - start_time
        status.write("正在恢复受保护术语")

        if function == "📝 中转英翻译":
            output_lang = detect_language(response)
            if output_lang == input_lang:
                status.update(label="处理失败", state="error", expanded=True)
                st.error("❌ 翻译校验失败：输出语言与输入相同，请重试")
                return

        final_output = restore_hard_terms(response, term_mapping)
        st.session_state[output_key] = final_output
        st.session_state[writing_term_count_key(section_name)] = len(term_mapping)
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
        status.update(label="处理完成", state="complete", expanded=False)

def render_writing_section(section_name: str, function: str, domain: str, reference_styles: List[str]) -> None:
    input_key = writing_input_key(section_name)
    output_key = writing_output_key(section_name)
    note_key = writing_note_key(section_name)
    term_count = st.session_state.get(writing_term_count_key(section_name), 0)
    previous_note = st.session_state.get(note_key, "")
    domain_profile = DOMAIN_PROFILES.get(domain, {})

    col_left, col_right = st.columns([1, 1], gap="large")
    with col_left:
        st.markdown(f"### 📥 原文输入 [{section_name}]")
        st.caption("把最原始的想法、草稿、老师意见或文献信息放这里。")
        st.text_area(
            "原文输入",
            key=input_key,
            height=380,
            label_visibility="collapsed",
        )
        literature_file = st.file_uploader(
            "上传文献/原稿",
            type=["pdf", "docx", "txt", "md"],
            key=f"literature_input_{section_name}",
        )
        if literature_file:
            extracted = extract_text(literature_file.read(), literature_file.name)
            if extracted and not extracted.startswith("解析"):
                st.caption(f"已读取文献：{literature_file.name}")
                if st.button("📚 填入原文输入", key=f"fill_literature_{section_name}", use_container_width=True):
                    st.session_state[input_key] = extracted[:12000]
        current_input = st.session_state.get(input_key, "")
        input_language = detect_language(current_input)
        st.caption(f"📊 {len(current_input)} 字符 | 语言: {'中文' if input_language == 'zh' else '英文' if input_language == 'en' else '待识别'}")
        btn1, btn2, btn3 = st.columns([3, 1, 1])
        with btn1:
            if st.button(f"✨ 执行 {function}", key=f"process_{section_name}", type="primary", use_container_width=True):
                handle_writing_process(section_name, function, domain, reference_styles)
        with btn2:
            if st.button("🗑️", key=f"clear_{section_name}"):
                st.session_state[input_key] = ""
                st.session_state[output_key] = ""
                st.session_state[note_key] = ""
        with btn3:
            if st.button("🔄", key=f"rerun_{section_name}"):
                handle_writing_process(section_name, function, domain, reference_styles)

    with col_right:
        st.markdown(f"### 🪄 处理结果 [{section_name}]")
        st.caption("右侧只展示处理后的结果，切换功能时不丢失。")
        st.text_area(
            "处理结果",
            key=output_key,
            height=380,
            disabled=True,
            label_visibility="collapsed",
        )
        previous_output = st.session_state.get(output_key, "")
        if previous_output:
            render_result_actions(section_name, st.session_state.get(input_key, ""), previous_output, function, domain)
            if previous_note:
                st.caption(previous_note)
            lock_hint = domain_profile.get("hard_lock_terms", [])[:2] + domain_profile.get("hard_lock_regex", [])[:1]
            if term_count:
                st.caption(f"已按当前学科自动保护 {term_count} 个核心锁项：{', '.join(lock_hint) if lock_hint else 'LaTeX / 术语 / 单位'}。")
            else:
                st.caption("已启用当前学科核心锁，LaTeX 公式与术语默认不改写。")
            meta = {
                "function": function,
                "section": section_name,
                "domain": domain,
                "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
            export1, export2 = st.columns([1, 1])
            with export1:
                if st.download_button(
                    "📥 导出 Word",
                    create_docx(previous_output, meta),
                    file_name=f"xueyan_{section_name}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key=f"export_docx_{section_name}",
                    use_container_width=True,
                ):
                    st.toast("Word 文档已准备下载")
            with export2:
                if function in MODIFICATION_FUNCTIONS:
                    if st.download_button(
                        "📋 导出 Redlining",
                        create_docx_with_redlines(st.session_state.get(input_key, ""), previous_output, meta),
                        file_name=f"xueyan_redline_{section_name}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"export_redline_{section_name}",
                        use_container_width=True,
                    ):
                        st.toast("Redlining 文档已准备下载")
        else:
            st.caption("这里会生成可直接拿去改稿的版本。")


def get_function_nav_index(current_function: str) -> int:
    for idx, item in enumerate(FUNCTION_NAV):
        if item["value"] == current_function:
            return idx
    return 0


def render_writing_engine_sidebar() -> List[str]:
    render_sidebar_brand()
    st.sidebar.radio("核心引擎", [WRITING_PAGE, FORMATTING_PAGE], key="engine_mode")
    st.sidebar.selectbox("统一功能栏", options=FUNCTION_NAV, index=get_function_nav_index(st.session_state.get("writing_function", list(FUNCTION_MATRIX.keys())[0])), format_func=lambda item: item["label"], key="writing_function_selector")
    st.session_state.writing_function = st.session_state["writing_function_selector"]["value"]
    st.sidebar.selectbox("🔬 学科大脑", DOMAIN_ORDER, key="writing_domain_selector")
    st.session_state.writing_domain = st.session_state["writing_domain_selector"]

    function = st.session_state.get("writing_function", list(FUNCTION_MATRIX.keys())[0])
    func_info = FUNCTION_MATRIX.get(function, {})
    st.sidebar.markdown("<div class='sidebar-section-note'>当前功能说明</div>", unsafe_allow_html=True)
    st.sidebar.caption(func_info.get("description", ""))
    for rule in func_info.get("rules", []):
        st.sidebar.caption(rule)

    st.sidebar.markdown("---")
    st.sidebar.subheader("📚 影子合著者：标杆文献")
    reference_files = st.sidebar.file_uploader(
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
        st.sidebar.success(f"✅ 已加载 {len(st.session_state.writing_reference_docs)} 篇标杆文献")

    st.sidebar.markdown("---")
    st.sidebar.subheader("📁 普通文件上传")
    uploaded_file = st.sidebar.file_uploader("上传待处理文件", type=["pdf", "docx"], key="writing_uploaded_file")
    if uploaded_file:
        uploaded_bytes = uploaded_file.read()
        extracted = extract_text(uploaded_bytes, uploaded_file.name)
        if extracted and not extracted.startswith("解析"):
            st.sidebar.success(f"✅ 提取 {len(extracted)} 字符")
            if st.sidebar.button("📥 填入当前板块", use_container_width=True, key="fill_current_section"):
                active = st.session_state.writing_active_section
                st.session_state[writing_input_key(active)] = extracted[:10000]

    preview = append_skill_preview()
    st.sidebar.markdown("---")
    st.sidebar.caption(f"规则包：{', '.join(f'{k}={v}' for k, v in preview.items())}")
    return reference_styles


def render_writing_engine() -> None:
    init_writing_state()
    reference_styles = render_writing_engine_sidebar()
    function = st.session_state.get("writing_function", list(FUNCTION_MATRIX.keys())[0])
    domain = st.session_state.get("writing_domain", DOMAIN_ORDER[0])
    active_section = st.session_state.get("writing_active_section", SECTION_NAMES[0])
    focus_hint = SECTIONS.get(active_section, {}).get("focus", "聚焦学术表达")

    header_left, header_right = st.columns([1, 1], gap="large")
    with header_left:
        st.markdown(
            f"""
<div class="workbench-card compact">
    <div class="workbench-title">
        <div>
            <h3>写作引擎工作台</h3>
            <p>保持左右双栏直达输入/输出，减少首屏干扰。</p>
        </div>
        <span class="workbench-chip">{active_section}</span>
    </div>
    <div class="subtle-kpi-row">
        <strong>当前功能</strong>：{function}<br>
        <strong>当前学科</strong>：{domain}<br>
        <strong>焦点</strong>：{focus_hint}
    </div>
</div>
""",
            unsafe_allow_html=True,
        )
    with header_right:
        st.selectbox(
            "当前写作板块",
            SECTION_NAMES,
            index=SECTION_NAMES.index(active_section),
            key="writing_active_section_selector",
        )
        st.session_state.writing_active_section = st.session_state["writing_active_section_selector"]
        st.caption("保持首屏直接进入双栏输入/输出区。")

    active_section = st.session_state.get("writing_active_section", SECTION_NAMES[0])
    render_writing_section(active_section, function, domain, reference_styles)
    with st.expander("⏰ 版本时光机", expanded=False):
        render_history_panel()


def render_history_panel() -> None:
    st.caption("当前学科核心锁已启用：LaTeX 公式、术语 regex 与关键缩写默认不改写。")
    st.markdown("---")
    st.subheader("⏰ 版本时光机")
    filter_section = st.selectbox("筛选板块", ["全部"] + list(SECTIONS.keys()), key="history_section_filter")
    filter_function = st.selectbox("筛选功能", ["全部"] + list(FUNCTION_MATRIX.keys()), key="history_function_filter")
    filtered_history = st.session_state.writing_history
    if filter_section != "全部":
        filtered_history = [h for h in filtered_history if h["section"] == filter_section]
    if filter_function != "全部":
        filtered_history = [h for h in filtered_history if h["function"] == filter_function]

    if not filtered_history:
        st.info("📭 暂无符合条件的记录")
        return

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
            with c3:
                if st.button("🗑️", key=f"delete_{entry['id']}"):
                    st.session_state.writing_history = [h for h in st.session_state.writing_history if h['id'] != entry['id']]
            left, right = st.columns(2)
            with left:
                st.markdown("**原始输入**")
                st.text_area("", entry["input"], height=150, key=f"orig_{entry['id']}", disabled=True)
            with right:
                st.markdown("**处理输出**")
                st.text_area("", entry["output"], height=150, key=f"out_{entry['id']}", disabled=True)


def render_formatting_engine_sidebar() -> None:
    render_sidebar_brand()
    st.sidebar.radio("核心引擎", [WRITING_PAGE, FORMATTING_PAGE], key="engine_mode")
    st.sidebar.selectbox("🔬 学科大脑", DOMAIN_ORDER, key="format_formatting_domain_selector")
    st.session_state.formatting_domain = st.session_state["format_formatting_domain_selector"]
    st.sidebar.selectbox(
        "📚 规则包",
        options=list(FORMAT_RULESET_LIBRARY.keys()),
        format_func=lambda key: get_format_ruleset_profile(key)["label"],
        key="format_ruleset",
    )
    render_sidebar_panel("格式对齐", "排版引擎与写作引擎状态完全隔离，只保留确定性审计与自动修复。")
    st.sidebar.caption("规则来源参考 thesis-skills 的 check/fix 闭环，但这里直接面向 Word 文档执行。")


def render_formatting_engine() -> None:
    render_formatting_engine_sidebar()
    domain = st.session_state.get("formatting_domain", DOMAIN_ORDER[0])
    ruleset = st.session_state.get("format_ruleset", "tsinghua-thesis")
    ruleset_profile = get_format_ruleset_profile(ruleset)
    guideline_summary = st.session_state.get("format_guideline_summary") or "尚未解析格式指南"
    target_doc_bytes = st.session_state.get("format_target_doc_bytes", b"")
    target_doc_name = st.session_state.get("format_target_doc_name", "")
    audit_report = st.session_state.get("format_audit_report", {})
    audit_count = len(audit_report.get("issues", []))
    has_output_docx = bool(st.session_state.get("format_output_docx_bytes"))

    st.markdown(
        """
<div class="workbench-card compact">
    <div class="workbench-title">
        <div>
            <h3>排版引擎工作台</h3>
            <p>按 thesis-skills 的检查 → 修复思路重组为 Word 三段式闭环：指南、上传、自动审计与导出。</p>
        </div>
        <span class="workbench-chip">Formatting</span>
    </div>
</div>
""",
        unsafe_allow_html=True,
    )

    with st.expander("📘 格式规则与校准说明", expanded=True):
        st.markdown(f"**当前规则包**：{ruleset_profile['label']}")
        st.caption(ruleset_profile["summary"])
        for item in ruleset_profile["guidance"]:
            st.markdown(f"- {item}")
        format_guideline_text = st.text_area(
            "补充格式说明",
            value=st.session_state.get("format_guideline_text", ""),
            height=160,
            key="format_guideline_text",
            placeholder="可追加学院细则：如摘要关键词 3-5 个、一级标题黑体三号、页眉含校名等。",
        )
        format_guideline_file = st.file_uploader(
            "上传补充指南",
            type=["txt", "md", "docx"],
            key="format_guideline_file",
        )
        if st.button("🔍 解析当前规则", use_container_width=True, key="format_parse_guideline"):
            merged_text = ruleset_profile["summary"] + "\n" + "\n".join(ruleset_profile["guidance"]) + "\n" + format_guideline_text
            parsed_rules = ensure_guideline_rules(merged_text, format_guideline_file)
            st.session_state.format_guideline_summary = parsed_rules["summary"]
            st.session_state.format_guideline_rules = parsed_rules
            st.success(parsed_rules["summary"])
        elif guideline_summary:
            st.info(guideline_summary)

    stage_upload, stage_audit = st.columns([1, 1], gap="large")

    with stage_upload:
        st.markdown("### ② 上传待修复 Word")
        format_target_doc = st.file_uploader("上传 .docx", type=["docx"], key="format_target_doc_uploader")
        if format_target_doc:
            st.session_state.format_target_doc_bytes = format_target_doc.getvalue()
            st.session_state.format_target_doc_name = format_target_doc.name
            st.session_state.format_output_docx_bytes = b""
            st.session_state.format_audit_report = {}
            st.session_state.format_pipeline_summary = ""
            st.session_state.format_pipeline_steps = []
            st.session_state.format_audit_rows = []
            st.session_state.format_audited_docx_bytes = st.session_state.format_target_doc_bytes
            st.session_state.format_audited_docx_name = format_target_doc.name
            target_doc_bytes = st.session_state.format_target_doc_bytes
            target_doc_name = st.session_state.format_target_doc_name
        if target_doc_bytes and target_doc_name:
            st.success(f"已载入：{target_doc_name}")
            extracted = extract_text(target_doc_bytes, target_doc_name)
            if extracted and not extracted.startswith("解析"):
                st.text_area("文档预览", extracted[:3200], height=360, disabled=True, key="format_target_doc_preview")
        else:
            st.info("等待上传需要自动修复的 Word 文档。")

    with stage_audit:
        st.markdown("### ③ 审计与自动修复")
        stat1, stat2 = st.columns(2, gap="small")
        with stat1:
            st.metric("规则状态", "已解析" if st.session_state.get("format_guideline_rules") else "待解析")
        with stat2:
            st.metric("问题数", audit_count)

        audit_col, fix_col = st.columns(2, gap="small")
        with audit_col:
            if st.button("🩺 开始审计", use_container_width=True, key="format_run_audit"):
                if not target_doc_bytes:
                    st.warning("请先上传 .docx 文档。")
                else:
                    guideline_file = st.session_state.get("format_guideline_file")
                    merged_text = ruleset_profile["summary"] + "\n" + "\n".join(ruleset_profile["guidance"]) + "\n" + st.session_state.get("format_guideline_text", "")
                    current_rules = st.session_state.get("format_guideline_rules") or ensure_guideline_rules(merged_text, guideline_file)
                    with st.status("审计进行中...", expanded=True) as status:
                        status.write("正在解析规则包与补充指南...")
                        time.sleep(0.05)
                        status.write("正在审计正文引用与参考文献对应关系...")
                        time.sleep(0.05)
                        status.write("正在检查中文标点、空格与数字单位间距...")
                        time.sleep(0.05)
                        status.write("正在校验标题层级、图表题注与正文引用...")
                        report = audit_docx_format(target_doc_bytes, current_rules)
                        st.session_state.format_audit_report = report
                        st.session_state.format_fix_options = report.get("fix_options", [])
                        st.session_state.format_last_filename = target_doc_name
                        st.session_state.format_pipeline_steps = report.get("pipeline_steps", [])
                        st.session_state.format_pipeline_summary = create_format_audit_summary(report)
                        st.session_state.format_audit_rows = build_format_audit_rows(report)
                        st.session_state.format_audited_docx_bytes = target_doc_bytes
                        st.session_state.format_audited_docx_name = target_doc_name
                        st.session_state.format_output_docx_bytes = b""
                        status.update(label="审计完成", state="complete", expanded=False)
                    audit_report = st.session_state.get("format_audit_report", {})
                    audit_count = len(audit_report.get("issues", []))
                    has_output_docx = False

        with fix_col:
            st.caption("先审计，再修复。")

        if audit_report:
            st.success(st.session_state.get("format_pipeline_summary") or create_format_audit_summary(audit_report))
            for step in st.session_state.get("format_pipeline_steps", []):
                st.caption(f"{step['label']}：{step['detail']}")
            audit_rows = st.session_state.get("format_audit_rows", [])
            if audit_rows:
                st.dataframe(audit_rows, use_container_width=True, hide_index=True)
            else:
                st.markdown(build_format_audit_table(audit_report))

            if st.button("🚀 开始自动修复并生成新文档", use_container_width=True, key="format_run_fix"):
                if not target_doc_bytes:
                    st.warning("请先上传 .docx 文档。")
                else:
                    guideline_file = st.session_state.get("format_guideline_file")
                    merged_text = ruleset_profile["summary"] + "\n" + "\n".join(ruleset_profile["guidance"]) + "\n" + st.session_state.get("format_guideline_text", "")
                    current_rules = st.session_state.get("format_guideline_rules") or ensure_guideline_rules(merged_text, guideline_file)
                    selected_fixes = [item["label"] for item in st.session_state.get("format_fix_options", [])]
                    with st.status("自动修复进行中...", expanded=True) as status:
                        status.write("正在加载最近一次审计结果...")
                        time.sleep(0.05)
                        status.write("正在对齐引用编号、参考文献与题注结构...")
                        time.sleep(0.05)
                        status.write("正在统一中文标点、空格、数字与单位间距...")
                        time.sleep(0.05)
                        status.write("正在应用确定性样式修复并生成 Word...")
                        output_bytes = apply_docx_fixes(
                            target_doc_bytes,
                            current_rules,
                            selected_fixes,
                            st.session_state.get("format_micro_tune_request", ""),
                        )
                        st.session_state.format_output_docx_bytes = output_bytes
                        st.session_state.format_last_filename = target_doc_name
                        status.update(label="自动修复完成", state="complete", expanded=False)
                    st.success("自动修复完成，可直接下载新文档。")
                    st.toast("已生成修复版 Word")
                    has_output_docx = bool(st.session_state.get("format_output_docx_bytes"))
        else:
            st.info("先点“开始审计”查看问题，再点下方“🚀 开始自动修复并生成新文档”生成修复版 Word。")

        if has_output_docx:
            export_name = build_export_filename(
                "xueyan_format_fixed",
                Path(st.session_state.get("format_last_filename") or target_doc_name).stem,
            )
            st.download_button(
                "📥 下载确定性修复文档 (Word)",
                st.session_state.get("format_output_docx_bytes", b""),
                file_name=export_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="format_download_fixed_docx",
                use_container_width=True,
            )
            st.caption("下载内容绑定自动修复后的 DOCX，不再混用自由生成预览文档。")

    format_chat_value = st.chat_input("输入微调要求（如：页脚字号大一点）", key="format_micro_tune_chat")
    if format_chat_value:
        st.session_state.format_micro_tune_request = format_chat_value
        st.toast("已记录本轮格式微调要求")
    if st.session_state.get("format_micro_tune_request"):
        st.caption(f"最近一次微调要求：{st.session_state.format_micro_tune_request}")


def main() -> None:
    trim_session_state()
    inject_custom_css()
    render_header()

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
