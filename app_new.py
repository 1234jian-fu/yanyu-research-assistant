"""
学研·工科科研助手 v4.0 (Xueyan OS)
双引擎架构：论文内容写作 + 确定性格式对齐
"""

import os
import re
import time
import json
import base64
import random
from html import escape
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from copy import deepcopy
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Dict, List, Tuple

import anthropic
import fitz  # pymupdf
import requests
import streamlit as st
import streamlit.components.v1 as components
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor
try:
    from pptx import Presentation
    from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
    from pptx.util import Inches, Pt as PptxPt
    PPTX_AVAILABLE = True
except ModuleNotFoundError:
    Presentation = None
    MSO_AUTO_SHAPE_TYPE = None
    PPTX_AVAILABLE = False

    def Inches(value):
        return value

    def PptxPt(value):
        return value
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
CLAUDE_MODEL_OPTIONS = [
    "claude-opus-4-6",
    "claude-sonnet-4-6",
]

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "").strip()
GEMINI_BASE_URL = os.getenv("GEMINI_BASE_URL", "https://new.lemonapi.site").rstrip("/")
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gpt-image-2")
GEMINI_MODEL_OPTIONS = [
    "gpt-image-2",
    "gemini-3.1-flash-image-preview",
    "gemini-2.5-flash-image-preview",
    "gemini-2.0-flash-preview-image-generation",
]

WRITING_PAGE = "✍️ 论文写作"
FORMATTING_PAGE = "📐 格式对齐"
VIZ_PAGE = "🎨 视觉实验室"
PPT_PAGE = "🗂️ PPT大师"

CONFIG = {
    "sections": {
        "摘要": {
            "focus": "开门见山，数据支撑",
            "max_words": 250,
            "goal": "用最短篇幅交代研究对象、方法、核心结果与贡献。",
            "must_include": ["研究对象/问题", "采用的方法或策略", "关键结果或数据趋势", "结论或贡献"],
            "avoid": ["空泛背景铺垫", "无数据支撑的形容词", "展开过多机理讨论", "引入与本文无关信息"],
            "output_shape": "优先压缩成一段完整摘要，必要时使用4句结构。",
            "input_hint": "请输入研究对象、方法、最关键结果和最终结论，适合直接压缩成摘要。",
            "output_checkpoints": ["是否同时包含对象、方法、结果、贡献？", "是否有定量结果或明确趋势？", "是否避免过长背景铺垫？"],
        },
        "引言": {
            "focus": "背景转折，研究空白",
            "max_words": 800,
            "goal": "建立研究背景、指出现有不足，并自然引出本文问题与贡献。",
            "must_include": ["研究背景", "现有工作不足", "研究空白或痛点", "本文切入点/贡献"],
            "avoid": ["提前展开结果细节", "把方法写成操作手册", "空洞口号式创新描述", "缺少问题转折"],
            "output_shape": "建议按背景→不足→问题→本文贡献的顺序组织段落。",
            "input_hint": "请输入研究背景、现有不足、文献空白和你本文的切入点。",
            "output_checkpoints": ["是否清楚指出研究空白？", "是否自然引出本文工作？", "是否避免提前泄露结果？"],
        },
        "方法": {
            "focus": "流程清晰，参数精确",
            "max_words": 1500,
            "goal": "把实验/计算/系统方法写得可复现、可核对、步骤清楚。",
            "must_include": ["材料或数据来源", "关键步骤/流程", "核心参数与条件", "评价指标或表征方法"],
            "avoid": ["宣传式语言", "结果导向表述", "省略关键参数", "步骤顺序混乱"],
            "output_shape": "优先使用流程化段落，必要时分成材料、步骤、表征/评价三个层次。",
            "input_hint": "请输入材料来源、实验步骤、关键参数、仪器条件和评价方法。",
            "output_checkpoints": ["参数是否完整？", "步骤顺序是否清晰？", "是否具备可复现性？"],
        },
        "结果": {
            "focus": "数据客观，图表引用",
            "max_words": 1200,
            "goal": "客观描述数据、趋势、对比与异常点，并和图表编号对应。",
            "must_include": ["关键数据或现象", "趋势与对比", "图表编号或结果载体", "异常点/变化点"],
            "avoid": ["无依据拔高结论", "只说好不说对比", "脱离图表编号", "把讨论写进结果"],
            "output_shape": "按图表或实验顺序组织，每段尽量包含结果事实+趋势判断。",
            "input_hint": "请输入图表编号、关键数据、主要趋势、对比对象和异常现象。",
            "output_checkpoints": ["是否引用了图表或数据来源？", "是否体现趋势和对比？", "是否避免提前做机理解释？"],
        },
        "讨论": {
            "focus": "深度解读，文献对比",
            "max_words": 1000,
            "goal": "解释结果背后的原因、机制与边界条件，并联系文献展开讨论。",
            "must_include": ["结果解释", "可能机制或原因", "与文献对照", "局限性或边界条件"],
            "avoid": ["机械重复结果描述", "缺少解释链条", "完全脱离文献", "无边界条件意识"],
            "output_shape": "优先写成解释型段落，突出因果链、机制链和文献对照。",
            "input_hint": "请输入你对结果的解释、可能机制、参考文献对照和局限性判断。",
            "output_checkpoints": ["是否解释了为什么？", "是否和文献形成对照？", "是否避免只是重复结果？"],
        },
        "结论": {
            "focus": "总结贡献，展望未来",
            "max_words": 300,
            "goal": "简洁总结工作贡献、主要发现与后续展望，不引入新细节。",
            "must_include": ["核心发现", "本文贡献", "意义或应用价值", "合理展望"],
            "avoid": ["引入新实验细节", "重复整段引言", "展开大段讨论", "夸张式总结"],
            "output_shape": "优先写成1-2段收束性文本，最后一句可给出展望。",
            "input_hint": "请输入你最想保留的核心发现、贡献总结和后续展望。",
            "output_checkpoints": ["是否只总结已出现内容？", "是否避免新事实？", "是否收束得足够简洁？"],
        },
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
        "🧩 文本降重": {
            "description": "同语种原创改写优化，保持观点与数据不变",
            "rules": ["❌ 禁止翻译", "❌ 禁止篡改事实", "✅ 输出优化后文本+修改说明"],
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
        "✨ 核心润色": ["✨ 表达润色", "🧩 文本降重", "🤖 去AI味 (Humanizer)", "🎯 精修模式", "📋 Redlining修订", "🔍 逻辑检查", "👨‍⚖️ Reviewer视角"],
        "🔄 翻译转换": ["📝 中转英翻译", "📝 引用验证", "🎨 图表规范检查", "🔄 版本对比"],
        "📝 影子写作": ["✍️ 逐段起草", "✍️ 影子写作", "📄 节节头脑风暴", "💡 研究想法构思", "🧠 ML论文写作", "📊 概念图设计"],
    },
    "function_nav": [
        {"label": "✨ 表达润色 (核心)", "value": "✨ 表达润色"},
        {"label": "🧩 文本降重", "value": "🧩 文本降重"},
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
        "text_dedup": """
## 文本降重协议（原创改写与规范表达优化）
### 核心目标
1. 保持原文核心观点、事实、数据、术语和学术结论不被篡改。
2. 在不改变原意前提下，通过句式变换、语序调整、段落重组、术语替换、逻辑重建提升表达质量与原创性。
3. 保持语言正式、清晰、连贯，符合学术写作风格。
4. 不编造数据、不新增不存在的实验结果、不虚构参考文献、不擅自扩大结论。
5. 专有名词、变量名、材料名、化学式、公式、参考编号必须准确保留。
6. 当用户意图为规避查重或掩盖抄袭时，不提供规避策略，改为执行原创改写与规范表达优化，并提醒保留必要引用。

### 支持的处理模式库
- 基础改写：调整语序、替换表达、减少重复措辞。
- 学术润色：增强正式性、客观性、严谨性。
- 句式重构：改变句法结构、长短句重组、主被动转换。
- 段落重构：重建段落内部逻辑与信息顺序。
- 压缩精炼：压缩冗余表达，保留关键信息。
- 扩写说明：补充逻辑衔接和必要解释，但不编造新事实。
- 关键词优化：在保证术语准确前提下优化同义表达。

### 输出要求
- 优先输出“优化后文本”。
- 然后输出“修改说明”，简要标注改动类型（句式重构/术语统一/逻辑顺序调整/压缩冗余等）。
- 若存在逻辑跳跃、指代不清、术语不统一、学术表达不规范、缺少必要引用支撑，需明确提示。
- 不输出“可降低多少重复率”等承诺。
- 不以“躲避检测”“骗过系统”为目标。
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
    section[data-testid="stSidebar"] {
        min-width: 20rem;
        background: linear-gradient(180deg, #f7fbff 0%, #f2f7fd 100%);
        border-right: 1px solid #dbe7f5;
    }
    section[data-testid=\"stSidebar\"] .block-container {
        padding-top: 0.75rem;
    }
    .xueyan-badge {
        display: flex;
        align-items: center;
        justify-content: center;
        padding: 0.62rem 0.85rem;
        margin: 0 0 0.75rem 0;
        border-radius: 14px;
        background: linear-gradient(180deg, rgba(255,255,255,0.96) 0%, rgba(243,248,255,0.96) 100%);
        color: #17365d;
        font-family: "Avenir Next", "PingFang SC", "Noto Sans SC", sans-serif;
        font-weight: 700;
        font-size: 0.92rem;
        letter-spacing: 0.08em;
        box-shadow: 0 8px 22px rgba(15, 23, 42, 0.04);
        border: 1px solid #d8e6f6;
    }
    .viz-language-muted {
        opacity: 0.5;
        filter: grayscale(0.15);
        pointer-events: none;
    }
    .sidebar-panel {
        padding: 0.68rem 0.8rem;
        margin: 0.28rem 0 0.72rem 0;
        border-radius: 14px;
        background: linear-gradient(180deg, rgba(255,255,255,0.88) 0%, rgba(247,251,255,0.88) 100%);
        border: 1px solid #dbe7f5;
        box-shadow: 0 8px 20px rgba(15, 23, 42, 0.04);
    }
    section[data-testid=\"stSidebar\"] [data-testid=\"stExpander\"] {
        border: 1px solid #dbe7f5;
        border-radius: 12px;
        background: rgba(255,255,255,0.88);
        overflow: hidden;
        margin-bottom: 0.55rem;
    }
    section[data-testid="stSidebar"] [data-testid="stExpander"] details summary {
        padding: 0.34rem 0.6rem;
        font-weight: 600;
    }
    section[data-testid=\"stSidebar\"] .stRadio > div {
        gap: 0.35rem;
    }
    .sidebar-panel-title {
        margin: 0 0 0.18rem 0;
        color: #10233f;
        font-size: 0.9rem;
        font-weight: 700;
    }
    .sidebar-panel-note {
        margin: 0;
        color: #64748b;
        font-size: 0.77rem;
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
        padding: 0.72rem 0.95rem;
        background: linear-gradient(180deg, rgba(255,255,255,0.98) 0%, rgba(245,249,255,0.98) 100%);
        color: #10233f;
        border-radius: 16px;
        margin-bottom: 0.55rem;
        box-shadow: 0 8px 24px rgba(15, 23, 42, 0.04);
        border: 1px solid #dbe7f5;
    }
    .main-header::after {
        display: none;
    }
    .main-header h1 {
        margin: 0;
        font-size: 1.22rem;
        letter-spacing: 0.01em;
    }
    .main-header p {
        margin: 0.18rem 0 0 0;
        color: #5a6f89;
        font-size: 0.83rem;
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
    .viz-lab-shell {
        display: block;
    }
    .viz-thin-header {
        padding: 0.72rem 0.9rem;
        margin-bottom: 0.65rem;
        background: linear-gradient(180deg, rgba(255,255,255,0.98) 0%, rgba(245,249,255,0.98) 100%);
        border: 1px solid #dbe7f5;
        border-radius: 16px;
        box-shadow: 0 8px 24px rgba(15, 23, 42, 0.04);
    }
    .viz-thin-header-row {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 1rem;
    }
    .viz-thin-header h3 {
        margin: 0;
        font-size: 1.02rem;
        color: #10233f;
    }
    .viz-thin-header p {
        margin: 0.16rem 0 0 0;
        font-size: 0.83rem;
        color: #5a6f89;
    }
    .viz-topbar {
        padding: 0.8rem 0.9rem 0.45rem 0.9rem;
        margin-bottom: 0.7rem;
        background: linear-gradient(180deg, #f8fbff 0%, #f3f8ff 100%);
        border: 1px solid #dbe7f5;
        border-radius: 16px;
        box-shadow: inset 0 1px 0 rgba(255,255,255,0.82);
    }
    .viz-topbar div[data-testid="stHorizontalBlock"] {
        gap: 0.55rem;
    }
    .viz-topbar [data-baseweb="select"] > div,
    .viz-topbar [data-baseweb="radio"] {
        background: rgba(255,255,255,0.78);
        border-radius: 12px;
    }
    .viz-main-card {
        padding: 0.95rem 1rem;
        margin-bottom: 0.75rem;
        border-radius: 18px;
        background: linear-gradient(180deg, rgba(255,255,255,0.98) 0%, rgba(248,251,255,0.98) 100%);
        border: 1px solid #d7e4f4;
        box-shadow: 0 10px 28px rgba(15, 23, 42, 0.05);
    }
    .viz-main-card + .viz-main-card {
        margin-top: 0.1rem;
    }
    .viz-main-card div[data-testid="stHorizontalBlock"] {
        gap: 0.65rem;
    }
    .viz-panel-label {
        font-size: 0.86rem;
        font-weight: 700;
        margin-bottom: 0.38rem;
        letter-spacing: 0.01em;
    }
    .viz-panel-note {
        margin: 0 0 0.45rem 0;
        color: #60748c;
        font-size: 0.79rem;
        line-height: 1.55;
    }
    .viz-secondary-card {
        padding: 0.82rem 0.9rem;
        margin-bottom: 0.72rem;
        border-radius: 16px;
        background: linear-gradient(180deg, #fbfdff 0%, #f7fbff 100%);
        border: 1px solid #e1ebf7;
    }
    .viz-secondary-card div[data-testid="stExpander"] {
        border: 1px solid #e4edf8;
        border-radius: 14px;
        background: rgba(255,255,255,0.74);
        margin-bottom: 0.52rem;
        overflow: hidden;
    }
    .viz-secondary-card details {
        background: transparent;
    }
    .viz-secondary-card summary {
        font-weight: 600;
    }
    .viz-secondary-card div[data-testid="stExpanderDetails"] {
        padding-top: 0.18rem;
    }
    .viz-compact-status {
        padding: 0.62rem 0.76rem;
        margin: 0.55rem 0 0.75rem 0;
        border-radius: 12px;
        background: linear-gradient(180deg, #f4f8ff 0%, #eef6ff 100%);
        border: 1px solid #d7e7fb;
        color: #3d5b80;
        font-size: 0.8rem;
    }
    .viz-status-pills {
        display: flex;
        flex-wrap: wrap;
        gap: 0.45rem;
        margin: 0 0 0.7rem 0;
    }
    .viz-status-pill {
        display: inline-flex;
        align-items: center;
        gap: 0.28rem;
        padding: 0.3rem 0.66rem;
        border-radius: 999px;
        background: linear-gradient(180deg, #f3f8ff 0%, #edf5ff 100%);
        border: 1px solid #d7e7fb;
        color: #33567e;
        font-size: 0.76rem;
        line-height: 1;
        box-shadow: inset 0 1px 0 rgba(255,255,255,0.82);
    }
    .viz-status-pill strong {
        color: #11345d;
        font-weight: 700;
    }
    .viz-quiet-block {
        margin-top: 0.35rem;
        color: #6a7c92;
        font-size: 0.78rem;
        line-height: 1.5;
    }
    .viz-preview-shell {
        width: 100%;
        border: 1px solid #dbe7f5;
        border-radius: 18px;
        background: linear-gradient(180deg, #ffffff 0%, #f6faff 100%);
        box-shadow: inset 0 1px 0 rgba(255,255,255,0.9), 0 10px 28px rgba(15, 23, 42, 0.05);
        padding: 0.9rem;
        overflow: hidden;
    }
    .viz-preview-frame {
        width: 100%;
        height: 100%;
        display: flex;
        align-items: center;
        justify-content: center;
        border-radius: 14px;
        background:
            radial-gradient(circle at top, rgba(191, 219, 254, 0.38), rgba(255,255,255,0.96)),
            linear-gradient(135deg, rgba(241,247,255,0.85) 0%, rgba(255,255,255,0.96) 100%);
        border: 1px dashed #c7d6ea;
    }
    .viz-preview-frame img {
        max-width: 100%;
        width: auto;
        height: auto;
        object-fit: contain;
        border-radius: 14px;
        box-shadow: 0 14px 30px rgba(15, 23, 42, 0.12);
    }
    .viz-preview-empty {
        max-width: 78%;
        text-align: center;
        color: #64748b;
        font-size: 0.92rem;
        line-height: 1.6;
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
    .lab-thin-header {
        padding: 0.74rem 0.95rem;
        margin-bottom: 0.7rem;
        background: linear-gradient(180deg, rgba(255,255,255,0.98) 0%, rgba(245,249,255,0.98) 100%);
        border: 1px solid #dbe7f5;
        border-radius: 16px;
        box-shadow: 0 8px 24px rgba(15, 23, 42, 0.04);
    }
    .lab-thin-header-row {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 1rem;
    }
    .lab-thin-header h3 {
        margin: 0;
        font-size: 1.02rem;
        color: #10233f;
    }
    .lab-thin-header p {
        margin: 0.16rem 0 0 0;
        font-size: 0.83rem;
        color: #5a6f89;
    }
    .lab-main-card {
        padding: 0.95rem 1rem;
        margin-bottom: 0.75rem;
        border-radius: 18px;
        background: linear-gradient(180deg, rgba(255,255,255,0.98) 0%, rgba(248,251,255,0.98) 100%);
        border: 1px solid #d7e4f4;
        box-shadow: 0 10px 28px rgba(15, 23, 42, 0.05);
    }
    .lab-secondary-card {
        padding: 0.82rem 0.9rem;
        margin-bottom: 0.72rem;
        border-radius: 16px;
        background: linear-gradient(180deg, #fbfdff 0%, #f7fbff 100%);
        border: 1px solid #e1ebf7;
    }
    .lab-panel-label {
        font-size: 0.86rem;
        font-weight: 700;
        margin-bottom: 0.38rem;
        color: #10233f;
        letter-spacing: 0.01em;
    }
    .lab-panel-note {
        margin: 0 0 0.45rem 0;
        color: #60748c;
        font-size: 0.79rem;
        line-height: 1.55;
    }
    .lab-status-pills {
        display: flex;
        flex-wrap: wrap;
        gap: 0.45rem;
        margin: 0 0 0.7rem 0;
    }
    .lab-status-pill {
        display: inline-flex;
        align-items: center;
        gap: 0.28rem;
        padding: 0.3rem 0.66rem;
        border-radius: 999px;
        background: linear-gradient(180deg, #f3f8ff 0%, #edf5ff 100%);
        border: 1px solid #d7e7fb;
        color: #33567e;
        font-size: 0.76rem;
        line-height: 1;
        box-shadow: inset 0 1px 0 rgba(255,255,255,0.82);
    }
    .lab-status-pill strong {
        color: #11345d;
        font-weight: 700;
    }
    .lab-quiet-note {
        color: #6a7c92;
        font-size: 0.78rem;
        line-height: 1.5;
    }
    .lab-secondary-card div[data-testid="stExpander"] {
        border: 1px solid #e4edf8;
        border-radius: 14px;
        background: rgba(255,255,255,0.74);
        margin-bottom: 0.52rem;
        overflow: hidden;
    }
    .lab-secondary-card summary {
        font-weight: 600;
    }
    .result-toolbar {
        margin: 0.55rem 0 0.15rem 0;
        padding: 0.68rem 0.8rem;
        background: linear-gradient(180deg, #f8fbff 0%, #f4f8ff 100%);
        border: 1px solid #dce8f6;
        border-radius: 12px;
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
MODIFICATION_FUNCTIONS = {"📋 Redlining修订", "✨ 表达润色", "🧩 文本降重", "🤖 去AI味 (Humanizer)", "🎯 精修模式"}
SHADOW_FUNCTIONS = {"✍️ 逐段起草", "💡 研究想法构思", "📄 节节头脑风暴", "✍️ 影子写作"}
VIZ_SCENE_PROMPTS = {
    "机理示意图": "show the core scientific mechanism, structural relationships, interaction pathways, key local zoom-ins, and cause-effect logic",
    "结构表征图": "focus on morphology, layered architecture, porosity, interfaces, crystallographic or nanoscale structural features",
    "合成流程图": "show a step-by-step synthesis or fabrication workflow with clear transitions between precursors, intermediates, and final products",
    "性能对比图解": "highlight comparative advantages, structure-property relationships, and visual evidence supporting performance claims",
    "界面反应图": "emphasize interface structure, interfacial reactions, boundary layers, and coupled transport or conversion processes",
    "传输路径图": "emphasize ion, electron, heat, mass, or charge transport pathways with directional clarity",
    "实验流程图": "show instruments, process stages, sample preparation, testing order, and workflow logic clearly",
    "逻辑框架图": "present conceptual nodes, relationships, hierarchy, and scientific reasoning in a structured visual map",
}
VIZ_STYLE_PROMPTS = {
    "科研 3D 渲染": "scientific 3D rendering, realistic material texture, layered depth, polished academic illustration",
    "BioRender 风格": "clean biomedical-style schematic, crisp icons, simplified but professional scientific composition",
    "Nature 图形摘要风格": "high-end journal graphical abstract style, concise layout, premium composition, strong clarity",
    "简约矢量风格": "minimal vector scientific illustration, clean edges, reduced clutter, publication-ready simplicity",
    "扁平化信息图风格": "flat infographic style, clear hierarchy, simplified geometry, explanatory visual balance",
    "深色高级感风格": "dark premium scientific visual style, cinematic contrast, glowing highlights, elegant composition",
    "高对比演示风格": "presentation-oriented, high contrast, visually striking, immediately readable on slides",
}
VIZ_DEFAULT_STATE = {
    "material_name": "",
    "component_tags_text": "",
    "component_tags": [],
    "usage": "论文主图",
    "scene": "机理示意图",
    "style": "科研 3D 渲染",
    "emphasis_points_text": "",
    "structure_notes": "",
    "description": "",
    "label_mode": "无文字版",
    "label_language": "中文",
    "label_content_mode": "自动生成后编辑（推荐）",
    "label_terms_text": "",
    "auto_generated_labels": [],
    "final_label_terms": [],
    "final_language_rule": "禁止任何文字标签",
    "info_density": "中",
    "aspect_ratio": "1:1",
    "logic_summary": "",
    "skill_prompt": "",
    "optimized_prompt": "",
    "hard_constraints": "",
    "base_prompt": "",
    "prompt_with_labels": "",
    "prompt_without_labels": "",
    "compact_prompt": "",
    "expanded_prompt": "",
    "final_prompt": "",
    "current_image_url": "",
    "current_image_bytes": b"",
    "current_result_id": None,
    "current_parent_id": None,
    "iteration_instruction": "",
    "local_area_hint": "",
    "inpaint_strength": 0.35,
    "preserve_composition": True,
    "edit_mode": "局部进化",
    "iteration_mode": "text-to-image",
    "seed_value": None,
    "current_seed": None,
    "current_image_seed": None,
    "style_reference_image": None,
    "style_reference_name": "",
    "style_strength": 0.6,
    "history": [],
    "active_history_id": None,
    "generation_counter": 0,
    "is_generating": False,
    "last_error": "",
}
PPT_PAGE_TYPES = ["封面页", "背景页", "问题定义页", "方法页", "流程页", "结果页", "对比页", "机理页", "结论页", "展望页"]
PPT_STYLE_MODES = ["academic-paperskills", "journal-briefing", "defense-clean"]
PPT_DEFAULT_LAYOUTS = {
    "左文右图": {"title": "上方标题 + 左文右图", "hint": "适合背景、方法、结果说明"},
    "上图下文": {"title": "上图下文", "hint": "适合流程、结果展示、结构说明"},
    "双栏对比": {"title": "双栏对比", "hint": "适合 before/after、模型对比、实验对比"},
    "大图重点说明": {"title": "大图重点说明", "hint": "适合单图强调、机理图配说明"},
    "结果+结论": {"title": "结果 + 结论", "hint": "适合实验结果与一句话结论并置"},
    "机理图说明页": {"title": "机理图说明页", "hint": "适合科研绘图与关键 bullet 配合"},
}
PPT_PAGE_TYPE_LAYOUTS = {
    "封面页": "大图重点说明",
    "背景页": "左文右图",
    "问题定义页": "左文右图",
    "方法页": "左文右图",
    "流程页": "上图下文",
    "结果页": "结果+结论",
    "对比页": "双栏对比",
    "机理页": "机理图说明页",
    "结论页": "结果+结论",
    "展望页": "左文右图",
}
PPT_DEFAULT_STATE = {
    "input_mode": "纯文字",
    "page_title": "",
    "page_type": "结果页",
    "page_goal": "",
    "raw_text": "",
    "extra_notes": "",
    "uploaded_source_name": "",
    "uploaded_source_text": "",
    "uploaded_images": [],
    "reference_images": [],
    "template_name": "默认科研模板",
    "template_analysis": {},
    "template_file_name": "",
    "style_mode": "academic-paperskills",
    "text_simplify_level": "中",
    "info_density": "中",
    "keep_original_images": True,
    "allow_external_support": False,
    "generate_aux_figure": False,
    "apply_template_layout": True,
    "processed_text_variants": {},
    "selected_layout": "结果+结论",
    "layout_analysis": {},
    "current_slide_markdown": "",
    "current_slide_struct": {},
    "current_slide_preview_html": "",
    "current_slide_snapshot": {},
    "current_slide_id": None,
    "current_parent_id": None,
    "micro_tune_request": "",
    "assembly_pages": [],
    "history": [],
    "active_history_id": None,
    "generation_counter": 0,
    "last_error": "",
}

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
    "viz_lab": deepcopy(VIZ_DEFAULT_STATE),
    "ppt_lab": deepcopy(PPT_DEFAULT_STATE),
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


def init_runtime_model_state() -> None:
    text_options = list(CLAUDE_MODEL_OPTIONS)
    image_options = list(GEMINI_MODEL_OPTIONS)
    if CLAUDE_MODEL not in text_options:
        text_options.insert(0, CLAUDE_MODEL)
    if GEMINI_MODEL not in image_options:
        image_options.insert(0, GEMINI_MODEL)
    st.session_state.setdefault("runtime_text_model_options", text_options)
    st.session_state.setdefault("runtime_image_model_options", image_options)
    st.session_state.setdefault("runtime_text_model", CLAUDE_MODEL)
    st.session_state.setdefault("runtime_image_model", GEMINI_MODEL)


def get_active_text_model() -> str:
    init_runtime_model_state()
    return st.session_state.get("runtime_text_model", CLAUDE_MODEL)


def get_active_image_model() -> str:
    init_runtime_model_state()
    return st.session_state.get("runtime_image_model", GEMINI_MODEL)


def init_viz_state() -> dict:
    init_runtime_model_state()
    current = st.session_state.get("viz_lab")
    if not isinstance(current, dict):
        current = deepcopy(VIZ_DEFAULT_STATE)
        st.session_state["viz_lab"] = current
    for key, default in VIZ_DEFAULT_STATE.items():
        current.setdefault(key, deepcopy(default))
    return current


def init_ppt_state() -> dict:
    init_runtime_model_state()
    current = st.session_state.get("ppt_lab")
    if not isinstance(current, dict):
        current = deepcopy(PPT_DEFAULT_STATE)
        st.session_state["ppt_lab"] = current
    for key, default in PPT_DEFAULT_STATE.items():
        current.setdefault(key, deepcopy(default))
    return current


def detect_ppt_input_mode(ppt_state: dict) -> str:
    if ppt_state.get("template_file_name"):
        return "PPT模板"
    if ppt_state.get("uploaded_source_name", "").lower().endswith(".pdf"):
        return "PDF"
    if ppt_state.get("uploaded_source_name", "").lower().endswith(".docx"):
        return "Word"
    if ppt_state.get("uploaded_source_text"):
        return "文档"
    if ppt_state.get("uploaded_images"):
        return "图片"
    return "纯文字"


def collect_uploaded_images(uploaded_files) -> List[dict]:
    images: List[dict] = []
    for file in uploaded_files or []:
        try:
            file_bytes = file.getvalue()
        except Exception:
            file_bytes = file.read()
        images.append({
            "name": file.name,
            "mime": getattr(file, "type", "image/png") or "image/png",
            "bytes": file_bytes,
        })
    return images


def collect_uploaded_image(uploaded_file) -> dict | None:
    if uploaded_file is None:
        return None
    try:
        file_bytes = uploaded_file.getvalue()
    except Exception:
        file_bytes = uploaded_file.read()
    return {
        "name": uploaded_file.name,
        "mime": getattr(uploaded_file, "type", "image/png") or "image/png",
        "bytes": file_bytes,
    }


def analyze_ppt_template(template_bytes: bytes, filename: str) -> dict:
    if not PPTX_AVAILABLE:
        return {
            "filename": filename,
            "error": "当前环境未安装 python-pptx，暂时无法分析模板。",
            "title_zone": "待分析",
            "content_layout": "待分析",
            "color_style": "待分析",
            "font_hierarchy": "待分析",
            "whitespace_style": "待分析",
        }
    try:
        prs = Presentation(BytesIO(template_bytes))
        width = round(prs.slide_width / 914400, 2)
        height = round(prs.slide_height / 914400, 2)
        sample_texts: List[str] = []
        title_like = 0
        text_like = 0
        for slide in prs.slides[:3]:
            for shape in slide.shapes:
                text = getattr(shape, "text", "").strip()
                if text:
                    sample_texts.append(text[:60])
                if getattr(shape, "has_text_frame", False):
                    if shape.height and shape.height < Inches(1.2):
                        title_like += 1
                    else:
                        text_like += 1
        palette = "科研蓝灰" if any("blue" in t.lower() for t in sample_texts) else "模板原生配色"
        return {
            "filename": filename,
            "slide_count": len(prs.slides),
            "page_size": f"{width} × {height} in",
            "title_zone": "上方标题区" if title_like >= text_like else "标题区不明显",
            "content_layout": "图文混排" if text_like else "大图主导",
            "color_style": palette,
            "font_hierarchy": "存在标题/正文层级" if title_like else "层级需手动判断",
            "whitespace_style": "中等留白",
            "sample_texts": sample_texts[:5],
            "has_slide_master": bool(getattr(prs, "slide_masters", [])),
        }
    except Exception as exc:
        return {
            "filename": filename,
            "error": str(exc),
            "title_zone": "待分析",
            "content_layout": "待分析",
            "color_style": "模板原生配色",
            "font_hierarchy": "待分析",
            "whitespace_style": "待分析",
        }


def normalize_ppt_input_sources(
    source_file,
    raw_text: str,
    uploaded_images: List[dict],
    reference_images: List[dict],
    template_file,
) -> dict:
    source_name = ""
    source_text = ""
    if source_file is not None:
        source_name = source_file.name
        source_bytes = source_file.getvalue()
        source_text = extract_text(source_bytes, source_name)
    template_name = ""
    template_analysis: dict = {}
    if template_file is not None:
        template_name = template_file.name
        template_analysis = analyze_ppt_template(template_file.getvalue(), template_name)
    return {
        "uploaded_source_name": source_name,
        "uploaded_source_text": source_text,
        "raw_text": raw_text,
        "uploaded_images": uploaded_images,
        "reference_images": reference_images,
        "template_file_name": template_name,
        "template_analysis": template_analysis,
        "input_mode": "纯文字",
    }


def build_ppt_page_context(ppt_state: dict) -> str:
    text_source = (ppt_state.get("raw_text") or "").strip()
    doc_source = (ppt_state.get("uploaded_source_text") or "").strip()
    merged = text_source or doc_source
    image_names = [item.get("name", "") for item in ppt_state.get("uploaded_images", [])]
    ref_names = [item.get("name", "") for item in ppt_state.get("reference_images", [])]
    return (
        f"页面标题：{ppt_state.get('page_title') or '未命名当前页'}\n"
        f"页面类型：{ppt_state.get('page_type', '结果页')}\n"
        f"本页目标：{ppt_state.get('page_goal') or '未填写'}\n"
        f"风格模式：{ppt_state.get('style_mode', 'academic-paperskills')}\n"
        f"文字简化强度：{ppt_state.get('text_simplify_level', '中')}\n"
        f"信息密度：{ppt_state.get('info_density', '中')}\n"
        f"输入来源：{detect_ppt_input_mode(ppt_state)}\n"
        f"原始文本：\n{merged[:6000]}\n\n"
        f"补充说明：\n{(ppt_state.get('extra_notes') or '').strip()}\n\n"
        f"上传图片：{', '.join(image_names) if image_names else '无'}\n"
        f"参考图：{', '.join(ref_names) if ref_names else '无'}"
    ).strip()


def create_local_ppt_variants(ppt_state: dict) -> dict:
    source = (ppt_state.get("raw_text") or ppt_state.get("uploaded_source_text") or "").strip()
    page_title = (ppt_state.get("page_title") or "").strip()
    if not source:
        return {
            "原文": "",
            "精简版": "",
            "标题版": page_title,
            "要点版": "",
            "结论先行版": "",
        }
    normalized = re.sub(r"\s+", " ", source).strip()
    sentences = [item.strip(" -•·\t") for item in re.split(r"(?<=[。！？.!?])\s+|\n+", source) if item.strip()]
    bullets = [f"- {item[:120].strip()}" for item in sentences[:5]]
    headline_source = page_title or (sentences[0][:26] if sentences else normalized[:26])
    conclusion = sentences[0] if sentences else normalized[:120]
    compact = " ".join(sentences[:3])[:360] if sentences else normalized[:360]
    return {
        "原文": source[:4000],
        "精简版": compact,
        "标题版": headline_source,
        "要点版": "\n".join(bullets),
        "结论先行版": f"结论：{conclusion[:120]}\n支撑信息：{'；'.join(sentences[1:4])[:220]}",
    }


def generate_ppt_text_variants(ppt_state: dict) -> dict:
    context = build_ppt_page_context(ppt_state)
    if not context.strip():
        return create_local_ppt_variants(ppt_state)
    prompt = f"""
你是科研汇报 PPT 单页工作台助手。只处理当前这一页，不要生成整套大纲。
请基于下面内容，输出 JSON 对象，必须包含 5 个键：原文、精简版、标题版、要点版、结论先行版。
要求：
1. 保持科研汇报口吻，中文输出。
2. 精简版适合放入单页 PPT。
3. 标题版控制在 24 个字以内。
4. 要点版写成 3-5 条短 bullet，每条单独一行，以“- ”开头。
5. 只返回 JSON，不要加解释。

当前页上下文：
{context}
""".strip()
    try:
        response = call_api(prompt, timeout=180)
        match = re.search(r"\{[\s\S]*\}", response)
        if match:
            data = json.loads(match.group(0))
            result = create_local_ppt_variants(ppt_state)
            for key in result.keys():
                value = data.get(key, result[key])
                result[key] = value.strip() if isinstance(value, str) else result[key]
            return result
    except Exception:
        pass
    return create_local_ppt_variants(ppt_state)


def resolve_ppt_layout(ppt_state: dict) -> dict:
    selected = ppt_state.get("selected_layout") or PPT_PAGE_TYPE_LAYOUTS.get(ppt_state.get("page_type", "结果页"), "结果+结论")
    layout = deepcopy(PPT_DEFAULT_LAYOUTS.get(selected, PPT_DEFAULT_LAYOUTS["结果+结论"]))
    layout.update({
        "name": selected,
        "page_type": ppt_state.get("page_type", "结果页"),
        "style_mode": ppt_state.get("style_mode", "academic-paperskills"),
        "apply_template_layout": ppt_state.get("apply_template_layout", True),
    })
    return layout


def choose_ppt_body_text(ppt_state: dict) -> str:
    variants = ppt_state.get("processed_text_variants") or {}
    density = ppt_state.get("info_density", "中")
    if density == "低":
        return variants.get("结论先行版") or variants.get("精简版") or variants.get("原文", "")
    if density == "高":
        return variants.get("原文") or variants.get("精简版", "")
    return variants.get("要点版") or variants.get("精简版") or variants.get("原文", "")


def build_ppt_slide_spec(ppt_state: dict) -> dict:
    variants = ppt_state.get("processed_text_variants") or create_local_ppt_variants(ppt_state)
    layout = resolve_ppt_layout(ppt_state)
    title = (ppt_state.get("page_title") or variants.get("标题版") or "未命名单页").strip()
    body_text = choose_ppt_body_text(ppt_state).strip()
    tags = [
        ppt_state.get("page_type", "结果页"),
        layout.get("name", "结果+结论"),
        ppt_state.get("style_mode", "academic-paperskills"),
        detect_ppt_input_mode(ppt_state),
        "补图开启" if ppt_state.get("generate_aux_figure") else "不补图",
    ]
    return {
        "title": title,
        "subtitle": ppt_state.get("page_goal", "").strip(),
        "body": body_text,
        "bullets": [line[2:].strip() for line in body_text.splitlines() if line.strip().startswith("- ")][:5],
        "layout": layout,
        "page_type": ppt_state.get("page_type", "结果页"),
        "style_mode": ppt_state.get("style_mode", "academic-paperskills"),
        "template_analysis": deepcopy(ppt_state.get("template_analysis") or {}),
        "template_name": ppt_state.get("template_file_name") or ppt_state.get("template_name") or "默认科研模板",
        "input_mode": detect_ppt_input_mode(ppt_state),
        "tags": tags,
        "aux_figure_prompt": "",
        "aux_figure_url": "",
        "aux_figure_bytes": b"",
        "notes": (ppt_state.get("extra_notes") or "").strip(),
        "source_images": deepcopy(ppt_state.get("uploaded_images") or []),
        "reference_images": deepcopy(ppt_state.get("reference_images") or []),
    }


def build_ppt_micro_tune_prompt(ppt_state: dict) -> str:
    spec = ppt_state.get("current_slide_struct") or {}
    return (
        f"当前页标题：{spec.get('title', '')}\n"
        f"当前布局：{(spec.get('layout') or {}).get('name', '')}\n"
        f"当前文案：\n{spec.get('body', '')}\n\n"
        f"用户微调要求：\n{ppt_state.get('micro_tune_request', '').strip()}"
    ).strip()


def generate_ppt_aux_figure(ppt_state: dict, spec: dict) -> dict:
    prompt = (
        f"Create a scientific presentation figure for a single PPT page. "
        f"Page type: {ppt_state.get('page_type', '结果页')}. "
        f"Layout: {(spec.get('layout') or {}).get('name', '结果+结论')}. "
        f"Theme: {spec.get('title', '')}. "
        f"Body: {spec.get('body', '')[:600]}. "
        f"Style: clean academic, blue-white-gray, high readability, presentation-ready."
    )
    try:
        image_url, image_bytes = call_gemini_image(prompt, aspect_ratio="16:9")
        return {
            "aux_figure_prompt": prompt,
            "aux_figure_url": image_url,
            "aux_figure_bytes": image_bytes,
            "error": "",
        }
    except Exception as exc:
        return {
            "aux_figure_prompt": prompt,
            "aux_figure_url": "",
            "aux_figure_bytes": b"",
            "error": str(exc),
        }


def render_ppt_slide_preview(spec: dict) -> str:
    body_html = "".join(f"<li>{escape(item)}</li>" for item in spec.get("bullets") or [])
    if not body_html:
        body_html = f"<p>{escape(spec.get('body', '')).replace(chr(10), '<br>')}</p>"
    tags_html = "".join(f"<span class='workbench-chip'>{escape(tag)}</span>" for tag in spec.get("tags", []))
    aux_html = ""
    if spec.get("aux_figure_url"):
        aux_html = f"<img src='{spec['aux_figure_url']}' style='width:100%; border-radius:14px; border:1px solid #dbe7f5; object-fit:cover;'/>"
    elif spec.get("source_images"):
        aux_html = f"<div style='padding:1rem;border:1px dashed #93c5fd;border-radius:14px;background:#eff6ff;color:#1e40af;'>已载入图片素材：{escape(spec['source_images'][0].get('name', 'image'))}</div>"
    else:
        aux_html = "<div style='padding:1rem;border:1px dashed #cbd5e1;border-radius:14px;background:#f8fafc;color:#475569;'>图片区占位：可上传原图 / 参考图 / 生成辅助图</div>"
    return f"""
    <div style="background:linear-gradient(180deg,#ffffff 0%,#f8fbff 100%);border:1px solid #dbe7f5;border-radius:18px;padding:1rem;box-shadow:0 10px 30px rgba(15,23,42,0.06);">
        <div style="display:flex;justify-content:space-between;gap:1rem;align-items:flex-start;margin-bottom:0.7rem;">
            <div>
                <div style="font-size:1.45rem;font-weight:700;color:#0f172a;line-height:1.25;">{escape(spec.get('title', '未命名单页'))}</div>
                <div style="font-size:0.92rem;color:#475569;margin-top:0.25rem;">{escape(spec.get('subtitle', ''))}</div>
            </div>
            <div style="display:flex;gap:0.35rem;flex-wrap:wrap;justify-content:flex-end;">{tags_html}</div>
        </div>
        <div style="display:grid;grid-template-columns:1.1fr 0.9fr;gap:1rem;align-items:start;">
            <div style="background:#ffffff;border-radius:14px;padding:0.9rem;border:1px solid #e2e8f0;color:#0f172a;min-height:260px;">
                <ul style="margin:0;padding-left:1.1rem;line-height:1.7;color:#334155;">{body_html}</ul>
            </div>
            <div>{aux_html}</div>
        </div>
    </div>
    """.strip()


def sanitize_for_json(value):
    if isinstance(value, bytes):
        return {
            "type": "bytes",
            "size": len(value),
        }
    if isinstance(value, dict):
        return {key: sanitize_for_json(val) for key, val in value.items()}
    if isinstance(value, list):
        return [sanitize_for_json(item) for item in value]
    if isinstance(value, tuple):
        return [sanitize_for_json(item) for item in value]
    return value


def append_ppt_history_entry(ppt_state: dict) -> None:
    spec = deepcopy(ppt_state.get("current_slide_struct") or {})
    if not spec:
        return
    entry_id = ppt_state.get("current_slide_id") or f"ppt-{ppt_state.get('generation_counter', 0)}"
    entry = {
        "id": entry_id,
        "parent_id": ppt_state.get("current_parent_id"),
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "page_title": ppt_state.get("page_title", ""),
        "page_type": ppt_state.get("page_type", "结果页"),
        "page_goal": ppt_state.get("page_goal", ""),
        "raw_text": ppt_state.get("raw_text", ""),
        "extra_notes": ppt_state.get("extra_notes", ""),
        "uploaded_source_name": ppt_state.get("uploaded_source_name", ""),
        "uploaded_source_text": ppt_state.get("uploaded_source_text", ""),
        "uploaded_images": deepcopy(ppt_state.get("uploaded_images", [])),
        "reference_images": deepcopy(ppt_state.get("reference_images", [])),
        "template_name": ppt_state.get("template_name", "默认科研模板"),
        "template_analysis": deepcopy(ppt_state.get("template_analysis", {})),
        "template_file_name": ppt_state.get("template_file_name", ""),
        "style_mode": ppt_state.get("style_mode", "academic-paperskills"),
        "text_simplify_level": ppt_state.get("text_simplify_level", "中"),
        "info_density": ppt_state.get("info_density", "中"),
        "keep_original_images": ppt_state.get("keep_original_images", True),
        "allow_external_support": ppt_state.get("allow_external_support", False),
        "generate_aux_figure": ppt_state.get("generate_aux_figure", False),
        "apply_template_layout": ppt_state.get("apply_template_layout", True),
        "processed_text_variants": deepcopy(ppt_state.get("processed_text_variants", {})),
        "selected_layout": ppt_state.get("selected_layout", "结果+结论"),
        "layout_analysis": deepcopy(ppt_state.get("layout_analysis", {})),
        "current_slide_markdown": ppt_state.get("current_slide_markdown", ""),
        "current_slide_struct": spec,
        "current_slide_preview_html": ppt_state.get("current_slide_preview_html", ""),
        "current_slide_snapshot": deepcopy(ppt_state.get("current_slide_snapshot", {})),
        "micro_tune_request": ppt_state.get("micro_tune_request", ""),
    }
    ppt_state["history"] = [item for item in ppt_state.get("history", []) if item.get("id") != entry_id]
    ppt_state["history"].insert(0, entry)
    ppt_state["history"] = ppt_state["history"][:30]
    ppt_state["active_history_id"] = entry_id


def restore_ppt_history_entry(ppt_state: dict, entry_id: str) -> None:
    for entry in ppt_state.get("history", []):
        if entry.get("id") != entry_id:
            continue
        for key in PPT_DEFAULT_STATE.keys():
            if key in ["assembly_pages", "history", "active_history_id", "generation_counter", "last_error"]:
                continue
            ppt_state[key] = deepcopy(entry.get(key, PPT_DEFAULT_STATE.get(key)))
        ppt_state["current_slide_id"] = entry.get("id")
        ppt_state["current_parent_id"] = entry.get("parent_id")
        ppt_state["active_history_id"] = entry.get("id")
        return


def add_current_slide_to_assembly(ppt_state: dict) -> None:
    spec = deepcopy(ppt_state.get("current_slide_struct") or {})
    if not spec:
        return
    slide_id = ppt_state.get("current_slide_id") or f"ppt-{ppt_state.get('generation_counter', 0)}"
    assembly_entry = {
        "id": slide_id,
        "page_title": spec.get("title", "未命名单页"),
        "page_type": spec.get("page_type", ppt_state.get("page_type", "结果页")),
        "layout_name": (spec.get("layout") or {}).get("name", ppt_state.get("selected_layout", "结果+结论")),
        "preview_html": ppt_state.get("current_slide_preview_html", ""),
        "slide_struct": spec,
        "markdown": ppt_state.get("current_slide_markdown", ""),
        "snapshot": deepcopy(ppt_state.get("current_slide_snapshot", {})),
    }
    ppt_state["assembly_pages"].append(assembly_entry)


def duplicate_assembly_slide(ppt_state: dict, slide_id: str) -> None:
    for index, slide in enumerate(ppt_state.get("assembly_pages", [])):
        if slide.get("id") != slide_id:
            continue
        duplicated = deepcopy(slide)
        duplicated["id"] = f"{slide_id}-copy-{len(ppt_state['assembly_pages']) + 1}"
        duplicated["page_title"] = f"{slide.get('page_title', '页面')}（副本）"
        ppt_state["assembly_pages"].insert(index + 1, duplicated)
        return


def remove_assembly_slide(ppt_state: dict, slide_id: str) -> None:
    ppt_state["assembly_pages"] = [slide for slide in ppt_state.get("assembly_pages", []) if slide.get("id") != slide_id]


def move_assembly_slide(ppt_state: dict, slide_id: str, direction: int) -> None:
    slides = ppt_state.get("assembly_pages", [])
    for index, slide in enumerate(slides):
        if slide.get("id") != slide_id:
            continue
        new_index = index + direction
        if 0 <= new_index < len(slides):
            slides[index], slides[new_index] = slides[new_index], slides[index]
        return


def export_pptx_deck(ppt_state: dict) -> bytes:
    if not PPTX_AVAILABLE:
        raise RuntimeError("当前环境未安装 python-pptx，暂时无法导出 PPTX。请先安装 requirements.txt 里的依赖。")
    prs = Presentation()
    for slide_info in ppt_state.get("assembly_pages", []):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        spec = slide_info.get("slide_struct") or {}
        title_box = slide.shapes.add_textbox(Inches(0.55), Inches(0.4), Inches(8.2), Inches(0.9))
        title_tf = title_box.text_frame
        title_tf.word_wrap = True
        title_p = title_tf.paragraphs[0]
        title_p.text = spec.get("title", slide_info.get("page_title", "未命名单页"))
        title_p.font.size = PptxPt(26)
        title_p.font.bold = True

        sub_box = slide.shapes.add_textbox(Inches(0.6), Inches(1.15), Inches(7.8), Inches(0.55))
        sub_tf = sub_box.text_frame
        sub_tf.paragraphs[0].text = spec.get("subtitle", "")
        sub_tf.paragraphs[0].font.size = PptxPt(12)

        body_box = slide.shapes.add_textbox(Inches(0.7), Inches(1.8), Inches(5.0), Inches(4.8))
        body_tf = body_box.text_frame
        body_tf.word_wrap = True
        bullets = spec.get("bullets") or []
        if bullets:
            first = body_tf.paragraphs[0]
            first.text = bullets[0]
            first.font.size = PptxPt(18)
            for item in bullets[1:5]:
                p = body_tf.add_paragraph()
                p.text = item
                p.font.size = PptxPt(18)
                p.level = 0
        else:
            body_tf.paragraphs[0].text = spec.get("body", "")[:900]
            body_tf.paragraphs[0].font.size = PptxPt(16)

        if spec.get("aux_figure_bytes"):
            slide.shapes.add_picture(BytesIO(spec["aux_figure_bytes"]), Inches(6.0), Inches(1.7), width=Inches(3.1), height=Inches(3.8))
        else:
            shape = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE, Inches(6.0), Inches(1.7), Inches(3.1), Inches(3.8))
            shape.text_frame.text = "图像区\n上传原图 / 参考图 / 辅助图"
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor(239, 246, 255)
            shape.line.color.rgb = RGBColor(147, 197, 253)

        footer_box = slide.shapes.add_textbox(Inches(0.65), Inches(6.8), Inches(8.2), Inches(0.3))
        footer_tf = footer_box.text_frame
        footer_tf.paragraphs[0].text = f"{slide_info.get('page_type', '结果页')} · {slide_info.get('layout_name', '结果+结论')}"
        footer_tf.paragraphs[0].font.size = PptxPt(10)
    if not prs.slides:
        prs.slides.add_slide(prs.slide_layouts[6])
    buffer = BytesIO()
    prs.save(buffer)
    buffer.seek(0)
    return buffer.read()


def parse_tag_text(raw_text: str) -> List[str]:
    parts = re.split(r"[\n,;，；]+", raw_text or "")
    merged: List[str] = []
    for part in parts:
        item = part.strip()
        if item and item not in merged:
            merged.append(item)
    return merged


def get_viz_components(viz_state: dict) -> List[str]:
    explicit = list(viz_state.get("component_tags", []))
    typed = parse_tag_text(viz_state.get("component_tags_text", ""))
    merged: List[str] = []
    for item in explicit + typed:
        if item not in merged:
            merged.append(item)
    viz_state["component_tags"] = merged
    return merged


def build_viz_label_terms(viz_state: dict) -> List[str]:
    label_mode = viz_state.get("label_mode", "无文字版")
    if label_mode != "有文字版":
        return []

    manual_terms = [line.strip() for line in re.split(r"[\n,，;；]+", viz_state.get("label_terms_text", "")) if line.strip()]
    auto_terms = [str(term).strip() for term in viz_state.get("auto_generated_labels", []) if str(term).strip()]
    content_mode = viz_state.get("label_content_mode", "自动生成后编辑（推荐）")

    if content_mode == "手动填写":
        return manual_terms
    if content_mode == "仅自动生成":
        return auto_terms

    merged: List[str] = []
    for term in manual_terms + auto_terms:
        if term and term not in merged:
            merged.append(term)
    return merged


def infer_viz_final_language_rule(viz_state: dict) -> str:
    label_mode = viz_state.get("label_mode", "无文字版")
    if label_mode != "有文字版":
        return "禁止任何文字标签"
    return "仅允许中文标签" if viz_state.get("label_language", "中文") == "中文" else "仅允许英文标签"


def suggest_viz_label_terms(viz_state: dict) -> List[str]:
    label_mode = viz_state.get("label_mode", "无文字版")
    if label_mode != "有文字版":
        return []

    material_text = viz_state.get("material_name", "").strip()
    components = get_viz_components(viz_state)
    scene = viz_state.get("scene", "").strip()
    usage = viz_state.get("usage", "").strip()
    emphasis_parts = [part.strip() for part in re.split(r"[\n,，;；]+", viz_state.get("emphasis_points_text", "")) if part.strip()]
    description = viz_state.get("description", "").strip()
    structure_notes = viz_state.get("structure_notes", "").strip()
    language = viz_state.get("label_language", "中文")

    candidates: List[str] = []
    if language == "中文":
        if material_text:
            candidates.append(material_text)
        candidates.extend(components[:2])
        candidates.extend(emphasis_parts[:3])
        for text in [scene, usage, structure_notes]:
            if text:
                candidates.append(text)
        if "离子" in description and "迁移路径" not in candidates:
            candidates.append("离子迁移路径")
        if ("界面" in description or "异质结" in description) and "界面反应区" not in candidates:
            candidates.append("界面反应区")
        if ("结构" in description or "层" in description) and "结构示意" not in candidates:
            candidates.append("结构示意")
    else:
        if material_text:
            candidates.append(material_text)
        candidates.extend(components[:2])
        candidates.extend(emphasis_parts[:3])
        scene_map = {
            "机理示意图": "Mechanism",
            "结构表征图": "Structure",
            "实验流程图": "Workflow",
            "对比结果图": "Comparison",
            "逻辑框架图": "Framework",
        }
        usage_map = {
            "论文主图": "Main Figure",
            "论文 TOC 图": "Graphical Abstract",
            "汇报展示": "Presentation Figure",
            "基金申请": "Proposal Figure",
            "教学示意": "Teaching Diagram",
            "社媒科普": "Science Visual",
        }
        for mapped in [scene_map.get(scene, scene), usage_map.get(usage, usage)]:
            if mapped:
                candidates.append(mapped)
        if "ion" in description.lower() or "离子" in description:
            candidates.append("Ion transport")
        if "interface" in description.lower() or "界面" in description:
            candidates.append("Interface region")
        if "structure" in description.lower() or "结构" in description:
            candidates.append("Structure")

    deduped: List[str] = []
    for term in candidates:
        clean = str(term).strip()
        if not clean or clean in deduped:
            continue
        deduped.append(clean)
        if len(deduped) >= 6:
            break
    return deduped


def build_viz_skill_prompt(viz_state: dict) -> str:
    material_text = viz_state.get("material_name", "").strip() or "unspecified material system"
    components = get_viz_components(viz_state)
    components_text = ", ".join(components) if components else "key material components"
    scene_text = VIZ_SCENE_PROMPTS.get(viz_state.get("scene", ""), viz_state.get("scene", "scientific illustration"))
    usage_text = viz_state.get("usage", "论文主图")
    emphasis_text = viz_state.get("emphasis_points_text", "").strip() or "highlight the core scientific message"
    structure_notes = viz_state.get("structure_notes", "").strip() or "maintain accurate structural relationships"
    description_text = viz_state.get("description", "").strip() or "show the target scientific content clearly"
    skill_hint = LOCAL_SKILLS.get("README", {}).get("description", "")
    skill_prefix = ""
    if skill_hint:
        skill_prefix = f"Skill guidance: {skill_hint[:120].strip()}. "
    return (
        f"{skill_prefix}Create a scientific figure for Gemini image generation. "
        f"Material system: {material_text}. "
        f"Key components: {components_text}. "
        f"Scene task: {scene_text}. "
        f"Usage context: {usage_text}. "
        f"Emphasis points: {emphasis_text}. "
        f"Structural notes: {structure_notes}. "
        f"Scientific description: {description_text}."
    )


def build_viz_optimized_prompt(viz_state: dict, skill_prompt: str) -> str:
    style_text = VIZ_STYLE_PROMPTS.get(viz_state.get("style", ""), viz_state.get("style", "academic illustration"))
    info_density_map = {"低": "low", "中": "medium", "高": "high"}
    info_density = info_density_map.get(viz_state.get("info_density", "中"), "medium")
    aspect_ratio = viz_state.get("aspect_ratio", "1:1")
    return (
        f"Preserve the original scientific meaning of this prompt and optimize it for Gemini image generation: {skill_prompt} "
        f"Enhance the visual clarity, composition hierarchy, and publication-grade storytelling without changing the material system or scientific relationships. "
        f"Visual style target: {style_text}. Information density: {info_density}. Aspect ratio: {aspect_ratio}. "
        "Favor clean academic composition, clear focal regions, balanced spacing, crisp arrows/callouts when needed, and accurate scientific structure depiction."
    )


def normalize_viz_label_terms(label_terms: List[str], label_language: str) -> List[str]:
    normalized: List[str] = []
    for term in label_terms:
        clean = str(term).strip()
        if not clean:
            continue
        if label_language == "中文":
            if not re.search(r"[\u4e00-\u9fff]", clean):
                continue
        else:
            if re.search(r"[\u4e00-\u9fff]", clean):
                continue
        if clean not in normalized:
            normalized.append(clean)
    return normalized


def build_viz_hard_constraints(viz_state: dict, label_terms: List[str]) -> str:
    label_mode = viz_state.get("label_mode", "无文字版")
    label_language = viz_state.get("label_language", "中文")
    aspect_ratio = viz_state.get("aspect_ratio", "1:1")
    ratio_constraint = f"Strict layout constraint: keep the composition suitable for a {aspect_ratio} figure ratio."

    if label_mode != "有文字版":
        return (
            f"{ratio_constraint} "
            "Absolute text prohibition: the image must contain zero visible text elements. "
            "Do not render any labels, legends, annotations, callouts, axis text, titles, numbers, Chinese characters, English words, abbreviations, or caption-like markings anywhere in the figure. "
            "Communicate only through shapes, arrows, icons, regions, and purely visual scientific elements."
        )

    normalized_terms = normalize_viz_label_terms(label_terms, label_language)
    joined_terms = "; ".join(normalized_terms)
    if label_language == "中文":
        label_constraint = (
            "Absolute language constraint: every visible label, annotation, legend, and callout must be written in Simplified Chinese only. "
            "Do not use any English words, Latin letters, pinyin, or mixed-language text anywhere in visible annotations."
        )
    else:
        label_constraint = (
            "Absolute language constraint: every visible label, annotation, legend, and callout must be written in English only. "
            "Do not use any Chinese characters or mixed-language text anywhere in visible annotations."
        )

    terms_constraint = ""
    if joined_terms:
        terms_constraint = f" Visible labels should preferentially use only these approved terms: {joined_terms}. Do not invent unrelated labels."

    return f"{ratio_constraint} {label_constraint}{terms_constraint}"


def build_viz_logic_summary(viz_state: dict) -> str:
    material_text = viz_state.get("material_name", "").strip() or "未指定材料"
    components = get_viz_components(viz_state)
    components_text = "、".join(components) if components else "未指定组成"
    usage = viz_state.get("usage", "未指定用途")
    scene = viz_state.get("scene", "未指定场景")
    style = viz_state.get("style", "未指定风格")
    label_mode = viz_state.get("label_mode", "未指定标注模式")
    label_language = viz_state.get("label_language", "未指定标签语言") if label_mode == "有文字版" else "关闭"
    final_language_rule = infer_viz_final_language_rule(viz_state)
    label_content_mode = viz_state.get("label_content_mode", "关闭") if label_mode == "有文字版" else "关闭"
    final_labels = build_viz_label_terms(viz_state)
    labels_preview = "、".join(final_labels[:4]) if final_labels else "无"
    return (
        f"材料：{material_text}｜组成：{components_text}｜用途：{usage}｜场景：{scene}｜风格：{style}｜标注：{label_mode}"
        f"｜标签语言：{label_language}｜最终语言约束：{final_language_rule}｜标签来源：{label_content_mode}｜标签预览：{labels_preview}"
    )


def build_viz_prompt_bundle(viz_state: dict) -> dict:
    skill_prompt = build_viz_skill_prompt(viz_state)
    optimized_prompt = build_viz_optimized_prompt(viz_state, skill_prompt)
    auto_generated_labels = normalize_viz_label_terms(suggest_viz_label_terms(viz_state), viz_state.get("label_language", "中文"))
    label_terms = normalize_viz_label_terms(build_viz_label_terms({**viz_state, "auto_generated_labels": auto_generated_labels}), viz_state.get("label_language", "中文"))
    final_language_rule = infer_viz_final_language_rule(viz_state)
    hard_constraints = build_viz_hard_constraints(viz_state, label_terms)
    prompt_with_labels = f"{skill_prompt} {optimized_prompt} {hard_constraints}".strip()
    prompt_without_labels = f"{skill_prompt} {optimized_prompt} {build_viz_hard_constraints({**viz_state, 'label_mode': '无文字版'}, [])}".strip()
    final_prompt = prompt_with_labels if viz_state.get("label_mode", "无文字版") == "有文字版" else prompt_without_labels
    compact_prompt = f"{viz_state.get('style', '')}; {viz_state.get('material_name', '')}; {viz_state.get('scene', '')}; {viz_state.get('description', '')}; Gemini-ready scientific figure"
    return {
        "skill_prompt": skill_prompt,
        "optimized_prompt": optimized_prompt,
        "hard_constraints": hard_constraints,
        "auto_generated_labels": auto_generated_labels,
        "final_label_terms": label_terms,
        "final_language_rule": final_language_rule,
        "base_prompt": skill_prompt,
        "prompt_with_labels": prompt_with_labels,
        "prompt_without_labels": prompt_without_labels,
        "compact_prompt": compact_prompt,
        "expanded_prompt": optimized_prompt,
        "final_prompt": final_prompt,
    }


def build_viz_iteration_prompt(viz_state: dict) -> str:
    instruction = viz_state.get("iteration_instruction", "").strip()
    local_area = viz_state.get("local_area_hint", "").strip() or "the user-specified local area"
    optimized_prompt = viz_state.get("optimized_prompt", "").strip() or viz_state.get("expanded_prompt", "").strip()
    hard_constraints = viz_state.get("hard_constraints", "").strip()
    label_terms = viz_state.get("final_label_terms") or build_viz_label_terms(viz_state)
    label_terms_clause = f"Keep these label terms consistent when visible: {'; '.join(label_terms)}. " if label_terms else ""
    composition_rule = "Maintain the original composition, spatial arrangement, materials, lighting, and camera perspective. "
    if not viz_state.get("preserve_composition", True):
        composition_rule = "Preserve the main scientific subject and overall readability, but allow moderate local layout adjustment. "
    return (
        "Use the provided reference image as the primary visual basis. "
        "First understand the current figure, including composition, object placement, label distribution, and scientific relationships. "
        f"{optimized_prompt} {hard_constraints} "
        f"{composition_rule}"
        f"Only modify this path or local area: {local_area}. "
        f"Requested local change: {instruction}. "
        f"{label_terms_clause}"
        "Keep all other scientific relationships, textures, labels outside the target area, and non-target regions unchanged."
    ).strip()


def get_viz_image_bytes(image_url: str) -> bytes | None:
    if not image_url:
        return None
    if image_url.startswith("data:"):
        try:
            _, base64_data = image_url.split(",", 1)
            return base64.b64decode(base64_data)
        except Exception:
            return None
    try:
        resp = requests.get(image_url, timeout=30)
        resp.raise_for_status()
        return resp.content
    except Exception:
        return None


def render_viz_image_slot(image_url: str = "", empty_text: str = "尚未生成图片。", container_height: int = 380, image_max_height: int = 350) -> None:
    safe_text = escape(empty_text)
    if image_url:
        html = f"""
        <div class=\"viz-preview-shell\" style=\"height:{container_height}px;\">
            <div class=\"viz-preview-frame\">
                <img src=\"{image_url}\" style=\"max-height:100%; width:auto; object-fit:contain;\" alt=\"viz preview\" />
            </div>
        </div>
        """
    else:
        html = f"""
        <div class=\"viz-preview-shell\" style=\"height:{container_height}px;\">
            <div class=\"viz-preview-frame\">
                <div class=\"viz-preview-empty\">{safe_text}</div>
            </div>
        </div>
        """
    st.html(html)


def normalize_gemini_prompt(prompt: str) -> str:
    cleaned = (prompt or "").replace("\r", " ").replace("\n", " ")
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    cleaned = cleaned.replace("“", '"').replace("”", '"').replace("‘", "'").replace("’", "'")
    return cleaned


def _is_gpt_image_model(model_name: str) -> bool:
    return model_name.startswith("gpt-image")


def _call_gpt_image(
    prompt: str,
    aspect_ratio: str = "1:1",
    seed: int | None = None,
) -> tuple[str, bytes, int | None]:
    """GPT Image models use OpenAI-compatible /v1/images/generations endpoint."""
    active_image_model = get_active_image_model()
    size_map = {
        "1:1": "1024x1024",
        "16:9": "1792x1024",
        "9:16": "1024x1792",
        "4:3": "1024x1024",
        "3:4": "1024x1024",
    }
    size = size_map.get(aspect_ratio, "1024x1024")
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {GEMINI_API_KEY}",
    }
    payload = {
        "model": active_image_model,
        "prompt": prompt,
        "n": 1,
        "size": size,
        "response_format": "b64_json",
    }
    resp = requests.post(
        f"{GEMINI_BASE_URL}/v1/images/generations",
        json=payload,
        headers=headers,
        timeout=120,
    )
    resp.raise_for_status()
    data = resp.json()
    images = data.get("data") or []
    if not images:
        raise ValueError(f"GPT Image returned no images: {json.dumps(data, ensure_ascii=False)[:500]}")
    b64_data = images[0].get("b64_json", "")
    if not b64_data:
        raise ValueError(f"GPT Image missing b64_json: {json.dumps(images[0], ensure_ascii=False)[:300]}")
    image_bytes = base64.b64decode(b64_data)
    image_url = f"data:image/png;base64,{b64_data}"
    return image_url, image_bytes, seed


def call_gemini_image(
    prompt: str,
    aspect_ratio: str = "1:1",
    reference_image_bytes: bytes | None = None,
    reference_mime: str = "image/png",
    seed: int | None = None,
    strength: float | None = None,
    style_reference_image: dict | None = None,
    style_strength: float | None = None,
) -> tuple[str, bytes | None, int | None]:
    active_image_model = get_active_image_model()

    if _is_gpt_image_model(active_image_model):
        return _call_gpt_image(prompt, aspect_ratio=aspect_ratio, seed=seed)

    safe_prompt = normalize_gemini_prompt(prompt)
    parts = [{"text": safe_prompt}]
    if reference_image_bytes:
        parts.append({
            "inline_data": {
                "mime_type": reference_mime,
                "data": base64.b64encode(reference_image_bytes).decode("utf-8"),
            }
        })
    if style_reference_image and style_reference_image.get("bytes"):
        parts.append({
            "inline_data": {
                "mime_type": style_reference_image.get("mime", "image/png"),
                "data": base64.b64encode(style_reference_image["bytes"]).decode("utf-8"),
            }
        })

    text_only_payload = {
        "contents": [{"parts": [{"text": safe_prompt}]}],
        "generationConfig": {
            "responseModalities": ["TEXT", "IMAGE"],
            "imageConfig": {
                "aspectRatio": aspect_ratio,
                "imageSize": "1K",
            },
        },
    }
    if seed is not None:
        text_only_payload["generationConfig"]["seed"] = seed

    payload = deepcopy(text_only_payload)
    if reference_image_bytes or (style_reference_image and style_reference_image.get("bytes")):
        payload["contents"] = [{"parts": parts}]
        if strength is not None:
            payload["generationConfig"]["imageStrength"] = strength
        if style_reference_image and style_reference_image.get("bytes"):
            payload["generationConfig"]["styleStrength"] = style_strength if style_strength is not None else 0.6
    headers = {
        "Content-Type": "application/json",
        "X-goog-api-key": GEMINI_API_KEY,
    }

    def _post_with_payload(active_payload: dict) -> requests.Response:
        resp = requests.post(
            f"{GEMINI_BASE_URL}/v1beta/models/{active_image_model}:generateContent",
            json=active_payload,
            headers=headers,
            timeout=90,
        )
        resp.raise_for_status()
        return resp

    data = None
    effective_seed = seed
    try:
        resp = _post_with_payload(payload)
        data = resp.json()
    except requests.HTTPError:
        fallback_payload = deepcopy(text_only_payload)
        resp = _post_with_payload(fallback_payload)
        data = resp.json()
        effective_seed = None

    candidates = data.get("candidates") or []
    if not candidates:
        raise ValueError(f"Unexpected Gemini response: {json.dumps(data, ensure_ascii=False)[:500]}")
    parts = ((candidates[0].get("content") or {}).get("parts")) or []
    for part in parts:
        inline_data = part.get("inlineData") or part.get("inline_data")
        if not inline_data:
            continue
        b64_data = inline_data.get("data", "")
        if not b64_data:
            continue
        image_bytes = base64.b64decode(b64_data)
        image_url = f"data:{inline_data.get('mimeType', inline_data.get('mime_type', 'image/png'))};base64,{b64_data}"
        return image_url, image_bytes, effective_seed

    finish_reason = candidates[0].get("finishReason", "")
    if finish_reason == "MALFORMED_FUNCTION_CALL" and payload != text_only_payload:
        retry_resp = _post_with_payload(text_only_payload)
        retry_data = retry_resp.json()
        retry_candidates = retry_data.get("candidates") or []
        retry_parts = ((retry_candidates[0].get("content") or {}).get("parts")) or [] if retry_candidates else []
        for part in retry_parts:
            inline_data = part.get("inlineData") or part.get("inline_data")
            if not inline_data:
                continue
            b64_data = inline_data.get("data", "")
            if not b64_data:
                continue
            image_bytes = base64.b64decode(b64_data)
            image_url = f"data:{inline_data.get('mimeType', inline_data.get('mime_type', 'image/png'))};base64,{b64_data}"
            return image_url, image_bytes, None
        raise ValueError(f"Gemini malformed function call after fallback: {json.dumps(retry_data, ensure_ascii=False)[:500]}")

    raise ValueError(f"Gemini response contained no image parts: {json.dumps(data, ensure_ascii=False)[:500]}")


def generate_viz_image(
    prompt: str,
    generation_index: int,
    aspect_ratio: str,
    reference_image_bytes: bytes | None = None,
    reference_mime: str = "image/png",
    seed: int | None = None,
    strength: float | None = None,
    style_reference_image: dict | None = None,
    style_strength: float | None = None,
) -> dict:
    try:
        image_url, image_bytes, effective_seed = call_gemini_image(
            prompt,
            aspect_ratio=aspect_ratio,
            reference_image_bytes=reference_image_bytes,
            reference_mime=reference_mime,
            seed=seed,
            strength=strength,
            style_reference_image=style_reference_image,
            style_strength=style_strength,
        )
        return {
            "id": f"viz-{generation_index}",
            "prompt": prompt,
            "image_url": image_url,
            "image_bytes": image_bytes,
            "seed": effective_seed,
            "error": "",
        }
    except Exception as exc:
        return {
            "id": f"viz-{generation_index}",
            "prompt": prompt,
            "image_url": "",
            "image_bytes": None,
            "seed": seed,
            "error": str(exc),
        }


def append_viz_history_entry(viz_state: dict, prompt_used: str, iteration_instruction: str = "") -> None:
    if not viz_state.get("current_image_url"):
        return
    entry_id = viz_state.get("current_result_id") or f"viz-{viz_state.get('generation_counter', 0)}"
    parent_id = viz_state.get("current_parent_id")
    viz_state["history"].insert(0, {
        "id": entry_id,
        "parent_id": parent_id,
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "material_name": viz_state.get("material_name", ""),
        "component_tags_text": viz_state.get("component_tags_text", ""),
        "component_tags": list(get_viz_components(viz_state)),
        "usage": viz_state.get("usage", ""),
        "scene": viz_state.get("scene", ""),
        "style": viz_state.get("style", ""),
        "emphasis_points_text": viz_state.get("emphasis_points_text", ""),
        "structure_notes": viz_state.get("structure_notes", ""),
        "description": viz_state.get("description", ""),
        "label_mode": viz_state.get("label_mode", ""),
        "label_language": viz_state.get("label_language", ""),
        "label_content_mode": viz_state.get("label_content_mode", ""),
        "label_terms_text": viz_state.get("label_terms_text", ""),
        "auto_generated_labels": list(viz_state.get("auto_generated_labels", [])),
        "final_label_terms": list(viz_state.get("final_label_terms", [])),
        "final_language_rule": viz_state.get("final_language_rule", ""),
        "info_density": viz_state.get("info_density", ""),
        "aspect_ratio": viz_state.get("aspect_ratio", "1:1"),
        "logic_summary": viz_state.get("logic_summary", ""),
        "base_prompt": viz_state.get("base_prompt", ""),
        "skill_prompt": viz_state.get("skill_prompt", ""),
        "optimized_prompt": viz_state.get("optimized_prompt", ""),
        "hard_constraints": viz_state.get("hard_constraints", ""),
        "prompt_with_labels": viz_state.get("prompt_with_labels", ""),
        "prompt_without_labels": viz_state.get("prompt_without_labels", ""),
        "compact_prompt": viz_state.get("compact_prompt", ""),
        "expanded_prompt": viz_state.get("expanded_prompt", ""),
        "final_prompt": viz_state.get("final_prompt", ""),
        "image_url": viz_state.get("current_image_url", ""),
        "image_bytes": viz_state.get("current_image_bytes", b""),
        "iteration_instruction": iteration_instruction,
        "local_area_hint": viz_state.get("local_area_hint", ""),
        "inpaint_strength": viz_state.get("inpaint_strength", 0.35),
        "preserve_composition": viz_state.get("preserve_composition", True),
        "edit_mode": viz_state.get("edit_mode", "局部进化"),
        "iteration_mode": viz_state.get("iteration_mode", "text-to-image"),
        "seed_value": viz_state.get("seed_value"),
        "current_seed": viz_state.get("current_seed"),
        "current_image_seed": viz_state.get("current_image_seed"),
        "style_reference_image": deepcopy(viz_state.get("style_reference_image")),
        "style_reference_name": viz_state.get("style_reference_name", ""),
        "style_strength": viz_state.get("style_strength", 0.6),
        "prompt_used": prompt_used,
    })
    viz_state["history"] = viz_state["history"][:30]
    viz_state["active_history_id"] = entry_id


def restore_viz_history_entry(viz_state: dict, entry_id: str) -> None:
    for entry in viz_state.get("history", []):
        if entry.get("id") != entry_id:
            continue
        for key in [
            "material_name", "component_tags_text", "usage", "scene", "style", "emphasis_points_text", "structure_notes",
            "description", "label_mode", "label_language", "label_content_mode", "label_terms_text", "info_density", "aspect_ratio", "logic_summary", "base_prompt",
            "skill_prompt", "optimized_prompt", "hard_constraints", "prompt_with_labels", "prompt_without_labels", "compact_prompt", "expanded_prompt", "final_prompt", "local_area_hint",
            "inpaint_strength", "preserve_composition", "edit_mode", "iteration_mode", "seed_value", "current_seed", "current_image_seed",
            "style_reference_image", "style_reference_name", "style_strength", "final_language_rule"
        ]:
            viz_state[key] = entry.get(key, deepcopy(VIZ_DEFAULT_STATE.get(key)))
        viz_state["component_tags"] = list(entry.get("component_tags", []))
        viz_state["auto_generated_labels"] = list(entry.get("auto_generated_labels", []))
        viz_state["final_label_terms"] = list(entry.get("final_label_terms", []))
        viz_state["current_result_id"] = entry.get("id")
        viz_state["current_parent_id"] = entry.get("parent_id")
        viz_state["current_image_url"] = entry.get("image_url", "")
        viz_state["current_image_bytes"] = entry.get("image_bytes", b"")
        viz_state["iteration_instruction"] = entry.get("iteration_instruction", "")
        viz_state["active_history_id"] = entry.get("id")
        return


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


def build_section_protocol(section: str) -> str:
    section_profile = SECTIONS.get(section, {})
    must_include = "\n".join(f"- {item}" for item in section_profile.get("must_include", [])) or "- 围绕当前板块核心任务展开。"
    avoid = "\n".join(f"- {item}" for item in section_profile.get("avoid", [])) or "- 避免越界到其他板块。"
    return f"""
## 板块写作协议
- 板块目标: {section_profile.get('goal', '保持学术表达稳定。')}
- 板块焦点: {section_profile.get('focus', '学术表达')}
- 推荐篇幅上限: {section_profile.get('max_words', '未设定')} 词
- 推荐输出形态: {section_profile.get('output_shape', '保持正式段落输出')}

### 本板块必须覆盖
{must_include}

### 本板块需要避免
{avoid}
"""


def build_function_section_guidance(function: str, section: str) -> str:
    guidance_map = {
        ("✨ 表达润色", "摘要"): "强化摘要压缩能力，优先保留研究对象、方法、结果和贡献四个信息点。",
        ("✨ 表达润色", "方法"): "润色时优先保证流程顺序、参数表达和可复现性，不把方法写成结果。",
        ("🤖 去AI味 (Humanizer)", "引言"): "去AI味时重点处理背景转场、研究空白引出和贡献落点，避免模板腔。",
        ("👨‍⚖️ Reviewer视角", "结果"): "重点检查数据是否充分、图表引用是否明确、结论是否超出了结果支撑。",
        ("👨‍⚖️ Reviewer视角", "方法"): "重点检查实验流程、参数完整性、评价指标和可复现性是否充分。",
        ("✍️ 影子写作", "讨论"): "优先模仿高质量论文的解释链条和文献对照方式，增强讨论深度。",
        ("🔍 逻辑检查", "结论"): "重点检查结论是否只总结已有发现，是否引入了新事实或夸张表述。",
    }
    return guidance_map.get((function, section), f"当前组合重点：在“{section}”板块下严格按“{function}”职责执行，并保持该板块应有的表达结构。")


def create_strict_prompt(function: str, input_text: str, section: str, domain: str, reference_styles: List[str], lang: str) -> str:
    skill_groups = append_skill_preview()
    domain_profile = DOMAIN_PROFILES.get(domain, {})
    section_profile = SECTIONS.get(section, {})
    prompt_rules = CONFIG["prompt_rules"]
    section_protocol = build_section_protocol(section)
    combo_guidance = build_function_section_guidance(function, section)
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
- 当前板块焦点: {section_profile.get('focus', '学术表达')}
- 当前板块增强点: {combo_guidance}

{section_protocol}

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

    if function == "🧩 文本降重":
        return base + prompt_rules["text_dedup"] + """

请执行同语种原创改写优化：
1. 保持核心观点、事实、数据、术语、结论不变。
2. 优先使用：基础改写 + 学术润色 + 句式重构 + 段落重构 + 压缩精炼 + 关键词优化（必要时扩写说明）。
3. 严禁提供规避查重/掩盖抄袭策略；如检测到此类意图，转为规范原创改写并提醒保留必要引用。
4. 严格保留 LaTeX、变量名、材料名、化学式、编号、参考标记。
5. 先输出“优化后文本”，再输出“修改说明”。
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
    active_text_model = get_active_text_model()
    message = client.messages.create(
        model=active_text_model,
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


def render_copy_text(text: str, key: str, button_label: str = "📋 一键复制结果") -> None:
    payload = json.dumps(text)
    button_id = f"copy_btn_html_{key}"
    status_id = f"copy_status_{key}"
    components.html(
        f"""
        <div style="margin:0 0 8px 0;">
            <button id="{button_id}" style="width:100%; padding:0.6rem 0.9rem; border:none; border-radius:10px; background:#2f6bff; color:white; font-weight:600; cursor:pointer;">
                {button_label}
            </button>
            <div id="{status_id}" style="font-size:12px; color:#5b6b82; margin-top:6px;"></div>
        </div>
        <script>
        const text = {payload};
        const btn = document.getElementById('{button_id}');
        const status = document.getElementById('{status_id}');
        if (btn && !btn.dataset.copyBound) {{
            btn.dataset.copyBound = '1';
            btn.addEventListener('click', async () => {{
                try {{
                    await navigator.clipboard.writeText(text);
                    if (status) status.textContent = '已复制到剪贴板';
                }} catch (e) {{
                    if (status) status.textContent = '复制失败，请手动复制';
                }}
            }});
        }}
        </script>
        """,
        height=58,
    )


def extract_primary_output(text: str, function: str) -> str:
    cleaned = text.strip()
    if function not in {"🤖 去AI味 (Humanizer)", "📋 Redlining修订", "🧩 文本降重"}:
        return cleaned

    patterns = [
        r"(?:^|\n)#+\s*优化后文本\s*[:：]?\s*\n([\s\S]*?)(?=\n#+\s*修改说明\s*[:：]?|\n修改说明\s*[:：]?|\Z)",
        r"(?:^|\n)\*\*优化后文本\*\*\s*[:：]?\s*\n([\s\S]*?)(?=\n\*\*修改说明\*\*\s*[:：]?|\n修改说明\s*[:：]?|\Z)",
        r"(?:^|\n)优化后文本\s*[:：]\s*\n?([\s\S]*?)(?=\n修改说明\s*[:：]?|\Z)",
        r"(?:^|\n)#+\s*修改后文本\s*[:：]?\s*\n([\s\S]*?)(?=\n#+\s*修改说明\s*[:：]?|\n修改说明\s*[:：]?|\Z)",
        r"(?:^|\n)\*\*修改后文本\*\*\s*[:：]?\s*\n([\s\S]*?)(?=\n\*\*修改说明\*\*\s*[:：]?|\n修改说明\s*[:：]?|\Z)",
        r"(?:^|\n)修改后文本\s*[:：]\s*\n?([\s\S]*?)(?=\n修改说明\s*[:：]?|\Z)",
    ]
    for pattern in patterns:
        match = re.search(pattern, cleaned, flags=re.MULTILINE)
        if match:
            candidate = match.group(1).strip()
            if candidate:
                return candidate
    return cleaned


def inject_custom_css() -> None:
    st.markdown(
        UI_CONFIG["primary_button_css"],
        unsafe_allow_html=True,
    )


def render_sidebar_brand() -> None:
    st.sidebar.markdown("<div class='xueyan-badge'>XUEYAN GOOD</div>", unsafe_allow_html=True)


def render_sidebar_panel(title: str, note: str) -> None:
    st.sidebar.markdown(
        f"<div class='sidebar-panel'><div class='sidebar-panel-title'>{title}</div><p class='sidebar-panel-note'>{note}</p></div>",
        unsafe_allow_html=True,
    )


def render_header() -> None:
    api_status = "API 已连接" if CLAUDE_API_KEY else "API 未配置"
    st.markdown(
        f"""
<div class="main-header">
    <h1>学研 · Xueyan Workstation</h1>
    <p>写作、排版、绘图与单页 PPT 在同一工作台内协同完成</p>
    <div class="lab-status-pills" style="margin-top:0.42rem; margin-bottom:0;">
        <span class="lab-status-pill"><strong>状态</strong>{api_status}</span>
        <span class="lab-status-pill"><strong>规则包</strong>{len(append_skill_preview())} 组</span>
        <span class="lab-status-pill"><strong>响应</strong>300s / 3次重试</span>
    </div>
</div>
""",
        unsafe_allow_html=True,
    )


def render_result_actions(section_name: str, current_input: str, previous_output: str, function: str, domain: str) -> None:
    st.markdown("<div class='result-toolbar'><strong>快捷操作</strong></div>", unsafe_allow_html=True)
    copy_button_id = f"copy_result_btn_{section_name}_{function}".replace(" ", "_")
    col1, col2 = st.columns([1, 1])
    with col1:
        render_copy_text(previous_output, copy_button_id, button_label="📋 一键复制结果")
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
        primary_output = extract_primary_output(final_output, function)
        st.session_state[output_key] = primary_output
        st.session_state[writing_term_count_key(section_name)] = len(term_mapping)
        if function == "🤖 去AI味 (Humanizer)":
            note_text = f"✅ 已针对【{section_name}】完成【{function}】，保持原语种，仅重构语序与节奏。"
        elif function == "📝 中转英翻译":
            note_text = f"✅ 已针对【{section_name}】完成【{function}】，已执行跨语种精准学术翻译并保留术语。"
        else:
            note_text = f"✅ 已针对【{section_name}】完成【{function}】，严格按单一职责执行。"
        st.session_state[note_key] = note_text
        section_focus = SECTIONS.get(section_name, {}).get("focus", "学术表达")
        section_goal = SECTIONS.get(section_name, {}).get("goal", "保持当前板块写作目标稳定。")
        combo_guidance = build_function_section_guidance(function, section_name)
        st.session_state.writing_history.insert(0, {
            "id": len(st.session_state.writing_history) + 1,
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "function": function,
            "section": section_name,
            "domain": domain,
            "input_lang": input_lang,
            "input": current_input,
            "output": primary_output,
            "elapsed": f"{elapsed:.1f}s",
            "section_focus": section_focus,
            "section_goal": section_goal,
            "combo_guidance": combo_guidance,
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
        st.markdown('<div class="lab-main-card">', unsafe_allow_html=True)
        st.markdown(f'<div class="lab-panel-label">原文输入 · {section_name}</div>', unsafe_allow_html=True)
        st.markdown('<div class="lab-panel-note">把草稿、老师意见或文献信息直接放这里。</div>', unsafe_allow_html=True)
        st.text_area(
            "原文输入",
            key=input_key,
            height=400,
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
        section_profile = SECTIONS.get(section_name, {})
        st.markdown(
            f"<div class=\"lab-status-pills\"><span class=\"lab-status-pill\"><strong>字数</strong>{len(current_input)}</span><span class=\"lab-status-pill\"><strong>语言</strong>{'中文' if input_language == 'zh' else '英文' if input_language == 'en' else '待识别'}</span></div>",
            unsafe_allow_html=True,
        )
        st.markdown(f"<div class='lab-quiet-note'>{section_profile.get('input_hint', '把最原始的想法、草稿、老师意见或文献信息放这里。')}</div>", unsafe_allow_html=True)
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
        st.markdown('</div>', unsafe_allow_html=True)

    with col_right:
        st.markdown('<div class="lab-main-card">', unsafe_allow_html=True)
        st.markdown(f'<div class="lab-panel-label">处理结果 · {section_name}</div>', unsafe_allow_html=True)
        st.markdown('<div class="lab-panel-note">这里只显示当前板块的处理结果。</div>', unsafe_allow_html=True)
        section_profile = SECTIONS.get(section_name, {})
        checkpoints = section_profile.get("output_checkpoints", [])
        if checkpoints:
            st.markdown(f"<div class='lab-quiet-note'>检查点：{' / '.join(checkpoints[:3])}</div>", unsafe_allow_html=True)
        st.text_area(
            "处理结果",
            key=output_key,
            height=400,
            disabled=True,
            label_visibility="collapsed",
        )
        previous_output = st.session_state.get(output_key, "")
        if previous_output:
            render_result_actions(section_name, st.session_state.get(input_key, ""), previous_output, function, domain)
            if previous_note:
                st.markdown(f"<div class='lab-quiet-note'>{previous_note}</div>", unsafe_allow_html=True)
            lock_hint = domain_profile.get("hard_lock_terms", [])[:2] + domain_profile.get("hard_lock_regex", [])[:1]
            if term_count:
                st.markdown(f"<div class='lab-quiet-note'>已保护 {term_count} 个核心锁项：{', '.join(lock_hint) if lock_hint else 'LaTeX / 术语 / 单位'}。</div>", unsafe_allow_html=True)
            else:
                st.markdown("<div class='lab-quiet-note'>已启用学科核心锁，LaTeX 公式与术语默认不改写。</div>", unsafe_allow_html=True)
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
            st.markdown("<div class='lab-quiet-note'>这里会生成可直接拿去改稿的版本。</div>", unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)


def get_function_nav_index(current_function: str) -> int:
    for idx, item in enumerate(FUNCTION_NAV):
        if item["value"] == current_function:
            return idx
    return 0


def render_writing_engine_sidebar() -> List[str]:
    render_sidebar_brand()
    st.sidebar.radio("核心引擎", [WRITING_PAGE, FORMATTING_PAGE, VIZ_PAGE, PPT_PAGE], key="engine_mode")
    st.sidebar.selectbox("统一功能栏", options=FUNCTION_NAV, index=get_function_nav_index(st.session_state.get("writing_function", list(FUNCTION_MATRIX.keys())[0])), format_func=lambda item: item["label"], key="writing_function_selector")
    st.session_state.writing_function = st.session_state["writing_function_selector"]["value"]
    st.sidebar.selectbox("🔬 学科大脑", DOMAIN_ORDER, key="writing_domain_selector")
    st.session_state.writing_domain = st.session_state["writing_domain_selector"]

    function = st.session_state.get("writing_function", list(FUNCTION_MATRIX.keys())[0])
    func_info = FUNCTION_MATRIX.get(function, {})
    st.sidebar.markdown("<div class='sidebar-section-note'>当前功能说明</div>", unsafe_allow_html=True)
    st.sidebar.caption(func_info.get("description", ""))
    active_section = st.session_state.get("writing_active_section", SECTION_NAMES[0])
    st.sidebar.caption("板块增强点：" + build_function_section_guidance(function, active_section))
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

    st.sidebar.markdown("---")
    st.sidebar.caption("规则包已加载，可按当前功能与学科直接调用。")
    return reference_styles


def render_writing_engine() -> None:
    init_writing_state()
    reference_styles = render_writing_engine_sidebar()
    function = st.session_state.get("writing_function", list(FUNCTION_MATRIX.keys())[0])
    domain = st.session_state.get("writing_domain", DOMAIN_ORDER[0])
    active_section = st.session_state.get("writing_active_section", SECTION_NAMES[0])
    focus_hint = SECTIONS.get(active_section, {}).get("focus", "聚焦学术表达")
    goal_hint = SECTIONS.get(active_section, {}).get("goal", "保持当前板块写作目标稳定。")
    combo_hint = build_function_section_guidance(function, active_section)

    st.markdown(
        f"""
<div class="lab-thin-header">
    <div class="lab-thin-header-row">
        <div>
            <h3>论文写作工作台</h3>
            <p>围绕当前板块，直接完成输入、处理和结果回看。</p>
        </div>
        <span class="workbench-chip">{active_section}</span>
    </div>
</div>
""",
        unsafe_allow_html=True,
    )

    top_left, top_right = st.columns([1.25, 0.95], gap="large")
    with top_left:
        st.markdown(
            f"<div class=\"lab-status-pills\"><span class=\"lab-status-pill\"><strong>功能</strong>{function}</span><span class=\"lab-status-pill\"><strong>学科</strong>{domain}</span><span class=\"lab-status-pill\"><strong>焦点</strong>{focus_hint}</span></div>",
            unsafe_allow_html=True,
        )
        st.markdown(f"<div class='lab-quiet-note'>板块目标：{goal_hint}<br>增强点：{combo_hint}</div>", unsafe_allow_html=True)
    with top_right:
        st.markdown('<div class="lab-secondary-card">', unsafe_allow_html=True)
        st.selectbox(
            "当前写作板块",
            SECTION_NAMES,
            index=SECTION_NAMES.index(active_section),
            key="writing_active_section_selector",
        )
        st.session_state.writing_active_section = st.session_state["writing_active_section_selector"]
        st.markdown("<div class='lab-quiet-note'>保持双栏主工作流，其他信息放到次级区域。</div>", unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

    active_section = st.session_state.get("writing_active_section", SECTION_NAMES[0])
    render_writing_section(active_section, function, domain, reference_styles)
    with st.expander("版本时光机", expanded=False):
        render_history_panel()


def render_history_panel() -> None:
    st.markdown("<div class='lab-quiet-note'>当前学科核心锁已启用：LaTeX 公式、术语 regex 与关键缩写默认不改写。</div>", unsafe_allow_html=True)
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
                meta_bits = [f"领域: {entry['domain']}", f"语言: {'中文' if entry['input_lang'] == 'zh' else '英文'}", f"耗时: {entry['elapsed']}"]
                st.markdown(f"<div class='lab-quiet-note'>{' | '.join(meta_bits)}</div>", unsafe_allow_html=True)
                if entry.get("section_focus"):
                    st.markdown(f"<div class='lab-quiet-note'>板块焦点：{entry['section_focus']}</div>", unsafe_allow_html=True)
                if entry.get("combo_guidance"):
                    st.markdown(f"<div class='lab-quiet-note'>组合增强：{entry['combo_guidance']}</div>", unsafe_allow_html=True)
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
    st.sidebar.radio("核心引擎", [WRITING_PAGE, FORMATTING_PAGE, VIZ_PAGE, PPT_PAGE], key="engine_mode")
    st.sidebar.selectbox("🔬 学科大脑", DOMAIN_ORDER, key="format_formatting_domain_selector")
    st.session_state.formatting_domain = st.session_state["format_formatting_domain_selector"]
    st.sidebar.selectbox(
        "📚 规则包",
        options=list(FORMAT_RULESET_LIBRARY.keys()),
        format_func=lambda key: get_format_ruleset_profile(key)["label"],
        key="format_ruleset",
    )
    render_sidebar_panel("格式对齐", "面向 Word 的确定性审计与自动修复。")
    st.sidebar.caption("指南、上传、审计与导出保持独立闭环。")


def render_viz_engine_sidebar() -> None:
    viz_state = init_viz_state()
    viz_state["logic_summary"] = build_viz_logic_summary(viz_state)
    render_sidebar_brand()
    st.sidebar.radio("核心引擎", [WRITING_PAGE, FORMATTING_PAGE, VIZ_PAGE, PPT_PAGE], key="engine_mode")
    render_sidebar_panel("视觉实验室", "材料科研单图生成、精修与版本迭代工作台。")
    active_image_model = get_active_image_model()
    image_options = st.session_state.get("runtime_image_model_options", list(GEMINI_MODEL_OPTIONS))
    st.sidebar.caption(f"当前绘图模型：{active_image_model}")
    st.sidebar.selectbox(
        "切换绘图模型",
        image_options,
        index=image_options.index(active_image_model) if active_image_model in image_options else 0,
        key="runtime_image_model",
    )
    st.sidebar.caption("网页切换仅影响当前会话。")
    with st.sidebar.expander("查看自动逻辑摘要", expanded=False):
        st.code(viz_state["logic_summary"] or "填写参数后，这里会更新摘要。", language=None)


def render_viz_history_panel(viz_state: dict) -> None:
    if not viz_state.get("history"):
        st.info("📭 暂无绘图历史")
        return
    for entry in viz_state["history"]:
        title = f"{entry['timestamp']} · {entry['scene']} · {entry['style']}"
        with st.expander(title, expanded=False):
            meta_bits = [f"链路：{entry.get('parent_id') or 'ROOT'} → {entry['id']}"]
            if entry.get("iteration_mode"):
                meta_bits.append(f"模式：{entry['iteration_mode']}")
            if entry.get("current_seed") is not None:
                meta_bits.append(f"Seed：{entry['current_seed']}")
            st.caption(" | ".join(meta_bits))
            st.caption(entry.get("description", "")[:140] or "无描述")
            render_viz_image_slot(entry.get("image_url", ""), empty_text="该历史版本没有可预览图片。", container_height=220, image_max_height=180)
            c1, c2 = st.columns([1, 1])
            with c1:
                if st.button("📥 恢复此版本", key=f"viz_restore_{entry['id']}", use_container_width=True):
                    restore_viz_history_entry(viz_state, entry["id"])
                    st.rerun()
            with c2:
                image_bytes = entry.get("image_bytes") or get_viz_image_bytes(entry.get("image_url", ""))
                if image_bytes:
                    st.download_button(
                        "⬇️ 下载当前图",
                        image_bytes,
                        file_name=f"xueyan_viz_{entry['id']}.png",
                        mime="image/png",
                        key=f"viz_hist_dl_{entry['id']}",
                        use_container_width=True,
                    )


def render_ppt_engine_sidebar() -> None:
    ppt_state = init_ppt_state()
    render_sidebar_brand()
    st.sidebar.radio("核心引擎", [WRITING_PAGE, FORMATTING_PAGE, VIZ_PAGE, PPT_PAGE], key="engine_mode")
    render_sidebar_panel("PPT 工作台", "围绕当前页完成提炼、预览与组装。")
    active_text_model = get_active_text_model()
    active_image_model = get_active_image_model()
    text_options = st.session_state.get("runtime_text_model_options", list(CLAUDE_MODEL_OPTIONS))
    image_options = st.session_state.get("runtime_image_model_options", list(GEMINI_MODEL_OPTIONS))
    st.sidebar.caption(f"当前文本模型：{active_text_model}")
    st.sidebar.selectbox(
        "切换文本模型",
        text_options,
        index=text_options.index(active_text_model) if active_text_model in text_options else 0,
        key="runtime_text_model",
    )
    st.sidebar.caption(f"当前辅助图模型：{active_image_model}")
    st.sidebar.selectbox(
        "切换绘图模型",
        image_options,
        index=image_options.index(active_image_model) if active_image_model in image_options else 0,
        key="runtime_image_model",
    )
    st.sidebar.caption("网页切换仅影响当前会话。")
    summary = (
        f"标题：{ppt_state.get('page_title') or '未命名'}\n"
        f"类型：{ppt_state.get('page_type', '结果页')}\n"
        f"布局：{ppt_state.get('selected_layout') or PPT_PAGE_TYPE_LAYOUTS.get(ppt_state.get('page_type', '结果页'), '结果+结论')}\n"
        f"输入：{detect_ppt_input_mode(ppt_state)}\n"
        f"组装页数：{len(ppt_state.get('assembly_pages', []))}"
    )
    with st.sidebar.expander("查看当前页摘要", expanded=False):
        st.code(summary, language=None)


def render_ppt_history_panel(ppt_state: dict) -> None:
    if not ppt_state.get("history"):
        st.info("📭 暂无当前页历史")
        return
    for entry in ppt_state["history"]:
        title = f"{entry['timestamp']} · {entry.get('page_type', '结果页')} · {entry.get('page_title') or '未命名'}"
        with st.expander(title, expanded=False):
            st.caption(f"链路: {entry.get('parent_id') or 'ROOT'} → {entry['id']}")
            st.caption(f"布局：{entry.get('selected_layout', '结果+结论')} | 风格：{entry.get('style_mode', 'academic-paperskills')}")
            if entry.get("current_slide_preview_html"):
                components.html(entry["current_slide_preview_html"], height=360)
            c1, c2 = st.columns(2)
            with c1:
                if st.button("📥 恢复当前页", key=f"ppt_restore_{entry['id']}", use_container_width=True):
                    restore_ppt_history_entry(ppt_state, entry["id"])
                    st.rerun()
            with c2:
                st.download_button(
                    "⬇️ 下载 JSON",
                    data=json.dumps(sanitize_for_json(entry.get("current_slide_struct") or {}), ensure_ascii=False, indent=2).encode("utf-8"),
                    file_name=f"xueyan_ppt_{entry['id']}.json",
                    mime="application/json",
                    key=f"ppt_hist_dl_{entry['id']}",
                    use_container_width=True,
                )


def render_viz_engine() -> None:
    viz_state = init_viz_state()
    render_viz_engine_sidebar()

    st.markdown(
        """
<div class="viz-lab-shell">
    <div class="viz-thin-header">
        <div class="viz-thin-header-row">
            <div>
                <h3>科研绘图工作台</h3>
                <p>围绕当前图，快速完成描述、生成与精修。</p>
            </div>
            <span class="workbench-chip">Visualization Lab</span>
        </div>
    </div>
</div>
""",
        unsafe_allow_html=True,
    )

    left_col, right_col = st.columns([0.92, 1.78], gap="large")

    with left_col:
        st.markdown('<div class="viz-secondary-card">', unsafe_allow_html=True)
        st.markdown('<div class="viz-panel-label">参数区</div>', unsafe_allow_html=True)
        st.markdown('<div class="viz-panel-note">高频参数已前置到右侧，这里保留完整设置能力。</div>', unsafe_allow_html=True)
        with st.expander("材料基础", expanded=True):
            viz_state["material_name"] = st.text_input(
                "材料名称",
                value=viz_state.get("material_name", ""),
                placeholder="如：MoS2 / 石墨烯复合材料 / 多孔氧化物",
                key="viz_material_name",
            )
            viz_state["component_tags_text"] = st.text_area(
                "物质组成 / 关键组分",
                value=viz_state.get("component_tags_text", ""),
                height=82,
                placeholder="每行一个，或用逗号分隔",
                key="viz_component_tags_text",
            )
            viz_state["structure_notes"] = st.text_area(
                "结构补充说明",
                value=viz_state.get("structure_notes", ""),
                height=78,
                placeholder="例如：核壳结构 / 多孔骨架 / 层状堆叠",
                key="viz_structure_notes",
            )
        with st.expander("场景与用途", expanded=True):
            usage_options = ["论文主图", "论文 TOC 图", "汇报展示", "基金申请", "教学示意", "社媒科普"]
            viz_state["usage"] = st.selectbox(
                "图像用途",
                usage_options,
                index=usage_options.index(viz_state.get("usage", "论文主图")) if viz_state.get("usage", "论文主图") in usage_options else 0,
                key="viz_usage",
            )
            scene_options = list(VIZ_SCENE_PROMPTS.keys())
            viz_state["scene"] = st.selectbox(
                "场景类型",
                scene_options,
                index=scene_options.index(viz_state.get("scene", scene_options[0])) if viz_state.get("scene", scene_options[0]) in scene_options else 0,
                key="viz_scene",
            )
            viz_state["emphasis_points_text"] = st.text_area(
                "强调重点",
                value=viz_state.get("emphasis_points_text", ""),
                height=78,
                placeholder="例如：突出离子扩散路径、界面反应区域",
                key="viz_emphasis_points_text",
            )
        with st.expander("风格扩展", expanded=False):
            style_reference_file = st.file_uploader(
                "上传风格参考图",
                type=["png", "jpg", "jpeg"],
                key="viz_style_reference_file",
            )
            if style_reference_file is not None:
                style_reference = collect_uploaded_image(style_reference_file)
                viz_state["style_reference_image"] = style_reference
                viz_state["style_reference_name"] = style_reference.get("name", "") if style_reference else ""
            elif viz_state.get("style_reference_name") and not viz_state.get("style_reference_image"):
                viz_state["style_reference_name"] = ""
            if viz_state.get("style_reference_image"):
                st.image(viz_state["style_reference_image"]["bytes"], caption=viz_state.get("style_reference_name") or "风格参考图", width=180)
            viz_state["style_strength"] = st.slider(
                "风格模仿强度",
                min_value=0.0,
                max_value=1.0,
                value=float(viz_state.get("style_strength", 0.6)),
                step=0.1,
                key="viz_style_strength",
            )
        with st.expander("标注内容", expanded=viz_state.get("label_mode") == "有文字版"):
            if viz_state.get("label_mode") == "有文字版":
                content_modes = ["自动生成后编辑（推荐）", "手动填写", "仅自动生成"]
                viz_state["label_content_mode"] = st.selectbox(
                    "标签内容来源",
                    content_modes,
                    index=content_modes.index(viz_state.get("label_content_mode", "自动生成后编辑（推荐）")) if viz_state.get("label_content_mode", "自动生成后编辑（推荐）") in content_modes else 0,
                    key="viz_label_content_mode",
                )
                suggested_terms = suggest_viz_label_terms(viz_state)
                viz_state["auto_generated_labels"] = suggested_terms
                if viz_state.get("label_content_mode") != "手动填写" and not viz_state.get("label_terms_text", "").strip():
                    viz_state["label_terms_text"] = "\n".join(suggested_terms)
                label_title = "中文标识内容" if viz_state.get("label_language") == "中文" else "英文标识内容"
                viz_state["label_terms_text"] = st.text_area(
                    label_title,
                    value=viz_state.get("label_terms_text", ""),
                    height=110,
                    placeholder="例如：离子迁移路径\n界面反应区\n结构示意",
                    key="viz_label_terms_text",
                )
                if suggested_terms:
                    st.caption(f"自动建议：{' ｜ '.join(suggested_terms)}")
            else:
                viz_state["final_label_terms"] = []
                st.caption("当前为无文字版；切换到有文字版后可编辑标签内容。")
        with st.expander("自动逻辑摘要", expanded=False):
            viz_state["logic_summary"] = build_viz_logic_summary(viz_state)
            st.caption(viz_state["logic_summary"] or "填写参数后，这里会自动生成摘要。")
        st.markdown('</div>', unsafe_allow_html=True)

    with right_col:
        style_options = list(VIZ_STYLE_PROMPTS.keys())
        st.markdown(
            """
<div class="viz-topbar">
    <div class="viz-topbar-title">
        <strong>核心参数</strong>
        <span>先设关键项，再生成结果。</span>
    </div>
</div>
""",
            unsafe_allow_html=True,
        )
        top1, top2, top3, top4, top5 = st.columns(5)
        with top1:
            viz_state["label_mode"] = st.radio(
                "标注模式",
                ["无文字版", "有文字版"],
                horizontal=True,
                index=0 if viz_state.get("label_mode", "无文字版") == "无文字版" else 1,
                key="viz_label_mode",
            )
        with top2:
            language_disabled = viz_state.get("label_mode", "无文字版") != "有文字版"
            if language_disabled:
                st.markdown('<div class="viz-language-muted">', unsafe_allow_html=True)
            viz_state["label_language"] = st.radio(
                "标签语言",
                ["中文", "英文"],
                horizontal=True,
                index=0 if viz_state.get("label_language", "中文") == "中文" else 1,
                key="viz_label_language",
            )
            if language_disabled:
                st.markdown('</div>', unsafe_allow_html=True)
        with top3:
            viz_state["style"] = st.selectbox(
                "风格",
                style_options,
                index=style_options.index(viz_state.get("style", style_options[0])) if viz_state.get("style", style_options[0]) in style_options else 0,
                key="viz_style",
            )
        with top4:
            density_options = ["低", "中", "高"]
            viz_state["info_density"] = st.selectbox(
                "信息密度",
                density_options,
                index=density_options.index(viz_state.get("info_density", "中")) if viz_state.get("info_density", "中") in density_options else 1,
                key="viz_info_density",
            )
        with top5:
            ratio_options = ["1:1", "4:3", "3:2", "16:9", "9:16"]
            viz_state["aspect_ratio"] = st.selectbox(
                "输出比例",
                ratio_options,
                index=ratio_options.index(viz_state.get("aspect_ratio", "1:1")) if viz_state.get("aspect_ratio", "1:1") in ratio_options else 0,
                key="viz_aspect_ratio",
            )

        st.markdown('<div class="viz-main-card">', unsafe_allow_html=True)
        st.markdown('<div class="viz-panel-label">图像描述</div>', unsafe_allow_html=True)
        st.markdown('<div class="viz-panel-note">直接写你想表达的科研图像内容。</div>', unsafe_allow_html=True)
        viz_state["description"] = st.text_area(
            "图像描述",
            value=viz_state.get("description", ""),
            height=230,
            placeholder="例如：展示层状材料中离子在层间扩散，并突出表面异质结对反应动力学的促进作用。",
            key="viz_description",
        )

        prompt_bundle = build_viz_prompt_bundle(viz_state)
        viz_state.update(prompt_bundle)

        action_col1, action_col2, action_col3 = st.columns([1.1, 0.95, 1.35])
        with action_col1:
            generate_clicked = st.button("生成图片", type="primary", use_container_width=True, key="viz_generate")
        with action_col2:
            regenerate_clicked = st.button("重新生成", use_container_width=True, key="viz_regenerate")
        with action_col3:
            st.markdown('<div class="viz-quiet-block">默认围绕当前图逐步迭代，适合持续微调。</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

        if generate_clicked or regenerate_clicked:
            if not viz_state["description"].strip():
                st.warning("请先填写图像描述。")
            elif not GEMINI_API_KEY:
                st.error("未配置 GEMINI_API_KEY。")
            else:
                prompt_to_use = viz_state.get("final_prompt") or (viz_state["prompt_with_labels"] if viz_state.get("label_mode") == "有文字版" else viz_state["prompt_without_labels"])
                viz_state["base_prompt"] = prompt_to_use
                viz_state["generation_counter"] += 1
                viz_state["current_parent_id"] = None if generate_clicked else viz_state.get("current_parent_id")
                seed_to_use = viz_state.get("current_seed")
                if seed_to_use is None:
                    seed_to_use = random.randint(100000, 999999)
                    viz_state["seed_value"] = seed_to_use
                    viz_state["current_seed"] = seed_to_use
                with st.spinner("正在生成图片..."):
                    result = generate_viz_image(
                        prompt_to_use,
                        viz_state["generation_counter"],
                        viz_state.get("aspect_ratio", "1:1"),
                        seed=seed_to_use,
                        style_reference_image=viz_state.get("style_reference_image"),
                        style_strength=float(viz_state.get("style_strength", 0.6)),
                    )
                if result.get("error"):
                    viz_state["last_error"] = result["error"]
                    st.error(result["error"])
                else:
                    viz_state["iteration_mode"] = "text-to-image"
                    viz_state["current_result_id"] = result["id"]
                    viz_state["current_image_url"] = result["image_url"]
                    viz_state["current_image_bytes"] = result.get("image_bytes") or b""
                    viz_state["current_seed"] = result.get("seed")
                    viz_state["current_image_seed"] = result.get("seed")
                    viz_state["last_error"] = ""
                    append_viz_history_entry(viz_state, prompt_to_use)

        st.markdown('<div class="viz-main-card">', unsafe_allow_html=True)
        st.markdown('<div class="viz-panel-label">当前结果</div>', unsafe_allow_html=True)
        st.markdown('<div class="viz-panel-note">当前版本会固定展示在这里，便于对照与继续修改。</div>', unsafe_allow_html=True)
        result_name = viz_state.get("current_result_id") or "未生成"
        seed_label = viz_state.get("current_image_seed")
        st.markdown(
            f"<div class=\"viz-status-pills\"><span class=\"viz-status-pill\"><strong>结果</strong>{result_name}</span><span class=\"viz-status-pill\"><strong>Seed</strong>{seed_label if seed_label is not None else '自动'}</span><span class=\"viz-status-pill\"><strong>模式</strong>{viz_state.get('iteration_mode') or 'text-to-image'}</span></div>",
            unsafe_allow_html=True,
        )
        if viz_state.get("current_image_url"):
            render_viz_image_slot(viz_state["current_image_url"], empty_text="尚未生成图片。", container_height=500, image_max_height=460)
            dcol1, dcol2 = st.columns([1, 1])
            with dcol1:
                st.download_button(
                    "下载当前图片",
                    data=viz_state.get("current_image_bytes") or get_viz_image_bytes(viz_state.get("current_image_url", "")) or b"",
                    file_name=f"viz-{viz_state.get('current_result_id') or 'current'}.png",
                    mime="image/png",
                    use_container_width=True,
                    key="viz_current_download",
                )
            with dcol2:
                st.markdown('<div class="viz-compact-status">当前结果已就绪，可直接继续修改。</div>', unsafe_allow_html=True)
        else:
            render_viz_image_slot("", empty_text="填写图像描述后点击“生成图片”。", container_height=500, image_max_height=460)
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div class="viz-secondary-card">', unsafe_allow_html=True)
        st.markdown('<div class="viz-panel-label">继续修改当前结果</div>', unsafe_allow_html=True)
        st.markdown('<div class="viz-panel-note">生成后可在这里做局部微调，不打断主流程。</div>', unsafe_allow_html=True)
        viz_state["iteration_instruction"] = st.text_input(
            "继续修改当前结果",
            value=viz_state.get("iteration_instruction", ""),
            placeholder="例如：把 Li+ 改为红色 / 把右上角放大框增强一点",
            key="viz_iteration_instruction",
        )
        tweak_col1, tweak_col2 = st.columns([1.2, 0.8])
        with tweak_col1:
            viz_state["local_area_hint"] = st.text_input(
                "修改路径 / 局部区域",
                value=viz_state.get("local_area_hint", ""),
                placeholder="例如：右上角放大框 > 箭头路径 / 中央主体颗粒",
                key="viz_local_area_hint",
            )
        with tweak_col2:
            viz_state["inpaint_strength"] = st.slider(
                "局部修改强度",
                min_value=0.2,
                max_value=0.5,
                value=float(viz_state.get("inpaint_strength", 0.35)),
                step=0.05,
                key="viz_inpaint_strength",
            )
        viz_state["preserve_composition"] = st.checkbox(
            "保持原始构图",
            value=viz_state.get("preserve_composition", True),
            key="viz_preserve_composition",
        )
        iteration_preview_prompt = build_viz_iteration_prompt(viz_state) if viz_state.get("iteration_instruction", "").strip() else ""
        viz_state["iteration_preview_prompt"] = iteration_preview_prompt
        apply_edit_clicked = st.button("应用修改", use_container_width=True, key="viz_apply_edit")
        st.markdown('<div class="viz-quiet-block">保持原始构图时，更适合做局部修正与版本迭代。</div>', unsafe_allow_html=True)
        if apply_edit_clicked:
            if not viz_state.get("current_image_url"):
                st.warning("请先生成当前图片。")
            elif not viz_state["iteration_instruction"].strip():
                st.warning("请先填写修改指令。")
            elif not GEMINI_API_KEY:
                st.error("未配置 GEMINI_API_KEY。")
            else:
                parent_id = viz_state.get("current_result_id")
                iteration_prompt = build_viz_iteration_prompt(viz_state)
                reference_image_bytes = viz_state.get("current_image_bytes") or get_viz_image_bytes(viz_state.get("current_image_url", ""))
                viz_state["generation_counter"] += 1
                seed_to_use = viz_state.get("current_seed") or random.randint(100000, 999999)
                viz_state["seed_value"] = seed_to_use
                viz_state["current_seed"] = seed_to_use
                strength_to_use = min(float(viz_state.get("inpaint_strength", 0.35)), 0.35)
                with st.spinner("正在基于当前结果生成新版本..."):
                    result = generate_viz_image(
                        iteration_prompt,
                        viz_state["generation_counter"],
                        viz_state.get("aspect_ratio", "1:1"),
                        reference_image_bytes=reference_image_bytes,
                        seed=seed_to_use,
                        strength=strength_to_use,
                        style_reference_image=viz_state.get("style_reference_image"),
                        style_strength=float(viz_state.get("style_strength", 0.6)),
                    )
                if result.get("error"):
                    viz_state["last_error"] = result["error"]
                    st.error(result["error"])
                else:
                    viz_state["base_prompt"] = iteration_prompt
                    viz_state["iteration_mode"] = "reference-refine"
                    viz_state["current_parent_id"] = parent_id
                    viz_state["current_result_id"] = result["id"]
                    viz_state["current_image_url"] = result["image_url"]
                    viz_state["current_image_bytes"] = result.get("image_bytes") or b""
                    viz_state["current_seed"] = result.get("seed")
                    viz_state["current_image_seed"] = result.get("seed")
                    viz_state["last_error"] = ""
                    append_viz_history_entry(viz_state, iteration_prompt, viz_state["iteration_instruction"])
        st.markdown('</div>', unsafe_allow_html=True)

        with st.expander("Prompt 预览", expanded=False):
            tab1, tab2, tab3, tab4 = st.tabs(["最终 Prompt", "Skills 基础稿", "优化增强稿", "硬约束"])
            with tab1:
                st.code(viz_state["final_prompt"], language="text")
            with tab2:
                st.code(viz_state.get("skill_prompt", ""), language="text")
            with tab3:
                st.code(viz_state.get("optimized_prompt", ""), language="text")
            with tab4:
                st.code(viz_state.get("hard_constraints", ""), language="text")
        with st.expander("局部修改 Prompt 预览", expanded=False):
            if iteration_preview_prompt:
                st.code(iteration_preview_prompt, language="text")
            else:
                st.info("填写修改框后，这里会显示最终发给 Gemini 的局部修改 Prompt。")
        with st.expander("自动逻辑摘要", expanded=False):
            viz_state["logic_summary"] = build_viz_logic_summary(viz_state)
            st.caption(viz_state["logic_summary"] or "填写参数后，这里会自动生成摘要。")
        with st.expander("版本历史", expanded=False):
            render_viz_history_panel(viz_state)


def render_ppt_engine() -> None:
    ppt_state = init_ppt_state()
    render_ppt_engine_sidebar()

    st.markdown(
        """
<div class="lab-thin-header">
    <div class="lab-thin-header-row">
        <div>
            <h3>PPT 工作台</h3>
            <p>围绕当前页，完成输入提炼、预览微调与逐页组装。</p>
        </div>
        <span class="workbench-chip">Single-page PPT</span>
    </div>
</div>
""",
        unsafe_allow_html=True,
    )

    panel1, panel2, panel3 = st.columns([1.0, 1.28, 0.95], gap="large")

    with panel1:
        st.markdown('<div class="lab-secondary-card">', unsafe_allow_html=True)
        st.markdown('<div class="lab-panel-label">输入与页面设置</div>', unsafe_allow_html=True)
        st.markdown('<div class="lab-panel-note">当前页素材与控制项放在这里，保持完整但弱化展示。</div>', unsafe_allow_html=True)
        with st.expander("输入来源", expanded=True):
            source_file = st.file_uploader("上传 PDF / Word", type=["pdf", "docx"], key="ppt_source_file")
            ppt_state["raw_text"] = st.text_area(
                "输入文字",
                value=ppt_state.get("raw_text", ""),
                height=180,
                placeholder="把当前页相关的文字、论文摘要、实验结果描述、老师意见放这里。",
                key="ppt_raw_text",
            )
            uploaded_images_files = st.file_uploader("上传图片", type=["png", "jpg", "jpeg"], accept_multiple_files=True, key="ppt_uploaded_images")
            reference_image_files = st.file_uploader("上传参考图", type=["png", "jpg", "jpeg"], accept_multiple_files=True, key="ppt_reference_images")
            template_file = st.file_uploader("上传 PPT 模板", type=["pptx"], key="ppt_template_file")

            normalized = normalize_ppt_input_sources(
                source_file,
                ppt_state.get("raw_text", ""),
                collect_uploaded_images(uploaded_images_files),
                collect_uploaded_images(reference_image_files),
                template_file,
            )
            for key, value in normalized.items():
                ppt_state[key] = value
            ppt_state["input_mode"] = detect_ppt_input_mode(ppt_state)
            if template_file is not None:
                ppt_state["template_name"] = template_file.name

        with st.expander("当前页基础信息", expanded=True):
            ppt_state["page_title"] = st.text_input("当前页标题", value=ppt_state.get("page_title", ""), key="ppt_page_title")
            ppt_state["page_type"] = st.selectbox(
                "页面类型",
                PPT_PAGE_TYPES,
                index=PPT_PAGE_TYPES.index(ppt_state.get("page_type", "结果页")) if ppt_state.get("page_type", "结果页") in PPT_PAGE_TYPES else 5,
                key="ppt_page_type",
            )
            ppt_state["page_goal"] = st.text_area("本页目标", value=ppt_state.get("page_goal", ""), height=90, key="ppt_page_goal")

        with st.expander("补充说明", expanded=False):
            ppt_state["extra_notes"] = st.text_area(
                "补充说明",
                value=ppt_state.get("extra_notes", ""),
                height=110,
                placeholder="例如：这一页重点突出实验趋势，不要写太满；保留原图；结论要先出现在右上角。",
                key="ppt_extra_notes",
            )
            st.markdown(f"<div class='lab-quiet-note'>已载入图片素材：{len(ppt_state.get('uploaded_images', []))} 张 | 参考图：{len(ppt_state.get('reference_images', []))} 张</div>", unsafe_allow_html=True)
            if ppt_state.get("uploaded_source_name"):
                st.markdown(f"<div class='lab-quiet-note'>文档来源：{ppt_state['uploaded_source_name']}</div>", unsafe_allow_html=True)

        with st.expander("生成控制项", expanded=False):
            c1, c2 = st.columns(2)
            with c1:
                ppt_state["text_simplify_level"] = st.selectbox("文字简化强度", ["低", "中", "高"], index=["低", "中", "高"].index(ppt_state.get("text_simplify_level", "中")), key="ppt_text_simplify_level")
                ppt_state["keep_original_images"] = st.checkbox("保留原图", value=ppt_state.get("keep_original_images", True), key="ppt_keep_original_images")
                ppt_state["generate_aux_figure"] = st.checkbox("生成辅助图", value=ppt_state.get("generate_aux_figure", False), key="ppt_generate_aux_figure")
            with c2:
                ppt_state["info_density"] = st.selectbox("信息密度", ["低", "中", "高"], index=["低", "中", "高"].index(ppt_state.get("info_density", "中")), key="ppt_info_density")
                ppt_state["allow_external_support"] = st.checkbox("补充资料", value=ppt_state.get("allow_external_support", False), key="ppt_allow_external_support")
                ppt_state["apply_template_layout"] = st.checkbox("按模板排版", value=ppt_state.get("apply_template_layout", True), key="ppt_apply_template_layout")
            ppt_state["style_mode"] = st.selectbox(
                "风格模式",
                PPT_STYLE_MODES,
                index=PPT_STYLE_MODES.index(ppt_state.get("style_mode", "academic-paperskills")) if ppt_state.get("style_mode", "academic-paperskills") in PPT_STYLE_MODES else 0,
                key="ppt_style_mode",
            )
        st.markdown('</div>', unsafe_allow_html=True)

    with panel2:
        st.markdown('<div class="lab-main-card">', unsafe_allow_html=True)
        st.markdown('<div class="lab-panel-label">当前页处理</div>', unsafe_allow_html=True)
        st.markdown('<div class="lab-panel-note">当前页的文本变体、布局适配与预览都集中在这里。</div>', unsafe_allow_html=True)
        variants = ppt_state.get("processed_text_variants") or create_local_ppt_variants(ppt_state)
        tabs = st.tabs(["原文", "精简版", "标题版", "要点版", "结论先行版"])
        variant_keys = ["原文", "精简版", "标题版", "要点版", "结论先行版"]
        for tab, key in zip(tabs, variant_keys):
            with tab:
                st.text_area(key, value=variants.get(key, ""), height=150, disabled=True, key=f"ppt_variant_preview_{key}")

        with st.expander("模板适配", expanded=False):
            if ppt_state.get("template_analysis"):
                analysis = ppt_state["template_analysis"]
                st.markdown(f"**模板**：{analysis.get('filename', ppt_state.get('template_file_name', 'PPT模板'))}")
                if analysis.get("error"):
                    st.warning(f"模板分析失败：{analysis['error']}")
                else:
                    st.markdown(f"- 标题区位置：{analysis.get('title_zone', '待分析')}")
                    st.markdown(f"- 图文区布局：{analysis.get('content_layout', '待分析')}")
                    st.markdown(f"- 配色风格：{analysis.get('color_style', '待分析')}")
                    st.markdown(f"- 字体层级：{analysis.get('font_hierarchy', '待分析')}")
                    st.markdown(f"- 留白风格：{analysis.get('whitespace_style', '待分析')}")
            else:
                st.info("未上传模板，自动回退到默认科研页版式。")
            layout_names = list(PPT_DEFAULT_LAYOUTS.keys())
            default_layout = PPT_PAGE_TYPE_LAYOUTS.get(ppt_state.get("page_type", "结果页"), "结果+结论")
            current_layout = ppt_state.get("selected_layout") or default_layout
            ppt_state["selected_layout"] = st.selectbox(
                "当前页布局",
                layout_names,
                index=layout_names.index(current_layout) if current_layout in layout_names else layout_names.index(default_layout),
                format_func=lambda key: f"{key}｜{PPT_DEFAULT_LAYOUTS[key]['hint']}",
                key="ppt_selected_layout",
            )
            ppt_state["layout_analysis"] = resolve_ppt_layout(ppt_state)

        action1, action2, action3 = st.columns([1, 1, 1])
        with action1:
            generate_slide_clicked = st.button("生成当前页", type="primary", use_container_width=True, key="ppt_generate_slide")
        with action2:
            regenerate_slide_clicked = st.button("重新生成当前页", use_container_width=True, key="ppt_regenerate_slide")
        with action3:
            add_to_ppt_clicked = st.button("加入 PPT", use_container_width=True, key="ppt_add_to_assembly")

        if generate_slide_clicked or regenerate_slide_clicked:
            content_exists = any([
                (ppt_state.get("raw_text") or "").strip(),
                (ppt_state.get("uploaded_source_text") or "").strip(),
                ppt_state.get("uploaded_images"),
                ppt_state.get("reference_images"),
            ])
            if not content_exists:
                st.warning("请先输入当前页素材。")
            else:
                ppt_state["generation_counter"] += 1
                ppt_state["current_parent_id"] = None if generate_slide_clicked else ppt_state.get("current_slide_id")
                with st.spinner("正在提炼当前页内容..."):
                    ppt_state["processed_text_variants"] = generate_ppt_text_variants(ppt_state)
                    spec = build_ppt_slide_spec(ppt_state)
                    if ppt_state.get("generate_aux_figure") and GEMINI_API_KEY:
                        aux_result = generate_ppt_aux_figure(ppt_state, spec)
                        if aux_result.get("error"):
                            ppt_state["last_error"] = aux_result["error"]
                        else:
                            spec.update(aux_result)
                            ppt_state["last_error"] = ""
                    ppt_state["current_slide_struct"] = spec
                    ppt_state["current_slide_markdown"] = f"# {spec['title']}\n\n{spec.get('body', '')}"
                    ppt_state["current_slide_preview_html"] = render_ppt_slide_preview(spec)
                    ppt_state["current_slide_snapshot"] = deepcopy(spec)
                    ppt_state["current_slide_id"] = f"ppt-{ppt_state['generation_counter']}"
                    append_ppt_history_entry(ppt_state)

        st.markdown('<div class="lab-status-pills">' +
            f'<span class="lab-status-pill"><strong>布局</strong>{ppt_state.get("selected_layout") or PPT_PAGE_TYPE_LAYOUTS.get(ppt_state.get("page_type", "结果页"), "结果+结论")}</span>' +
            f'<span class="lab-status-pill"><strong>类型</strong>{ppt_state.get("page_type", "结果页")}</span>' +
            f'<span class="lab-status-pill"><strong>输入</strong>{detect_ppt_input_mode(ppt_state)}</span>' +
            '</div>', unsafe_allow_html=True)
        if ppt_state.get("current_slide_preview_html"):
            components.html(ppt_state["current_slide_preview_html"], height=420)
            st.markdown(f"<div class='lab-quiet-note'>标签：{' / '.join((ppt_state.get('current_slide_struct') or {}).get('tags', []))}</div>", unsafe_allow_html=True)
        else:
            st.info("尚未生成当前页。先在左侧输入素材，再点“生成当前页”。")

        ppt_state["micro_tune_request"] = st.text_area(
            "微调当前页",
            value=ppt_state.get("micro_tune_request", ""),
            height=90,
            placeholder="例如：标题更学术化；要点减到 3 条；把结论提前；图片区更突出。",
            key="ppt_micro_tune_request",
        )
        tune1, tune2, tune3 = st.columns(3)
        with tune1:
            micro_tune_clicked = st.button("微调当前页", use_container_width=True, key="ppt_micro_tune")
        with tune2:
            copy_slide_clicked = st.button("复制当前页文案", use_container_width=True, key="ppt_copy_slide")
        with tune3:
            download_slide_ready = bool(ppt_state.get("current_slide_struct"))
            if download_slide_ready:
                st.download_button(
                    "下载当前页",
                    data=json.dumps(sanitize_for_json(ppt_state.get("current_slide_struct") or {}), ensure_ascii=False, indent=2).encode("utf-8"),
                    file_name=build_export_filename("xueyan_ppt_slide", ppt_state.get("page_title") or "current-slide", suffix=".json"),
                    mime="application/json",
                    use_container_width=True,
                    key="ppt_download_slide_json",
                )

        if micro_tune_clicked:
            if not ppt_state.get("current_slide_struct"):
                st.warning("请先生成当前页。")
            elif not ppt_state.get("micro_tune_request", "").strip():
                st.warning("请先输入微调要求。")
            else:
                ppt_state["generation_counter"] += 1
                ppt_state["current_parent_id"] = ppt_state.get("current_slide_id")
                prompt = build_ppt_micro_tune_prompt(ppt_state)
                variants = deepcopy(ppt_state.get("processed_text_variants") or {})
                variants["精简版"] = (variants.get("精简版") or choose_ppt_body_text(ppt_state)) + f"\n\n微调要求：{ppt_state['micro_tune_request']}"
                ppt_state["processed_text_variants"] = variants
                spec = build_ppt_slide_spec(ppt_state)
                spec["notes"] = prompt
                if ppt_state.get("current_slide_struct", {}).get("aux_figure_url"):
                    spec["aux_figure_url"] = ppt_state["current_slide_struct"].get("aux_figure_url", "")
                    spec["aux_figure_bytes"] = ppt_state["current_slide_struct"].get("aux_figure_bytes", b"")
                ppt_state["current_slide_struct"] = spec
                ppt_state["current_slide_markdown"] = f"# {spec['title']}\n\n{spec.get('body', '')}"
                ppt_state["current_slide_preview_html"] = render_ppt_slide_preview(spec)
                ppt_state["current_slide_snapshot"] = deepcopy(spec)
                ppt_state["current_slide_id"] = f"ppt-{ppt_state['generation_counter']}"
                append_ppt_history_entry(ppt_state)

        if copy_slide_clicked and ppt_state.get("current_slide_markdown"):
            render_copy_text(ppt_state["current_slide_markdown"], "ppt_copy_current_slide_btn")
            st.button("📋 复制当前页文案", key="ppt_copy_current_slide_btn", use_container_width=True)

        if add_to_ppt_clicked:
            if not ppt_state.get("current_slide_struct"):
                st.warning("请先生成当前页。")
            else:
                add_current_slide_to_assembly(ppt_state)
                st.success("已加入 PPT 组装区。")

        with st.expander("当前页历史", expanded=False):
            render_ppt_history_panel(ppt_state)
        st.markdown('</div>', unsafe_allow_html=True)

    with panel3:
        st.markdown('<div class="lab-secondary-card">', unsafe_allow_html=True)
        st.markdown('<div class="lab-panel-label">输出与 PPT 组装</div>', unsafe_allow_html=True)
        st.markdown('<div class="lab-panel-note">导出、页序调整和组装管理都放在这里。</div>', unsafe_allow_html=True)
        assembly_pages = ppt_state.get("assembly_pages", [])
        export_ready = bool(assembly_pages)
        export_col, status_col = st.columns([1.2, 1])
        with export_col:
            if PPTX_AVAILABLE and export_ready:
                deck_bytes = export_pptx_deck(ppt_state)
                st.download_button(
                    "导出 PPTX",
                    data=deck_bytes,
                    file_name=build_export_filename("xueyan_ppt_deck", ppt_state.get("page_title") or "research-deck", suffix=".pptx"),
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True,
                    key="ppt_export_deck_top",
                )
            elif PPTX_AVAILABLE:
                st.button("导出 PPTX", disabled=True, use_container_width=True, key="ppt_export_deck_disabled")
            else:
                st.button("导出 PPTX", disabled=True, use_container_width=True, key="ppt_export_deck_missing_dep")
        with status_col:
            state_text = "已具备导出条件" if export_ready else "至少先加入 1 页后才能导出"
            st.markdown(f"<div class='lab-quiet-note'>{state_text}</div>", unsafe_allow_html=True)
        st.markdown(f"<div class=\"lab-status-pills\"><span class=\"lab-status-pill\"><strong>已加入页数</strong>{len(assembly_pages)}</span>{f'<span class=\"lab-status-pill\"><strong>当前页</strong>{ppt_state["current_slide_id"]}</span>' if ppt_state.get("current_slide_id") else ''}</div>", unsafe_allow_html=True)
        if not PPTX_AVAILABLE:
            st.warning("当前环境缺少 python-pptx，暂时不能导出 PPTX。先安装 requirements.txt 依赖后即可使用。")
        if not assembly_pages:
            st.info("当前还没有加入 PPT 的页面。")
        for idx, slide in enumerate(assembly_pages, start=1):
            title = slide.get("page_title") or f"页面 {idx}"
            with st.expander(f"{idx}. {title}", expanded=False):
                new_title = st.text_input("页面重命名", value=title, key=f"ppt_assembly_title_{slide['id']}")
                slide["page_title"] = new_title
                st.markdown(f"<div class='lab-quiet-note'>类型：{slide.get('page_type', '结果页')} | 布局：{slide.get('layout_name', '结果+结论')}</div>", unsafe_allow_html=True)
                if slide.get("preview_html"):
                    components.html(slide["preview_html"], height=260)
                c1, c2 = st.columns(2)
                with c1:
                    if st.button("上移", key=f"ppt_move_up_{slide['id']}", use_container_width=True):
                        move_assembly_slide(ppt_state, slide["id"], -1)
                        st.rerun()
                with c2:
                    if st.button("下移", key=f"ppt_move_down_{slide['id']}", use_container_width=True):
                        move_assembly_slide(ppt_state, slide["id"], 1)
                        st.rerun()
                c3, c4 = st.columns(2)
                with c3:
                    if st.button("复制页面", key=f"ppt_dup_{slide['id']}", use_container_width=True):
                        duplicate_assembly_slide(ppt_state, slide["id"])
                        st.rerun()
                with c4:
                    if st.button("删除页面", key=f"ppt_del_{slide['id']}", use_container_width=True):
                        remove_assembly_slide(ppt_state, slide["id"])
                        st.rerun()

        if assembly_pages and PPTX_AVAILABLE:
            deck_bytes = export_pptx_deck(ppt_state)
            st.download_button(
                "导出 PPTX（底部）",
                data=deck_bytes,
                file_name=build_export_filename("xueyan_ppt_deck", ppt_state.get("page_title") or "research-deck", suffix=".pptx"),
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
                key="ppt_export_deck_bottom",
            )
        st.markdown('</div>', unsafe_allow_html=True)


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
    elif st.session_state.engine_mode == FORMATTING_PAGE:
        render_formatting_engine()
    elif st.session_state.engine_mode == VIZ_PAGE:
        render_viz_engine()
    else:
        render_ppt_engine()

    st.markdown(
        """
---
<div style="text-align: center; color: #64748b; font-size: 0.8rem; padding: 2rem 0;">
    <p><strong>🧪 学研·工科科研助手 v4.0</strong></p>
    <p>四引擎架构 · 写作协作 OS · 确定性格式对齐 · 视觉实验室 · PPT大师</p>
</div>
""",
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
