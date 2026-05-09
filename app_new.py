"""
学研·工科科研助手 v4.0 (YanYu OS)
15+功能矩阵 · 板块锚定 · 零篡位执行 · 影子合著者 · 语言基因深度提取
"""

import os
import re
import time
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Dict, List, Tuple, Optional
from collections import Counter
import json

import anthropic
import fitz  # pymupdf
import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml.shared import OxmlElement
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type

# ── Configuration ────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="学研·工科科研助手 v4.0",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Claude API 配置 (使用 Claude Code 相同接口)
CLAUDE_API_KEY = (os.getenv("ANTHROPIC_API_KEY") or os.getenv("ANTHROPIC_AUTH_TOKEN") or "").strip()
CLAUDE_BASE_URL = os.getenv("ANTHROPIC_BASE_URL", "https://api.anthropic.com").rstrip("/")
CLAUDE_MODEL = os.getenv("ANTHROPIC_MODEL", "claude-sonnet-4-6")

# ── Session State ───────────────────────────────────────────────────────────
DEFAULT_STATES = {
    "history": [],              # 历史时光机（按板块+功能分类）
    "reference_docs": {},       # 标杆文献库 {filename: style_analysis}
    "current_input": "",        # 当前输入
    "loaded_skills": False,     # Skills加载状态
}

for key, default in DEFAULT_STATES.items():
    if key not in st.session_state:
        st.session_state[key] = default

# ── 常量定义 ─────────────────────────────────────────────────────────────────

# 15+ 功能矩阵（单一职责）
FUNCTION_MATRIX = {
    "📝 中转英翻译": {
        "description": "仅执行CN→EN语种转换，禁止润色",
        "rules": ["❌ 禁止自行润色", "❌ 禁止改变原意", "✅ 仅语言转换"]
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

# 板块定义
SECTIONS = {
    "摘要": {"focus": "开门见山，数据支撑", "max_words": 250},
    "引言": {"focus": "背景转折，研究空白", "max_words": 800},
    "方法": {"focus": "流程清晰，参数精确", "max_words": 1500},
    "结果": {"focus": "数据客观，图表引用", "max_words": 1200},
    "讨论": {"focus": "深度解读，文献对比", "max_words": 1000},
    "结论": {"focus": "总结贡献，展望未来", "max_words": 300},
}

# 学科领域与术语硬锁
DOMAINS = {
    "🔋 能源电池": {
        "hard_lock": [r'\$Li\+\$', r'\$Na\+\$', 'NCM523', 'NCM622', 'capacity retention', 'intercalation'],
        "focus": "电化学性能"
    },
    "🏗️ 固废/土木": {
        "hard_lock": ['GGBS', '水化动力学', 'compressive strength', 'C-S-H', 'pozzolanic'],
        "focus": "材料性能"
    },
    "🔩 机械/材料": {
        "hard_lock": ['tensile strength', 'grain boundary', 'dislocation', 'microstructure'],
        "focus": "力学性能"
    },
    "🧪 催化/化工": {
        "hard_lock": ['TOF', 'turnover frequency', 'BET', 'heterogeneous catalysis'],
        "focus": "催化性能"
    },
    "🌊 环境工程": {
        "hard_lock": ['COD', 'BOD', 'MBR', 'activated sludge', 'removal efficiency'],
        "focus": "处理效果"
    },
    "📱 电子半导体": {
        "hard_lock": ['MOSFET', 'bandgap', 'GaN', 'SiC', 'carrier mobility'],
        "focus": "电学性能"
    },
    "🧬 生物材料": {
        "hard_lock": ['hydrogel', 'cell adhesion', 'biocompatibility', 'MTT assay'],
        "focus": "生物相容性"
    },
    "🧠 机器学习": {
        "hard_lock": ['CNN', 'Transformer', 'overfitting', 'gradient descent', 'F1-score'],
        "focus": "模型性能"
    },
    "💻 软件工程": {
        "hard_lock": ['CI/CD', 'Agile', 'DevOps', 'API', 'microservices'],
        "focus": "工程实践"
    },
}

# ── 深度技能加载系统 ────────────────────────────────────────────────────────

@st.cache_resource
def load_local_skills() -> Dict[str, str]:
    """递归加载本地学术库"""
    skills = {}
    base_path = Path("./awesome-ai-research-writing")

    if not base_path.exists():
        return skills

    for md_file in base_path.rglob("*.md"):
        try:
            content = md_file.read_text(encoding="utf-8")
            key = md_file.stem
            skills[key] = content
        except:
            continue

    return skills

LOCAL_SKILLS = load_local_skills()

# Humanizer 核心规则（硬编码注入）
HUMANIZER_RULES = """
## Humanizer 核心规则（零篡位）

### AI 痕迹识别清单
1. **过度强调意义**: "具有...意义", "crucial", "critical", "important", "significant"
2. **AI 常用词**: "delve", "landscape", "realm", "leverage", "underscore", "utilize"
3. **破折号滥用**: 解释性插入语过多（—which, —that）
4. **三点式堆砌**: A, B, and C 结构过多
5. **促销腔**: "exciting", "promising", "novel", "groundbreaking"
6. **空洞-ing分析**: "indicating", "suggesting", "implying" 滥用

### 人味注入策略
- 承认不确定性: "suggest", "may", "potentially", "appears to"
- 节奏变化: 长短句交替，避免单调
- 自然过渡: 删除机械连接词（Firstly, Secondly, Furthermore）
- 简洁动词: show, use, find, make (而非 demonstrate, utilize, discover, fabricate)

### 语序优化原则
- **中文输入**: 优化中文语序，保持中文输出
- **英文输入**: 增强英文节奏感，保持英文输出
- **零翻译**: 严格禁止语种转换
"""

# ML Paper Writing Checklist（仅用于 Results 板块验证）
ML_CHECKLIST = """
## ML Paper Writing Checklist (仅 Results 板块)

### 引用验证
- [ ] 引用格式统一（NeurIPS/ICML 格式）
- [ ] 所有方法引用有明确出处
- [ ] 基线模型正确引用

### 图表规范
- [ ] 色盲友好配色（viridis, magma）
- [ ] 坐标轴标签清晰
- [ ] 图例位置合理
- [ ] 误差线标注

### 数据完整性
- [ ] 超参数表完整
- [ ] 实验设置可复现
- [ ] 统计显著性标注
"""


# ── 术语硬锁保护系统 v2.0 ────────────────────────────────────────────────────────

def protect_hard_terms(text: str, domain: str) -> Tuple[str, Dict[str, str]]:
    """硬锁保护工科术语，绝对禁止改动 v2.0"""
    protected = text
    mapping = {}
    counter = 0

    # 1. 保护 LaTeX 公式（多模式）
    # 行内公式: $...$
    for match in re.finditer(r'\$[^$]+?\$', text):
        placeholder = f"__LOCK_{counter}__"
        mapping[placeholder] = match.group()
        protected = protected.replace(match.group(), placeholder, 1)
        counter += 1

    # 块级公式: $$...$$ 或 \[...\]
    for match in re.finditer(r'\$\$[^$]+?\$\$|\\\[.*?\\\]', text, re.DOTALL):
        placeholder = f"__LOCK_{counter}__"
        mapping[placeholder] = match.group()
        protected = protected.replace(match.group(), placeholder, 1)
        counter += 1

    # LaTeX命令: \frac{}{}, \sum_{}^{}, \int_{}^{}, etc.
    for match in re.finditer(r'\\[a-zA-Z]+(?:\{[^}]*\}|\[[^\]]*\]|_[^{}s]|[_^]{[^}]*})?', text):
        placeholder = f"__LOCK_{counter}__"
        mapping[placeholder] = match.group()
        protected = protected.replace(match.group(), placeholder, 1)
        counter += 1

    # 2. 保护化学式和材料名称
    chemical_patterns = [
        r'\b[A-Z][a-z]?\d*(?:_[\d-]+|[+\-]?\d*)?\b',  # Li+, Na+, NCM523
        r'\b[A-Z][a-z]?\d*(?:_[\d-]+|[+\-]?\d*)?\s*[A-Z][a-z]?\d*',  # LiCoO2, NaMnO2
    ]
    for pattern in chemical_patterns:
        for match in re.finditer(pattern, text):
            # 排除普通单词
            if len(match.group()) > 2 and any(c.isdigit() or c in '_+-' for c in match.group()):
                placeholder = f"__LOCK_{counter}__"
                mapping[placeholder] = match.group()
                protected = protected.replace(match.group(), placeholder, 1)
                counter += 1

    # 3. 保护领域特定术语
    if domain in DOMAINS:
        for term in DOMAINS[domain]["hard_lock"]:
            if isinstance(term, str):
                # 精确匹配
                if term in protected:
                    placeholder = f"__LOCK_{counter}__"
                    mapping[placeholder] = term
                    protected = protected.replace(term, placeholder, 1)
                    counter += 1

    # 4. 保护测量单位和数值组合
    unit_patterns = [
        r'\d+\s*(?:mAh·g[^-1]|mAh g[^-1]|mA h[^-1]|A h[^-1])',
        r'\d+\s*(?:°C|K|MPa|GPa|kPa|Pa)',
        r'\d+\s*(?:wt%|vol%|at%|mol%)',
        r'\d+\s*(?:nm|μm|mm|cm|m|km)',
        r'\d+\s*(?:Hz|kHz|MHz|GHz)',
        r'\d+\s*(?:S cm[^-1]|S m[^-1])',
    ]
    for pattern in unit_patterns:
        for match in re.finditer(pattern, text):
            placeholder = f"__LOCK_{counter}__"
            mapping[placeholder] = match.group()
            protected = protected.replace(match.group(), placeholder, 1)
            counter += 1

    return protected, mapping


def restore_hard_terms(text: str, mapping: Dict[str, str]) -> str:
    """恢复硬锁保护的术语"""
    result = text
    for placeholder, original in mapping.items():
        result = result.replace(placeholder, original)
    return result


# ── 语言智能识别 ─────────────────────────────────────────────────────────────

def detect_language(text: str) -> str:
    """智能识别输入语言"""
    # 统计中文字符
    chinese_chars = len(re.findall(r'[\u4e00-\u9fff]', text))
    total_chars = len(text.strip())

    if total_chars == 0:
        return "unknown"

    chinese_ratio = chinese_chars / total_chars

    if chinese_ratio > 0.3:
        return "zh"
    else:
        return "en"


# ── 影子合著者：标杆文献分析 v2.0 ───────────────────────────────────────────────

def extract_language_genes(text: str) -> Dict:
    """深度提取文献的语言基因"""
    genes = {
        "sentence_length_pattern": [],
        "connecting_words": Counter(),
        "academic_phrases": Counter(),
        "voice_pattern": {"passive": 0, "active": 0},
        "citation_style": Counter(),
        "complexity_markers": Counter(),
    }

    sentences = [s.strip() for s in text.split('.') if s.strip()]
    for sent in sentences[:100]:  # 分析前100句
        # 句长模式
        words = len(sent.split())
        genes["sentence_length_pattern"].append(words)

        # 连接词统计
        connectors = [
            'however', 'therefore', 'furthermore', 'moreover', 'additionally',
            'consequently', 'subsequently', 'meanwhile', 'nevertheless',
            'thus', 'hence', 'accordingly', 'otherwise', 'moreover'
        ]
        for conn in connectors:
            if re.search(rf'\b{conn}\b', sent.lower()):
                genes["connecting_words"][conn] += 1

        # 学术短语
        phrases = [
            'it is worth noting', 'it should be mentioned', 'previous studies',
            'recent work', 'to the best of', 'knowledge', 'suggest that',
            'indicate that', 'demonstrate that', 'reveal that'
        ]
        for phrase in phrases:
            if phrase in sent.lower():
                genes["academic_phrases"][phrase] += 1

        # 主动/被动语态
        if re.search(r'\b(was|were)\s+\w+ed\b', sent, re.IGNORECASE):
            genes["voice_pattern"]["passive"] += 1
        else:
            genes["voice_pattern"]["active"] += 1

        # 引用风格
        citations = re.findall(r'\[\d+\]|\([^)]+\d+[^)]*\)|\w+\s+et\s+al\.', sent)
        for cite in citations:
            genes["citation_style"][cite[:20]] += 1

        # 复杂性标记
        complexity = [
            'although', 'while', 'despite', 'whereas', 'not only but also',
            'either or', 'neither nor', 'whether or'
        ]
        for marker in complexity:
            if re.search(rf'\b{marker}\b', sent.lower()):
                genes["complexity_markers"][marker] += 1

    # 计算统计数据
    if genes["sentence_length_pattern"]:
        genes["avg_sentence_length"] = sum(genes["sentence_length_pattern"]) / len(genes["sentence_length_pattern"])
        genes["sentence_length_std"] = (sum((x - genes["avg_sentence_length"])**2 for x in genes["sentence_length_pattern"]) /
                                       len(genes["sentence_length_pattern"]))**0.5
    else:
        genes["avg_sentence_length"] = 0
        genes["sentence_length_std"] = 0

    return genes


def analyze_reference_paper(docx_bytes: bytes, filename: str) -> str:
    """分析标杆文献的写作风格 v2.0 - 语言基因深度提取"""
    try:
        doc = Document(BytesIO(docx_bytes))
        paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
        text = "\n".join(paragraphs[:80])  # 取前80段

        # 提取风格特征
        lang = detect_language(text)
        genes = extract_language_genes(text)

        # 生成风格分析报告
        style_analysis = f"""
## 🧬 标杆文献语言基因提取: {filename}

### 📊 基础信息
- **段落数**: {len(paragraphs)}
- **语言**: {"中文" if lang == "zh" else "英文"}
- **平均句长**: {genes.get('avg_sentence_length', 0):.1f} 词/句
- **句长波动**: {genes.get('sentence_length_std', 0):.1f} (标准差)

### 🎵 叙事节奏
- **平均句长**: {genes.get('avg_sentence_length', 0):.1f} 词/句
- **句式多样性**: {"高" if genes.get('sentence_length_std', 0) > 10 else "中" if genes.get('sentence_length_std', 0) > 5 else "低"}

### 🔗 连接词偏好
{', '.join([f'{k}({v})' for k, v in genes['connecting_words'].most_common(5)]) if genes['connecting_words'] else '无明显偏好'}

### 🎓 学术短语
{', '.join([f'{k}({v})' for k, v in genes['academic_phrases'].most_common(5)]) if genes['academic_phrases'] else '无常用短语'}

### 📢 语态倾向
- **主动**: {genes['voice_pattern']['active']} 句
- **被动**: {genes['voice_pattern']['passive']} 句
- **主动比**: {genes['voice_pattern']['active'] / max(1, genes['voice_pattern']['active'] + genes['voice_pattern']['passive']) * 100:.1f}%

### 📚 引用风格
{', '.join([k for k, v in genes['citation_style'].most_common(3)]) if genes['citation_style'] else '未检测到明确模式'}

### 🧠 复杂性标记
{', '.join([f'{k}({v})' for k, v in genes['complexity_markers'].most_common(5)]) if genes['complexity_markers'] else '无明显复杂结构'}

### 📝 样本文本
{text[:500]}...

### 🎯 模仿策略
1. **句长控制**: 目标平均句长 {genes.get('avg_sentence_length', 0):.0f} 词，波动 ±{genes.get('sentence_length_std', 0):.0f}
2. **连接词**: 优先使用 {', '.join([k for k, v in genes['connecting_words'].most_common(3)]) if genes['connecting_words'] else '自然过渡'}
3. **语态**: {"偏好主动语态" if genes['voice_pattern']['active'] > genes['voice_pattern']['passive'] * 1.5 else "平衡主动/被动"}
4. **复杂性**: {"高复杂句式" if sum(genes['complexity_markers'].values()) > 5 else "中等复杂度"}
"""
        return style_analysis
    except Exception as e:
        return f"分析失败: {e}"


# ── 零篡位 Prompt 生成器 ───────────────────────────────────────────────────

def create_strict_prompt(
    function: str,
    input_text: str,
    section: str,
    domain: str,
    reference_styles: List[str],
    lang: str
) -> str:
    """创建单一职责 prompt，严格功能隔离 v2.0"""

    # 语言锁定强制声明
    lang_lock = {
        "zh": """## ⚠️ 语言锁定（强制执行）
1. **检测到中文输入**：全程使用中文处理
2. **零翻译**：严格禁止任何中→英翻译行为
3. **输出语言锁定**：所有输出必须是中文
4. **术语保护**：英文专有名词保持原样""",
        "en": """## ⚠️ Language Lock (Strict Enforcement)
1. **English Input Detected**: Process entirely in English
2. **Zero Translation**: Strictly forbidden from translating to any other language
3. **Output Language Lock**: All output must be in English
4. **Term Protection**: Preserve non-English technical terms as-is"""
    }

    # 基础注入
    base_injection = f"""
# XueYan OS v4.0 - 零篡位执行系统
## 🔒 当前配置
- **功能模式**: {function}
- **论文板块**: {section}
- **研究领域**: {domain}
- **输入语言**: {"中文" if lang == "zh" else "英文"}
- **术语保护**: 已激活硬锁保护

## 📚 本地学术库
已加载 {len(LOCAL_SKILLS)} 个学术 Skills

{lang_lock.get(lang, '')}
"""

    # 添加 Humanizer 规则
    if "Humanizer" in function or "去AI味" in function:
        base_injection += f"""
## 🎯 Humanizer 核心规则（{lang.upper()}专用）
### AI 痕迹识别清单
1. **过度强调意义**: "具有...意义", "crucial", "critical", "important", "significant"
2. **AI 常用词**: "delve", "landscape", "realm", "leverage", "underscore", "utilize"
3. **破折号滥用**: 解释性插入语过多（—which, —that）
4. **三点式堆砌**: A, B, and C 结构过多
5. **促销腔**: "exciting", "promising", "novel", "groundbreaking"
6. **空洞-ing分析**: "indicating", "suggesting", "implying" 滥用

### 人味注入策略
- **承认不确定性**: "suggest", "may", "potentially", "appears to"
- **节奏变化**: 长短句交替，避免单调
- **自然过渡**: 删除机械连接词（Firstly, Secondly, Furthermore）
- **简洁动词**: show, use, find, make (而非 demonstrate, utilize, discover, fabricate)

### 语序优化原则
- **当前语言**: {"中文" if lang == "zh" else "英文"}
- **优化目标**: {"优化中文语序，保持中文输出" if lang == "zh" else "增强英文节奏感，保持英文输出"}
- **零翻译**: 严格禁止语种转换
"""
    else:
        base_injection += HUMANIZER_RULES + "\n"

    # 如果是 Results 板块且有 ML 写作需求，添加 checklist
    if section == "结果" and "ML" in domain:
        base_injection += ML_CHECKLIST + "\n"

    # 添加标杆文献风格（影子合著者：仅限起草/构思类功能）
    shadow_functions = {"✍️ 逐段起草", "💡 研究想法构思", "📄 节节头脑风暴"}
    if reference_styles and function in shadow_functions:
        base_injection += "\n## 📚 影子合著者：标杆文献语言基因\n"
        base_injection += "\n".join(reference_styles)
        base_injection += "\n\n**模仿指令**: 严格模仿上述标杆文献的句长、连接词、语态和复杂性偏好。\n"

    # 板块特定指导
    section_info = SECTIONS.get(section, {})
    if section_info:
        base_injection += f"""
## 📝 {section} 板块写作要求
- **核心重点**: {section_info.get('focus', '')}
- **建议长度**: {section_info.get('max_words', '')} 词
- **写作策略**: 针对该板块的叙事逻辑进行优化
"""

    # 功能特定规则（零篡位）
    func_rules = FUNCTION_MATRIX.get(function, {})
    if func_rules:
        base_injection += f"""
## 🎯 {function} 功能规则（单一职责）
{chr(10).join(f'- {rule}' for rule in func_rules.get('rules', []))}

**⛔ 零篡位强制声明**: 严格按照上述规则执行，绝不越界到其他功能领域。
"""

    # 功能特定模板
    function_templates = {
        "📝 中转英翻译": f"""{base_injection}

# 📝 任务：中转英翻译（CN→EN ONLY）
## 输入文本
{input_text}

## 输出要求
1. ✅ 仅执行 CN→EN 语言转换
2. ✅ 保持原意完整
3. ✅ 术语准确性
4. ✅ 学术语体
5. ❌ 禁止额外润色
6. ❌ 禁止改变原意

## 输出格式
```
[翻译后的英文文本]
```

## 关键术语对照
[列出5-10个核心术语的翻译]
""",

        "✨ 表达润色": f"""{base_injection}

# ✨ 任务：表达润色（同语言优化）
## 输入文本
{input_text}

## 润色要求
1. ✅ 提升学术地道性
2. ✅ 替换高级动词
3. ✅ {"保持中文，优化中文" if lang == "zh" else "保持英文，优化英文"}
4. ✅ 保持原意
5. ❌ 严格禁止翻译

## 输出格式
```
[润色后的文本]
```

## 修改说明
[列出主要修改点：动词替换、结构调整等]
""",

        "🔍 逻辑检查": f"""{base_injection}

# 🔍 任务：逻辑检查（不修改文本）
## 输入文本
{input_text}

## 检查要点
1. ✅ 因果链条完整性
2. ✅ 论据与论点一致性
3. ✅ 逻辑衔接流畅性
4. ✅ 术语使用准确性

## 输出格式
```
✅ 逻辑检查通过
```
或
```
❌ 发现以下问题:
1. [具体问题]
2. [具体问题]
...
```

**⚠️ 注意**: 仅检查，不修改文本。
""",

        "🤖 去AI味 (Humanizer)": f"""{base_injection}

# 🤖 任务：去AI化处理（零翻译）
## 输入文本
{input_text}

## 去AI化要求
{"1. ✅ 优化中文语序，保持中文输出" if lang == "zh" else "1. ✅ 增强英文节奏感，保持英文输出"}
2. ✅ 消除 AI 常用词
3. ✅ 删除机械连接词
4. ✅ 节奏变化
5. ✅ 承认不确定性
6. ❌ 严格禁止翻译

## 输出格式
```
[去AI化后的文本]
```

## 修改说明
[列出3-5个关键修改点及理由]
""",

        "👨‍⚖️ Reviewer视角": f"""{base_injection}

# 👨‍⚖️ 任务：Reviewer视角审视
## 输入文本
{input_text}

## 审视要点
1. ✅ 学术规范性
2. ✅ 论证强度
3. ✅ 表达清晰度
4. ✅ 潜在问题

## 输出格式
## 📝 审稿意见
[总体评价]

## 🔍 具体建议
1. [具体建议]
2. [具体建议]
...

## ✅ 优点
[列出优点]
""",

        "✍️ 逐段起草": f"""{base_injection}

# ✍️ 任务：逐段起草（影子写作模式）
## 输入大纲
{input_text}

## 起草要求
1. ✅ 将大纲扩充为正式段落
2. ✅ 模仿标杆文献风格（句长、连接词、语态）
3. ✅ 符合 {section} 板块特点
4. ✅ 保持学术严谨性
5. ✅ {"中文输入输出中文，英文输入输出英文" if lang != "unknown" else "保持原语言"}

## 输出格式
```
[扩充后的正式段落]
```

## 模仿说明
[说明模仿了标杆文献的哪些特征：句长、连接词、语态等]
""",

        "🎯 精修模式": f"""{base_injection}

# 🎯 任务：深度精修（顶刊水准）
## 输入文本
{input_text}

## 精修要求
1. ✅ 达到顶刊出版水准
2. ✅ 模仿标杆文献风格
3. ✅ 符合 {section} 板块特点
4. ✅ 术语精确使用
5. ✅ {"保持原语言，优化原语言" if lang != "unknown" else "保持原语言"}

## 输出格式
```
[精修后的文本]
```

## 修改说明
[详细说明修改内容和理由]
""",

        "📝 引用验证": f"""{base_injection}

# 📝 任务：引用格式验证
## 输入文本
{input_text}

## 验证要点
1. ✅ 引用格式统一
2. ✅ 引用完整性
3. ✅ 引用相关性

## 输出格式
```
✅ 引用格式验证通过
```
或
```
❌ 发现以下问题:
1. [具体问题]
2. [具体问题]
...
```
""",

        "🎨 图表规范检查": f"""{base_injection}

# 🎨 任务：图表描述规范检查
## 输入文本
{input_text}

## 检查要点
1. ✅ 图表引用完整性
2. ✅ 图表描述清晰度
3. ✅ 色盲友好性
4. ✅ 坐标轴标注

## 输出格式
```
✅ 图表描述规范
```
或
```
❌ 发现以下问题:
1. [具体问题]
2. [具体问题]
...
```
""",

        "📋 Redlining修订": f"""{base_injection}

# 📋 任务：Redlining修订（显示修改痕迹）
## 输入文本
{input_text}

## 修订要求
1. ✅ 显示修改痕迹
2. ✅ 提供对比视图
3. ✅ 保持原语言
4. ✅ 说明修改理由

## 输出格式
```
## 修改后文本
[修改后的文本]

## 修改说明
[详细列出所有修改点]
```
""",
    }

    # 默认模板
    default_template = f"""{base_injection}

# 🎯 任务：{function}
## 输入文本
{input_text}

## 输出要求
严格按照 {function} 功能规则执行
{"保持原语言，绝不翻译" if lang != "unknown" else ""}

## 输出格式
```
[处理结果]
```

## 修改说明
[简要说明处理内容]
"""

    return function_templates.get(function, default_template)


# ── 原生 API 调用（带重试）──────────────────────────────────────────────────

def get_client():
    if not CLAUDE_API_KEY:
        st.error("❌ 未检测到 ANTHROPIC_AUTH_TOKEN")
        st.stop()
    return anthropic.Anthropic(
        api_key=CLAUDE_API_KEY,
        base_url=CLAUDE_BASE_URL if CLAUDE_BASE_URL != "https://api.anthropic.com" else None
    )


@retry(
    stop=stop_after_attempt(3),
    wait=wait_exponential(multiplier=1, min=4, max=60),
    retry=retry_if_exception_type((anthropic.APITimeoutError, anthropic.InternalServerError)),
)
def call_api(prompt: str, timeout: int = 180) -> str:
    client = get_client()
    message = client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=8192,
        temperature=0.2,
        timeout=timeout,
        messages=[{"role": "user", "content": prompt}]
    )
    return message.content[0].text


# ── 文件处理 ────────────────────────────────────────────────────────────────

def extract_text(file_bytes: bytes, filename: str) -> str:
    """提取文件文本"""
    try:
        if filename.endswith('.pdf'):
            doc = fitz.open(stream=file_bytes, filetype="pdf")
            return "\n\n".join([p.get_text() for p in doc])
        elif filename.endswith('.docx'):
            doc = Document(BytesIO(file_bytes))
            return "\n\n".join([p.text for p in doc.paragraphs if p.text.strip()])
    except Exception as e:
        return f"解析失败: {e}"
    return ""


def create_docx_with_redlines(original: str, revised: str, metadata: dict) -> bytes:
    """创建带修订痕迹的Word文档"""
    doc = Document()
    doc.add_heading(f"学研·修订模式 - {metadata['function']}", 0)

    # 元信息
    p = doc.add_paragraph()
    p.add_run(f"板块: {metadata['section']}\n")
    p.add_run(f"领域: {metadata['domain']}\n")
    p.add_run(f"时间: {metadata['timestamp']}\n")

    doc.add_heading("原文", 1)
    doc.add_paragraph(original)

    doc.add_heading("修改后", 1)
    doc.add_paragraph(revised)

    # 简单差异标注（使用颜色）
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

    doc_io = BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io.read()


def create_docx(content: str, metadata: dict) -> bytes:
    """创建 Word 文档"""
    doc = Document()
    doc.add_heading(f"学研·工科科研助手 - {metadata['function']}", 0)

    # 元信息
    p = doc.add_paragraph()
    p.add_run(f"板块: {metadata['section']}\n")
    p.add_run(f"领域: {metadata['domain']}\n")
    p.add_run(f"时间: {metadata['timestamp']}\n")

    doc.add_heading("处理结果", 1)
    doc.add_paragraph(content)

    doc_io = BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io.read()


# ── UI Styling ───────────────────────────────────────────────────────────────

st.markdown("""
<style>
    .stApp { background: linear-gradient(135deg, #f8fafc 0%, #e2e8f0 100%); }
    .main-header {
        text-align: center; padding: 2rem 0;
        background: linear-gradient(135deg, #1e3a8a 0%, #3b82f6 100%);
        color: white; border-radius: 12px; margin-bottom: 1.5rem;
        box-shadow: 0 10px 25px rgba(30, 58, 138, 0.2);
    }
    .section-tab { background: white; padding: 0.5rem 1rem; border-radius: 8px; font-weight: 600; }
    .comparison-box { background: white; padding: 1.5rem; border-radius: 8px; border-left: 4px solid #3b82f6; min-height: 400px; }
    .stButton>button[kind="primary"] {
        background: linear-gradient(135deg, #2563eb, #3b82f6);
        color: white; border: none; padding: 0.75rem 2rem;
        font-weight: 600; border-radius: 8px;
    }
</style>
""", unsafe_allow_html=True)


# ── Header ──────────────────────────────────────────────────────────────────

st.markdown("""
<div class="main-header">
    <h1>🧪 学研·工科科研助手 v4.0</h1>
    <p style="margin: 0.5rem 0 0 0; opacity: 0.9;">15+功能矩阵 · 板块锚定 · 零篡位执行 · 影子合著者 · 语言基因深度提取</p>
</div>
""", unsafe_allow_html=True)

# API 状态栏
col1, col2, col3 = st.columns(3)
with col1:
    st.success(f"✅ API 就绪" if CLAUDE_API_KEY else "❌ API 未配置")
with col2:
    st.info(f"📚 {len(LOCAL_SKILLS)} 个学术 Skills")
with col3:
    st.success("⏱️ 180s 超时 | 3次重试")

# ── 侧边栏：15+功能矩阵 + 影子合著者 ─────────────────────────────────────────

with st.sidebar:
    st.markdown("---")
    st.subheader("🎯 15+ 功能矩阵")

    function = st.selectbox(
        "选择功能",
        list(FUNCTION_MATRIX.keys()),
        help="每个功能严格单一职责，零篡位执行"
    )

    # 显示当前功能规则
    func_info = FUNCTION_MATRIX.get(function, {})
    if func_info:
        st.caption(f"📋 {func_info.get('description', '')}")
        for rule in func_info.get('rules', []):
            st.caption(rule)

    st.markdown("---")
    st.subheader("🔬 学科领域")

    domain = st.selectbox(
        "研究领域",
        list(DOMAINS.keys()),
        help="选择领域以激活术语硬锁保护"
    )

    st.markdown("---")
    st.subheader("📚 影子合著者：标杆文献")

    reference_files = st.file_uploader(
        "上传标杆文献 (1-3篇)",
        type=["pdf", "docx"],
        accept_multiple_files=True,
        help="AI将学习其写作风格并模仿"
    )

    reference_styles = []
    if reference_files:
        for ref_file in reference_files:
            if ref_file.name not in st.session_state.reference_docs:
                with st.spinner(f"分析 {ref_file.name}..."):
                    file_bytes = ref_file.read()
                    text = extract_text(file_bytes, ref_file.name)
                    if text and not text.startswith("解析"):
                        style = analyze_reference_paper(file_bytes, ref_file.name)
                        st.session_state.reference_docs[ref_file.name] = style
                        st.success(f"✅ {ref_file.name}")

        # 收集所有标杆文献风格
        reference_styles = list(st.session_state.reference_docs.values())

    if st.session_state.reference_docs:
        st.success(f"✅ 已加载 {len(st.session_state.reference_docs)} 篇标杆文献")

    st.markdown("---")
    st.subheader("📁 普通文件上传")

    uploaded_file = st.file_uploader("上传待处理文件", type=["pdf", "docx"])

    if uploaded_file:
        extracted = extract_text(uploaded_file.read(), uploaded_file.name)
        if extracted and not extracted.startswith("解析"):
            st.success(f"✅ 提取 {len(extracted)} 字符")
            if st.button("📥 填入编辑器"):
                st.session_state.current_input = extracted[:10000]
                st.rerun()

# ── 顶部：板块锚点 Tabs ─────────────────────────────────────────────────────

st.markdown("---")
section_tabs = st.tabs([
    "📝 摘要",
    "📘 引言",
    "🔬 方法",
    "📊 结果",
    "💡 讨论",
    "🏁 结论"
])

# 确定当前板块
current_section = section_names[0]
for i, tab in enumerate(section_tabs):
    with tab:
        st.session_state[f"active_section_{i}"] = section_names[i]

# 通过 query params 或 session state 确定当前激活的 tab
# Streamlit tabs 无法直接获取激活索引，用 radio 隐式跟踪
if "selected_section_idx" not in st.session_state:
    st.session_state.selected_section_idx = 0

for i, tab in enumerate(section_tabs):
    with tab:
        if st.button("📌 锚定此板块", key=f"anchor_{i}", help="点击后下方操作将针对此板块"):
            st.session_state.selected_section_idx = i
            st.rerun()

current_section = section_names[st.session_state.selected_section_idx]

# ── 主区域：左右分栏对比 ─────────────────────────────────────────────────────

st.markdown("---")
col_left, col_right = st.columns([1, 1])

with col_left:
    st.markdown(f"### 📝 原文输入 [{current_section}]")
    current_input = st.text_area(
        "",
        value=st.session_state.get("current_input", ""),
        height=400,
        label_visibility="collapsed",
        key=f"input_{current_section}"
    )
    st.session_state.current_input = current_input

    st.caption(f"📊 {len(current_input)} 字符 | 语言: {'中文' if detect_language(current_input) == 'zh' else '英文'}")

    st.markdown("---")
    btn_col1, btn_col2, btn_col3 = st.columns([3, 1, 1])
    with btn_col1:
        process_button = st.button(
            f"✨ 执行 {function}",
            type="primary",
            use_container_width=True
        )
    with btn_col2:
        if st.button("🗑️"):
            st.session_state.current_input = ""
            st.rerun()
    with btn_col3:
        if st.button("🔄"):
            st.rerun()

with col_right:
    st.markdown(f"### 👁️ 处理结果 [{current_section}]")
    result_placeholder = st.empty()
    modification_note = st.empty()

    if process_button and current_input.strip():
        input_lang = detect_language(current_input)

        with st.spinner(f"🔄 处理中 [{function}] | 板块: {current_section} | 领域: {domain} | 语言: {'中文' if input_lang == 'zh' else '英文'}..."):
            try:
                # 硬锁保护
                protected_input, term_mapping = protect_hard_terms(current_input, domain)

                # 创建零篡位 prompt
                full_prompt = create_strict_prompt(
                    function, protected_input, current_section, domain, reference_styles, input_lang
                )

                # 调用 API
                start_time = time.time()
                response = call_api(full_prompt, timeout=180)
                elapsed = time.time() - start_time

                if response:
                    # 恢复术语
                    final_output = restore_hard_terms(response, term_mapping)

                    # 保存到历史
                    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    history_entry = {
                        "id": len(st.session_state.history) + 1,
                        "timestamp": timestamp,
                        "function": function,
                        "section": current_section,
                        "domain": domain,
                        "input_lang": input_lang,
                        "input": current_input,
                        "output": final_output,
                        "elapsed": f"{elapsed:.1f}s"
                    }
                    st.session_state.history.insert(0, history_entry)
                    st.session_state.history = st.session_state.history[:100]

                    # 显示结果
                    result_placeholder.markdown(final_output)

                    # 修改说明
                    modification_note.caption(
                        f"✅ 已针对【{current_section}】完成【{function}】，"
                        f"主要优化了语序和表达，使论证更具{'真人节奏感' if 'Humanizer' in function else '学术专业性'}"
                    )

                    # 导出按钮（支持redlining）
                    st.markdown("---")
                    export_col1, export_col2 = st.columns(2)

                    # 标准导出
                    doc_bytes = create_docx(final_output, {
                        "function": function,
                        "section": current_section,
                        "domain": domain,
                        "timestamp": timestamp
                    })
                    with export_col1:
                        st.download_button(
                            "📥 导出 Word",
                            doc_bytes,
                            file_name=f"yanyu_{current_section}_{timestamp.replace(':', '-')}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )

                    # Redlining导出（仅限修改类功能）
                    if function in ["✨ 表达润色", "🤖 去AI味 (Humanizer)", "🎯 精修模式"]:
                        redline_bytes = create_docx_with_redlines(current_input, final_output, {
                            "function": function,
                            "section": current_section,
                            "domain": domain,
                            "timestamp": timestamp
                        })
                        with export_col2:
                            st.download_button(
                                "📋 导出 Redlining",
                                redline_bytes,
                                file_name=f"yanyu_redline_{current_section}_{timestamp.replace(':', '-')}.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                            )

            except Exception as e:
                st.error(f"❌ 处理失败: {e}")
                if "timeout" in str(e).lower():
                    st.warning("💡 建议：缩短文本或增加超时时间")

# ── 版本时光机 ─────────────────────────────────────────────────────────────

st.markdown("---")
st.subheader("⏰ 版本时光机")

# 按板块筛选
filter_section = st.selectbox(
    "筛选板块",
    ["全部"] + list(SECTIONS.keys()),
    index=0
)

# 按功能筛选
filter_function = st.selectbox(
    "筛选功能",
    ["全部"] + list(FUNCTION_MATRIX.keys()),
    index=0
)

# 过滤历史
filtered_history = st.session_state.history
if filter_section != "全部":
    filtered_history = [h for h in filtered_history if h["section"] == filter_section]
if filter_function != "全部":
    filtered_history = [h for h in filtered_history if h["function"] == filter_function]

if filtered_history:
    for entry in filtered_history[:20]:
        with st.expander(
            f"#{entry['id']} · {entry['function']} · {entry['section']} · {entry['timestamp']}",
            expanded=False
        ):
            col1, col2, col3 = st.columns([3, 1, 1])
            with col1:
                st.caption(f"领域: {entry['domain']} | 语言: {'中文' if entry['input_lang'] == 'zh' else '英文'} | 耗时: {entry['elapsed']}")
            with col2:
                if st.button("📥 恢复", key=f"restore_{entry['id']}"):
                    st.session_state.current_input = entry['input']
                    st.rerun()
            with col3:
                if st.button("🗑️", key=f"delete_{entry['id']}"):
                    st.session_state.history = [h for h in st.session_state.history if h['id'] != entry['id']]
                    st.rerun()

            # 对比视图
            col_orig, col_out = st.columns(2)
            with col_orig:
                st.markdown("**原始输入**")
                st.text_area("", entry['input'], height=150, key=f"orig_{entry['id']}", disabled=True)
            with col_out:
                st.markdown("**处理输出**")
                st.markdown(entry['output'])
            st.markdown("---")
else:
    st.info("📭 暂无符合条件的记录")

# ── Footer ─────────────────────────────────────────────────────────────────

st.markdown("""
---
<div style="text-align: center; color: #64748b; font-size: 0.8rem; padding: 2rem 0;">
    <p><strong>🧪 学研·工科科研助手 v4.0</strong></p>
    <p>15+功能矩阵 · 板块锚定 · 零篡位执行 · 影子合著者 · 语言基因深度提取</p>
    <p>🔬 9大学科领域 · 📚 本地学术库集成 · ⏰ 版本时光机 · 📋 Redlining支持</p>
</div>
""", unsafe_allow_html=True)
