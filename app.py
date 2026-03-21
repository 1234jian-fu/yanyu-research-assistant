import base64
import json
import os
import re
import time
from datetime import datetime
from io import BytesIO

import fitz  # pymupdf
import requests
import streamlit as st
import streamlit.components.v1 as components

# ── Configuration ────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="研语科研工作站",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# Claude API 配置 (逻辑模型 - PDF 参数提取)
try:
    CLAUDE_API_KEY = st.secrets["CLAUDE_API_KEY"]
except (FileNotFoundError, KeyError):
    CLAUDE_API_KEY = os.getenv("NANOBANANA_API_KEY", "").strip()

CLAUDE_BASE_URL = os.getenv("CLAUDE_BASE_URL", "https://aiapi.aixia.tech/v1").rstrip("/")
CLAUDE_MODEL = "claude-sonnet-4-6"

# Gemini API 配置 (绘图模型)
GEMINI_API_KEY = "sk-bw08yrg3DOCNqjFSlXjnm5o3QOwcZo4B31STzzrLddpibmCp"
GEMINI_BASE_URL = "https://new.lemonapi.site/v1"
GEMINI_MODEL = "gemini-3-flash"

SCALE_OPTIONS = ["微观晶格", "界面机理", "介观形貌", "宏观器件"]
STYLE_MATRIX = {
    "Nature/Science 顶刊风": "minimalist, high-contrast scientific illustration, clean lines, professional journal quality",
    "高级 3D 渐变风": "cinematic 3D render, octane render, volumetric lighting, subsurface scattering, depth of field",
    "学术扁平线稿": "crisp technical line art, monochrome, dashed pathways, schematic diagram style",
}

if "image_history" not in st.session_state:
    st.session_state.image_history = []
if "last_image_url" not in st.session_state:
    st.session_state.last_image_url = ""
if "last_prompt" not in st.session_state:
    st.session_state.last_prompt = ""
if "legend_params" not in st.session_state:
    st.session_state.legend_params = []
if "pdf_params" not in st.session_state:
    st.session_state.pdf_params = []
if "svg_mode" not in st.session_state:
    st.session_state.svg_mode = False  # SVG 降级模式标志

# ── Helpers ───────────────────────────────────────────────────────────────────

def call_claude(model: str, content, temperature: float = 0.2, max_tokens: int = 2048) -> str:
    """调用 Claude API 用于逻辑处理和 PDF 参数提取"""
    payload = {
        "model": model,
        "temperature": temperature,
        "top_p": 0.9,
        "max_tokens": max_tokens,
        "messages": [{"role": "user", "content": content}],
    }
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {CLAUDE_API_KEY}",
        "anthropic-version": "2023-06-01",
    }
    resp = requests.post(f"{CLAUDE_BASE_URL}/messages", json=payload, headers=headers, timeout=120)
    resp.raise_for_status()
    data = resp.json()
    return data.get("content", [{}])[0].get("text", "")


def call_gemini_image(prompt: str, max_retries: int = 3) -> str:
    """调用 Gemini 3 Flash 生成科研图像，支持 URL 或 base64 返回，带重试机制"""
    payload = {
        "model": GEMINI_MODEL,
        "prompt": prompt,
        "size": "1024x1024",
        "quality": "hd",
        "n": 1,
    }
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {GEMINI_API_KEY}",
    }

    # 重试机制
    for attempt in range(max_retries):
        try:
            resp = requests.post(
                f"{GEMINI_BASE_URL.rstrip('/')}/images/generations",
                json=payload,
                headers=headers,
                timeout=60
            )
            resp.raise_for_status()
            data = resp.json()

            # 尝试获取 URL
            image_url = data.get("data", [{}])[0].get("url", "")

            # 如果没有 URL，检查是否有 base64 数据
            if not image_url:
                b64_data = data.get("data", [{}])[0].get("b64_json", "")
                if b64_data:
                    # 将 base64 存储为 data URI
                    image_url = f"data:image/png;base64,{b64_data}"

            return image_url

        except requests.exceptions.HTTPError as e:
            # 503/500 服务暂时不可用，等待后重试
            if e.response.status_code in [500, 502, 503, 504]:
                if attempt < max_retries - 1:
                    wait_time = (attempt + 1) * 2  # 递增等待时间
                    st.warning(f"服务繁忙，{wait_time}秒后重试... ({attempt + 1}/{max_retries})")
                    time.sleep(wait_time)
                    continue
            raise

        except requests.exceptions.RequestException as exc:
            if attempt < max_retries - 1:
                time.sleep(2)
                continue
            raise


def compose_gemini_prompt(
    material: str,
    description: str,
    style_label: str,
    scale: str,
    legend: list[str],
    modification: str | None,
) -> str:
    """构建 Gemini 图像生成 Prompt，包含专家级学术视觉约束"""
    style_intro = STYLE_MATRIX.get(style_label, STYLE_MATRIX["Nature/Science 顶刊风"])
    legend_text = " ".join(legend) if legend else "battery cathode anode electrolyte interface layers"

    mod_text = f"\nAdditional Modification Request: {modification}" if modification else ""

    prompt = f"""High-impact scientific illustration for Nature Materials cover.

Visual Subject: 3D volumetric rendering, micro-spherical {material} particles with metallic texture, wrapped in a translucent blue-tinted Se-C nanofiber network.

Research Context: {scale} scale - {description}.

Style Reference: {style_intro}.

Legend Parameters: {legend_text}.{mod_text}

Lighting & Environment: Laboratory white background, soft global illumination, octane render style, no text or watermarks on the image.

Technical Specifications: Show mixed ionic-electronic conduction pathways, Li+ diffusion channels with subtle glow effect, electron transport networks, and precise interface engineering between cathode and electrolyte. Ultra-high 8K resolution, photorealistic materials science visualization."""
    return prompt


def generate_image(description: str, material: str, style_label: str, scale: str, legend: list[str], modification: str | None):
    """使用 Gemini 3 Flash 生成科研图像，失败时自动降级到 Claude SVG 模式"""
    prompt = compose_gemini_prompt(material, description, style_label, scale, legend, modification)

    # 如果已经在 SVG 模式，直接使用 SVG
    if st.session_state.svg_mode:
        return generate_svg_fallback(description, material, style_label, scale, legend, modification)

    # 尝试使用 Gemini 生图
    with st.spinner(f"正在调用 {GEMINI_MODEL} 生成 8K 高清科研图…"):
        try:
            image_url = call_gemini_image(prompt, max_retries=3)
            timestamp = datetime.utcnow().isoformat() + "Z"
            st.session_state.image_history.insert(0, {
                "timestamp": timestamp,
                "prompt": prompt,
                "url": image_url,
                "type": "image"
            })
            st.session_state.image_history = st.session_state.image_history[:10]
            st.session_state.last_image_url = image_url
            st.session_state.last_prompt = prompt
            return image_url, prompt
        except Exception as exc:
            # 智能降级：切换到 SVG 模式
            st.warning(f"⚠️ 生图服务繁忙: {exc}")
            st.info("🔄 自动切换至 Claude 矢量线稿模式…")
            st.session_state.svg_mode = True
            return generate_svg_fallback(description, material, style_label, scale, legend, modification)


def generate_svg_fallback(description: str, material: str, style_label: str, scale: str, legend: list[str], modification: str | None):
    """Claude SVG 降级方案：生成矢量线稿"""
    style_intro = STYLE_MATRIX.get(style_label, STYLE_MATRIX["Nature/Science 顶刊风"])
    legend_text = " ".join(legend) if legend else "battery cathode anode electrolyte interface layers"
    mod_text = f" Additional note: {modification}." if modification else ""

    # 构建 SVG 专用 prompt
    svg_prompt = f"""You are a top-tier scientific illustrator. Create a professional SVG diagram for a Nature/Science journal.

Subject: {material} - {description}
Scale: {scale}
Style: {style_intro}
Legend: {legend_text}.{mod_text}

Requirements:
- Clean SVG code with proper viewBox
- Electric cyan (#00F5FF) accents on deep gray (#1A1A1A) background
- Clear labels for all components
- Scale bar included
- Output ONLY the <svg>...</svg> code, no explanations"""

    with st.spinner("正在调用 Claude Sonnet 4.6 生成矢量线稿…"):
        try:
            svg_output = call_claude(CLAUDE_MODEL, [{"type": "text", "text": svg_prompt}])
            match = re.search(r"<svg[\s\S]*?</svg>", svg_output, re.IGNORECASE)
            svg_code = match.group(0) if match else svg_output.strip()

            timestamp = datetime.utcnow().isoformat() + "Z"
            st.session_state.image_history.insert(0, {
                "timestamp": timestamp,
                "prompt": svg_prompt,
                "svg": svg_code,
                "type": "svg"
            })
            st.session_state.image_history = st.session_state.image_history[:10]
            st.session_state.last_svg = svg_code
            st.session_state.last_prompt = svg_prompt

            return f"svg://{svg_code}", svg_prompt

        except Exception as exc:
            st.error(f"SVG 生成也失败了: {exc}")
            return None, svg_prompt


def render_pdf_images(pdf_bytes: bytes, max_pages: int = 6):
    """渲染 PDF 页面为 base64 图像，用于 Claude Vision 分析"""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    page_images = []
    for i, page in enumerate(doc):
        if i >= max_pages:
            break
        mat = fitz.Matrix(2.0, 2.0)
        pix = page.get_pixmap(matrix=mat)
        img_b64 = base64.b64encode(pix.tobytes("png")).decode()
        page_images.append(
            {
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": "image/png",
                    "data": img_b64,
                },
            }
        )
    doc.close()
    return page_images


def extract_parameters_from_pdf(pdf_bytes: bytes):
    """使用 Claude Vision 从 PDF 提取实验参数"""
    if not CLAUDE_API_KEY:
        return [], "Missing Claude API key"
    if not CLAUDE_BASE_URL:
        return [], "Missing Claude base URL"

    page_images = render_pdf_images(pdf_bytes)
    if not page_images:
        return [], "无法渲染 PDF 页面"

    vision_prompt = (
        "Extract experimental parameters labeled or shown on these PDF pages (temperature, concentration, cutoff voltage, "
        "current density, N/P ratio, cycle number, capacity, electrolyte formula, etc.). Return a JSON array like ["
        "{\"参数\":\"...\", \"数值\":\"...\", \"单位\":\"...\", \"备注\":\"...\"}, ...]. Do not include any extra explanation."
    )
    content = page_images + [{"type": "text", "text": vision_prompt}]
    try:
        raw = call_claude(CLAUDE_MODEL, content)
    except Exception as exc:
        return [], str(exc)

    match = re.search(r"\[.*\]", raw, re.DOTALL)
    if not match:
        return [], "未在 Claude 响应中找到 JSON"

    try:
        params = json.loads(match.group(0))
        return params, ""
    except json.JSONDecodeError as exc:
        return [], f"JSON 解析失败：{exc}"


def format_legend(params: list[dict]) -> list[str]:
    """格式化参数为图注"""
    if not params:
        return []
    summary = []
    for key in ("温度", "浓度", "截止电压", "N/P比"):
        for item in params:
            if (item.get("参数") or item.get("param")) == key:
                value = item.get("数值") or item.get("value", "")
                unit = item.get("单位") or item.get("unit", "")
                summary.append(f"{key}: {value}{unit}")
                break
    return summary


def get_image_bytes(image_url: str) -> bytes | None:
    """将 image_url (URL 或 data URI) 转换为字节，用于下载"""
    if image_url.startswith("svg://"):
        # SVG 模式：提取 SVG 代码
        svg_code = image_url.replace("svg://", "")
        return svg_code.encode("utf-8")
    elif image_url.startswith("data:"):
        # data URI 格式: data:image/png;base64,{base64_data}
        _, base64_data = image_url.split(",", 1)
        return base64.b64decode(base64_data)
    else:
        # 远程 URL
        try:
            resp = requests.get(image_url, timeout=30)
            resp.raise_for_status()
            return resp.content
        except Exception:
            return None


# ── Layout ─────────────────────────────────────────────────────────────────────

st.markdown("""<style>
body { background-color: #0A0A0A; }
section.main { background-color: #0A0A0A; }
</style>""", unsafe_allow_html=True)

st.markdown('# 研语科研工作站 · Se修饰课题组', unsafe_allow_html=True)

# API 状态检查
claude_status = "✅ Claude 就绪" if CLAUDE_API_KEY else "❌ 缺少 Claude API Key"
gemini_status = "✅ Gemini 就绪" if GEMINI_API_KEY else "❌ 缺少 Gemini API Key"
st.caption(f"绘图引擎: {GEMINI_MODEL} {gemini_status} | 逻辑引擎: {CLAUDE_MODEL} {claude_status}")

col1, col2, col3 = st.columns([1, 2, 1])

with col1:
    st.subheader("参数引擎")
    material = st.text_input("研究材料", value="NCM523 动态电极体系")
    research_scale = st.selectbox("研究尺度", SCALE_OPTIONS)
    style_label = st.radio("视觉预设 · Style Matrix", list(STYLE_MATRIX.keys()))
    legend_field = st.text_area("当前图注摘要", value="\n".join(st.session_state.legend_params), height=100)
    st.info("同步后的参数会自动填入图注摘要，可在生成前微调。")

with col3:
    st.subheader("书语助手")
    pdf_file = st.file_uploader("上传文献 (支持扫描图像)", type="pdf")
    status_holder = st.empty()
    param_table = st.empty()
    if pdf_file:
        try:
            params, error = extract_parameters_from_pdf(pdf_file.read())
            if error:
                status_holder.error(error)
            else:
                st.session_state.pdf_params = params
                status_holder.success("参数提取完毕：可点击一键同步至图注。")
        except Exception as exc:
            status_holder.error(f"PDF 识别失败：{exc}")
    if st.session_state.pdf_params:
        df = [
            {"参数": item.get("参数") or item.get("param"), "数值": item.get("数值") or item.get("value"), "单位": item.get("单位") or item.get("unit")}
            for item in st.session_state.pdf_params
        ]
        param_table.table(df)
        if st.button("↗ 一键同步至图注"):
            st.session_state.legend_params = format_legend(st.session_state.pdf_params)
            status_holder.success("图注已同步，已加入生成 Prompt。")
    else:
        param_table.info("尚无参数，上传 PDF 可自动提取温度、浓度、截止电压、N/P 比等。")

with col2:
    st.subheader("创作画布")

    # SVG 模式提示
    if st.session_state.svg_mode:
        st.warning("🔄 当前为 SVG 矢量线稿模式（Gemini 服务繁忙时自动启用）")
        if st.button("🎨 切换回 Gemini 8K 模式", key="switch_to_gemini"):
            st.session_state.svg_mode = False
            st.rerun()

    description = st.text_area("机理描述", height=150, placeholder="请描述 Se-C 网络如何包覆 NCM523 颗粒、电子 & Li+ 传输路径、以及电池界面细节…")
    generate_button = st.button("🎨 生成科研图", key="generate", type="primary")
    image_placeholder = st.empty()

    if generate_button:
        if not description.strip():
            st.warning("描述不能为空，请输入想要表现的机理。")
        else:
            image_url, prompt = generate_image(
                description=description,
                material=material,
                style_label=style_label,
                scale=research_scale,
                legend=st.session_state.legend_params,
                modification=None,
            )
            if image_url:
                # 根据类型显示图像
                if image_url.startswith("svg://"):
                    svg_code = image_url.replace("svg://", "")
                    components.html(svg_code, height=600)
                    st.success("✨ 矢量线稿生成完成！（已切换至 SVG 降级模式）")
                    st.download_button("⬇️ 下载 SVG", svg_code.encode("utf-8"), file_name=f"yanyu_diagram_{datetime.now().strftime('%Y%m%d_%H%M%S')}.svg", mime="image/svg+xml")
                else:
                    image_placeholder.image(image_url, use_container_width=True)
                    st.success("✨ 8K 科研图生成完成！")
                    image_bytes = get_image_bytes(image_url)
                    if image_bytes:
                        st.download_button("⬇️ 下载图像", image_bytes, file_name=f"yanyu_research_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png", mime="image/png")

    # 显示上次生成的图像
    elif st.session_state.last_image_url:
        if st.session_state.last_image_url.startswith("svg://"):
            svg_code = st.session_state.last_image_url.replace("svg://", "")
            image_placeholder.html(svg_code, height=600)
        elif st.session_state.last_svg:
            image_placeholder.html(st.session_state.last_svg, height=600)
        else:
            image_placeholder.image(st.session_state.last_image_url, use_container_width=True)

    # 修改意见输入
    chat_note = st.chat_input("💬 输入修改意见（如：'增强 Li+ 扩散路径的发光效果'）")
    if chat_note:
        if not description.strip():
            st.warning("请先输入初始描述，再补充修改意见。")
        else:
            image_url, prompt = generate_image(
                description=description,
                material=material,
                style_label=style_label,
                scale=research_scale,
                legend=st.session_state.legend_params,
                modification=chat_note,
            )
            if image_url:
                if image_url.startswith("svg://"):
                    svg_code = image_url.replace("svg://", "")
                    image_placeholder.html(svg_code, height=600)
                    st.success("根据修改意见重新生成 SVG 完成。")
                else:
                    image_placeholder.image(image_url, use_container_width=True)
                    st.success("根据修改意见重新生成完成。")

    # 生成历史
    with st.expander("📜 生成历史", expanded=False):
        for entry in st.session_state.image_history:
            entry_type = entry.get("type", "image")
            type_label = "📐 SVG" if entry_type == "svg" else "🖼️ 8K"
            st.write(f"{type_label} {entry['timestamp']} · {entry['prompt'][:60]}...")
            col_a, col_b = st.columns([3, 1])
            with col_a:
                st.button("使用此 Prompt", key=f"reuse_{entry['timestamp']}", on_click=lambda e=entry: st.session_state.__setitem__("last_prompt", e["prompt"]))
            with col_b:
                if entry_type == "svg" and entry.get("svg"):
                    st.download_button("下载", entry["svg"].encode("utf-8"), file_name=f"yanyu_history_{entry['timestamp']}.svg", mime="image/svg+xml", key=f"dl_{entry['timestamp']}")
                elif entry.get("url"):
                    image_bytes = get_image_bytes(entry["url"])
                    if image_bytes:
                        st.download_button("下载", image_bytes, file_name=f"yanyu_history_{entry['timestamp']}.png", mime="image/png", key=f"dl_{entry['timestamp']}")
                    else:
                        st.caption("无法下载")
            st.markdown("---")

    if st.session_state.last_prompt:
        with st.expander("📋 查看 Prompt", expanded=False):
            st.text_area("完整 Prompt", st.session_state.last_prompt, height=200, disabled=True)

st.caption("YanYu Research Assistant · 绘图: Gemini 3 Flash | 逻辑: Claude Sonnet 4.6")
