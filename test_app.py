"""测试 Streamlit 是否能启动"""
import streamlit as st
import os

st.set_page_config(page_title="测试", layout="wide")

st.title("✅ 应用启动成功！")

st.write("环境变量:")
st.write(f"ANTHROPIC_AUTH_TOKEN: {'✅ 已配置' if os.getenv('ANTHROPIC_AUTH_TOKEN') else '❌ 未配置'}")
st.write(f"ANTHROPIC_BASE_URL: {os.getenv('ANTHROPIC_BASE_URL', '未设置')}")
st.write(f"ANTHROPIC_MODEL: {os.getenv('ANTHROPIC_MODEL', '未设置')}")

st.write("Python 包:")
try:
    import anthropic
    st.write(f"✅ anthropic: {anthropic.__version__}")
except:
    st.write("❌ anthropic 导入失败")

try:
    import fitz
    st.write(f"✅ pymupdf: {fitz.__version__}")
except:
    st.write("❌ pymupdf 导入失败")

try:
    from docx import Document
    st.write("✅ python-docx: 导入成功")
except:
    st.write("❌ python-docx 导入失败")

st.write("文件系统:")
from pathlib import Path
skills_path = Path("./awesome-ai-research-writing")
st.write(f"awesome-ai-research-writing 文件夹: {'✅ 存在' if skills_path.exists() else '❌ 不存在'}")

st.success("🎉 所有基础组件正常！")
