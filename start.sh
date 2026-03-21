#!/bin/bash
# 研语·工科科研助手 v4.0 - 快速启动脚本

set -e

echo "🧪 研语·工科科研助手 v4.0"
echo "====================================="
echo ""

# 检查 Python
if ! command -v python3 &> /dev/null; then
    echo "❌ Python 未安装"
    echo "请访问 https://www.python.org/downloads/"
    exit 1
fi

# 检查依赖
echo "🔍 检查依赖..."
if [ ! -f "requirements.txt" ]; then
    echo "❌ requirements.txt 未找到"
    exit 1
fi

# 安装依赖（如果需要）
if ! python3 -c "import streamlit" 2>/dev/null; then
    echo "📦 安装依赖..."
    pip install -r requirements.txt
fi

# 检查环境变量
if [ -z "$ANTHROPIC_AUTH_TOKEN" ]; then
    echo ""
    echo "⚠️  警告: ANTHROPIC_AUTH_TOKEN 未设置"
    echo ""
    echo "请设置环境变量:"
    echo "  export ANTHROPIC_AUTH_TOKEN=你的API密钥"
    echo ""
    read -p "是否继续启动? (y/n): " continue
    if [ "$continue" != "y" ]; then
        exit 1
    fi
fi

# 启动应用
echo ""
echo "🚀 启动应用..."
echo ""
streamlit run app_new.py "$@"
