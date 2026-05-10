#!/bin/bash
# 研语·工科科研助手 v4.0 - 一键部署脚本

echo "🧪 研语·工科科研助手 v4.0 - 部署向导"
echo "====================================="
echo ""

# 检查 git
if ! command -v git &> /dev/null; then
    echo "❌ 请先安装 Git: https://git-scm.com/downloads"
    exit 1
fi

# 显示部署选项
echo "请选择部署平台："
echo "1) Streamlit Community Cloud (推荐 - 免费)"
echo "2) Hugging Face Spaces (推荐 - 免费)"
echo "3) Railway (灵活 - 需信用卡)"
echo "4) 本地运行"
echo ""
read -p "请输入选项 (1-4): " choice

case $choice in
    1)
        echo ""
        echo "🚀 部署到 Streamlit Community Cloud"
        echo "====================================="
        echo ""

        # 初始化 git
        echo "📦 初始化 Git 仓库..."
        if [ ! -d ".git" ]; then
            git init
        fi

        # 添加所有文件
        echo "➕ 添加文件..."
        git add .
        git commit -m "YanYu v4.0 - 工科科研助手" || echo "没有新的更改"

        # 询问 GitHub 用户名
        echo ""
        echo "请输入你的 GitHub 用户名:"
        read username

        echo ""
        echo "请输入仓库名称 (默认: xueyan-research-assistant):"
        read reponame
        reponame=${reponame:-xueyan-research-assistant}

        # 创建远程仓库
        echo ""
        echo "📝 创建远程仓库..."
        echo "1. 在浏览器打开: https://github.com/new"
        echo "2. 创建新仓库: $reponame"
        echo "3. 点击 'Create repository'"
        echo ""
        read -p "按回车继续..."

        # 添加远程仓库
        git remote add origin "https://github.com/$username/$reponame.git" 2>/dev/null || git remote set-url origin "https://github.com/$username/$reponame.git"

        # 推送
        echo ""
        echo "📤 推送代码到 GitHub..."
        git branch -M main
        git push -u origin main

        echo ""
        echo "✅ 代码已推送到 GitHub!"
        echo ""
        echo "🎯 下一步:"
        echo "1. 访问: https://share.streamlit.io/"
        echo "2. 点击 'New app'"
        echo "3. 选择仓库: $username/$reponame"
        echo "4. 主文件: app_new.py"
        echo "5. 添加环境变量 (Secrets):"
        echo "   ANTHROPIC_AUTH_TOKEN = 你的API密钥"
        echo "   ANTHROPIC_BASE_URL = 你的API地址 (可选)"
        echo "   ANTHROPIC_MODEL = claude-sonnet-4-6 (可选)"
        echo "6. 点击 Deploy"
        echo ""
        echo "🎉 部署完成！几分钟后获得公网地址"
        ;;

    2)
        echo ""
        echo "🚀 部署到 Hugging Face Spaces"
        echo "================================"
        echo ""
        echo "步骤 1：创建 Space"
        echo "------------------"
        echo "1. 访问：https://huggingface.co/spaces"
        echo "2. 点击 'Create new Space'"
        echo "3. SDK 选择: Streamlit"
        echo "4. License: MIT"
        echo "5. Hardware: CPU basic (免费)"
        echo ""
        read -p "输入你的 Space 名称: " space_name
        read -p "输入你的 Hugging Face 用户名: " hf_username

        echo ""
        echo "步骤 2：克隆并部署"
        echo "------------------"
        echo "执行以下命令："
        echo ""
        echo "git clone https://huggingface.co/spaces/$hf_username/$space_name"
        echo "cd $space_name"
        echo "cp ../app_new.py app.py"
        echo "cp ../requirements.txt ."
        echo "git add ."
        echo "git commit -m 'Deploy 研语 v4.0'"
        echo "git push"
        echo ""
        echo "然后在 Space 的 Settings → Secrets 添加环境变量"
        echo ""
        echo "或者直接在网页端上传文件："
        echo "1. 访问你的 Space 页面"
        echo "2. 点击 Files"
        echo "3. 上传 app_new.py (重命名为 app.py)"
        echo "4. 上传 requirements.txt"
        echo "5. 在 Settings → Secrets 添加环境变量"
        ;;

    3)
        echo ""
        echo "🚀 部署到 Railway"
        echo "================="
        echo ""
        echo "注意：Railway 需要信用卡验证（即使使用免费额度）"
        echo ""
        echo "步骤 1：安装 Railway CLI"
        echo "npm install -g railway"
        echo ""
        echo "步骤 2：登录并部署"
        echo "railway login"
        echo "railway init"
        echo "railway up"
        echo ""
        echo "然后在 Railway Dashboard 配置环境变量"
        ;;

    4)
        echo ""
        echo "💻 本地运行"
        echo "==========="
        echo ""
        echo "安装依赖："
        if [ -f "requirements.txt" ]; then
            pip install -r requirements.txt
        fi
        echo ""
        echo "启动应用："
        streamlit run app_new.py
        ;;

    *)
        echo "❌ 无效选项"
        exit 1
        ;;
esac

echo ""
echo "✅ 部署指南完成！"
echo ""
echo "📖 更多信息请查看：DEPLOY.md"

