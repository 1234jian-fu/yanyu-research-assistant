---
title: 学研
emoji: 🧪
colorFrom: red
colorTo: red
sdk: streamlit
app_file: app_new.py
pinned: false
short_description: 学研·工科科研助手
license: mit
---

# 🧪 学研·工科科研助手 v4.0

> 15+功能矩阵 · 板块锚定 · 零篡位执行 · 影子合著者 · 语言基因深度提取

[![Streamlit](https://img.shields.io/badge/Streamlit-1.28+-red?logo=streamlit)](https://streamlit.io/)
[![Python](https://img.shields.io/badge/Python-3.9+-blue?logo=python)](https://www.python.org/)

## 🚀 一键部署到 Streamlit Cloud

### 步骤 1：推送到 GitHub

```bash
cd /c/Users/29582/imgweb
git init
git add .
git commit -m "XueYan v4.0 - 专业级工科科研助手"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/xueyan-research-assistant.git
git push -u origin main
```

### 步骤 2：部署到 Streamlit Cloud

1. 访问：https://streamlit.io/cloud
2. 点击 "New app"
3. 选择你的 GitHub 仓库
4. 配置：
   - **Repository**: `xueyan-research-assistant`
   - **Branch**: `main`
   - **Main file path**: `app_new.py`
5. **环境变量**（Secrets）：
   ```
   ANTHROPIC_AUTH_TOKEN=sk-your-api-key-here
   ANTHROPIC_BASE_URL=https://aiapi.aixia.tech
   ANTHROPIC_MODEL=claude-sonnet-4-6
   ```
6. 点击 "Deploy"

### 步骤 3：访问应用

等待 2-3 分钟部署完成，获得公网地址：
`https://xueyan-research-assistant.streamlit.app`

---

## 🌐 其他部署平台

### Hugging Face Spaces（推荐）

1. 访问：https://huggingface.co/spaces
2. "Create new Space"
3. 选择 "Streamlit" SDK
4. 上传代码或连接 GitHub
5. 在 Secrets 中添加环境变量
6. 获得：`https://huggingface.co/spaces/YOUR_USERNAME/xueyan`

### Railway（支持 Docker）

```bash
# 安装 Railway CLI
npm install -g railway

# 登录并部署
railway login
railway init
railway up
```

---

## 🔧 本地运行

```bash
# 安装依赖
pip install -r requirements.txt

# 运行应用
streamlit run app_new.py
```

访问：http://localhost:8503

---

## 📋 功能列表

### 核心功能
- 📝 **中转英翻译**：CN→EN语种转换，零润色
- ✨ **表达润色**：提升学术地道性，同语言优化
- 🔍 **逻辑检查**：检查因果链条和衔接
- 🤖 **去AI味 (Humanizer)**：消除AI痕迹，模仿真人语序
- 👨‍⚖️ **Reviewer视角**：模拟审稿人视角审视
- ✍️ **逐段起草**：将大纲扩充为正式段落
- 🎯 **精修模式**：深度精修，达到顶刊水准
- 📋 **Redlining修订**：带修订痕迹的修改建议

### 辅助功能
- 💡 研究想法构思
- 🧠 ML论文写作（NeurIPS/ICML级别）
- 📊 概念图设计
- 📄 节节头脑风暴
- 📝 引用验证
- 🎨 图表规范检查
- 🔄 版本对比

---

## 🎓 学科领域

🔋 能源电池 | 🏗️ 固废/土木 | 🔩 机械/材料 | 🧪 催化/化工
🌊 环境工程 | 📱 电子半导体 | 🧬 生物材料 | 🧠 机器学习 | 💻 软件工程

---

## 📄 License

MIT License
