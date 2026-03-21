# 🚀 快速部署指南

## 📋 部署前准备

### 1. 检查环境
```bash
# 运行环境检查
python check_env.py
```

### 2. 安装依赖
```bash
pip install -r requirements.txt
```

### 3. 本地测试
```bash
streamlit run app_new.py
```

---

## 🎯 三种部署方式

### 方式 1: Streamlit Community Cloud ⭐ 最简单

**优点**: 免费、官方支持、一键部署

#### 步骤:

1. **推送代码到 GitHub**
   ```bash
   # Windows
   deploy.bat

   # Linux/Mac
   bash deploy.sh
   ```

2. **部署到 Streamlit Cloud**
   - 访问: https://share.streamlit.io/
   - 点击 "New app"
   - 选择你的 GitHub 仓库
   - 主文件: `app_new.py`
   - 添加环境变量 (Secrets):
     ```
     ANTHROPIC_AUTH_TOKEN = 你的API密钥
     ANTHROPIC_BASE_URL = 你的API地址 (可选)
     ANTHROPIC_MODEL = claude-sonnet-4-6 (可选)
     ```
   - 点击 "Deploy"

3. **等待部署完成** (2-3分钟)

4. **获得公网地址**: `https://你的仓库名.streamlit.app`

---

### 方式 2: Hugging Face Spaces ⭐⭐ 推荐

**优点**: 完全免费、全球 CDN、支持自定义域名

#### 步骤:

1. **创建 Space**
   - 访问: https://huggingface.co/spaces
   - 点击 "Create new Space"
   - SDK: 选择 `Streamlit`
   - License: `MIT`
   - Hardware: `CPU basic` (免费)

2. **上传文件**
   - 方式 A: 网页上传
     - 在 Space 页面点击 "Files"
     - 上传 `app_new.py` (重命名为 `app.py`)
     - 上传 `requirements.txt`

   - 方式 B: Git 推送
     ```bash
     git clone https://huggingface.co/spaces/你的用户名/空间名
     cd 空间名
     cp ../app_new.py app.py
     cp ../requirements.txt .
     git add .
     git commit -m "Deploy 研语 v4.0"
     git push
     ```

3. **配置 Secrets**
   - Space 页面 → Settings → Secrets
   - 添加环境变量

4. **自动部署**: 推送后自动构建和部署

5. **访问地址**: `https://huggingface.co/spaces/你的用户名/空间名`

---

### 方式 3: Railway ⭐⭐⭐ 灵活

**优点**: 支持 Docker、持久化、更灵活

#### 步骤:

1. **安装 Railway CLI**
   ```bash
   npm install -g railway
   ```

2. **登录**
   ```bash
   railway login
   ```

3. **初始化项目**
   ```bash
   railway init
   ```

4. **部署**
   ```bash
   railway up
   ```

5. **配置环境变量**
   - Railway Dashboard
   - Variables 标签
   - 添加环境变量

6. **访问**: Railway 会提供一个公网地址

---

## 🔧 环境变量配置

| 变量名 | 说明 | 必需 | 默认值 |
|--------|------|------|--------|
| `ANTHROPIC_AUTH_TOKEN` | Claude API Key | ✅ | - |
| `ANTHROPIC_BASE_URL` | API Base URL | ❌ | https://api.anthropic.com |
| `ANTHROPIC_MODEL` | 模型名称 | ❌ | claude-sonnet-4-6 |

---

## 📊 平台对比

| 平台 | 价格 | 难度 | 优势 | 劣势 |
|------|------|------|------|------|
| **Streamlit Cloud** | 免费 | ⭐ 简单 | 官方支持、一键部署 | 功能受限 |
| **Hugging Face** | 免费 | ⭐ 简单 | 全球 CDN、自定义域名 | 构建稍慢 |
| **Railway** | $5/月起 | ⭐⭐ 中等 | 灵活、Docker 支持 | 需信用卡 |

---

## ⚡ 快速开始（推荐新手）

### 最简单的 3 步：

1. **Fork 项目到你的 GitHub**

2. **访问 Streamlit Cloud**
   ```
   https://share.streamlit.io/
   ```

3. **点击 "New app" → 选择你的仓库 → 添加 API Key → Deploy**

**完成！** 🎉

---

## 🆘 常见问题

### Q1: 部署后无法访问 API？
**A**: 检查环境变量是否正确配置，确保 `ANTHROPIC_AUTH_TOKEN` 已添加。

### Q2: 上传文件功能不工作？
**A**: 某些平台需要特殊配置，建议使用 Hugging Face Spaces。

### Q3: 如何自定义域名？
**A**:
- Streamlit Cloud: 不支持
- Hugging Face: 支持，在 Space Settings 中配置
- Railway: 支持，在 Dashboard 中配置

### Q4: 部署失败怎么办？
**A**:
1. 检查 `requirements.txt` 是否正确
2. 查看部署日志
3. 确保所有依赖都支持目标平台
4. 尝试在本地运行 `python check_env.py`

### Q5: 免费方案有限制吗？
**A**:
- **Streamlit Cloud**: 每月 750 小时
- **Hugging Face**: 基本无限制
- **Railway**: $5 免费额度/月

---

## 📞 获取帮助

- **文档**: [DEPLOY.md](DEPLOY.md)
- **问题反馈**: [GitHub Issues](https://github.com/你的用户名/你的仓库/issues)
- **Streamlit 文档**: https://docs.streamlit.io/
- **Hugging Face 文档**: https://huggingface.co/docs/hub/spaces

---

## ✅ 部署检查清单

部署前请确认：

- [ ] Python 版本 >= 3.9
- [ ] 所有依赖已安装
- [ ] API Key 已获取
- [ ] 代码已推送到 GitHub
- [ ] 环境变量已配置
- [ ] 本地测试通过

---

**🧪 研语·工科科研助手** - 让科研写作更高效
