# 研语·工科科研助手 - 部署指南

## 🚀 推荐部署方案（免费）

### 方案 1：Streamlit Community Cloud ⭐ 推荐
**优点**：官方支持、免费、专为 Streamlit 设计、一键部署

#### 部署步骤：

1. **准备项目文件**
```bash
cd C:\Users\29582\imgweb
```

2. **创建 requirements.txt**
```txt
streamlit>=1.28.0
anthropic>=0.18.0
pymupdf>=1.23.0
python-docx>=0.8.11
tenacity>=8.2.0
```

3. **创建 README.md**
```markdown
# 研语·工科科研助手

15+功能矩阵 · 板块锚定 · 零篡位执行 · 影子合著者

## 运行方式
```bash
streamlit run app_new.py
```

## 环境变量
- `ANTHROPIC_AUTH_TOKEN`: Claude API Key
- `ANTHROPIC_BASE_URL`: API Base URL (可选)
- `ANTHROPIC_MODEL`: 模型名称 (可选)
```

4. **推送到 GitHub**
```bash
# 初始化 git 仓库（如果还没有）
git init
git add .
git commit -m "Initial commit: 研语 v4.0"

# 创建 GitHub 仓库后
git remote add origin https://github.com/你的用户名/仓库名.git
git branch -M main
git push -u origin main
```

5. **部署到 Streamlit Community Cloud**
   - 访问：https://share.streamlit.io/
   - 点击 "New app"
   - 关联你的 GitHub 仓库
   - 选择 `imgweb/app_new.py` 作为主文件
   - 在 Settings → Secrets 添加环境变量：
     - `ANTHROPIC_AUTH_TOKEN`: 你的 API Key
     - `ANTHROPIC_BASE_URL`: 你的 API Base URL
   - 点击 "Deploy"

---

### 方案 2：Hugging Face Spaces ⭐⭐ 强烈推荐
**优点**：完全免费、支持 Streamlit、内置 Python 环境、全球 CDN

#### 部署步骤：

1. **创建 Spaces 项目**
   - 访问：https://huggingface.co/spaces
   - 点击 "Create new Space"
   - License: MIT
   - SDK: Streamlit
   - Hardware: CPU basic (免费)

2. **上传项目文件**
   ```bash
   # 安装 git lfs
   git lfs install

   # 克隆 Space 仓库
   git clone https://huggingface.co/spaces/你的用户名/空间名
   cd 空间名

   # 复制项目文件
   cp /c/Users/29582/imgweb/app_new.py app.py
   cp /c/Users/29582/imgweb/requirements.txt .

   # 提交
   git add .
   git commit -m "Deploy 研语 v4.0"
   git push
   ```

3. **配置 Secrets**
   - 在 Space 页面 → Settings → Repository → Secrets
   - 添加：
     - `ANTHROPIC_AUTH_TOKEN`
     - `ANTHROPIC_BASE_URL`

4. **自动部署**：推送后会自动构建和部署

---

### 方案 3：Railway ⭐⭐⭐
**优点**：支持 Docker、持久化、更灵活

#### 部署步骤：

1. **创建 Dockerfile**
```dockerfile
FROM python:3.11-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 8501

CMD ["streamlit", "run", "app_new.py", "--server.port=8501", "--server.address=0.0.0.0"]
```

2. **创建 railway.toml**
```toml
[build]
builder = "DOCKERFILE"

[deploy]
startCommand = "streamlit run app_new.py --server.port=8501 --server.address=0.0.0.0"
healthcheckPath = "/_stcore/health"
healthcheckTimeout = 300
```

3. **部署**
   - 访问：https://railway.app/
   - 点击 "New Project" → "Deploy from GitHub repo"
   - 选择你的仓库
   - 在 Variables 添加环境变量
   - 自动部署

---

## 🔧 环境变量配置

无论哪种方案，都需要配置以下环境变量：

| 变量名 | 说明 | 必需 |
|--------|------|------|
| `ANTHROPIC_AUTH_TOKEN` | Claude API Key | ✅ |
| `ANTHROPIC_BASE_URL` | API Base URL | ❌ |
| `ANTHROPIC_MODEL` | 模型名称 | ❌ |

---

## 📊 部署方案对比

| 方案 | 价格 | 难度 | 适合人群 |
|------|------|------|----------|
| **Streamlit Community Cloud** | 免费 | ⭐ 简单 | 所有用户 |
| **Hugging Face Spaces** | 免费 | ⭐ 简单 | 科研/AI项目 |
| **Railway** | $5/月起 | ⭐⭐ 中等 | 需要更多控制 |

---

## ⚡ 快速开始（推荐新手）

**最简单的方式**：使用 Streamlit Community Cloud

1. Fork 这个项目到你的 GitHub
2. 访问 https://share.streamlit.io/
3. 点击 "Deploy" → 选择你的仓库
4. 添加 API Key 到 Secrets
5. 完成！🎉

---

## 🛠️ 本地运行

```bash
cd C:\Users\29582\imgweb
pip install -r requirements.txt
streamlit run app_new.py
```

---

## 📝 注意事项

1. **API Key 安全**：永远不要将 API Key 提交到公开仓库
2. **使用量限制**：免费方案可能有并发和时长限制
3. **文件上传**：某些平台对上传文件大小有限制
4. **会话持久性**：免费方案重启后会丢失 session 数据

---

## 🆘 常见问题

### Q: 部署后无法访问 API？
A: 检查环境变量是否正确配置，确保 `ANTHROPIC_AUTH_TOKEN` 已添加。

### Q: 上传文件功能不工作？
A: 某些平台需要特殊配置，建议使用 Hugging Face Spaces。

### Q: 如何自定义域名？
A: Streamlit Community Cloud 不支持自定义域名，建议用 Hugging Face 或 Railway。

---

## 📞 支持

如有问题，请查看：
- [Streamlit 文档](https://docs.streamlit.io/)
- [Hugging Face Spaces 文档](https://huggingface.co/docs/hub/spaces)
- [Railway 文档](https://docs.railway.app/)
