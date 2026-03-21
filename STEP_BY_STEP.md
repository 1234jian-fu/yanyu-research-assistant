# 🚀 研语·工科科研助手 - 零基础部署指南

## 📖 准备工作（开始前必读）

### 你需要：
- ✅ 一个 GitHub 账号（如果没有，去 https://github.com 注册）
- ✅ 一个 Claude API Key
- ✅ 电脑上已安装 Git（如果没有，去 https://git-scm.com/downloads 下载）

---

## 🎯 推荐方案：Streamlit Community Cloud（完全免费）

这是最简单、最适合新手的方案！

---

## 📝 第一步：准备代码文件

### 1.1 检查文件完整性

在你的 `C:\Users\29582\imgweb` 文件夹中，确保有以下文件：

```
imgweb/
├── app_new.py                    ✅ 必需
├── requirements.txt              ✅ 必需
├── README.md                     ✅ 必需
├── awesome-ai-research-writing/  ✅ 必需（文件夹）
└── DEPLOY.md                     可选
```

**怎么检查**：
1. 打开文件资源管理器
2. 进入 `C:\Users\29582\imgweb`
3. 确认这些文件都在

---

## 📝 第二步：创建 GitHub 仓库

### 2.1 登录 GitHub

1. 打开浏览器，访问：https://github.com/
2. 点击右上角 **"Sign in"** 登录（或 **"Sign up"** 注册）

### 2.2 创建新仓库

1. 登录后，点击右上角的 **"+"** 号
2. 选择 **"New repository"**
3. 填写信息：
   - **Repository name**: `yanyu-research-assistant`（或你喜欢的名字）
   - **Description**: `研语·工科科研助手 - 15+功能矩阵科研写作工具`
   - 选择 **"Public"**（公开）
   - ⚠️ **不要**勾选 "Add a README file"
   - ⚠️ **不要**勾选其他选项
4. 点击绿色按钮 **"Create repository"**

### 2.3 复制仓库地址

创建后，你会看到一个页面，复制你的仓库地址：
```
https://github.com/你的用户名/yanyu-research-assistant.git
```
记住这个地址！

---

## 📝 第三步：推送代码到 GitHub

### 3.1 打开命令行

**Windows 用户**：
1. 按 `Win + R` 键
2. 输入 `cmd` 并回车
3. 在命令行中输入：
   ```bash
   cd C:\Users\29582\imgweb
   ```

**或者更简单的方式**：
1. 在文件资源管理器中打开 `C:\Users\29582\imgweb` 文件夹
2. 在文件夹空白处**按住 Shift + 右键**
3. 选择 **"在此处打开 PowerShell 窗口"** 或 **"在此处打开命令窗口"**

### 3.2 运行部署脚本

在命令行中，输入：

**Windows 用户**：
```bash
deploy.bat
```

**或者手动执行**：
```bash
git init
git add .
git commit -m "YanYu v4.0 - 工科科研助手"
git branch -M main
git remote add origin https://github.com/你的用户名/yanyu-research-assistant.git
git push -u origin main
```

### 3.3 输入 GitHub 凭据

如果提示输入用户名和密码：
- **用户名**: 输入你的 GitHub 用户名
- **密码**: 输入你的 **Personal Access Token**（不是 GitHub 密码！）

**如何获取 Personal Access Token**：
1. GitHub → 右上角头像 → **Settings**
2. 左侧菜单最下方 → **Developer settings**
3. **Personal access tokens** → **Tokens (classic)**
4. **Generate new token** → **Generate new token (classic)**
5. Note: `Streamlit Deploy`
6. Expiration: 选择有效期（如 90 days）
7. 勾选 **`repo`** 权限
8. 点击 **Generate token**
9. 复制生成的 token（只显示一次！）

---

## 📝 第四步：部署到 Streamlit Cloud

### 4.1 访问 Streamlit Cloud

1. 打开浏览器，访问：https://share.streamlit.io/
2. 点击右上角 **"Sign in"**
3. 使用你的 GitHub 账号登录并授权

### 4.2 创建新应用

1. 登录后，点击右上角 **"New app"**
2. 填写信息：
   - **Repository**: 点击下拉框，选择 `yanyu-research-assistant`
   - **Branch**: 选择 `main`
   - **Main file path**: 输入 `app_new.py`
3. 点击 **"Advanced settings"**
4. 在 **"Secrets"** 部分，添加环境变量（见下一步）
5. 点击蓝色按钮 **"Deploy"**

### 4.3 配置环境变量（重要！）

在 **"Advanced settings"** → **"Secrets"** 部分，点击 **"+ New secret"**，添加以下变量：

**第一个变量**：
- **Key**: `ANTHROPIC_AUTH_TOKEN`
- **Value**: 粘贴你的 Claude API Key
- 点击 **"Add"**

**第二个变量**（可选）：
- **Key**: `ANTHROPIC_BASE_URL`
- **Value**: 你的 API 地址（如果有）
- 点击 **"Add"**

**第三个变量**（可选）：
- **Key**: `ANTHROPIC_MODEL`
- **Value**: `claude-sonnet-4-6`
- 点击 **"Add"**

### 4.4 开始部署

添加完环境变量后，点击 **"Deploy"** 按钮。

---

## 📝 第五步：等待部署完成

### 5.1 部署过程

你会看到以下信息：
```
Deploying your app...
This usually takes 2-3 minutes.
```

**等待 2-3 分钟**，页面会自动刷新。

### 5.2 部署成功

如果成功，你会看到：
```
✅ Your app is live!
```

并且有一个链接，比如：
```
https://yanyu-research-assistant.streamlit.app
```

### 5.3 测试应用

1. 点击链接打开应用
2. 尝试使用各项功能
3. 如果一切正常，恭喜你！🎉

---

## 📝 第六步：分享应用

### 6.1 获取应用链接

你的应用链接格式为：
```
https://你的仓库名.streamlit.app
```

例如：
```
https://yanyu-research-assistant.streamlit.app
```

### 6.2 分享给他人

直接复制这个链接，发送给任何人，他们都可以使用！

---

## ❌ 常见问题排查

### 问题 1: 推送时提示 "authentication failed"

**原因**: 密码错误或使用了错误的凭据

**解决**:
1. 不要使用 GitHub 密码
2. 使用 Personal Access Token（见第三步 3.3）

### 问题 2: 部署失败，显示 "Error"

**原因**: 代码有错误或依赖有问题

**解决**:
1. 检查 `app_new.py` 是否存在
2. 检查 `requirements.txt` 是否正确
3. 查看部署日志（点击 "View logs"）

### 问题 3: 应用可以打开，但功能不工作

**原因**: 环境变量配置错误

**解决**:
1. 检查 `ANTHROPIC_AUTH_TOKEN` 是否正确
2. 确保没有多余的空格或换行
3. 尝试重新部署

### 问题 4: 上传文件功能不工作

**原因**: Streamlit Cloud 的限制

**解决**:
1. 这是已知限制
2. 考虑使用 Hugging Face Spaces（见备用方案）

---

## 🔄 备用方案：Hugging Face Spaces

如果 Streamlit Cloud 不适合你，可以尝试 Hugging Face Spaces：

### 步骤：

1. **创建 Space**
   - 访问：https://huggingface.co/spaces
   - 点击 **"Create new Space"**
   - **SDK**: 选择 `Streamlit`
   - **Space name**: 输入名称（如 `yanyu-assistant`）
   - **License**: 选择 `MIT`
   - **Hardware**: 选择 `CPU basic` (免费)
   - 点击 **"Create Space"**

2. **上传文件**
   - 在 Space 页面，点击 **"Files"** 标签
   - 点击 **"Add file"** → **"Upload files"**
   - 上传 `app_new.py`（重命名为 `app.py`）
   - 上传 `requirements.txt`
   - 点击 **"Commit changes to main"**

3. **配置 Secrets**
   - 点击 **"Settings"** 标签
   - 找到 **"Repository secrets"** 部分
   - 添加环境变量（同第四步 4.3）

4. **等待构建**
   - 页面会自动开始构建
   - 等待 3-5 分钟
   - 构建完成后，应用自动上线

5. **访问应用**
   - 地址格式：`https://huggingface.co/spaces/你的用户名/你的空间名`

---

## ✅ 成功标志

当你看到以下情况之一，说明部署成功：

1. **Streamlit Cloud**:
   - 页面显示 "Your app is live!"
   - 有一个可点击的链接

2. **Hugging Face Spaces**:
   - 页面顶部显示 "Running"
   - 可以看到应用界面

3. **测试**:
   - 可以在浏览器中打开应用
   - 可以选择功能和输入文本
   - 点击按钮后能正常响应

---

## 🎉 恭喜！

如果你已经成功部署，恭喜你！🎊

现在你可以：
- ✅ 通过链接访问你的应用
- ✅ 分享链接给其他人使用
- ✅ 随时更新代码（只需 git push）

---

## 📞 还需要帮助？

如果遇到问题：
1. 检查上面的"常见问题排查"
2. 查看 [DEPLOY.md](DEPLOY.md) 详细文档
3. 查看 [QUICKSTART.md](QUICKSTART.md) 快速指南

---

**🧪 祝你部署顺利！**
