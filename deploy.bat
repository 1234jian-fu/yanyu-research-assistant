@echo off
REM 研语·工科科研助手 v4.0 - Windows 部署脚本

echo =====================================
echo 研语·工科科研助手 v4.0 - 部署向导
echo =====================================
echo.

REM 显示部署选项
echo 请选择部署平台：
echo 1) Streamlit Community Cloud (推荐 - 免费)
echo 2) Hugging Face Spaces (推荐 - 免费)
echo 3) 本地运行
echo.
set /p choice="请输入选项 (1-3): "

if "%choice%"=="1" goto streamlit
if "%choice%"=="2" goto huggingface
if "%choice%"=="3" goto local
goto invalid

:streamlit
echo.
echo ======================================
echo 部署到 Streamlit Community Cloud
echo ======================================
echo.

REM 检查 git
where git >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo ❌ 请先安装 Git: https://git-scm.com/downloads
    pause
    exit /b 1
)

echo 步骤 1：初始化 Git 仓库
echo ------------------------
if not exist ".git" (
    git init
    echo ✅ Git 仓库已初始化
) else (
    echo ✅ Git 仓库已存在
)

echo.
echo 步骤 2：添加文件
echo -----------------
git add .
git commit -m "YanYu v4.0 - 工科科研助手" 2>nul || echo 没有新的更改

echo.
echo 步骤 3：连接 GitHub
echo -------------------
set /p username="请输入你的 GitHub 用户名: "
set /p reponame="请输入仓库名称 (默认: xueyan-research-assistant): "
if "%reponame%"=="" set reponame=xueyan-research-assistant

echo.
echo 在浏览器中打开以下链接创建仓库：
echo https://github.com/new
echo 仓库名称: %reponame%
echo.
pause

git remote add origin https://github.com/%username%/%reponame%.git 2>nul || git remote set-url origin https://github.com/%username%/%reponame%.git
git branch -M main

echo.
echo 步骤 4：推送代码
echo ----------------
git push -u origin main

echo.
echo ======================================
echo ✅ 代码已推送到 GitHub！
echo ======================================
echo.
echo 下一步：
echo 1. 访问: https://share.streamlit.io/
echo 2. 点击 'New app'
echo 3. 选择仓库: %username%/%reponame%
echo 4. 主文件: app_new.py
echo 5. 添加环境变量 (Secrets)：
echo    ANTHROPIC_AUTH_TOKEN = 你的API密钥
echo    ANTHROPIC_BASE_URL = 你的API地址 (可选)
echo    ANTHROPIC_MODEL = claude-sonnet-4-6 (可选)
echo 6. 点击 Deploy
echo.
echo 🎉 部署完成！几分钟后获得公网地址
echo.
pause
exit /b 0

:huggingface
echo.
echo ======================================
echo 部署到 Hugging Face Spaces
echo ======================================
echo.
echo 步骤 1：创建 Space
echo ------------------
echo 1. 访问：https://huggingface.co/spaces
echo 2. 点击 'Create new Space'
echo 3. SDK 选择: Streamlit
echo 4. License: MIT
echo 5. Hardware: CPU basic (免费)
echo.
set /p space_name="输入你的 Space 名称: "
set /p hf_username="输入你的 Hugging Face 用户名: "

echo.
echo 步骤 2：准备文件
echo ----------------
echo 在项目文件夹中，将以下文件准备好：
echo - 将 app_new.py 重命名为 app.py
echo - 准备 requirements.txt
echo.
echo 步骤 3：上传文件
echo ----------------
echo 1. 访问你的 Space 页面
echo 2. 点击 Files 标签
echo 3. 上传 app.py 和 requirements.txt
echo 4. 在 Settings → Secrets 添加环境变量：
echo    ANTHROPIC_AUTH_TOKEN = 你的API密钥
echo    ANTHROPIC_BASE_URL = 你的API地址 (可选)
echo.
pause
exit /b 0

:local
echo.
echo ==================
echo 本地运行
echo ==================
echo.

if exist "requirements.txt" (
    echo 安装依赖...
    pip install -r requirements.txt
)

echo.
echo 启动应用...
streamlit run app_new.py

pause
exit /b 0

:invalid
echo ❌ 无效选项
pause
exit /b 1
