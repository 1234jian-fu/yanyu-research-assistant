#!/usr/bin/env python3
"""
研语·工科科研助手 v4.0 - 环境验证脚本
检查部署所需的所有依赖和配置
"""

import sys
import os
import subprocess
from pathlib import Path

def check_python_version():
    """检查 Python 版本"""
    print("🔍 检查 Python 版本...")
    version = sys.version_info
    if version >= (3, 9):
        print(f"   ✅ Python {version.major}.{version.minor}.{version.micro}")
        return True
    else:
        print(f"   ❌ Python 版本过低: {version.major}.{version.minor}.{version.micro}")
        print("   需要 Python 3.9 或更高版本")
        return False

def check_dependencies():
    """检查依赖包"""
    print("\n🔍 检查依赖包...")
    required = {
        'streamlit': '1.28.0',
        'anthropic': '0.18.0',
        'fitz': '1.23.0',  # pymupdf
        'docx': '0.8.11',  # python-docx
        'tenacity': '8.2.0',
    }

    all_ok = True
    for package, min_version in required.items():
        try:
            if package == 'fitz':
                import fitz
            elif package == 'docx':
                import docx
            else:
                __import__(package)
            print(f"   ✅ {package}")
        except ImportError:
            print(f"   ❌ {package} 未安装")
            all_ok = False

    return all_ok

def check_files():
    """检查必需文件"""
    print("\n🔍 检查必需文件...")
    required = [
        'app_new.py',
        'requirements.txt',
        'awesome-ai-research-writing/',
    ]

    all_ok = True
    for item in required:
        path = Path(item)
        if path.exists():
            print(f"   ✅ {item}")
        else:
            print(f"   ❌ {item} 缺失")
            all_ok = False

    return all_ok

def check_env_vars():
    """检查环境变量"""
    print("\n🔍 检查环境变量...")
    env_vars = [
        'ANTHROPIC_AUTH_TOKEN',
        'ANTHROPIC_BASE_URL',
        'ANTHROPIC_MODEL',
    ]

    has_token = False
    for var in env_vars:
        value = os.environ.get(var)
        if value:
            print(f"   ✅ {var} = ***{value[-4:]}")
            if var == 'ANTHROPIC_AUTH_TOKEN':
                has_token = True
        else:
            if var == 'ANTHROPIC_AUTH_TOKEN':
                print(f"   ⚠️  {var} 未设置 (必需)")
            else:
                print(f"   ℹ️  {var} 未设置 (可选)")

    return has_token

def check_git():
    """检查 Git"""
    print("\n🔍 检查 Git...")
    try:
        result = subprocess.run(['git', '--version'],
                              capture_output=True,
                              text=True)
        if result.returncode == 0:
            print(f"   ✅ {result.stdout.strip()}")
            return True
        else:
            print("   ❌ Git 未安装")
            return False
    except FileNotFoundError:
        print("   ❌ Git 未安装")
        return False

def main():
    """主函数"""
    print("=" * 50)
    print("🧪 研语·工科科研助手 v4.0 - 环境验证")
    print("=" * 50)

    results = {
        'Python 版本': check_python_version(),
        '依赖包': check_dependencies(),
        '必需文件': check_files(),
        '环境变量': check_env_vars(),
        'Git': check_git(),
    }

    print("\n" + "=" * 50)
    print("📊 验证结果汇总")
    print("=" * 50)

    for name, passed in results.items():
        status = "✅ 通过" if passed else "❌ 未通过"
        print(f"{name}: {status}")

    print("\n" + "=" * 50)

    if all(results.values()):
        print("🎉 所有检查通过！可以部署或运行应用")
        print("\n运行命令:")
        print("  streamlit run app_new.py")
        return 0
    else:
        print("⚠️  部分检查未通过，请处理上述问题")
        if not results['依赖包']:
            print("\n安装依赖:")
            print("  pip install -r requirements.txt")
        if not results['环境变量']:
            print("\n设置环境变量:")
            print("  export ANTHROPIC_AUTH_TOKEN=你的API密钥")
        return 1

if __name__ == '__main__':
    sys.exit(main())
