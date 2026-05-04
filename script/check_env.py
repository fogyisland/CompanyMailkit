# -*- coding: utf-8 -*-
"""检查并安装所需的环境依赖"""
import sys
import subprocess
import os

def check_python():
    """检查Python版本"""
    print(f"Python version: {sys.version}", flush=True)
    if sys.version_info < (3, 6):
        print("ERROR: 需要Python 3.6或更高版本", flush=True)
        return False
    return True

def check_pywin32():
    """检查pywin32是否已安装"""
    try:
        import win32com.client
        import pythoncom
        print("pywin32: 已安装", flush=True)
        return True
    except ImportError:
        print("pywin32: 未安装", flush=True)
        return False

def install_pywin32():
    """安装pywin32"""
    print("正在安装pywin32...", flush=True)
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pywin32"])
        print("pywin32 安装成功", flush=True)
        return True
    except Exception as e:
        print(f"安装失败: {e}", flush=True)
        return False

def check_environment():
    """检查所有环境依赖"""
    print("=== 检查环境依赖 ===", flush=True)

    # 检查Python
    if not check_python():
        return False

    # 检查pywin32
    if not check_pywin32():
        print("尝试安装...", flush=True)
        if not install_pywin32():
            print("请手动运行: pip install pywin32", flush=True)
            return False
        # 重新检查
        if not check_pywin32():
            return False

    print("=== 环境检查通过 ===", flush=True)
    return True

if __name__ == "__main__":
    success = check_environment()
    sys.exit(0 if success else 1)
