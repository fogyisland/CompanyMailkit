# -*- coding: utf-8 -*-
"""
PST/OST 离线提取 - WSL/Linux 方案
使用 Linux libpst 工具在 WSL 中提取

前提条件:
1. 已安装 WSL (Windows Subsystem for Linux)
2. 在 WSL 中安装了 libpst 工具: sudo apt-get install libpst-utils
"""
import subprocess
import os
import sys
import re
import shutil
import email.header

def run_wsl_command(cmd):
    """运行WSL命令"""
    try:
        # 使用 wsl.exe 执行命令，设置UTF-8编码
        result = subprocess.run(
            ["wsl", "-e", "bash", "-c", cmd],
            capture_output=True,
            timeout=300,
            encoding='utf-8',
            errors='replace'
        )
        return result.returncode == 0, result.stdout, result.stderr
    except FileNotFoundError:
        return False, "", "WSL未安装。请运行: wsl --install"
    except Exception as e:
        return False, "", str(e)

def check_wsl_installed():
    """检查WSL和readpst工具是否可用"""
    success, stdout, stderr = run_wsl_command("which readpst")
    if not success:
        return False, "libpst-utils未安装。请在WSL中运行: sudo apt-get install libpst-utils"
    return True, "WSL和工具已就绪"

def extract_with_wsl(pst_path, output_dir):
    """使用WSL中的readpst提取"""
    print("=" * 60)
    print("PST/OST 离线提取 - WSL/Linux 方案")
    print("=" * 60)

    # 检查WSL
    print("检查WSL环境...")
    success, msg = check_wsl_installed()
    if not success:
        print(f"错误: {msg}")
        print("\n请先安装WSL和libpst-utils:")
        print("1. 打开PowerShell(管理员): wsl --install")
        print("2. 安装完成后在WSL中运行: sudo apt-get update")
        print("3. 然后运行: sudo apt-get install libpst-utils")
        print("\n或者使用方案1 (需要安装Outlook)")
        return False

    print("WSL环境就绪")

    if not os.path.exists(pst_path):
        print(f"错误: 文件不存在: {pst_path}")
        return False

    os.makedirs(output_dir, exist_ok=True)

    # 将路径转换为WSL路径
    # C:\Users\... -> /mnt/c/Users/...
    wsl_pst = pst_path.replace("\\", "/")
    if ":" in wsl_pst:
        drive = wsl_pst[0].lower()
        wsl_pst = f"/mnt/{drive}" + wsl_pst[2:]

    wsl_output = output_dir.replace("\\", "/")
    if ":" in wsl_output:
        drive = wsl_output[0].lower()
        wsl_output = f"/mnt/{drive}" + wsl_output[2:]

    # 确保输出目录存在
    run_wsl_command(f"mkdir -p '{wsl_output}'")

    # 使用readpst提取
    # -M: 提取为MH格式 (单独文件)
    # -e: 保留文件扩展名 (.eml)
    # -o: 输出目录
    cmd = f"readpst -M -e -o '{wsl_output}' '{wsl_pst}'"

    print(f"正在提取: {pst_path}")
    print(f"命令: {cmd}")

    success, stdout, stderr = run_wsl_command(cmd)

    print(stdout)
    if stderr:
        print("警告:", stderr)

    # 重命名文件，使用邮件主题
    print("正在重命名文件...")
    renamed_count = rename_eml_files(output_dir)

    # 检查输出
    files = os.listdir(output_dir)
    eml_count = len([f for f in files if f.endswith('.eml')])

    print(f"\n提取完成!")
    print(f"输出目录: {output_dir}")
    print(f"提取邮件数: {eml_count}")
    if renamed_count > 0:
        print(f"已重命名: {renamed_count} 个文件")

    return True

def rename_eml_files(output_dir):
    """重命名EML文件，使用邮件主题作为文件名"""
    renamed_count = 0

    # 遍历所有子目录
    for root, dirs, files in os.walk(output_dir):
        for filename in files:
            if not filename.endswith('.eml'):
                continue

            filepath = os.path.join(root, filename)
            try:
                # 读取邮件
                with open(filepath, 'r', encoding='utf-8', errors='replace') as f:
                    content = f.read()

                # 提取Subject
                subject = extract_subject(content)
                if not subject:
                    continue

                # 清理文件名
                safe_subject = clean_filename(subject)
                if len(safe_subject) > 100:
                    safe_subject = safe_subject[:100]

                # 新文件名
                name, ext = os.path.splitext(filename)
                new_filename = f"{safe_subject}{ext}"
                new_filepath = os.path.join(root, new_filename)

                # 处理重名
                if new_filepath != filepath:
                    counter = 1
                    while os.path.exists(new_filepath):
                        new_filename = f"{safe_subject}_{counter}{ext}"
                        new_filepath = os.path.join(root, new_filename)
                        counter += 1

                    os.rename(filepath, new_filepath)
                    renamed_count += 1

            except Exception as e:
                print(f"处理 {filename} 失败: {e}")
                continue

    return renamed_count

def extract_subject(content):
    """从邮件内容中提取Subject"""
    try:
        lines = content.split('\n')
        for line in lines:
            line = line.strip()
            if line.lower().startswith('subject:'):
                subject = line[8:].strip()
                # 解码RFC 2047编码的主题
                decoded = email.header.decode_header(subject)
                result = ''
                for part, encoding in decoded:
                    if isinstance(part, bytes):
                        result += part.decode(encoding or 'utf-8', errors='replace')
                    else:
                        result += part
                return result.strip()
    except:
        pass
    return None

def clean_filename(name):
    """清理文件名中的非法字符"""
    invalid_chars = '<>:"/\\|?*'
    for c in invalid_chars:
        name = name.replace(c, '_')
    # 移除多余空格
    name = '_'.join(name.split())
    return name or "无主题"

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("用法: python extract_pst_wsl.py <PST文件路径> <输出目录>")
        print("\n前提条件:")
        print("1. 安装WSL: 在PowerShell中运行 'wsl --install'")
        print("2. 在WSL中安装工具: sudo apt-get install libpst-utils")
    else:
        pst_path = sys.argv[1]
        output_dir = sys.argv[2]
        extract_with_wsl(pst_path, output_dir)
