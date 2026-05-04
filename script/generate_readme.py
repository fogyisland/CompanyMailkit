# -*- coding: utf-8 -*-
"""生成项目文档"""
import os
import glob
from datetime import datetime

def get_project_files(root_dir):
    """获取项目源文件"""
    patterns = ['*.cs', '*.py', '*.csproj', '*.md', '*.txt']
    files = []

    for pattern in patterns:
        for f in glob.glob(os.path.join(root_dir, pattern)):
            # 排除bin和obj目录
            if 'bin' not in f and 'obj' not in f:
                files.append(f)

    return sorted(files)

def get_file_info(filepath):
    """获取文件信息"""
    try:
        size = os.path.getsize(filepath)
        with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
            lines = f.readlines()
        return len(lines), size
    except:
        return 0, 0

def generate_readme(project_dir, output_file):
    """生成README.md"""
    files = get_project_files(project_dir)

    # 按类型分组
    cs_files = [f for f in files if f.endswith('.cs')]
    py_files = [f for f in files if f.endswith('.py')]
    other_files = [f for f in files if f.endswith(('.csproj', '.md', '.txt'))]

    lines = []

    # 标题
    lines.append("# MailConverter 邮件转换工具")
    lines.append("")
    lines.append(f"**生成时间**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append("")

    # 项目概述
    lines.append("## 项目概述")
    lines.append("")
    lines.append("MailConverter 是一款多功能的邮件转换工具，支持以下功能：")
    lines.append("")
    lines.append("- **EML → PST** - 将Foxmail导出的EML目录转换为PST文件")
    lines.append("- **OST → PST** - 将Outlook的OST格式转换为PST文件")
    lines.append("- **提取邮件** - 从PST/OST文件中提取邮件导出为EML格式")
    lines.append("- **IMAP 收件** - 从IMAP服务器收取邮件到本地PST文件")
    lines.append("")

    # 项目结构
    lines.append("## 项目结构")
    lines.append("")

    # C#文件
    lines.append("### C# 源文件")
    lines.append("")
    for f in cs_files:
        rel_path = os.path.relpath(f, project_dir)
        file_lines, size = get_file_info(f)
        lines.append(f"- `{rel_path}` ({file_lines} 行)")
    lines.append("")

    # Python文件
    lines.append("### Python 脚本")
    lines.append("")
    for f in py_files:
        rel_path = os.path.relpath(f, project_dir)
        file_lines, size = get_file_info(f)
        lines.append(f"- `{rel_path}` ({file_lines} 行)")
    lines.append("")

    # 其他文件
    lines.append("### 其他文件")
    lines.append("")
    for f in other_files:
        rel_path = os.path.relpath(f, project_dir)
        file_lines, size = get_file_info(f)
        lines.append(f"- `{rel_path}` ({file_lines} 行)")
    lines.append("")

    # 技术栈
    lines.append("## 技术栈")
    lines.append("")
    lines.append("- **.NET Framework 4.8** - 目标框架")
    lines.append("- **C#** - 编程语言")
    lines.append("- **WinForms** - 用户界面")
    lines.append("- **MimeKit 4.0.0** - EML/MIME解析")
    lines.append("- **MailKit 4.0.0** - IMAP客户端")
    lines.append("- **Serilog** - 日志系统")
    lines.append("")

    # 环境要求
    lines.append("## 环境要求")
    lines.append("")
    lines.append("- Windows 7 或更高版本")
    lines.append("- .NET Framework 4.8")
    lines.append("- Python 3.6+")
    lines.append("- pywin32 (Python库)")
    lines.append("- Microsoft Outlook")
    lines.append("")

    # 构建
    lines.append("## 构建项目")
    lines.append("")
    lines.append("```bash")
    lines.append("cd src/MailConverter")
    lines.append("dotnet build")
    lines.append("```")
    lines.append("")

    # 写入文件
    content = '\n'.join(lines)
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(content)

    print(f"README.md 已生成: {output_file}")
    print(f"包含 {len(cs_files)} 个C#文件, {len(py_files)} 个Python文件")

if __name__ == "__main__":
    # 获取脚本所在目录
    script_dir = os.path.dirname(os.path.abspath(__file__))
    readme_path = os.path.join(script_dir, "README.md")
    generate_readme(script_dir, readme_path)
