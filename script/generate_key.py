# -*- coding: utf-8 -*-
"""
MailConverter 激活码生成工具
用于管理员生成用户激活码
"""
import sys
import hashlib
from datetime import datetime

def generate_activation_code(email, name, install_date):
    """生成激活码"""
    # 清理输入
    email = email.strip().lower()
    name = name.strip()
    install_str = install_date.strftime("%Y-%m-%d") if isinstance(install_date, datetime) else install_date

    # 组合原始字符串
    raw_string = f"{email}|{name}|{install_str}|MailConverter_v1"

    # 使用SHA256生成哈希
    sha256_hash = hashlib.sha256(raw_string.encode('utf-8')).hexdigest()

    # 转换为大写，添加分隔符 (格式: XXXX-XXXX-XXXX-XXXX)
    # C#版本: 取前19个字符 (前18字符 + 位置17的分隔符 = 19)
    # 字节0-3: X2 * 4 = 8字符 -> 4字符 + -
    # 字节4-7: X2 * 4 = 8字符 -> 4字符 + -
    # 字节8-11: X2 * 4 = 8字符 -> 4字符 + -
    # 字节12-15: X2 * 4 = 8字符 -> 4字符
    # 总共: 8 + 1 + 8 + 1 + 8 + 1 + 8 = 35... 不对

    # 重新实现: 按照C#的逻辑
    # 字节0-15，每个字节转2个十六进制字符
    # 每4个字节(8个十六进制字符)加一个横杠
    hex_str = sha256_hash[:16]  # 取前16个十六进制字符 (8字节)

    code = ""
    for i in range(0, len(hex_str), 1):
        code += hex_str[i]
        if i > 0 and (i + 1) % 4 == 0 and i < len(hex_str) - 1:
            code += "-"

    code = code.upper()

    # 添加校验位
    code = add_check_digit(code)

    return code

def add_check_digit(code):
    """添加校验位"""
    code_without_dash = code.replace("-", "")
    total = 0
    for c in code_without_dash:
        if c.isalpha():
            total += ord(c) - ord('A') + 10
        else:
            total += int(c)
    check = (100 - (total % 100)) % 100
    return f"{code}-{check:02d}"

def verify_activation_code(email, name, install_date, activation_code):
    """验证激活码"""
    expected = generate_activation_code(email, name, install_date)
    return expected.upper() == activation_code.upper()

if __name__ == "__main__":
    if len(sys.argv) < 4:
        print("MailConverter Activation Code Generator")
        print("=" * 50)
        print("Usage: generate_key.py <email> <name> <install_date>")
        print("Example: generate_key.py test@example.com John 2026-03-19")
        print()
        print("Install date format: YYYY-MM-DD")
        sys.exit(1)

    email = sys.argv[1]
    name = sys.argv[2]
    install_date = sys.argv[3]

    # 验证日期格式
    try:
        datetime.strptime(install_date, "%Y-%m-%d")
    except ValueError:
        print("Error: Invalid date format, use YYYY-MM-DD")
        sys.exit(1)

    # 生成激活码
    code = generate_activation_code(email, name, install_date)

    print()
    print("=" * 50)
    print("MailConverter Activation Code Generated Successfully!")
    print("=" * 50)
    print(f"Email: {email}")
    print(f"Name: {name}")
    print(f"Install Date: {install_date}")
    print(f"Activation Code: {code}")
    print("=" * 50)
    print()
    print("Please send the activation code to the user")
