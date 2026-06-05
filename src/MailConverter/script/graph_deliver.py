# -*- coding: utf-8 -*-
"""
通过 Microsoft Graph API 投递 EML 邮件到 Office 365 邮箱。
与 ews_deliver.py 风格保持一致 (Python 子进程, C# 调用)。

用法:
    graph_deliver.py <access_token> <user_email> <eml_path> [target_folder]

target_folder 可选，支持 "/" 分隔的嵌套路径，例如 "MyImport/Mail1"。
"""
import sys
import os
import json
import urllib.request
import urllib.parse
import urllib.error
import email
from email.header import decode_header
from email.utils import getaddresses


GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def log(msg):
    print(msg, flush=True)


def decode_mime_header(value):
    """解码 RFC 2047 编码的头部"""
    if not value:
        return ""
    try:
        parts = decode_header(value)
        decoded = []
        for content, charset in parts:
            if isinstance(content, bytes):
                try:
                    decoded.append(content.decode(charset or "utf-8", errors="replace"))
                except (LookupError, UnicodeDecodeError):
                    decoded.append(content.decode("utf-8", errors="replace"))
            else:
                decoded.append(content)
        return "".join(decoded)
    except Exception:
        return value


def parse_address(value):
    """解析邮件地址字符串, 返回 (name, address) 列表"""
    if not value:
        return []
    try:
        decoded = decode_mime_header(value)
        addrs = getaddresses([decoded])
        result = []
        for name, addr in addrs:
            if not addr:
                continue
            result.append({"name": name or "", "address": addr})
        return result
    except Exception:
        return []


def find_or_create_folder(token, user_email, folder_path):
    """
    按 "/ " 分隔的路径逐级查找/创建文件夹, 返回最后一级 folderId。
    folder_path 可以是 "Inbox" (well-known) 或 "MyImport/Mail1" (自定义嵌套)。
    """
    parts = [p.strip() for p in folder_path.replace("\\", "/").split("/") if p.strip()]
    if not parts:
        return "inbox"

    # well-known 文件夹映射
    wellknown = {
        "inbox": "inbox", "收件箱": "inbox",
        "sent": "sent", "sentitems": "sent", "已发送": "sent",
        "deleted": "deleteditems", "deleteditems": "deleteditems", "已删除": "deleteditems",
        "drafts": "drafts", "草稿": "drafts",
        "junk": "junkemail", "junkemail": "junkemail", "垃圾邮件": "junkemail",
        "archive": "archive", "archivefolderroot": "archive", "存档": "archive",
    }

    parent_id = None  # None 表示根目录
    current_path = ""

    for i, seg in enumerate(parts):
        current_path = seg if not current_path else current_path + "/" + seg

        # 第一段如果是 well-known, 直接用
        if i == 0 and seg.lower() in wellknown:
            parent_id = wellknown[seg.lower()]
            continue

        # 在当前 parent 下查找
        escaped = seg.replace("'", "''")
        filter_expr = f"displayName eq '{escaped}'"
        if parent_id is None:
            url = f"{GRAPH_BASE}/users/{urllib.parse.quote(user_email)}/mailFolders?{urllib.parse.urlencode({'$filter': filter_expr})}"
        else:
            url = f"{GRAPH_BASE}/users/{urllib.parse.quote(user_email)}/mailFolders/{parent_id}/childFolders?{urllib.parse.urlencode({'$filter': filter_expr})}"

        req = urllib.request.Request(url, headers={"Authorization": f"Bearer {token}"})
        with urllib.request.urlopen(req) as resp:
            data = json.loads(resp.read().decode("utf-8"))

        if data.get("value"):
            parent_id = data["value"][0]["id"]
        else:
            # 创建
            body = json.dumps({"displayName": seg}).encode("utf-8")
            if parent_id is None:
                create_url = f"{GRAPH_BASE}/users/{urllib.parse.quote(user_email)}/mailFolders"
            else:
                create_url = f"{GRAPH_BASE}/users/{urllib.parse.quote(user_email)}/mailFolders/{parent_id}/childFolders"
            req = urllib.request.Request(create_url, data=body, method="POST",
                                         headers={"Authorization": f"Bearer {token}",
                                                  "Content-Type": "application/json"})
            with urllib.request.urlopen(req) as resp:
                created = json.loads(resp.read().decode("utf-8"))
            parent_id = created.get("id")
            log(f"  [Folder] Created: {current_path}")

    return parent_id


def parse_eml_file(eml_path):
    """
    解析 EML 文件, 返回清洗后的字段
    使用 Python 标准库 email, 容忍畸形头.
    """
    with open(eml_path, "rb") as f:
        raw = f.read()

    msg = email.message_from_bytes(raw)

    subject = decode_mime_header(msg.get("Subject", ""))
    from_addrs = parse_address(msg.get("From", ""))
    to_addrs = parse_address(msg.get("To", ""))
    cc_addrs = parse_address(msg.get("Cc", ""))
    bcc_addrs = parse_address(msg.get("Bcc", ""))

    # 提取正文
    body_html = ""
    body_text = ""
    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            disp = str(part.get("Content-Disposition", ""))
            if "attachment" in disp.lower():
                continue
            if ctype == "text/html":
                try:
                    body_html = part.get_payload(decode=True).decode(
                        part.get_content_charset() or "utf-8", errors="replace")
                except Exception:
                    pass
            elif ctype == "text/plain" and not body_text:
                try:
                    body_text = part.get_payload(decode=True).decode(
                        part.get_content_charset() or "utf-8", errors="replace")
                except Exception:
                    pass
    else:
        try:
            payload = msg.get_payload(decode=True)
            if payload:
                body_text = payload.decode(msg.get_content_charset() or "utf-8", errors="replace")
        except Exception:
            pass

    from_addr = from_addrs[0] if from_addrs else {"name": "", "address": "no-reply@booming.one"}

    # 时间解析:
    # 1) 优先 X-Original-Received-Time (PST 提取时写入, 原始接收时间)
    # 2) 否则用 Date 头 (发件时间, 接近真实接收时间)
    import datetime
    received_dt = None
    x_received = msg.get("X-Original-Received-Time")
    if x_received:
        try:
            received_dt = email.utils.parsedate_to_datetime(x_received)
        except Exception:
            pass
    if received_dt is None:
        date_str = msg.get("Date")
        if date_str:
            try:
                received_dt = email.utils.parsedate_to_datetime(date_str)
            except Exception:
                pass
    if received_dt is None:
        received_dt = datetime.datetime.now(datetime.timezone.utc)

    # 转为 ISO 8601 (Graph 要求 UTC 带 Z)
    if received_dt.tzinfo is None:
        received_dt = received_dt.replace(tzinfo=datetime.timezone.utc)
    received_iso = received_dt.astimezone(datetime.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    return subject, from_addr, to_addrs, cc_addrs, bcc_addrs, body_html, body_text, received_iso


def build_message_object(subject, from_addr, to_addrs, cc_addrs, bcc_addrs, body_html, body_text, received_iso):
    """构造 Graph Message JSON 对象 (清洁, 无畸形头)
    重要: Graph API 的 receivedDateTime/sentDateTime 顶层字段是只读,
    服务器强制覆盖为当前时间. 必须用 singleValueExtendedProperties
    设置 PidTagMessageDeliveryTime (0x0E06) 才能保留原始时间.
    """
    msg = {
        "subject": subject or "(无主题)",
        "isDraft": False,
        "from": {
            "emailAddress": {
                "name": from_addr.get("name", ""),
                "address": from_addr.get("address", "")
            }
        }
    }

    if body_html:
        msg["body"] = {"contentType": "HTML", "content": body_html}
    elif body_text:
        msg["body"] = {"contentType": "Text", "content": body_text}
    else:
        msg["body"] = {"contentType": "Text", "content": ""}

    if to_addrs:
        msg["toRecipients"] = [
            {"emailAddress": {"name": a["name"], "address": a["address"]}} for a in to_addrs
        ]
    if cc_addrs:
        msg["ccRecipients"] = [
            {"emailAddress": {"name": a["name"], "address": a["address"]}} for a in cc_addrs
        ]
    if bcc_addrs:
        msg["bccRecipients"] = [
            {"emailAddress": {"name": a["name"], "address": a["address"]}} for a in bcc_addrs
        ]

    # 关键: 扩展属性设置时间字段
    # - SystemTime 0x0e06 (PidTagMessageDeliveryTime): 投递时间 -> receivedDateTime
    # - SystemTime 0x0039 (PidTagClientSubmitTime): 提交时间 -> sentDateTime
    # - SystemTime 0x002a (PidTagReceiptTime): 收据时间, 部分 Outlook 排序会参考
    # - Integer 0x0e07 (PidTagMessageFlags):
    #     MSGFLAG_READ=0x01, MSGFLAG_SENT=0x02, MSGFLAG_UNSENT=0x08
    #     收件箱邮件: 0x01 (已读) 或 0x00 (未读)
    #     已发送邮件: 0x02 (已发送) - 不设此位会显示为草稿
    #     当前默认 0x01 (收件箱场景). 已发送场景需调用方传 is_sent 参数.
    msg["singleValueExtendedProperties"] = [
        {"id": "SystemTime 0x0e06", "value": received_iso},
        {"id": "SystemTime 0x0039", "value": received_iso},
        {"id": "SystemTime 0x002a", "value": received_iso},
        {"id": "Integer 0x0e07", "value": "1"}
    ]

    return msg


def graph_post(url, token, payload, content_type="application/json"):
    body = payload if isinstance(payload, bytes) else json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(url, data=body, method="POST",
                                 headers={"Authorization": f"Bearer {token}",
                                          "Content-Type": content_type})
    return urllib.request.urlopen(req)


def graph_upload_mime(token, user_email, folder_id, eml_path):
    """
    方案 A: 直接用 message/rfc822 上传原始 EML (服务器端解析)
    微软官方支持, 但对畸形头严格. 失败时返回 False 让调用方 fallback.
    """
    with open(eml_path, "rb") as f:
        eml_bytes = f.read()
    url = f"{GRAPH_BASE}/users/{urllib.parse.quote(user_email)}/mailFolders/{folder_id}/messages"
    try:
        resp = graph_post(url, token, eml_bytes, content_type="message/rfc822")
        return True, f"HTTP {resp.status} (MIME upload)"
    except urllib.error.HTTPError as e:
        return False, f"HTTP {e.code}: {e.read().decode('utf-8', errors='replace')[:200]}"


def graph_post_message(token, user_email, folder_id, message_obj):
    """
    方案 B: 构造清洁的 Message JSON 对象 (推荐, 容忍畸形 EML)
    """
    url = f"{GRAPH_BASE}/users/{urllib.parse.quote(user_email)}/mailFolders/{folder_id}/messages"
    try:
        resp = graph_post(url, token, message_obj, content_type="application/json")
        return True, f"HTTP {resp.status} (Message object)"
    except urllib.error.HTTPError as e:
        return False, f"HTTP {e.code}: {e.read().decode('utf-8', errors='replace')[:200]}"


def deliver_to_graph(token, user_email, eml_path, target_folder=None):
    log(f"  [Graph] User: {user_email}")
    log(f"  [Graph] EML: {os.path.basename(eml_path)}")
    log(f"  [Graph] Target: {target_folder or '(root)'}")

    # 1. 解析 EML (用 Python email 库, 容忍畸形头)
    try:
        subject, from_addr, to_addrs, cc_addrs, bcc_addrs, body_html, body_text, received_iso = parse_eml_file(eml_path)
    except Exception as e:
        return False, f"EML 解析失败: {e}"

    log(f"  [Graph] Subject: {subject[:60]}")
    log(f"  [Graph] From: {from_addr.get('address', '')}")
    log(f"  [Graph] ReceivedDateTime: {received_iso}")
    log(f"  [Graph] To: {len(to_addrs)}, Cc: {len(cc_addrs)}")

    # 2. 解析/创建目标文件夹
    folder_id = "inbox"
    if target_folder:
        try:
            folder_id = find_or_create_folder(token, user_email, target_folder)
        except urllib.error.HTTPError as e:
            return False, f"文件夹解析失败: HTTP {e.code}: {e.read().decode('utf-8', errors='replace')[:200]}"
        except Exception as e:
            return False, f"文件夹解析失败: {e}"
    log(f"  [Graph] FolderId: {folder_id}")

    # 3. 先尝试方案 A (raw MIME), 失败则方案 B (Message 对象)
    ok, msg = graph_upload_mime(token, user_email, folder_id, eml_path)
    if ok:
        log(f"  [Graph] OK via raw MIME: {msg}")
        return True, msg

    log(f"  [Graph] raw MIME failed ({msg}), fallback to Message object...")
    message_obj = build_message_object(subject, from_addr, to_addrs, cc_addrs, bcc_addrs, body_html, body_text, received_iso)
    ok, msg = graph_post_message(token, user_email, folder_id, message_obj)
    if ok:
        log(f"  [Graph] OK via Message object: {msg}")
        return True, msg

    return False, f"两种方案均失败: MIME={msg}"


if __name__ == "__main__":
    if len(sys.argv) < 4:
        print("Usage: graph_deliver.py <access_token> <user_email> <eml_path> [target_folder]", file=sys.stderr)
        sys.exit(1)

    token = sys.argv[1]
    user_email = sys.argv[2]
    eml_path = sys.argv[3]
    target_folder = sys.argv[4] if len(sys.argv) > 4 else None

    if not os.path.exists(eml_path):
        print(f"EML file not found: {eml_path}", file=sys.stderr)
        sys.exit(1)

    success, message = deliver_to_graph(token, user_email, eml_path, target_folder)
    print(f"RESULT: {'SUCCESS' if success else 'FAILED'}: {message}")
    sys.exit(0 if success else 1)
