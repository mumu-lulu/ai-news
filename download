#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AI资讯每日邮件推送脚本（跨平台通用版：Windows / Linux / macOS）
仅依赖 Python 3.8+ 标准库，无需安装任何第三方包。

功能：
  1. 抓取当天最新 AI 资讯（量子位头条 + IT之家RSS按AI关键词过滤）
  2. 排版为 HTML 邮件
  3. 通过 SMTP 发送到指定邮箱

用法（在本脚本所在目录执行）：
  python ai_news_sender.py test    # 发送一封测试邮件（单条资讯）
  python ai_news_sender.py once    # 立即抓取并发送一次今日资讯
  python ai_news_sender.py daemon  # 常驻模式：每天 08:00（北京时间）自动发送

配置：同目录下 smtp_config.json
  {
    "smtp_host": "smtp.qq.com",
    "smtp_port": 465,
    "sender_email": "你的邮箱",
    "auth_code": "SMTP授权码（非登录密码）",
    "receiver_email": "收件邮箱"
  }

Windows 定时推荐（任务计划程序，管理员CMD执行一次即可）：
  schtasks /Create /TN "AI资讯每日推送" /SC DAILY /ST 08:00 ^
    /TR "python D:\\path\\to\\ai_news_sender.py once"
"""
import json
import os
import re
import smtplib
import sys
import time
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from datetime import datetime, date, timedelta, timezone
from email.header import Header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr
from pathlib import Path

# Windows 控制台中文输出保护
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

BASE = Path(__file__).resolve().parent
CONFIG_FILE = BASE / "smtp_config.json"
LAST_SENT_FILE = BASE / ".last_sent"
SEND_HOUR = 8          # 每天早上8点
TZ = timezone(timedelta(hours=8))   # Asia/Shanghai

UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}


def load_config():
    """优先读取环境变量（适配GitHub Actions/云函数的Secrets注入），否则读取本地smtp_config.json"""
    file_cfg = {}
    if CONFIG_FILE.exists():
        with open(CONFIG_FILE, encoding="utf-8") as f:
            file_cfg = json.load(f)

    def pick(env_key, file_key, default=None):
        val = os.environ.get(env_key) or file_cfg.get(file_key) or default
        if val is None:
            sys.exit(f"[配置缺失] 请设置环境变量 {env_key}，或在 {CONFIG_FILE} 中配置 {file_key}")
        return val

    return {
        "smtp_host": pick("SMTP_HOST", "smtp_host", "smtp.qq.com"),
        "smtp_port": pick("SMTP_PORT", "smtp_port", 465),
        "sender_email": pick("SENDER_EMAIL", "sender_email"),
        "auth_code": pick("AUTH_CODE", "auth_code"),
        "receiver_email": pick("RECEIVER_EMAIL", "receiver_email"),
    }


def http_get(url, timeout=15, retries=2):
    """带重试的HTTP GET，返回文本"""
    for i in range(retries):
        try:
            req = urllib.request.Request(url, headers=UA)
            return urllib.request.urlopen(req, timeout=timeout).read().decode("utf-8", "ignore")
        except Exception as e:
            print(f"[警告] 抓取失败({i+1}/{retries}) {url}: {e}")
            time.sleep(2)
    return ""


def now_cn():
    return datetime.now(TZ)


def fetch_qbitai_headlines(limit=6):
    """抓取量子位首页最新文章标题与链接（AI垂直资讯）"""
    html = http_get("https://www.qbitai.com")
    if not html:
        return []
    items, seen = [], set()
    pat = re.compile(r'<a[^>]*href="(https://www\.qbitai\.com/20\d{2}/\d{1,2}/\d+\.html)"[^>]*>(.*?)</a>', re.S)
    for m in pat.finditer(html):
        url = m.group(1)
        title = re.sub(r"<[^>]+>", "", m.group(2))
        title = re.sub(r"\d{4}-\d{2}-\d{2}", "", title).strip()
        if url not in seen and len(title) >= 8:
            seen.add(url)
            items.append({"title": title, "url": url, "source": "量子位"})
    return items[:limit]


AI_KEYWORDS = ("AI", "人工智能", "大模型", "DeepSeek", "OpenAI", "GPT", "智能体",
               "Agent", "芯片", "算力", "Claude", "Gemini", "文心", "通义", "豆包",
               "智谱", "机器人", "具身", "多模态", "LLM", "Sora", "Copilot", "千问")


def fetch_tech_rss(limit=4):
    """从IT之家RSS抓取并按AI关键词过滤（补充资讯）"""
    xml_text = http_get("https://www.ithome.com/rss/")
    if not xml_text:
        return []
    items = []
    try:
        root = ET.fromstring(xml_text)
        for it in root.findall(".//item"):
            title = (it.findtext("title") or "").strip()
            link = (it.findtext("link") or "").strip()
            desc = re.sub(r"<[^>]+>", "", it.findtext("description") or "").strip()[:120]
            if title and link:
                items.append({"title": title, "url": link, "source": "IT之家", "snippet": desc})
    except ET.ParseError as e:
        print(f"[警告] RSS解析失败: {e}")
        return []
    ai_items = [x for x in items if any(k.lower() in x["title"].lower() for k in AI_KEYWORDS)]
    return (ai_items or items)[:limit]


def build_digest():
    return fetch_qbitai_headlines(), fetch_tech_rss()


def render_html(today_str, headlines, search_news, is_test=False, single=None):
    """排版 HTML 邮件"""
    if is_test and single:
        head_item = f'''
        <div class="card">
          <h3>🔥 {single["title"]}</h3>
          <p>{single.get("summary", "")}</p>
          <a href="{single["url"]}">阅读原文 →</a>
        </div>'''
        more_block = ""
    else:
        cards = "".join(f'''
        <div class="card">
          <h3>📌 {it["title"]}</h3>
          <a href="{it["url"]}">阅读原文 →</a>
        </div>''' for it in headlines)
        head_item = f'<h2>今日头条资讯</h2>{cards}' if cards else '<p>今日未抓取到头条。</p>'
        rows = ""
        for it in search_news:
            snippet = f'<span class="snippet">{it.get("snippet","")}</span>' if it.get("snippet") else ""
            rows += f'<li><a href="{it["url"]}">{it["title"]}</a> {snippet}</li>'
        more_block = f'<h2>更多 AI 动态</h2><ul class="list">{rows}</ul>' if rows else ""

    test_banner = '<p class="banner">✉️ 这是一封测试邮件，用于验证邮件推送通道。收到即代表发送成功！</p>' if is_test else ""
    return f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<style>
  body {{ font-family: "PingFang SC","Microsoft YaHei",Arial,sans-serif; background:#f5f6fa; margin:0; padding:24px; color:#2c3e50; }}
  .wrap {{ max-width:640px; margin:0 auto; background:#ffffff; border-radius:12px; overflow:hidden; border:1px solid #e8eaf0; }}
  .header {{ background:linear-gradient(135deg,#4f6ef7,#7a5cf0); color:#fff; padding:26px 28px; }}
  .header h1 {{ margin:0; font-size:21px; }}
  .header p {{ margin:8px 0 0; font-size:13px; opacity:.9; }}
  .content {{ padding:24px 28px; }}
  h2 {{ font-size:16px; border-left:4px solid #4f6ef7; padding-left:10px; margin:22px 0 12px; }}
  .card {{ background:#f8f9fd; border:1px solid #eef0f6; border-radius:10px; padding:14px 16px; margin-bottom:12px; }}
  .card h3 {{ margin:0 0 6px; font-size:15px; line-height:1.5; }}
  .card a {{ color:#4f6ef7; font-size:13px; text-decoration:none; }}
  .banner {{ background:#fff8e6; border:1px solid #ffe9b0; color:#8a6d1a; border-radius:8px; padding:10px 14px; font-size:13px; }}
  .list {{ padding-left:18px; }}
  .list li {{ margin-bottom:10px; font-size:14px; line-height:1.6; }}
  .list a {{ color:#2c3e50; text-decoration:none; }}
  .snippet {{ color:#8a94a6; font-size:12px; }}
  .footer {{ text-align:center; color:#a0a8b8; font-size:12px; padding:16px; border-top:1px solid #eef0f6; }}
</style>
</head>
<body>
<div class="wrap">
  <div class="header">
    <h1>🤖 每日 AI 资讯速递</h1>
    <p>{today_str} · 自动推送</p>
  </div>
  <div class="content">
    {test_banner}
    {head_item}
    {more_block}
  </div>
  <div class="footer">本邮件由自动化脚本定时发送 · 资讯来源于公开网络</div>
</div>
</body>
</html>'''


def send_mail(config, subject, html):
    msg = MIMEMultipart("alternative")
    msg["From"] = formataddr((str(Header("AI资讯速递", "utf-8")), config["sender_email"]))
    msg["To"] = config["receiver_email"]
    msg["Subject"] = Header(subject, "utf-8")
    msg.attach(MIMEText(html, "html", "utf-8"))

    host = config.get("smtp_host", "smtp.qq.com")
    port = int(config.get("smtp_port", 465))
    if port == 465:
        server = smtplib.SMTP_SSL(host, port, timeout=30)
    else:
        server = smtplib.SMTP(host, port, timeout=30)
        server.starttls()
    try:
        server.login(config["sender_email"], config["auth_code"])
        server.sendmail(config["sender_email"], [config["receiver_email"]], msg.as_string())
    finally:
        server.quit()
    print(f"[OK] 邮件已发送: {subject} -> {config['receiver_email']}")


def already_sent_today():
    today = now_cn().strftime("%Y-%m-%d")
    if not LAST_SENT_FILE.exists():
        return False
    return LAST_SENT_FILE.read_text(encoding="utf-8").strip() == today


def mark_sent():
    LAST_SENT_FILE.write_text(now_cn().strftime("%Y-%m-%d"), encoding="utf-8")


def job_daily():
    today_str = now_cn().strftime("%Y-%m-%d")
    print(f"[INFO] 开始抓取今日AI资讯 {today_str}")
    headlines, more = build_digest()
    if not headlines and not more:
        print("[警告] 未抓取到资讯，本次跳过")
        return False
    html = render_html(today_str, headlines, more)
    send_mail(load_config(), f"每日AI资讯速递 · {today_str}", html)
    mark_sent()
    return True


def job_test():
    config = load_config()
    today_str = now_cn().strftime("%Y-%m-%d")
    single = {
        "title": "突发！DeepSeek新模型发布，超级性价比斩杀一切",
        "url": "https://www.qbitai.com/2026/10/504067.html",
        "summary": "DeepSeek 于10月17日晚发布新模型 V4.1-Flash-1017：代码、Agent任务等能力明显提升，同时任务运行时间与输出token消耗反而下降，官方称为“性能—费用的帕累托改进”；同步增强 DeepSeek Harness 0.2.1，并开放自研搜索API，网页端/APP/API已同步切换新模型。",
    }
    html = render_html(today_str, [], [], is_test=True, single=single)
    send_mail(config, f"【测试邮件】每日AI资讯速递 · {today_str}", html)


def job_daemon():
    print(f"[INFO] 定时任务已启动：每天 {SEND_HOUR:02d}:00（北京时间）推送AI资讯。Ctrl+C 退出。")
    last_day = None
    while True:
        n = now_cn()
        today = n.strftime("%Y-%m-%d")
        sec = n.hour * 3600 + n.minute * 60 + n.second
        window_start = SEND_HOUR * 3600
        if last_day != today and window_start <= sec < window_start + 3600:
            if not already_sent_today():
                try:
                    job_daily()
                except Exception as e:
                    print(f"[错误] 发送失败，30分钟后重试: {e}")
                    time.sleep(1800)
                    continue
            last_day = today
        time.sleep(30)


if __name__ == "__main__":
    mode = sys.argv[1] if len(sys.argv) > 1 else ""
    if mode == "test":
        job_test()
    elif mode == "once":
        job_daily()
    elif mode == "daemon":
        job_daemon()
    else:
        print("用法: python ai_news_sender.py [test|once|daemon]")
