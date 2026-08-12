#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Project Copy-Save: Wechat Project Copy-Save (Markdown + Word + 英文文件名)
功能：读取剪贴板 -> 清洗 -> 翻译文件名 -> 转MD -> 转Word (通用开源版)
"""

import os
import sys
import re
import time
import json
import shutil
import subprocess
import requests
import pyperclip
import html2text
from bs4 import BeautifulSoup
from datetime import datetime
from pathlib import Path
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# 加载 .env
from dotenv import load_dotenv
load_dotenv(Path(__file__).parent / ".env")

# ================= 配置 =================
BASE = Path(__file__).parent
PREPROCESS_PATH = BASE / "preprocess_md.py"
FORMAT_EXPERT_PATH = BASE / "format_expert.py"
GOLDEN_TEMPLATE = BASE / "黄金模板.docx"
CONFIG_PATH = BASE / "config.json"

# 请求头
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.60 Safari/537.36 MicroMessenger/7.0.20.1781(0x6700143B) NetType/WIFI MiniProgramEnv/Windows WindowsWechat/WMPF XWEB/6939"
}

# ================= 路径配置与初始化 =================
def load_config():
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def save_config(config_data):
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(config_data, f, indent=4, ensure_ascii=False)
    except Exception as e:
        print(f"写入配置失败: {e}")

config = load_config()
SAVE_DIR = config.get("word_save_dir", "")
OBSIDIAN_BASE = config.get("obsidian_base_dir", "")
DOCX_ARCHIVE_DIR = config.get("docx_archive_dir", "")

def initialize_paths():
    global SAVE_DIR, OBSIDIAN_BASE
    if not SAVE_DIR:
        desktop_path = Path(os.path.expanduser("~/Desktop"))
        default_save_dir = desktop_path / "微信公众号归档"
        try:
            default_save_dir.mkdir(parents=True, exist_ok=True)
            SAVE_DIR = str(default_save_dir.resolve())
            OBSIDIAN_BASE = ""
            config["word_save_dir"] = SAVE_DIR
            config["obsidian_base_dir"] = OBSIDIAN_BASE
            save_config(config)
            print("⚙️ 首次启动：已自动在桌面创建“微信公众号归档”文件夹作为保存路径！")
        except Exception as e:
            print(f"⚠️ 自动创建目录失败: {e}")

def get_pandoc_executable():
    """获取内置或本地同级目录或系统环境中的 pandoc 路径"""
    if hasattr(sys, "_MEIPASS"):
        local_pandoc = os.path.join(sys._MEIPASS, "pandoc.exe")
        if os.path.exists(local_pandoc):
            return local_pandoc
    local_pandoc_sibling = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pandoc.exe")
    if os.path.exists(local_pandoc_sibling):
        return local_pandoc_sibling
    return "pandoc"

# ===== 文章自动分类逻辑 =====
CATEGORY_RULES = [
    {
        "folder": r"1.AI技术汇编",
        "tag": "AI技术",
        "keywords": [
            "AI", "大模型", "Claude", "Gemini", "GPT", "人工智能", "机器学习",
            "算法", "编程", "Python", "自动化工具", "API"
        ]
    },
    {
        "folder": r"2.HSE工作笔记库",
        "tag": "HSE",
        "keywords": [
            "安全", "吊装", "危大", "风险", "事故", "隐患", "整改",
            "监理", "应急", "消防", "职业健康", "安全培训", "安全生产", "防护"
        ]
    },
    {
        "folder": r"3.中英文术语库",
        "tag": "中英文术语",
        "keywords": [
            "术语", "翻译", "对照", "词表", "双语", "词汇", "缩略语", "词典"
        ]
    },
    {
        "folder": r"4.施工技术汇编",
        "tag": "施工技术",
        "keywords": [
            "施工", "基坑", "钢结构", "脚手架", "高处作业", "模板工程",
            "港口", "码头", "桩基", "混凝土", "土方", "防水", "机电安装", "施工方案"
        ]
    },
]

DEFAULT_FOLDER = r"公众号归档"

def get_target_folder_and_tags(title: str, content: str = "") -> tuple:
    """根据标题和正文前200字，判断归档路径和标签"""
    search_text = title + content[:200]
    tags = ["待分类"]
    folder_name = DEFAULT_FOLDER
    
    for rule in CATEGORY_RULES:
        matched = False
        for kw in rule["keywords"]:
            if kw.lower() in search_text.lower():
                folder_name = rule["folder"]
                if rule["folder"] == r"2.HSE工作笔记库" and "吊装" in search_text:
                    tags = ["HSE", "吊装"]
                else:
                    tags = [rule["tag"]]
                matched = True
                break
        if matched:
            break
            
    if not OBSIDIAN_BASE:
        return "", tags
        
    folder_path = os.path.join(OBSIDIAN_BASE, folder_name)
    return folder_path, tags

def clean_filename(title):
    """清理文件名中的非法字符"""
    return re.sub(r'[\\/:*?"<>|]', '_', title).strip()

def get_session():
    """创建带有重试机制的会话"""
    session = requests.Session()
    retry = Retry(
        total=3, 
        read=3, 
        connect=3, 
        backoff_factor=1, 
        status_forcelist=[500, 502, 503, 504]
    )
    adapter = HTTPAdapter(max_retries=retry)
    session.mount('http://', adapter)
    session.mount('https://', adapter)
    session.headers.update(HEADERS)
    return session

def save_article(url):
    print(f"🔗 检测到链接: {url}")
    print("⏳ 正在请求文章内容 (已启用重试机制)...")
    
    try:
        session = get_session()
        resp = session.get(url, timeout=30)
        if resp.status_code != 200:
             print(f"❌ 服务器返回错误: {resp.status_code}")
             return False
    except Exception as e:
        print(f"❌ 下载失败: {e}")
        print("💡 建议：网络可能不稳定，请稍后重试。")
        return False

    soup = BeautifulSoup(resp.content, "html.parser")
    
    # 1. 提取元数据
    try:
        title = soup.find("meta", property="og:title")["content"]
    except:
        try:
            title = soup.find(id="activity-name").get_text(strip=True)
        except:
            title = "未知标题_" + str(int(time.time()))
    
    try:
        profile_name = soup.find(id="js_name").get_text(strip=True)
    except:
        profile_name = "公众号"

    print(f"📄 标题: {title}")
    
    title_final = clean_filename(title)
    
    # 2. 提取并修复正文
    content_div = soup.find(id="js_content")
    if not content_div:
        content_div = soup.find(class_="rich_media_content")
        
    if not content_div:
        print("⚠️ 未找到正文区域，可能不是标准的公众号文章。")
        return False

    # 3. 预处理 (Anti-Lazyload)
    imgs = content_div.find_all("img")
    count = 0
    for img in imgs:
        if img.get("data-src"):
            img["src"] = img["data-src"]
            del img["data-src"]
            count += 1
    print(f"🖼️ 已保留图片链接: {count} 张")

    # 4. 转换为 Markdown
    converter = html2text.HTML2Text()
    converter.ignore_links = False
    converter.ignore_images = False
    converter.ignore_tables = False
    converter.body_width = 0 
    
    html_str = str(content_div)
    md_content = converter.handle(html_str)
    
    # 自动获取分类文件夹和标签
    archive_dir, tags = get_target_folder_and_tags(title, md_content)
    tags_str = "[" + ", ".join(tags) + "]"

    # 5. 组装最终文档 (顶部自动注入 Front Matter 格式)
    today_str = datetime.now().strftime("%Y-%m-%d")
    final_md = f"""---
title: {title}
date: {today_str}
source_url: {url}
tags: {tags_str}
summary: 
---

# {title}

> **来源**: {profile_name}
> **归档日期**: {today_str}
> **原文链接**: [{url}]({url})

---

{md_content}

---
*Created by Project Copy-Save*
"""

    # 6. 保存文件 (Markdown)
    if not os.path.exists(SAVE_DIR):
        os.makedirs(SAVE_DIR)

    filename_base = f"[{today_str}] {title_final}"
    filename_md = f"{filename_base}.md"
    filename_docx = f"{filename_base}.docx"
    
    file_path_md = os.path.join(SAVE_DIR, filename_md)
    file_path_docx = os.path.join(SAVE_DIR, filename_docx)
    
    if os.path.exists(file_path_md):
        print("⚠️ 文件已存在，自动覆盖...")

    with open(file_path_md, "w", encoding="utf-8") as f:
        f.write(final_md)
        
    print(f"✅ Markdown 保存成功: {filename_md}")
    
    # 7. 三步管线：预处理 → Pandoc → 精确排版
    print("⏳ 正在生成 Word 文档...")
    try:
        escaped_md = file_path_md.replace(".md", "_escaped.md")
        temp_docx  = file_path_md.replace(".md", "_temp.docx")

        # Step 1: 预处理 Markdown
        try:
            import preprocess_md
            preprocess_md.preprocess_markdown(file_path_md, escaped_md)
            print("   [1/3] Markdown 预处理完成")
        except Exception as e:
            print(f"   [1/3] 预处理失败: {e}")
            escaped_md = file_path_md
            print("   ↳ 跳过预处理，使用原始 Markdown")

        # Step 2: Pandoc 转 docx
        pandoc_bin = get_pandoc_executable()
        try:
            pandoc_cmd = [pandoc_bin, escaped_md, "-o", temp_docx]
            if GOLDEN_TEMPLATE.exists():
                pandoc_cmd.extend(["--reference-doc", str(GOLDEN_TEMPLATE)])
            subprocess.run(pandoc_cmd, check=True, capture_output=True, text=True)
            print("   [2/3] Pandoc 转换完成")
        except FileNotFoundError:
            print("⚠️ 未找到 pandoc，仅保存 Markdown。安装: choco install pandoc")
            return True
        except subprocess.CalledProcessError as e:
            print(f"   [2/3] Pandoc 失败: {e.stderr}")
            # 无模板重试
            pandoc_cmd = [pandoc_bin, escaped_md, "-o", temp_docx]
            subprocess.run(pandoc_cmd, check=True, capture_output=True, text=True)
            print("   ↳ 无模板重试成功")

        # Step 3: format_expert.py 精确排版
        try:
            import format_expert
            formatter = format_expert.DocumentFormatter(temp_docx, file_path_docx, add_pagenum=True)
            formatter.format()
            print("   [3/3] 排版处理完成")
        except Exception as e:
            print(f"   [3/3] 排版失败: {e}")
            shutil.copy(temp_docx, file_path_docx)
            print("   ↳ 回退到基础 Pandoc 版本")

        # 清理临时文件
        for f in [escaped_md, temp_docx]:
            if os.path.exists(f) and f != file_path_md:
                os.remove(f)

        print(f"✅ Word 保存成功: {filename_docx}")
        print(f"📍 文件位置: {SAVE_DIR}")

        # 根据标题和正文关键词，动态判断目标分类归档目录
        if OBSIDIAN_BASE:
            try:
                if not os.path.exists(archive_dir):
                    os.makedirs(archive_dir)
                archive_md_path = os.path.join(archive_dir, filename_md)
                shutil.copy2(file_path_md, archive_md_path)
                print(f"📂 已同步复制 Markdown 中间文件至归档目录: {archive_md_path}")

                # 成功复制到归档目录后，删除原输出目录下的 md 临时文件
                if os.path.exists(archive_md_path) and os.path.exists(file_path_md):
                    os.remove(file_path_md)
            except Exception as e:
                print(f"⚠️ 同步归档 Obsidian 失败: {e}")
        else:
            print("ℹ️ 未启用 Obsidian 归档，跳过同步。")

        if DOCX_ARCHIVE_DIR:
            try:
                if not os.path.exists(DOCX_ARCHIVE_DIR):
                    os.makedirs(DOCX_ARCHIVE_DIR)
                archive_docx_path = os.path.join(DOCX_ARCHIVE_DIR, filename_docx)
                shutil.copy2(file_path_docx, archive_docx_path)
                print(f" 已同步复制 Word 文件至归档目录: {archive_docx_path}")
                if os.path.exists(archive_docx_path) and os.path.exists(file_path_docx):
                    os.remove(file_path_docx)
            except Exception as e:
                print(f" 同步归档 Word 文件失败: {e}")
        else:
            print("ℹ 未启用 Word 归档，跳过同步。")

    except Exception as e:
        print(f"❌ 转换出错: {e}")
        return False

    return True

def main():
    print("="*40)
    print("      Project Copy-Save | 一键归档 (通用开源版)")
    print("="*40)
    
    initialize_paths()
    
    try:
        content = pyperclip.paste()
    except Exception as e:
        print(f"❌ 读取剪贴板失败: {e}")
        return

    if not content:
        print("📭 剪贴板为空！")
        return
        
    url_match = re.search(r'(https://mp\.weixin\.qq\.com/s/[a-zA-Z0-9_\-]+)', content)
    
    if not url_match:
         if "mp.weixin.qq.com" in content and "http" in content:
             url = content.strip()
         else:
             print("📭 未发现有效链接。")
             return
    else:
        url = url_match.group(0)

    success = save_article(url)
    
    if success:
        time.sleep(2)
    else:
        print("\n运行未完全成功，请检查错误提示...")

if __name__ == "__main__":
    main()
