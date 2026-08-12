#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
微信公众号一键归档系统 GUI 客户端 (通用开源版)
包含首次运行路径配置向导，以及智能检测与安装 Pandoc 环境向导（支持 winget 失败自动双重兜底）。
"""

import os
import sys
import re
import time
import json
import shutil
import subprocess
import threading
import queue
import uuid
import tempfile
import requests
import pyperclip
import html2text
import winreg
import webbrowser
from bs4 import BeautifulSoup
from datetime import datetime
from pathlib import Path
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

import tkinter as tk
from tkinter import messagebox, filedialog
from tkinter.scrolledtext import ScrolledText

# ================= 配置 =================
BASE = Path(__file__).parent
PREPROCESS_PATH = BASE / "preprocess_md.py"
FORMAT_EXPERT_PATH = BASE / "format_expert.py"
GOLDEN_TEMPLATE = BASE / "黄金模板.docx"
CONFIG_PATH = BASE / "config.json"

# 加载 .env
from dotenv import load_dotenv
load_dotenv(BASE / ".env")

# 请求头
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/100.0.4896.60 Safari/537.36 MicroMessenger/7.0.20.1781(0x6700143B) NetType/WIFI MiniProgramEnv/Windows WindowsWechat/WMPF XWEB/6939"
}

# 分类规则
CATEGORY_RULES = [
    {
        "folder": r"1.AI技术汇编",
        "tag": "AI技术",
        "keywords": [
            "AI", "人工智能", "Claude", "Gemini", "GPT", "ChatGPT",
            "大模型", "LLM", "DeepSeek", "Copilot", "机器学习",
            "深度学习", "神经网络", "提示词", "Prompt", "Agent",
            "自动化", "MCP", "Python", "编程", "代码", "开发者",
            "Antigravity", "Claude Code", "API", "插件", "工具系统",
            "知识库", "Obsidian", "NotebookLM", "数字化", "智能化"
        ]
    },
    {
        "folder": r"2.HSE工作笔记库",
        "tag": "HSE",
        "keywords": [
            "HSE", "安全", "危大", "风险", "隐患", "事故",
            "吊装", "起重", "吊耳", "索具", "卸扣",
            "脚手架", "模板支架", "高处作业", "临边防护",
            "机械", "设备", "特种作业", "特种设备",
            "环保", "环境", "废水", "废气", "噪音", "固废",
            "职业健康", "劳保", "防护用品", "应急", "消防",
            "安全培训", "安全管理", "安全规范", "安全检查",
            "违章", "整改", "监理", "验收"
        ]
    },
    {
        "folder": r"3.中英文术语库",
        "tag": "中英文术语",
        "keywords": [
            "中英", "英译", "中英文", "对照", "术语",
            "词汇", "词表", "词汇表", "翻译", "Glossary", "terminology"
        ]
    },
    {
        "folder": r"4.施工技术汇编",
        "tag": "施工技术",
        "keywords": [
            "施工管理", "施工技术", "施工工艺", "施工方案",
            "基坑", "深基坑", "土方", "围护",
            "港口", "码头", "散货", "泊位", "护岸", "海工",
            "建筑工程", "土建", "主体结构", "混凝土",
            "钢结构", "钢筋", "焊接", "预埋",
            "悬挑模架", "爬架", "塔吊", "施工电梯",
            "测量", "放线", "质量控制", "工程管理",
            "进度计划", "工期", "竣工", "验收"
        ]
    },
]

DEFAULT_FOLDER = r"公众号归档"

# ================= 辅助函数 =================
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

def get_target_folder_and_tags(obsidian_base, title: str, content: str = "") -> tuple:
    """根据标题和正文前200字判断自动分类文件夹和对应标签"""
    raw_text = title + content[:200]
    # 剔除 Markdown 图片链接语法（![]()），避免链接乱码子串误触发关键词匹配
    search_text = re.sub(r'!\[.*?\]\([^)]*\)', '', raw_text)
    tags = ["待分类"]
    folder_name = DEFAULT_FOLDER
    
    for rule in CATEGORY_RULES:
        matched = False
        for kw in rule["keywords"]:
            # 对纯大写短词（如 AI）使用全词正则匹配，避免被英文单词子串误伤
            if kw == "AI":
                kw_matched = bool(re.search(r'\bAI\b', search_text))
            else:
                kw_matched = kw.lower() in search_text.lower()
            if kw_matched:
                folder_name = rule["folder"]
                if rule["folder"] == r"2.HSE工作笔记库" and "吊装" in search_text:
                    tags = ["HSE", "吊装"]
                else:
                    tags = [rule["tag"]]
                matched = True
                break
        if matched:
            break
            
    if not obsidian_base:
        return "", tags
        
    folder_path = os.path.join(obsidian_base, folder_name)
    return folder_path, tags

def add_hover(widget, normal_bg, hover_bg):
    """为 Tkinter 控件绑定悬停高亮动画效果"""
    widget.bind("<Enter>", lambda e: widget.config(bg=hover_bg))
    widget.bind("<Leave>", lambda e: widget.config(bg=normal_bg))

def load_config():
    """加载配置文件"""
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}

def save_config(config):
    """保存配置文件"""
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=4, ensure_ascii=False)
    except Exception as e:
        print(f"写入配置文件失败: {e}")

# ================= 动态环境变量检测逻辑 =================
def refresh_path():
    """从 Windows 注册表动态刷新当前进程的 PATH 环境变量，避免进程重启"""
    try:
        # 读取系统级 PATH
        sys_path = ""
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Control\Session Manager\Environment") as key:
                sys_path, _ = winreg.QueryValueEx(key, "PATH")
        except Exception:
            pass

        # 读取用户级 PATH
        user_path = ""
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Environment") as key:
                user_path, _ = winreg.QueryValueEx(key, "PATH")
        except Exception:
            pass

        combined = sys_path
        if user_path:
            combined = combined + ";" + user_path
            
        os.environ["PATH"] = os.path.expandvars(combined)
    except Exception as e:
        print(f"动态更新环境变量 PATH 失败: {e}")

def get_pandoc_executable():
    """获取内置或本地同级目录或系统环境中的 pandoc 路径"""
    # 1. 优先使用打包内置的 pandoc
    if hasattr(sys, "_MEIPASS"):
        local_pandoc = os.path.join(sys._MEIPASS, "pandoc.exe")
        if os.path.exists(local_pandoc):
            return local_pandoc
    
    # 2. 其次尝试脚本同级目录下的 pandoc.exe
    local_pandoc_sibling = os.path.join(os.path.dirname(os.path.abspath(__file__)), "pandoc.exe")
    if os.path.exists(local_pandoc_sibling):
        return local_pandoc_sibling

    # 3. 最后回退使用系统 PATH 中的 pandoc
    return "pandoc"

def check_pandoc_installed():
    """由于内置了 Pandoc，此项始终为 True"""
    return True

# ================= GUI 主窗口 =================
class WechatGrabberApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("微信公众号一键归档系统 (开源通用版)")
        self.geometry("650x500")
        self.resizable(False, False)
        
        # 窗口居中
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - 650) // 2
        y = (screen_height - 500) // 2
        self.geometry(f"650x500+{x}+{y}")
        
        # 现代暗系主题配色
        self.bg_color = "#1e293b"      # slate 800
        self.card_bg = "#0f172a"       # slate 900
        self.text_color = "#f8fafc"    # slate 50
        self.accent_color = "#10b981"  # emerald 500
        self.accent_hover = "#059669"  # emerald 600
        self.btn_bg = "#475569"        # slate 600
        self.btn_hover = "#334155"     # slate 700
        self.log_bg = "#0f172a"        # slate 900
        self.log_fg = "#34d399"        # emerald 400
        
        self.configure(bg=self.bg_color)
        self.setup_ui()
        
        # 消息队列，用于线程间通信更新 UI
        self.queue = queue.Queue()
        self.check_queue_loop()
        
        # 读取配置并检查路径初始化
        self.config = load_config()
        self.word_save_dir = self.config.get("word_save_dir", "")
        self.obsidian_base_dir = self.config.get("obsidian_base_dir", "")
        
        # 如果路径未设置，直接在桌面创建默认文件夹并静默配置
        if not self.word_save_dir:
            self.setup_paths_default()
        else:
            self.write_log("⚙️ 配置已加载:")
            self.write_log(f"- Word 保存路径: {self.word_save_dir}")
            self.write_log(f"- Obsidian 归档路径: {self.obsidian_base_dir or '未启用'}")
            
        # 启动后延迟读取剪贴板并尝试填入链接
        self.after(1000, self.auto_paste_clipboard)

    def setup_ui(self):
        # 1. 顶部标题栏
        header_frame = tk.Frame(self, bg=self.bg_color, pady=12)
        header_frame.pack(fill="x")
        title_label = tk.Label(
            header_frame, 
            text="微信公众号一键归档系统", 
            font=("Microsoft YaHei", 16, "bold"), 
            bg=self.bg_color, 
            fg=self.accent_color
        )
        title_label.pack(side="left", padx=(25, 0))

        # ⚙️ 设置按钮 (挂载在右上角)
        btn_settings = tk.Button(
            header_frame, 
            text="⚙️ 设置", 
            command=self.setup_paths_manual,
            font=("Microsoft YaHei", 9), 
            bg=self.btn_bg, 
            fg=self.text_color, 
            relief="flat", 
            padx=12, 
            pady=4,
            activebackground=self.btn_hover,
            activeforeground=self.text_color
        )
        btn_settings.pack(side="right", padx=25)
        add_hover(btn_settings, self.btn_bg, self.btn_hover)

        # 2. 链接输入栏
        input_frame = tk.Frame(self, bg=self.bg_color, padx=25, pady=5)
        input_frame.pack(fill="x")
        
        tk.Label(
            input_frame, 
            text="微信文章链接 (URL):", 
            font=("Microsoft YaHei", 10), 
            bg=self.bg_color, 
            fg=self.text_color
        ).pack(anchor="w", pady=(0, 4))
        
        entry_row = tk.Frame(input_frame, bg=self.bg_color)
        entry_row.pack(fill="x")
        
        self.url_entry = tk.Entry(
            entry_row, 
            font=("Microsoft YaHei", 10), 
            bg=self.card_bg, 
            fg=self.text_color, 
            insertbackground=self.text_color,
            relief="flat", 
            bd=6
        )
        self.url_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        
        btn_paste = tk.Button(
            entry_row, 
            text="粘贴", 
            command=self.manual_paste,
            font=("Microsoft YaHei", 9), 
            bg=self.btn_bg, 
            fg=self.text_color, 
            relief="flat", 
            padx=12, 
            pady=3,
            activebackground=self.btn_hover,
            activeforeground=self.text_color
        )
        btn_paste.pack(side="left", padx=2)
        add_hover(btn_paste, self.btn_bg, self.btn_hover)
        
        btn_clear = tk.Button(
            entry_row, 
            text="清空", 
            command=self.clear_input,
            font=("Microsoft YaHei", 9), 
            bg=self.btn_bg, 
            fg=self.text_color, 
            relief="flat", 
            padx=12, 
            pady=3,
            activebackground=self.btn_hover,
            activeforeground=self.text_color
        )
        btn_clear.pack(side="left", padx=2)
        add_hover(btn_clear, self.btn_bg, self.btn_hover)

        # 3. 居中大号一键保存按钮
        btn_save_frame = tk.Frame(self, bg=self.bg_color, pady=12)
        btn_save_frame.pack(fill="x")
        
        self.btn_save = tk.Button(
            btn_save_frame, 
            text="一键保存", 
            command=self.start_capture,
            font=("Microsoft YaHei", 12, "bold"), 
            bg=self.accent_color, 
            fg=self.text_color, 
            relief="flat", 
            width=24, 
            pady=6,
            activebackground=self.accent_hover,
            activeforeground=self.text_color
        )
        self.btn_save.pack(anchor="center")
        add_hover(self.btn_save, self.accent_color, self.accent_hover)

        # 4. 实时日志窗口
        log_frame = tk.Frame(self, bg=self.bg_color, padx=25, pady=5)
        log_frame.pack(fill="both", expand=True)
        
        tk.Label(
            log_frame, 
            text="实时抓取进度日志:", 
            font=("Microsoft YaHei", 10), 
            bg=self.bg_color, 
            fg=self.text_color
        ).pack(anchor="w", pady=(0, 4))
        
        self.log_text = ScrolledText(
            log_frame, 
            bg=self.log_bg, 
            fg=self.log_fg, 
            insertbackground=self.log_fg, 
            font=("Consolas", 10), 
            state="disabled", 
            relief="flat",
            bd=5
        )
        self.log_text.pack(fill="both", expand=True)

        # 5. 底部状态栏
        self.status_var = tk.StringVar(value="最后保存文件：无")
        self.status_bar = tk.Label(
            self, 
            textvariable=self.status_var, 
            bg=self.card_bg, 
            fg="#94a3b8", 
            font=("Microsoft YaHei", 9), 
            anchor="w", 
            padx=15, 
            pady=5
        )
        self.status_bar.pack(fill="x", side="bottom")

    def setup_paths_default(self):
        """首次启动：自动在桌面创建“微信公众号归档”作为保存路径"""
        desktop_path = Path(os.path.expanduser("~/Desktop"))
        default_save_dir = desktop_path / "微信公众号归档"
        try:
            default_save_dir.mkdir(parents=True, exist_ok=True)
            self.word_save_dir = str(default_save_dir.resolve())
            self.obsidian_base_dir = "" # 默认跳过
            self.config["word_save_dir"] = self.word_save_dir
            self.config["obsidian_base_dir"] = self.obsidian_base_dir
            save_config(self.config)
            self.write_log("⚙️ 首次运行自动配置已完成:")
            self.write_log(f"- Word 默认保存路径: {self.word_save_dir}")
            self.write_log("ℹ️ 已自动在桌面创建“微信公众号归档”文件夹，您可以在右上角「设置」中修改此路径。")
        except Exception as e:
            self.write_log(f"⚠️ 首次运行自动创建目录失败: {e}")

        # 自动在桌面创建快捷方式（仅限打包后的 EXE 环境运行时触发，且仅首次创建）
        try:
            import win32com.client as wscript
            shell = wscript.Dispatch("WScript.Shell")
            shortcut_path = str(desktop_path / "微信公众号一键归档.lnk")
            if not os.path.exists(shortcut_path) and getattr(sys, "frozen", False):
                target_exe = sys.executable
                shortcut = shell.CreateShortCut(shortcut_path)
                shortcut.TargetPath = target_exe
                shortcut.WorkingDirectory = os.path.dirname(target_exe)
                shortcut.IconLocation = target_exe
                shortcut.Description = "微信公众号一键归档系统 - 双击启动"
                shortcut.save()
                self.write_log("🖥️ 已在桌面自动创建『微信公众号一键归档』快捷方式！")
        except Exception as e:
            self.write_log(f"ℹ️ 桌面快捷方式创建跳过: {e}")

    def setup_paths_manual(self):
        """手动选择修改保存路径"""
        # 1. 引导选择 Word 保存路径
        word_dir = filedialog.askdirectory(title="选择【Word文档】保存目标文件夹")
        if not word_dir:
            return # 取消
        
        # 2. 引导选择 Obsidian 归档路径 (可选)
        obsidian_dir = filedialog.askdirectory(title="选择【Obsidian库】的主目录文件夹 (可选，取消或直接关闭将不启用同步)")
        
        # 3. 更新配置值并保存
        self.word_save_dir = os.path.abspath(word_dir)
        self.obsidian_base_dir = os.path.abspath(obsidian_dir) if obsidian_dir else ""
        
        self.config["word_save_dir"] = self.word_save_dir
        self.config["obsidian_base_dir"] = self.obsidian_base_dir
        save_config(self.config)
        
        self.write_log("⚙️ 保存路径已更新:")
        self.write_log(f"- Word 保存路径: {self.word_save_dir}")
        self.write_log(f"- Obsidian 归档路径: {self.obsidian_base_dir or '未启用'}")
        messagebox.showinfo("成功", "保存路径配置已成功更新！")

    def auto_paste_clipboard(self):
        """启动时自动读取剪贴板"""
        try:
            content = pyperclip.paste()
            if content:
                url_match = re.search(r'(https://mp\.weixin\.qq\.com/s/[a-zA-Z0-9_\-]+)', content)
                if url_match:
                    url = url_match.group(0)
                    self.url_entry.delete(0, tk.END)
                    self.url_entry.insert(0, url)
                    self.write_log(f"📋 启动时已自动读取剪贴板链接: {url}")
                elif "mp.weixin.qq.com" in content and "http" in content:
                    url = content.strip()
                    self.url_entry.delete(0, tk.END)
                    self.url_entry.insert(0, url)
                    self.write_log(f"📋 启动时已自动读取剪贴板链接: {url}")
        except Exception as e:
            self.write_log(f"⚠️ 启动读取剪贴板失败: {e}")

    def manual_paste(self):
        """手动粘贴剪贴板"""
        try:
            content = pyperclip.paste()
            if content:
                self.url_entry.delete(0, tk.END)
                self.url_entry.insert(0, content.strip())
                self.write_log("📋 已成功粘贴剪贴板内容")
            else:
                messagebox.showinfo("提示", "当前剪贴板为空！")
        except Exception as e:
            self.write_log(f"⚠️ 手动粘贴失败: {e}")

    def clear_input(self):
        """清空输入框"""
        self.url_entry.delete(0, tk.END)
        self.write_log("🧹 已清空 URL 输入框")

    def write_log(self, text):
        """向日志队列放入一条消息"""
        self.queue.put(("log", f"{text}\n"))

    def check_queue_loop(self):
        """循环处理消息队列以更新 GUI"""
        try:
            while True:
                msg_type, val = self.queue.get_nowait()
                if msg_type == "log":
                    self.log_text.config(state="normal")
                    self.log_text.insert(tk.END, val)
                    self.log_text.see(tk.END)
                    self.log_text.config(state="disabled")
                elif msg_type == "status":
                    self.status_var.set(val)
                elif msg_type == "finished":
                    success = val
                    self.btn_save.config(state="normal", bg=self.accent_color)
                    self.url_entry.config(state="normal")
                    if success:
                        messagebox.showinfo("成功", "文章一键归档已顺利完成！")
                    else:
                        messagebox.showerror("错误", "抓取失败，请检查网络或输入的链接！")
        except queue.Empty:
            pass
        self.after(100, self.check_queue_loop)

    def start_capture(self):
        """开始执行抓取工作"""
        if not self.word_save_dir:
            self.setup_paths_default()
            
        url = self.url_entry.get().strip()
        if not url:
            messagebox.showwarning("警告", "请输入微信文章链接！")
            return
        
        # 禁用输入及保存按钮防止重复触发
        self.btn_save.config(state="disabled", bg="#475569")
        self.url_entry.config(state="disabled")
        
        # 清空日志区
        self.log_text.config(state="normal")
        self.log_text.delete("1.0", tk.END)
        self.log_text.config(state="disabled")
        
        # 启动后台工作线程
        threading.Thread(target=self.worker_thread_run, args=(url,), daemon=True).start()

    def worker_thread_run(self, url):
        """后台线程的抓取核心逻辑"""
        self.write_log(f"检测到链接: {url}")
        
        try:
            session = get_session()
            resp = session.get(url, timeout=30)
            if resp.status_code != 200:
                self.write_log(f"❌ 服务器响应失败: {resp.status_code}")
                self.queue.put(("finished", False))
                return
        except Exception as e:
            self.write_log(f"❌ 网络请求异常: {e}")
            self.queue.put(("finished", False))
            return

        soup = BeautifulSoup(resp.content, "html.parser")
        
        # 提取文章标题
        try:
            title = soup.find("meta", property="og:title")["content"]
        except:
            try:
                title = soup.find(id="activity-name").get_text(strip=True)
            except:
                title = "微信文章_" + str(int(time.time()))
        
        try:
            profile_name = soup.find(id="js_name").get_text(strip=True)
        except:
            profile_name = "公众号"

        self.write_log(f"标题：{title}")
        title_final = clean_filename(title)

        # 提取正文内容
        content_div = soup.find(id="js_content")
        if not content_div:
            content_div = soup.find(class_="rich_media_content")
        if not content_div:
            self.write_log("⚠️ 无法定位微信正文区域")
            self.queue.put(("finished", False))
            return

        # 预处理懒加载图片
        imgs = content_div.find_all("img")
        img_count = 0
        for img in imgs:
            if img.get("data-src"):
                img["src"] = img["data-src"]
                del img["data-src"]
                img_count += 1
        
        # 转换为 Markdown
        converter = html2text.HTML2Text()
        converter.ignore_links = False
        converter.ignore_images = False
        converter.ignore_tables = False
        converter.body_width = 0
        md_content = converter.handle(str(content_div))

        # 判断自动分类文件夹和对应标签
        archive_dir, tags = get_target_folder_and_tags(self.obsidian_base_dir, title, md_content)
        if self.obsidian_base_dir:
            folder_basename = os.path.basename(archive_dir)
            self.write_log(f"分类至：{folder_basename}文件夹")
        else:
            self.write_log("ℹ️ 未启用 Obsidian 归档，跳过自动分类匹配")

        # 生成 Front Matter 注入的 markdown
        today_str = datetime.now().strftime("%Y-%m-%d")
        tags_str = "[" + ", ".join(tags) + "]"
        
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
*Created by Wechat Archive GUI System*
"""

        # 保存 Markdown 临时文件
        filename_base = f"[{today_str}] {title_final}"
        filename_md = f"{filename_base}.md"
        filename_docx = f"{filename_base}.docx"
        
        file_path_md = os.path.join(self.word_save_dir, filename_md)
        file_path_docx = os.path.join(self.word_save_dir, filename_docx)
        
        try:
            with open(file_path_md, "w", encoding="utf-8") as f:
                f.write(final_md)
        except Exception as e:
            self.write_log(f"❌ 写入 Markdown 临时文件失败: {e}")
            self.queue.put(("finished", False))
            return

        # 调用排版管线转换 Word
        # 【关键修复】临时文件使用安全的 UUID 路径，避免文章标题中的特殊字符（日文/方括号等）导致 pandoc 路径解析失败
        safe_id = uuid.uuid4().hex[:12]
        tmp_dir = tempfile.gettempdir()
        escaped_md = os.path.join(tmp_dir, f"wcat_{safe_id}_escaped.md")
        temp_docx  = os.path.join(tmp_dir, f"wcat_{safe_id}_temp.docx")
        
        try:
            # 1. 预处理 Markdown
            try:
                import preprocess_md
                preprocess_md.preprocess_markdown(file_path_md, escaped_md)
            except Exception as e:
                self.write_log(f"   [警告] Markdown 预处理出错 (跳过): {e}")
                escaped_md = file_path_md

            # 2. Pandoc 转换并加载黄金模板
            pandoc_bin = get_pandoc_executable()
            try:
                pandoc_cmd = [pandoc_bin, escaped_md, "-o", temp_docx]
                if GOLDEN_TEMPLATE.exists():
                    pandoc_cmd.extend(["--reference-doc", str(GOLDEN_TEMPLATE)])
                subprocess.run(pandoc_cmd, check=True, capture_output=True, text=True)
            except FileNotFoundError:
                self.write_log("❌ 运行内置 pandoc 失败！")
                self.queue.put(("finished", False))
                return
            except subprocess.CalledProcessError as e:
                self.write_log(f"   [警告] Pandoc 应用黄金模板失败，正在回退无模板转换: {e.stderr}")
                subprocess.run([pandoc_bin, escaped_md, "-o", temp_docx], check=True, capture_output=True, text=True)

            # 3. 精确排版规整 (format_expert)
            try:
                import format_expert
                formatter = format_expert.DocumentFormatter(temp_docx, file_path_docx, add_pagenum=True)
                formatter.format()
            except Exception as e:
                self.write_log(f"   [警告] 精确排版工具失败，回退到普通排版 Word: {e}")
                shutil.copy(temp_docx, file_path_docx)

            # 清理临时排版产物
            for f in [escaped_md, temp_docx]:
                if os.path.exists(f) and f != file_path_md:
                    os.remove(f)

            self.write_log("Word 保存成功")

        except Exception as e:
            self.write_log(f"❌ 转换 Docx 出现异常: {e}")
            self.queue.put(("finished", False))
            return

        # 4. 归档同步到 Obsidian 库，清理临时文件
        if self.obsidian_base_dir:
            try:
                os.makedirs(archive_dir, exist_ok=True)
                archive_md_path = os.path.join(archive_dir, filename_md)
                shutil.copy2(file_path_md, archive_md_path)
                
                # 删除 Word 保存路径下的 MD 临时文件
                if os.path.exists(archive_md_path) and os.path.exists(file_path_md):
                    os.remove(file_path_md)
                
                self.write_log("已同步至 Obsidian")
            except Exception as e:
                self.write_log(f"⚠️ 同步归档 Obsidian 失败: {e}")
        else:
            self.write_log("ℹ️ 未同步至 Obsidian (未配置/已跳过 Obsidian 归档)")

        # 更新状态栏并通知完成
        self.queue.put(("status", f"最后保存文件：{filename_docx}"))
        self.queue.put(("finished", True))

if __name__ == "__main__":
    app = WechatGrabberApp()
    app.mainloop()
