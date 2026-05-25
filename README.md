# 微信公众号一键归档与自动化排版系统

> **Wechat Grabber & Auto-Typesetting System**
> 一键复制微信链接，自动抓取文章并转换为排版精美、符合规范的 **Word (.docx)** 文档，并可同步归档为带 Front Matter 的 **Obsidian Markdown** 笔记。

---

## 🌟 功能特点

1. **多重界面支持**：提供现代化暗色主题 GUI 桌面软件（`app.py`）以及轻量命令行终端（`clip_save.py`）。
2. **剪贴板自动识别**：启动时或点击粘贴时，程序自动提取剪贴板中的微信公众号链接。
3. **公文级自动排版**：三步渲染管线（预处理 -> Pandoc 转换 -> 格式精修），自动应用内置的“黄金模板”，输出符合正式排版标准的 Word 文档。
4. **智能分类与归档**：根据文章标题和正文关键词，自动匹配标签并按分类（如：AI技术、HSE工作笔记、中英文术语、施工技术等）同步归档至 Obsidian 库。
5. **Front Matter 注入**：为导出的 Markdown 笔记自动插入标准的头部元数据（Title、Date、Source URL、Tags、Summary 等）。
6. **无损图文保存**：自动清洗并修复懒加载图片，确保 Word 与 Markdown 里的图片能够无损展示，规避防盗链失效问题。
7. **首发向导配置**：零硬编码路径，首次启动程序会自动引导用户配置工作路径，并安全地保存在本地 `config.json` 中。

---

## 📂 目录结构

* `app.py`：桌面 GUI 应用程序主脚本。
* `clip_save.py`：命令行 CLI 应用程序主脚本。
* `preprocess_md.py`：Markdown 段落与数字序号预处理脚本（管线第一步）。
* `format_expert.py`：Word 样式与层级精确精修脚本（管线第三步）。
* `黄金模板.docx`：Pandoc 排版参考文档（样式基础）。
* `requirements.txt`：Python 项目依赖列表。
* `.env.example`：环境变量配置文件模板。
* `.gitignore`：Git 忽略清单。

---

## 🛠️ 安装与使用方法

项目支持 **直接运行源码** 或 **编译打包为 `.exe` 独立运行**。

### 方式一：直接运行源码

#### 1. 前置依赖准备
* **Python 3.10+**：请确保系统已安装 Python 并在环境变量中。
* **Pandoc**：
  * Windows (PowerShell/CMD): `winget install jgm.pandoc` 或使用 `choco install pandoc`
  * macOS: `brew install pandoc`
  * Linux: `sudo apt install pandoc`

#### 2. 安装 Python 依赖
进入项目根目录，运行以下命令（推荐在虚拟环境中操作）：
```bash
# 创建并激活虚拟环境 (可选)
python -m venv venv
source venv/bin/activate  # Linux/macOS
venv\Scripts\activate     # Windows

# 安装库依赖
pip install -r requirements.txt
```

#### 3. 运行程序
```bash
# 启动 GUI 桌面端
python app.py

# 启动 CLI 命令行端
python clip_save.py
```

---

### 方式二：打包编译为独立运行的 `.exe` (Windows)

如果您想将该工具分享给没有 Python 环境的同事，可以将其打包为 `.exe` 可执行程序：

1. 在开发环境中安装 PyInstaller：
   ```bash
   pip install pyinstaller
   ```
2. 执行打包命令：
   ```bash
   pyinstaller --noconsole --onefile --add-data "黄金模板.docx;." --add-data "preprocess_md.py;." --add-data "format_expert.py;." --name="微信公众号一键归档系统" app.py
   ```
3. 打包完成后，在生成的 `dist/` 文件夹下即可找到 `微信公众号一键归档系统.exe`。只需将其与 `黄金模板.docx`、`preprocess_md.py`、`format_expert.py` 放在同一目录下即可分发使用。

---

## 🚀 详细使用步骤

1. **复制文章链接**：在微信上打开任意公众号文章，点击右上角 “复制链接” 或在空白处右键选择 “复制地址”。
2. **运行程序**：双击启动 GUI 程序（或批处理脚本）。
3. **初次配置（仅首次启动）**：
   * 程序将弹窗提示您选择 **Word 文档保存文件夹**（必填，生成的 Word 文件存放处）。
   * 接着提示您选择 **Obsidian 库主文件夹**（可选，不选择或取消将跳过 Markdown 自动同步归档，Markdown 将默认保留在 Word 同级目录下）。
4. **一键保存**：程序启动后会自动读取您刚才复制的微信链接并自动粘贴在输入框内。确认无误后点击居中瞩目的 **“一键保存”** 按钮。
5. **归档完成**：日志框将实时显示抓取、转换及分类状态。完成后会弹窗提示成功，即可在目标目录查看排版完美的 `.docx` 和 `.md` 文件。

---

## 📝 自动分类匹配规则

抓取系统会通过扫描文章标题与正文的前 200 字，自动进行匹配并分类存入以下子文件夹中：
* **`1.AI技术汇编`** (标签 `[AI技术]`)：包含 AI、人工智能、LLM、大模型、Prompt 提示词等。
* **`2.HSE工作笔记库`** (标签 `[HSE]`/`[HSE, 吊装]`)：包含 HSE、安全、吊装、设备检查、事故通报等。
* **`3.中英文术语库`** (标签 `[中英文术语]`)：包含 中英、对照、英译、翻译、Glossary 术语表等。
* **`4.施工技术汇编`** (标签 `[施工技术]`)：包含 施工技术、基坑、混凝土、钢结构、工程管理等。
* **`公众号归档`** (标签 `[待分类]`，兜底)：若未匹配到任何关键词，将存入此兜底文件夹下。

---

## ⚠️ 注意事项

1. **Pandoc 环境变量**：请确保系统已安装 Pandoc 且 `pandoc` 命令在系统环境变量 `PATH` 中可被全局调用，否则无法转换为 Word 格式。
2. **黄金模板依赖**：程序同级目录下必须有 `黄金模板.docx`，否则转换出的 Word 排版样式将回退为系统默认的无主题状态。
3. **网络状况**：微信文章的图片下载和转换需保持网络连接畅通。
