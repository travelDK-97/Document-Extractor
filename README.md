# 文档结构化提取工具 (Document Structuring Tool)

这是一个基于纯本地离线环境的自动化文档解析工具。支持批量将 PDF（含扫描件）、DOCX、DOC、WPS 等多种格式的文档，进行结构化段落解析，并统一导出为 Markdown 文件和 Excel 汇总报表。

## ✨ 核心特性

- **🔒 纯本地运行：** 不依赖任何外部 API，保护敏感文档隐私，适合政企合规要求。
- **👁️ 智能 OCR 介入：** 遇到纯图片或扫描版 PDF 时，自动触发本地 RapidOCR 引擎进行文字提取。
- **📑 多格式统一处理：** 自动调用系统原生 Office 组件隐式转换 `.doc` 和 `.wps` 老旧格式。
- **📊 自动化结构提取：** 基于中式公文规范的正则启发式算法，自动识别「一级标题」、「二级标题」等层级，并最终生成 SQLite 数据库与 Excel 数据报表。
- **🖥️ 双模式支持：** 提供开箱即用的 GUI 图形界面和 Pipeline 核心代码。

## ⚠️ 环境依赖与限制 (非常重要)

1. **操作系统：** 本项目依赖 `win32com` 接口处理老旧文档，**仅支持 Windows 系统**。
2. **前置软件：** 运行本程序的电脑**必须已安装 Microsoft Word 或 WPS Office**，否则无法解析 `.doc` 和 `.wps` 文件。
3. **Python 版本：** 推荐使用 Python 3.10 - 3.12（如果是自行打包，请避免使用 3.13 以免触发 tkinter 兼容性 Bug）。

## 🚀 快速开始

### 1. 克隆仓库并安装依赖

\`\`\`bash
git clone https://github.com/travelDK-97/Document-Extractor.git
cd 你的仓库名

# 建议在虚拟环境中运行
pip install -r requirements.txt
\`\`\`

### 2. 运行方式

**方式一：使用图形界面 (GUI)**
直接运行界面脚本，通过可视化的方式选择输入/输出文件夹。
\`\`\`bash
python app_gui.py
\`\`\`

**方式二：代码调用 (Pipeline)**
打开 `main_pipeline.py`，修改文件底部的 `INPUT_FOLDER` 和 `OUTPUT_FOLDER` 等路径，然后直接运行：
\`\`\`bash
python main_pipeline.py
\`\`\`

## 📂 输出物说明

程序运行结束后，在您指定的输出目录会生成：
1. **[文件名]_结构化提取.md：** 每份源文件对应的结构化 Markdown 文本。
2. **解析数据库.db：** 包含所有段落结构化数据的 SQLite 本地数据库。
3. **提取结果汇总报表.xlsx：** 方便非技术人员查阅和筛选的 Excel 汇总表格。

## 🤝 贡献与反馈
如果你在处理特定奇葩格式的文档时遇到 Bug，或者有更好的正则解析优化思路，欢迎提交 Issue 或 Pull Request！