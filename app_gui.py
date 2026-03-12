import os
import re
import time
import sqlite3
import threading
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox
from pathlib import Path

# ----------------- 核心依赖库 -----------------
import fitz
import docx
import numpy as np
from PIL import Image
from rapidocr_onnxruntime import RapidOCR
import win32com.client
import pandas as pd

# ==========================================
# 0. 界面日志重定向 (让 print 输出到 GUI 面板)
# ==========================================
class PrintRedirector:
    def __init__(self, text_widget):
        self.text_widget = text_widget

    def write(self, text):
        self.text_widget.insert(tk.END, text)
        self.text_widget.see(tk.END) # 自动滚动到最底部
        self.text_widget.update()

    def flush(self):
        pass

# ==========================================
# 1. 核心处理逻辑 
# ==========================================
def convert_to_docx(input_path, output_docx_path):
    word_app = None
    try:
        word_app = win32com.client.DispatchEx("Word.Application")
    except Exception:
        try:
            word_app = win32com.client.DispatchEx("KWPS.Application")
        except Exception:
            raise RuntimeError("未检测到 Word 或 WPS")
    word_app.Visible = 0
    word_app.DisplayAlerts = 0
    try:
        doc = word_app.Documents.Open(input_path)
        doc.SaveAs(output_docx_path, 16)
        doc.Close()
    finally:
        word_app.Quit()

def extract_raw_from_docx(file_path):
    doc = docx.Document(file_path)
    return "\n".join([para.text.strip() for para in doc.paragraphs if para.text.strip()])

def extract_from_pdf(file_path, ocr_engine):
    doc = fitz.open(file_path)
    full_text = []
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        page_text = page.get_text("text").strip()
        if len(page_text) < 50:
            zoom_matrix = fitz.Matrix(1.5, 1.5)
            pix = page.get_pixmap(matrix=zoom_matrix)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            ocr_result, _ = ocr_engine(np.array(img))
            if ocr_result:
                page_text = "\n".join([line[1] for line in ocr_result])
        if page_text:
            full_text.append(page_text)
    doc.close()
    return "\n".join(full_text)

def format_text_to_markdown(text):
    if not text: return ""
    md_lines = []
    for line in text.split('\n'):
        line = line.strip()
        if not line: continue
        is_short = len(line) < 40
        if re.match(r'^[一二三四五六七八九十百千万]+[、]', line) and is_short:
            md_lines.append(f"\n# {line}")
        elif re.match(r'^[(（][一二三四五六七八九十百千万]+[)）]', line) and is_short:
            md_lines.append(f"\n## {line}")
        elif re.match(r'^\d+[\.．]', line) and is_short:
            md_lines.append(f"\n### {line}")
        elif re.match(r'^(第[一二三四五六七八九十百千万]+[章节部分篇])', line) and is_short:
            md_lines.append(f"\n# {line}")
        else:
            md_lines.append(line)
    return re.sub(r'\n{3,}', '\n\n', "\n".join(md_lines)).strip()

def extract_sections_for_db(md_text):
    records = []
    current_title = "前言/未分类"
    current_content = []
    for line in md_text.split('\n'):
        if line.startswith('# '):
            if current_content:
                records.append({"title": current_title, "content": "\n".join(current_content).strip()})
            current_title = line[2:].strip()
            current_content = []
        else:
            current_content.append(line)
    if current_content:
        records.append({"title": current_title, "content": "\n".join(current_content).strip()})
    return records

# ==========================================
# 2. 调度引擎 (适配 GUI)
# ==========================================
def run_pipeline(input_dir, output_dir, btn_start):
    try:
        print(">>> 正在初始化本地 OCR AI 引擎，请稍候...")
        ocr_engine = RapidOCR()
        print(">>> 引擎加载完成！开始执行任务...\n" + "-"*40)
        
        input_path = Path(input_dir)
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)
        
        db_path = output_path / "解析数据库.db"
        excel_path = output_path / "提取结果汇总报表.xlsx"
        
        conn = sqlite3.connect(db_path)
        cursor = conn.cursor()
        cursor.execute('''CREATE TABLE IF NOT EXISTS parsed_sections 
                          (id INTEGER PRIMARY KEY AUTOINCREMENT, file_name TEXT, section_title TEXT, content TEXT, process_time REAL)''')
        conn.commit()
        
        success_count, fail_count = 0, 0
        total_start = time.time()
        
        files = [f for f in input_path.rglob("*") if f.is_file() and not f.name.startswith('.') and f.suffix.lower() in ['.pdf', '.docx', '.doc', '.wps']]
        
        for idx, file_path in enumerate(files):
            print(f"\n[{idx+1}/{len(files)}] 正在处理: {file_path.name}")
            file_start = time.time()
            ext = file_path.suffix.lower()
            try:
                raw_text = ""
                if ext == '.docx':
                    raw_text = extract_raw_from_docx(str(file_path))
                elif ext in ['.doc', '.wps']:
                    temp_docx = output_path / f"temp_{file_path.stem}.docx"
                    convert_to_docx(str(file_path.absolute()), str(temp_docx.absolute()))
                    raw_text = extract_raw_from_docx(str(temp_docx))
                    if temp_docx.exists(): os.remove(temp_docx)
                elif ext == '.pdf':
                    raw_text = extract_from_pdf(str(file_path), ocr_engine)
                
                final_md = format_text_to_markdown(raw_text)
                
                with open(output_path / (file_path.stem + "_结构化提取.md"), "w", encoding="utf-8") as f:
                    f.write(final_md)
                    
                db_records = extract_sections_for_db(final_md)
                file_cost = round(time.time() - file_start, 2)
                
                for record in db_records:
                    cursor.execute("INSERT INTO parsed_sections (file_name, section_title, content, process_time) VALUES (?, ?, ?, ?)",
                                   (file_path.name, record['title'], record['content'], file_cost))
                conn.commit()
                
                success_count += 1
                print(f"  √ 完成 | 耗时 {file_cost} 秒")
                
            except Exception as e:
                print(f"  × 失败 | 报错: {str(e)}")
                fail_count += 1
                
        # 导出 Excel
        print("\n" + "-"*40)
        print(">>> 正在生成 Excel 报表...")
        df = pd.read_sql_query("SELECT file_name AS '文件名', section_title AS '一级标题', content AS '具体内容', process_time AS '解析耗时(秒)' FROM parsed_sections", conn)
        df.to_excel(excel_path, index=False, engine='openpyxl')
        conn.close()
        
        total_cost = round(time.time() - total_start, 2)
        print(">>> 任务全部完成！")
        print(f">>> 成功: {success_count} 份 | 失败: {fail_count} 份 | 总耗时: {total_cost} 秒")
        print(f">>> 报表已保存至: {output_path.absolute()}")
        messagebox.showinfo("完成", f"处理完毕！\n成功 {success_count} 份，失败 {fail_count} 份。\n请前往输出目录查看 Excel 报表。")

    except Exception as e:
        print(f"\n严重错误: {str(e)}")
        messagebox.showerror("错误", f"发生严重错误:\n{str(e)}")
    finally:
        btn_start.config(state=tk.NORMAL, text="开始解析")

# ==========================================
# 3. GUI 界面构建
# ==========================================
def select_input(entry_widget):
    folder = filedialog.askdirectory()
    if folder:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, folder)

def select_output(entry_widget):
    folder = filedialog.askdirectory()
    if folder:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, folder)

def start_process(entry_in, entry_out, btn_start):
    in_dir = entry_in.get()
    out_dir = entry_out.get()
    if not in_dir or not out_dir:
        messagebox.showwarning("提示", "请先选择输入和输出文件夹！")
        return
    
    btn_start.config(state=tk.DISABLED, text="正在处理中，请稍候...")
    threading.Thread(target=run_pipeline, args=(in_dir, out_dir, btn_start), daemon=True).start()

def main():
    root = tk.Tk()
    root.title("文档结构化提取工具 (测试版V1.0)")
    root.geometry("680x580") # 稍微加高了窗口以容纳说明区
    
    # 顶部标题
    tk.Label(root, text="文档结构化提取工具 (测试版V1.0)", font=("微软雅黑", 14, "bold"), fg="#2E86C1").pack(pady=(15, 5))
    
    # ---------------------------------------------------------
    # 【新增】使用说明区 (LabelFrame)
    # ---------------------------------------------------------
    help_frame = tk.LabelFrame(root, text=" 📌 快速使用说明 ", font=("微软雅黑", 10, "bold"), fg="#333333", padx=15, pady=10)
    help_frame.pack(fill=tk.X, padx=20, pady=5)
    
    instructions = (
        "1. 准备：将待处理的 PDF、DOCX、DOC、WPS 文件集中放入一个文件夹中。\n"
        "2. 路径：依次点击下方按钮，设置【待处理文档目录】和【结果保存目录】。\n"
        "3. 运行：点击绿色「开始解析」按钮。扫描件会自动触发离线识别，耗时稍长请耐心等待。\n"
        "4. 产出：结束后，输出目录会生成 单独的结构化 Markdown 文件 及 一份完整的 Excel 数据汇总表。\n"
        "（测试版生成数据仅供查阅参考，正式采用信息时需要核查）"
    )
    # justify=tk.LEFT 让多行文本左对齐
    tk.Label(help_frame, text=instructions, font=("微软雅黑", 9), justify=tk.LEFT, anchor="w", fg="#1624EB").pack(fill=tk.X)
    # ---------------------------------------------------------

    # 路径选择区
    frame_path = tk.Frame(root)
    frame_path.pack(pady=10, fill=tk.X, padx=20)
    
    tk.Label(frame_path, text="待处理文档目录:", font=("微软雅黑", 9)).grid(row=0, column=0, sticky=tk.W, pady=5)
    entry_in = tk.Entry(frame_path, width=55)
    entry_in.grid(row=0, column=1, padx=10)
    tk.Button(frame_path, text="浏览...", command=lambda: select_input(entry_in)).grid(row=0, column=2)
    
    tk.Label(frame_path, text="结果保存目录:", font=("微软雅黑", 9)).grid(row=1, column=0, sticky=tk.W, pady=5)
    entry_out = tk.Entry(frame_path, width=55)
    entry_out.grid(row=1, column=1, padx=10)
    tk.Button(frame_path, text="浏览...", command=lambda: select_output(entry_out)).grid(row=1, column=2)
    
    # 启动按钮
    btn_start = tk.Button(root, text="开始结构化解析", font=("微软雅黑", 11, "bold"), bg="#4CAF50", fg="white",
                          command=lambda: start_process(entry_in, entry_out, btn_start))
    btn_start.pack(pady=10, ipadx=30, ipady=5)
    
    # 日志输出区
    tk.Label(root, text="处理日志:", font=("微软雅黑", 9)).pack(anchor=tk.W, padx=20)
    text_log = scrolledtext.ScrolledText(root, width=80, height=12, font=("Consolas", 9), bg="#1E1E1E", fg="#D4D4D4")
    text_log.pack(padx=20, pady=5, fill=tk.BOTH, expand=True)
    
    # 替换系统原生 print
    import sys
    sys.stdout = PrintRedirector(text_log)
    sys.stderr = PrintRedirector(text_log)
    
    root.mainloop()

if __name__ == "__main__":
    main()