import os
import re
import time
import sqlite3
from pathlib import Path
import fitz  # PyMuPDF
import docx
import numpy as np
from PIL import Image
from rapidocr_onnxruntime import RapidOCR
import win32com.client
import pandas as pd

# ==========================================
# 1. 初始化引擎与数据库
# ==========================================
print("正在加载本地 OCR 模型 (RapidOCR)...")
ocr_engine = RapidOCR()
print("模型加载完成！\n" + "="*40)

def init_db(db_path):
    """初始化 SQLite 数据库和表结构"""
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    # 创建一个名为 parsed_sections 的表
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS parsed_sections (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            file_name TEXT,
            section_title TEXT,
            content TEXT,
            process_time REAL
        )
    ''')
    conn.commit()
    return conn

# ==========================================
# 2. 格式转换与文本提取
# ==========================================
def convert_to_docx(input_path, output_docx_path):
    """调用本机 Office 将老格式隐式转为 .docx"""
    word_app = None
    try:
        word_app = win32com.client.DispatchEx("Word.Application")
    except Exception:
        try:
            word_app = win32com.client.DispatchEx("KWPS.Application")
        except Exception:
            raise RuntimeError("未检测到可调用的 Word 或 WPS 软件")

    word_app.Visible = 0
    word_app.DisplayAlerts = 0
    try:
        doc = word_app.Documents.Open(input_path)
        doc.SaveAs(output_docx_path, 16)
        doc.Close()
    except Exception as e:
        raise RuntimeError(f"Office 转换失败: {e}")
    finally:
        word_app.Quit()

def extract_raw_from_docx(file_path):
    """提取 DOCX 纯文本，不再依赖不靠谱的 Word 官方样式"""
    doc = docx.Document(file_path)
    paragraphs = [para.text.strip() for para in doc.paragraphs if para.text.strip()]
    return "\n".join(paragraphs)

def extract_from_pdf(file_path):
    """提取 PDF 文本与 OCR"""
    doc = fitz.open(file_path)
    full_text = []
    total_pages = len(doc)
    
    for page_num in range(total_pages):
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

# ==========================================
# 3. 结构化引擎与解析器
# ==========================================
def format_text_to_markdown(text):
    """基于中式公文规范的正则启发式结构化"""
    if not text: return ""
    lines = text.split('\n')
    md_lines = []
    
    for line in lines:
        line = line.strip()
        if not line: continue
        is_short_enough = len(line) < 40
            
        if re.match(r'^[一二三四五六七八九十百千万]+[、]', line) and is_short_enough:
            md_lines.append(f"\n# {line}")
        elif re.match(r'^[(（][一二三四五六七八九十百千万]+[)）]', line) and is_short_enough:
            md_lines.append(f"\n## {line}")
        elif re.match(r'^\d+[\.．]', line) and is_short_enough:
            md_lines.append(f"\n### {line}")
        elif re.match(r'^(第[一二三四五六七八九十百千万]+[章节部分篇])', line) and is_short_enough:
            md_lines.append(f"\n# {line}")
        else:
            md_lines.append(line)
            
    raw_md = "\n".join(md_lines)
    cleaned_md = re.sub(r'\n{3,}', '\n\n', raw_md)
    return cleaned_md.strip()

def extract_sections_for_db(md_text):
    """将 Markdown 解析为一级标题及其内容的字典列表"""
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
# 4. 导出为 Excel
# ==========================================
def export_db_to_excel(db_path, excel_path):
    """从 SQLite 读取所有数据并导出为漂亮的 Excel"""
    try:
        conn = sqlite3.connect(db_path)
        # 使用 pandas 直接读取 SQL 为 DataFrame
        query = "SELECT file_name AS '文件名', section_title AS '一级标题', content AS '具体内容', process_time AS '解析耗时(秒)' FROM parsed_sections"
        df = pd.read_sql_query(query, conn)
        
        # 导出为 Excel
        df.to_excel(excel_path, index=False, engine='openpyxl')
        print(f"    √ 成功导出 Excel 报表至: {excel_path}")
    except Exception as e:
        print(f"    × 导出 Excel 失败: {e}")
    finally:
        conn.close()

# ==========================================
# 5. 主控制流水线 (包含计时逻辑)
# ==========================================
def batch_process_pipeline(input_dir, output_dir, db_path, excel_path):
    # 记录总任务开始时间
    total_start_time = time.time()
    
    input_path = Path(input_dir)
    output_path = Path(output_dir)
    
    if not input_path.exists():
        input_path.mkdir(parents=True)
        print(f"提示：未找到输入文件夹，已自动创建 '{input_dir}'，请放入文档后重试。")
        return

    output_path.mkdir(parents=True, exist_ok=True)
    
    # 初始化数据库连接
    db_conn = init_db(db_path)
    db_cursor = db_conn.cursor()
    
    success_count, fail_count = 0, 0
    
    for file_path in input_path.rglob("*"):
        if not file_path.is_file() or file_path.name.startswith('.') or file_path.suffix.lower() not in ['.pdf', '.docx', '.doc', '.wps']:
            continue
            
        ext = file_path.suffix.lower()
        print(f"\n[{success_count + fail_count + 1}] 开始处理: {file_path.name}")
        
        # 记录单文件开始时间
        file_start_time = time.time()
        
        try:
            # 1. 众生平等：全量提取纯文本
            raw_text = ""
            if ext == '.docx':
                raw_text = extract_raw_from_docx(str(file_path))
                
            elif ext in ['.doc', '.wps']:
                temp_docx = output_path / f"temp_{file_path.stem}.docx"
                convert_to_docx(str(file_path.absolute()), str(temp_docx.absolute()))
                raw_text = extract_raw_from_docx(str(temp_docx))
                if temp_docx.exists(): os.remove(temp_docx)
                
            elif ext == '.pdf':
                raw_text = extract_from_pdf(str(file_path))
            
            # 2. 核心清洗：不管是 Word 还是 PDF，统一交由正则引擎去判定层级
            final_md = format_text_to_markdown(raw_text)
            
            # 3. 独立保留 MD 文件
            md_out_file = output_path / (file_path.stem + "_结构化提取.md")
            with open(md_out_file, "w", encoding="utf-8") as f:
                f.write(final_md)
                
            # 4. 提取结构化列表并入库
            db_records = extract_sections_for_db(final_md)
            
            # 计算该文件的处理耗时
            file_cost_time = round(time.time() - file_start_time, 2)
            
            for record in db_records:
                db_cursor.execute(
                    "INSERT INTO parsed_sections (file_name, section_title, content, process_time) VALUES (?, ?, ?, ?)",
                    (file_path.name, record['title'], record['content'], file_cost_time)
                )
            db_conn.commit()
            
            success_count += 1
            print(f"    √ 处理完毕 | 入库段落: {len(db_records)} 个 | 耗时: {file_cost_time} 秒")
            
        except Exception as e:
            file_cost_time = round(time.time() - file_start_time, 2)
            print(f"    × [错误] 处理失败 ({file_cost_time} 秒): {e}")
            fail_count += 1

    # 关闭数据库
    db_conn.close()
    
    # 导出 Excel
    print("\n" + "="*40)
    print("正在生成数据报表...")
    export_db_to_excel(db_path, excel_path)

    # 打印总耗时
    total_cost_time = round(time.time() - total_start_time, 2)
    print("="*40)
    print(f"批量任务结束！")
    print(f"成功: {success_count} 份, 失败: {fail_count} 份。")
    print(f"总计耗时: {total_cost_time} 秒")
    print(f"Markdown文件保存在: {output_path.absolute()}")
    print("="*40)

# ==========================================
# 6. 执行区 (配置相对路径)
# ==========================================
if __name__ == "__main__":
    # 配置你的本地路径
    INPUT_FOLDER  = "待提取的文件夹"            # 把你的 PDF/Word 放这里
    OUTPUT_FOLDER = "提取结果存放"  # 生成的单篇 md 会保存在这里
    DB_FILE       = "documents_data.db"    # 本地 SQLite 数据库文件
    EXCEL_FILE    = "提取结果汇总.xlsx" # 最终导出的 Excel 汇总文件
    
    batch_process_pipeline(INPUT_FOLDER, OUTPUT_FOLDER, DB_FILE, EXCEL_FILE)