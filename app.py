from flask import Flask, request, jsonify, render_template, send_file
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT, WD_LINE_SPACING
from docx.shared import Length
import os
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Border, Side, Alignment, Font

app = Flask(__name__)

# 確保輸出目錄存在
OUTPUT_DIR = 'output_files'
if not os.path.exists(OUTPUT_DIR):
    os.makedirs(OUTPUT_DIR)

# 新增函數：找到符合條件的欄位名稱
def find_matching_column(df, keywords):
    """
    在DataFrame的欄位中尋找包含指定關鍵字的欄位名稱
    keywords 可以是單個字串或字串列表
    返回找到的第一個匹配欄位名稱，如果沒找到返回None
    """
    if isinstance(keywords, str):
        keywords = [keywords]

    keywords = [k.lower() for k in keywords]

    for col in df.columns:
        col_lower = str(col).lower()
        if any(keyword in col_lower for keyword in keywords):
            return col
    return None

def set_document_orientation_and_font(doc, is_landscape=True):
    section = doc.sections[0]
    if is_landscape:
        # 設定為橫式
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width, section.page_height = section.page_height, section.page_width
    else:
        # 設定為直式
        section.orientation = WD_ORIENT.PORTRAIT

    # 設定邊界為 1.27cm
    section.left_margin = Cm(1.27)
    section.right_margin = Cm(1.27)
    section.top_margin = Cm(1.27)
    section.bottom_margin = Cm(1.27)

    # 設定預設字型大小
    style = doc.styles['Normal']
    font = style.font
    font.size = Pt(12)

def set_paragraph_format(paragraph):
    """設置段落格式：無間距，固定行高16pt"""
    paragraph_format = paragraph.paragraph_format
    paragraph_format.space_after = Pt(0)  # 段落後間距為0
    paragraph_format.space_before = Pt(0)  # 段落前間距為0
    paragraph_format.line_spacing = Pt(16)  # 固定行高16pt
    paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY  # 設置為固定行高

def add_empty_lines(doc, count):
    for _ in range(count):
        paragraph = doc.add_paragraph('')
        set_paragraph_format(paragraph)

def estimate_line_count(text, max_chars_per_line):
    """估算文字會佔用幾行"""
    if not text:
        return 1
    return -(-len(text) // max_chars_per_line)  # 向上取整除法

def create_word_files(df):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_paths = {}
    LINES_PER_PAGE = 30
    START_LINE = 20
    MAX_CHARS_PER_LINE = 200

    # 找到對應的欄位名稱，支援多個關鍵字
    xiazai_col = find_matching_column(df, '祈福牌位')
    chaojian_col = find_matching_column(df, ['超薦牌位', '超渡牌位'])  # 支援兩種寫法
    gongde_col = find_matching_column(df, '功德主')

    # 1. 消災牌位 (直式)
    has_xiazai = xiazai_col is not None and df[xiazai_col].notna().any()
    if has_xiazai:
        doc1 = Document()
        set_document_orientation_and_font(doc1, is_landscape=False)
        add_empty_lines(doc1, START_LINE - 1)

        current_line = START_LINE
        for _, row in df.iterrows():
            if pd.notna(row[xiazai_col]):
                content_text = str(row[xiazai_col]).replace('\n', ' ')
                content = f"{row['姓名']}\t{content_text}"

                name_length = len(row['姓名'])
                content_length = len(content_text)
                total_length = name_length + content_length + 1
                estimated_lines = max(1, -(-total_length // MAX_CHARS_PER_LINE))

                if current_line + estimated_lines > LINES_PER_PAGE:
                    doc1.add_page_break()
                    add_empty_lines(doc1, START_LINE - 1)
                    current_line = START_LINE

                paragraph = doc1.add_paragraph(content)
                set_paragraph_format(paragraph)
                current_line += estimated_lines

        file_paths['消災牌位'] = os.path.join(OUTPUT_DIR, f'消災牌位_{timestamp}.docx')
        doc1.save(file_paths['消災牌位'])

    # 2. 超薦牌位 (直式)
    has_chaojian = chaojian_col is not None and df[chaojian_col].notna().any()
    if has_chaojian:
        doc2 = Document()
        set_document_orientation_and_font(doc2, is_landscape=False)
        add_empty_lines(doc2, START_LINE - 1)

        current_line = START_LINE
        for _, row in df.iterrows():
            if pd.notna(row[chaojian_col]):
                content_text = str(row[chaojian_col]).replace('\n', ' ')
                # 使用 " | " 作為分隔符號，並在姓名前加上"陽上："
                content = f"陽上：{row['姓名']} | {content_text}"

                # 計算行數時需要考慮新的格式
                name_length = len(f"陽上：{row['姓名']}")
                content_length = len(content_text)
                separator_length = 3  # " | " 的長度
                total_length = name_length + separator_length + content_length
                estimated_lines = max(1, -(-total_length // MAX_CHARS_PER_LINE))

                if current_line + estimated_lines > LINES_PER_PAGE:
                    doc2.add_page_break()
                    add_empty_lines(doc2, START_LINE - 1)
                    current_line = START_LINE

                paragraph = doc2.add_paragraph(content)
                set_paragraph_format(paragraph)
                current_line += estimated_lines

        file_paths['超薦牌位'] = os.path.join(OUTPUT_DIR, f'超薦牌位_{timestamp}.docx')
        doc2.save(file_paths['超薦牌位'])

    # 3. 功德主 (橫式)
    has_gongde = gongde_col is not None and df[gongde_col].notna().any()
    if has_gongde:
        doc3 = Document()
        set_document_orientation_and_font(doc3, is_landscape=True)
        doc3.add_heading('功德主', 0)

        # 找到管理者註記事項欄位
        note_col = find_matching_column(df, '管理者註記事項')

        # 修改表格欄位，增加管理者註記事項
        table = doc3.add_table(rows=1, cols=5)  # 改為 5 欄
        table.style = 'Table Grid'
        header_cells = table.rows[0].cells
        headers = ['姓名', 'Email', '行動電話', '功德主', '管理者註記事項']  # 新增管理者註記事項

        for i, header in enumerate(headers):
            header_cells[i].text = header

        for _, row in df.iterrows():
            if pd.notna(row[gongde_col]):
                row_cells = table.add_row().cells
                row_cells[0].text = str(row['姓名'])
                row_cells[1].text = str(row['Email'])
                row_cells[2].text = str(row['行動電話'])
                row_cells[3].text = str(row[gongde_col])
                # 添加管理者註記事項
                row_cells[4].text = str(row[note_col]) if note_col and pd.notna(row[note_col]) else ''

        file_paths['功德主'] = os.path.join(OUTPUT_DIR, f'功德主_{timestamp}.docx')
        doc3.save(file_paths['功德主'])

    return file_paths, has_xiazai, has_chaojian, has_gongde

@app.route('/')
def index():
    return render_template('index.html')

def create_participant_excel(df):
    """創建參加者 Excel 檔案"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    # 找到相關欄位
    number_col = find_matching_column(df, '項次')
    name_col = find_matching_column(df, '姓名')
    # 修改：支援多個可能的活動欄位名稱
    activity_col = find_matching_column(df, ['參加項目', '參與課程'])
    note_col = find_matching_column(df, '管理者註記事項')

    if not all([number_col, name_col, activity_col]):
        return None

    # 篩選現場參加者（包含"現場上課"或"到場參加"）
    attendance_mask = df[activity_col].str.contains('現場上課|到場參加', case=False, na=False)
    full_participants = df[attendance_mask].copy()

    if full_participants.empty:
        return None

    # 重設索引並從1開始編號
    full_participants = full_participants.reset_index(drop=True)
    full_participants.index = full_participants.index + 1

    # 創建新的 Excel 工作簿
    wb = Workbook()
    ws = wb.active
    ws.title = "全程參加者名單"

    # 添加標題行
    ws.merge_cells('A1:E1')
    title_cell = ws.cell(row=1, column=1)
    title_cell.value = "到場參加"  # 設定標題文字
    title_cell.font = Font(size=24, name='Microsoft JhengHei')  # 設定字型
    title_cell.alignment = Alignment(horizontal='left', vertical='center')  # 置中對齊


    # 設定欄位標題（第二行）
    headers = ['項次', '姓名', '法本', '便當', '管理者註記事項']
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=2, column=col, value=header)
        cell.font = Font(size=14, name='Microsoft JhengHei')  # 改用微軟正黑體

    # 填入資料（從第三行開始）
    for idx, row in full_participants.iterrows():
        current_row = idx + 2  # 因為第一行是標題，第二行是欄位名稱
        # 填入每個欄位並設定字型
        cells = [
            (1, idx),  # 項次
            (2, row[name_col]),  # 姓名
            (3, ''),  # 法本
            (4, ''),  # 便當
            (5, row[note_col] if note_col and pd.notna(row[note_col]) else '')  # 管理者註記事項
        ]

        for col, value in cells:
            cell = ws.cell(row=current_row, column=col, value=value)
            cell.font = Font(size=14)  # 設定資料行的字型大小為14

    # 設定格式
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # 設定欄寬
    ws.column_dimensions['A'].width = 10  # 項次
    ws.column_dimensions['B'].width = 15  # 姓名
    ws.column_dimensions['C'].width = 15  # 法本
    ws.column_dimensions['D'].width = 15  # 便當
    ws.column_dimensions['E'].width = 30  # 管理者註記事項

    # 設定列高
    ws.row_dimensions[1].height = 40  # 第一行高度
    # 設定第二行開始的列高為20
    for row in range(2, len(full_participants) + 3):  # +3是因為包含標題行和欄位標題行
        ws.row_dimensions[row].height = 20

    # 應用格式到所有使用的單元格
    for row in ws.iter_rows(min_row=1, max_row=len(full_participants)+2,
                          min_col=1, max_col=5):
        for cell in row:
            cell.border = thin_border
            if cell.row > 1:  # 除了第一行的標題外
                cell.alignment = Alignment(horizontal='center', vertical='center')
                # 確保所有單元格都使用微軟正黑體
                if not cell.font or cell.font.name != 'Microsoft JhengHei':
                    current_font = cell.font or Font()
                    cell.font = Font(
                        name='Microsoft JhengHei',
                        size=current_font.size or 14,
                        bold=current_font.bold,
                        italic=current_font.italic
                    )

    # 儲存檔案
    file_path = os.path.join(OUTPUT_DIR, f'全程參加者名單_{timestamp}.xlsx')
    wb.save(file_path)
    return file_path

@app.route('/process_excel', methods=['POST'])
def process_excel():
    try:
        app.logger.info('開始處理上傳檔案')

        # 初始化 response_data
        response_data = {
            'message': '處理完成',
            'files': {}
        }

        # 獲取活動類型
        activity_type = request.form.get('activityType', 'both')
        app.logger.info(f'活動類型: {activity_type}')

        if not os.path.exists(OUTPUT_DIR):
            os.makedirs(OUTPUT_DIR)
            app.logger.info('創建輸出目錄')

        if 'file' not in request.files:
            app.logger.error('沒有檔案')
            return jsonify({'error': '沒有檔案'}), 400

        file = request.files['file']
        if file.filename == '':
            app.logger.error('沒有選擇檔案')
            return jsonify({'error': '沒有選擇檔案'}), 400

        try:
            app.logger.info('開始讀取 Excel 檔案')
            df = pd.read_excel(file, dtype={'行動電話': str})
            if '行動電話' in df.columns:
                df['行動電話'] = df['行動電話'].apply(lambda x: str(x).zfill(10) if pd.notna(x) else '')
            app.logger.info('Excel 檔案讀取完成')
        except Exception as e:
            app.logger.error(f'Excel 讀取失敗: {str(e)}')
            return jsonify({'error': f'Excel 讀取失敗: {str(e)}'}), 500

        # 檢查基本欄位
        required_base_columns = ['姓名', 'Email', '行動電話']
        missing_base_columns = [col for col in required_base_columns if col not in df.columns]
        if missing_base_columns:
            app.logger.error(f'缺少基本欄位：{", ".join(missing_base_columns)}')
            return jsonify({'error': f'缺少基本欄位：{", ".join(missing_base_columns)}'}), 400

        # 根據活動類型檢查必要欄位
        if activity_type == 'both':
            # 檢查是否至少有一個相關欄位（法會相關）
            xiazai_col = find_matching_column(df, '祈福牌位')
            chaojian_col = find_matching_column(df, ['超薦牌位', '超渡牌位'])
            gongde_col = find_matching_column(df, '功德主')

            if not any([xiazai_col, chaojian_col, gongde_col]):
                error_msg = '必須至少包含「祈福牌位」、「超薦牌位（或超渡牌位）」或「功德主」其中一個相關欄位'
                app.logger.error(error_msg)
                return jsonify({'error': error_msg}), 400

            # 檢查是否至少有一個欄位有資料
            has_xiazai_data = xiazai_col is not None and df[xiazai_col].notna().any()
            has_chaojian_data = chaojian_col is not None and df[chaojian_col].notna().any()
            has_gongde_data = gongde_col is not None and df[gongde_col].notna().any()

            if not (has_xiazai_data or has_chaojian_data or has_gongde_data):
                error_msg = '必須至少填寫「祈福牌位」、「超薦牌位」或「功德主」其中一項資料'
                app.logger.error(error_msg)
                return jsonify({'error': error_msg}), 400

            # 創建法會相關文件
            try:
                app.logger.info('開始創建 Word 檔案')
                file_paths, has_xiazai, has_chaojian, has_gongde = create_word_files(df)
                app.logger.info('Word 檔案創建完成')

                if has_xiazai:
                    response_data['files']['xiazai'] = os.path.basename(file_paths['消災牌位'])
                if has_chaojian:
                    response_data['files']['chaojian'] = os.path.basename(file_paths['超薦牌位'])
                if has_gongde:
                    response_data['files']['gongde'] = os.path.basename(file_paths['功德主'])

            except Exception as e:
                app.logger.error(f'Word 檔案創建失敗: {str(e)}')
                return jsonify({'error': f'Word 檔案創建失敗: {str(e)}'}), 500

        # 創建參加者名單（兩種活動類型都需要）
        try:
            participant_excel = create_participant_excel(df)
            if participant_excel:
                response_data['files']['participant'] = os.path.basename(participant_excel)
        except Exception as e:
            app.logger.error(f'參加者名單 Excel 創建失敗: {str(e)}')

        app.logger.info('處理完成，返回結果')
        return jsonify(response_data), 200

    except Exception as e:
        app.logger.error(f'處理過程發生錯誤：{str(e)}')
        return jsonify({'error': str(e)}), 500

@app.route('/download/<filename>')
def download_file(filename):
    try:
        if not os.path.exists(OUTPUT_DIR):
            return jsonify({'error': '輸出目錄不存在'}), 400

        file_path = os.path.join(OUTPUT_DIR, filename)
        if not os.path.exists(file_path):
            return jsonify({'error': '檔案不存在'}), 404

        return send_file(
            file_path,
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        return str(e), 400

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=33080, debug=True)
