import streamlit as st
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
import io
import re

# --- 全角/半角 変換用テーブル ---
HALF2FULL = str.maketrans(
    '0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ',
    '０１２３４５６７８９ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ'
)
FULL2HALF = str.maketrans(
    '０１２３４５６７８９ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ',
    '0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ'
)
BORDER_CHARS = ['┌', '└', '│', '─', '├', '┤', '┬', '┴', '┼']

def convert_char_width(text, mode):
    if mode == "半角に統一（推奨）":
        return text.translate(FULL2HALF)
    elif mode == "全角に統一":
        return text.translate(HALF2FULL)
    return text

def set_font(run, font_name, east_asia_font_name, size_pt=None, bold=False):
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), east_asia_font_name)
    run.font.color.rgb = RGBColor(0, 0, 0)
    if size_pt:
        run.font.size = Pt(size_pt)
    if bold:
        run.bold = True

def add_toc(doc):
    paragraph = doc.add_paragraph()
    run = paragraph.add_run("目　次")
    set_font(run, 'Arial', 'MS ゴシック', size_pt=14, bold=True)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    p = doc.add_paragraph()
    run = p.add_run()
    fldChar = OxmlElement('w:fldChar')
    fldChar.set(qn('w:fldCharType'), 'begin')
    run._element.append(fldChar)
    
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = 'TOC \\o "1-3" \\h \\z \\u'
    run._element.append(instrText)
    
    fldChar = OxmlElement('w:fldChar')
    fldChar.set(qn('w:fldCharType'), 'end')
    run._element.append(fldChar)
    doc.add_page_break()

def set_update_fields(doc):
    try:
        settings = doc.settings.element
        updateFields = OxmlElement('w:updateFields')
        updateFields.set(qn('w:val'), 'true')
        settings.append(updateFields)
    except Exception:
        pass

def is_csv_block(lines):
    if not lines: return False
    if any(any(c in line for c in BORDER_CHARS) for line in lines):
        return False
    return any(',' in line for line in lines if line.strip())

def create_word_doc(text, selected_font, width_mode, add_body_space, indent_normal, art_indent_val, indent_kou, hanging_kou, indent_gou, hanging_gou):
    doc = docx.Document()
    set_update_fields(doc)
    
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(10.5)
    style._element.rPr.rFonts.set(qn('w:eastAsia'), selected_font)
    style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    style.paragraph_format.space_after = Pt(0)

    toc_added = False
    art_buffer = []
    csv_buffer = []
    block_buffer = []
    in_code_block = False
    in_csv_table = False
    
    def flush_art(lines_to_flush):
        if not lines_to_flush: return
        
        # 【大復活】四角囲み（枠付きテーブル）の中に格納し、文字の折り返しを防ぐ
        table = doc.add_table(rows=1, cols=1)
        table.style = 'Table Grid'
        
        # 表（四角い箱）そのものを指定された幅だけ字下げする
        tbl_pr = table._element.xpath('w:tblPr')
        if tbl_pr:
            tbl_ind = OxmlElement('w:tblInd')
            indent_twips = int(10.5 * art_indent_val * 20) # 1em = 10.5pt = 210 twips
            tbl_ind.set(qn('w:w'), str(indent_twips))
            tbl_ind.set(qn('w:type'), 'dxa')
            tbl_pr[0].append(tbl_ind)
            
        cell = table.cell(0, 0)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        p.paragraph_format.line_spacing = Pt(11)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.snap_to_grid = False
        p.paragraph_format.left_indent = Pt(0) # 箱の中身は字下げしない（箱ごと動いているため）
        
        for i, line in enumerate(lines_to_flush):
            if i > 0: p.add_run('\n')
            # 【ご要望反映】図表内のテキストも全角/半角の変換対象にする
            converted_line = convert_char_width(line, width_mode)
            run = p.add_run(converted_line)
            # 空白が無視されないようにXMLレベルで保護
            for t in run._element.findall('.//w:t', namespaces=run._element.nsmap):
                t.set(qn('xml:space'), 'preserve')
            # 箱からはみ出さないようフォントを8.5ptに設定
            set_font(run, 'MS Gothic', 'ＭＳ ゴシック', size_pt=8.5)
            
        lines_to_flush.clear()

    def flush_csv(lines_to_flush):
        if not lines_to_flush: return
        valid_lines = [l for l in lines_to_flush if l.strip()]
        if not valid_lines: 
            lines_to_flush.clear()
            return
            
        max_cols = max(len(row.split(',')) for row in valid_lines)
        if max_cols <= 1:
            for row in valid_lines:
                p = doc.add_paragraph()
                p.paragraph_format.left_indent = Pt(10.5 * indent_normal)
                run = p.add_run(convert_char_width(row.strip(), width_mode))
                set_font(run, 'Arial', selected_font, size_pt=10.5)
            lines_to_flush.clear()
            return

        table = doc.add_table(rows=len(valid_lines), cols=max_cols)
        table.style = 'Table Grid'
        table.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r_idx, row_text in enumerate(valid_lines):
            cells_data = row_text.split(',')
            for c_idx, cell_data in enumerate(cells_data):
                if c_idx < len(table.columns):
                    cell = table.cell(r_idx, c_idx)
                    cell.text = cell_data.strip()
                    for p in cell.paragraphs:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        for run in p.runs:
                            run.text = convert_char_width(run.text, width_mode)
                            set_font(run, 'Arial', selected_font, size_pt=9)
        lines_to_flush.clear()

    def flush_code_block():
        if not block_buffer: return
        if is_csv_block(block_buffer):
            flush_csv(block_buffer)
        else:
            flush_art(block_buffer)

    for i, line in enumerate(text.split('\n')):
        raw_line = line.rstrip('\r\n')
        raw_line = re.sub(r'\[ \t]?', '', raw_line)
        raw_line = re.sub(r'\[\d+\][ \t]?', '', raw_line)
        
        line_strip = raw_line.strip()
        
        is_heading = bool(re.match(r'^第[0-9１２３４５６７８９０一二三四五六七八九十百]+[章条]', line_strip)) or line_strip in ["就　業　規　則", "就業規則", "賃　金　規　程", "賃金規程"]
        
        if '```' in line_strip or '｀｀｀' in line_strip:
            if in_code_block:
                flush_code_block()
                in_code_block = False
            else:
                if csv_buffer: flush_csv(csv_buffer); in_csv_table = False
                if art_buffer: flush_art(art_buffer)
                in_code_block = True
            continue

        if in_code_block:
            if is_heading:
                flush_code_block()
                in_code_block = False
                pass 
            else:
                if "The following table:" in line_strip: continue
                if line_strip.lower() in ['text', 'csv', 'markdown']: continue
                block_buffer.append(raw_line)
                continue

        if not line_strip:
            if csv_buffer: flush_csv(csv_buffer); in_csv_table = False
            if art_buffer: flush_art(art_buffer)
            continue

        if "The following table:" in line_strip:
            if art_buffer: flush_art(art_buffer)
            in_csv_table = True
            continue
            
        if in_csv_table:
            if is_heading:
                flush_csv(csv_buffer)
                in_csv_table = False
            else:
                csv_buffer.append(raw_line)
                continue

        if any(c in raw_line for c in BORDER_CHARS):
            art_buffer.append(raw_line)
            continue
        else:
            if art_buffer: flush_art(art_buffer)

        line_strip = convert_char_width(line_strip, width_mode)

        if line_strip in ["就　業　規　則", "就業規則", "賃　金　規　程", "賃金規程", "育児・介護休業規程", "退職金規程"]:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(24)
            run = p.add_run(line_strip)
            set_font(run, 'Arial', 'MS ゴシック', size_pt=16, bold=True)
            if not toc_added:
                add_toc(doc)
                toc_added = True

        elif re.match(r'^第[0-9１２３４５６７８９０一二三四五六七八九十百]+章', line_strip):
            if not toc_added:
                add_toc(doc)
                toc_added = True
            p = doc.add_paragraph(style='Heading 1')
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(18)
            p.paragraph_format.space_after = Pt(12)
            run = p.add_run(line_strip)
            set_font(run, 'Arial', 'MS ゴシック', size_pt=13, bold=True)
            
        elif re.match(r'^第[0-9１２３４５６７８９０一二三四五六七八九十百]+条', line_strip):
            if not toc_added:
                add_toc(doc)
                toc_added = True
            p = doc.add_paragraph(style='Heading 2')
            p.paragraph_format.space_before = Pt(6)
            run = p.add_run(line_strip)
            set_font(run, 'Arial', 'MS ゴシック', size_pt=10.5, bold=True)
        
        elif re.match(r'^([０-９0-9]+[．\.]|\([０-９0-9]+\))', line_strip):
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Pt(10.5 * (indent_kou + hanging_kou))
            p.paragraph_format.first_line_indent = Pt(-10.5 * hanging_kou)
            run = p.add_run(line_strip)
            set_font(run, 'Arial', selected_font, size_pt=10.5)

        elif re.match(r'^[①-⑳]', line_strip):
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Pt(10.5 * (indent_gou + hanging_gou))
            p.paragraph_format.first_line_indent = Pt(-10.5 * hanging_gou)
            run = p.add_run(line_strip)
            set_font(run, 'Arial', selected_font, size_pt=10.5)

        else:
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Pt(10.5 * indent_normal)
            
            output_text = line_strip
            if add_body_space and output_text:
                output_text = '　' + output_text
                
            run = p.add_run(output_text)
            set_font(run, 'Arial', selected_font, size_pt=10.5)
            
    if block_buffer: flush_code_block()
    if art_buffer: flush_art(art_buffer)
    if csv_buffer: flush_csv(csv_buffer)
        
    if not toc_added:
        add_toc(doc)
        
    return doc

def get_html_table(csv_lines, width_mode):
    csv_lines = [l for l in csv_lines if l.strip()]
    if not csv_lines: return ""
    max_cols = max(len(row.split(',')) for row in csv_lines)
    if max_cols <= 1:
        return "".join([f"<div style='margin-left: 1.0em;'>{convert_char_width(row.strip(), width_mode)}</div>" for row in csv_lines])

    html = "<table style='width:100%; border-collapse: collapse; margin: 10px 0; font-size: 12px;'>"
    for row in csv_lines:
        html += "<tr>"
        for cell in row.split(','):
            html += f"<td style='border: 1px solid #999; padding: 4px 8px; text-align: center;'>{convert_char_width(cell.strip(), width_mode)}</td>"
        html += "</tr>"
    html += "</table>"
    return html

def get_html_art(art_lines, indent_art, width_mode):
    if not art_lines: return ""
    converted_lines = [convert_char_width(line, width_mode) for line in art_lines]
    html = f"<div style='background-color: #fff3cd; color: #856404; padding: 4px 8px; font-size: 11px; margin-top: 10px; border-radius: 3px; margin-left: {indent_art}em; width: fit-content;'>"
    html += "⚠️ 本ツールでは可能な限り自動整形を行っておりますが、出力後はユーザー側で適宜微調整をお願いいたします。</div>"
    # 【変更】プレビューでも「四角囲み（枠線）」が見えるようにCSSを設定
    html += f"<pre style='border: 1px solid #999; background-color: #fff; padding: 10px; font-family: \"MS Gothic\", \"ＭＳ ゴシック\", monospace; font-size: 13px; line-height: 1.2; margin: 0 0 10px; margin-left: {indent_art}em; white-space: pre; overflow-x: auto;'>{chr(10).join(converted_lines)}</pre>"
    return html

# --- スクショに基づいた完全な初期値設定 ---
def reset_settings():
    st.session_state.font_choice = "Arial"
    st.session_state.width_mode_choice = "半角に統一（推奨）"
    st.session_state.add_body_space_choice = False  # スクショ通りOFF
    st.session_state.in_normal = 2.5                # 番号なし本文：2.50
    st.session_state.ik = 1.0                       # 項の字下げ：1.00
    st.session_state.hk = 1.5                       # 項の突き出し：1.50
    st.session_state.ig = 2.5                       # 号の字下げ：2.50
    st.session_state.hg = 1.0                       # 号の突き出し：1.00
    st.session_state.art_type = "本文（番号なし）の開始位置に合わせる"

def main():
    st.set_page_config(page_title="就業規則 Word変換", layout="wide")
    st.title("📄 就業規則プロフェッショナル整形ツール")
    
    if "initialized" not in st.session_state:
        reset_settings()
        st.session_state.initialized = True

    st.sidebar.header("⚙️ 書式・インデント設定")
    
    st.sidebar.button("🔄 推奨設定（初期値）に戻す", on_click=reset_settings, use_container_width=True)
    st.sidebar.markdown("---")

    selected_font = st.sidebar.selectbox("基本フォント", options=["MS 明朝", "MS ゴシック", "Meiryo", "Yu Mincho", "Yu Gothic", "Arial"], key="font_choice")
    width_mode = st.sidebar.radio("英数字の表記統一", options=["半角に統一（推奨）", "全角に統一", "変換しない（元のまま）"], key="width_mode_choice")
    
    st.sidebar.markdown("---")
    st.sidebar.write("**本文（番号なし）の設定**")
    add_body_space = st.sidebar.checkbox("本文の先頭に全角スペース(字下げ)を自動挿入", key="add_body_space_choice")
    indent_normal = st.sidebar.slider("「番号なし本文」の全体字下げ", min_value=0.0, max_value=5.0, step=0.5, key="in_normal")
    
    st.sidebar.markdown("---")
    st.sidebar.write("**項・号の設定**")
    indent_kou = st.sidebar.slider("項（１．等）の字下げ位置", min_value=0.0, max_value=5.0, step=0.5, key="ik")
    hanging_kou = st.sidebar.slider("項の突き出し幅（ぶら下げ）", min_value=0.0, max_value=5.0, step=0.5, key="hk")
    
    indent_gou = st.sidebar.slider("号（① 等）の字下げ位置", min_value=0.0, max_value=5.0, step=0.5, key="ig")
    hanging_gou = st.sidebar.slider("号の突き出し幅（ぶら下げ）", min_value=0.0, max_value=5.0, step=0.5, key="hg")

    st.sidebar.markdown("---")
    st.sidebar.write("**表（ツリー図）の配置設定**")
    art_indent_type = st.sidebar.radio(
        "表の左端をどこに合わせますか？",
        options=[
            "左端（0字下げ）に合わせる", 
            "本文（番号なし）の開始位置に合わせる", 
            "項（１．等）の番号位置に合わせる", 
            "号（① 等）の番号位置に合わせる"
        ],
        key="art_type"
    )

    calculated_art_indent = 0.0
    if art_indent_type == "本文（番号なし）の開始位置に合わせる":
        calculated_art_indent = indent_normal
    elif art_indent_type == "項（１．等）の番号位置に合わせる":
        calculated_art_indent = indent_kou
    elif art_indent_type == "号（① 等）の番号位置に合わせる":
        calculated_art_indent = indent_gou

    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.write("▼ ここにテキストを貼り付け")
        text_input = st.text_area("就業規則テキスト", height=600, label_visibility="collapsed")
    
    with col2:
        st.write("▼ インデント簡易プレビュー（全行表示）")
        if text_input:
            preview_html = f"<div style='font-family: \"{selected_font}\", sans-serif; font-size: 14px; line-height: 1.6; border: 1px solid #ddd; padding: 10px; height: 600px; overflow-y: auto;'>"
            
            art_buffer = []
            csv_buffer = []
            block_buffer = []
            in_code_block = False
            in_csv_table = False
            
            for line in text_input.split('\n'):
                raw_line = line.rstrip('\r\n')
                raw_line = re.sub(r'\[ \t]?', '', raw_line)
                raw_line = re.sub(r'\[\d+\][ \t]?', '', raw_line)
                line_strip = raw_line.strip()
                
                is_heading = bool(re.match(r'^第[0-9１２３４５６７８９０一二三四五六七八九十百]+[章条]', line_strip)) or line_strip in ["就　業　規　則", "就業規則", "賃　金　規　程", "賃金規程", "育児・介護休業規程", "退職金規程"]
                
                if '```' in line_strip or '｀｀｀' in line_strip:
                    if in_code_block:
                        if is_csv_block(block_buffer): preview_html += get_html_table(block_buffer, width_mode)
                        else: preview_html += get_html_art(block_buffer, calculated_art_indent, width_mode)
                        block_buffer.clear()
                        in_code_block = False
                    else:
                        if csv_buffer: preview_html += get_html_table(csv_buffer, width_mode); csv_buffer.clear(); in_csv_table = False
                        if art_buffer: preview_html += get_html_art(art_buffer, calculated_art_indent, width_mode); art_buffer.clear()
                        in_code_block = True
                    continue

                if in_code_block:
                    if is_heading:
                        if is_csv_block(block_buffer): preview_html += get_html_table(block_buffer, width_mode)
                        else: preview_html += get_html_art(block_buffer, calculated_art_indent, width_mode)
                        block_buffer.clear()
                        in_code_block = False
                    else:
                        if "The following table:" in line_strip: continue
                        if line_strip.lower() in ['text', 'csv', 'markdown']: continue
                        block_buffer.append(raw_line)
                        continue

                if not line_strip:
                    if csv_buffer: preview_html += get_html_table(csv_buffer, width_mode); csv_buffer.clear(); in_csv_table = False
                    if art_buffer: preview_html += get_html_art(art_buffer, calculated_art_indent, width_mode); art_buffer.clear()
                    continue

                if "The following table:" in line_strip:
                    if art_buffer: preview_html += get_html_art(art_buffer, calculated_art_indent, width_mode); art_buffer.clear()
                    in_csv_table = True
                    continue
                    
                if in_csv_table:
                    if is_heading:
                        preview_html += get_html_table(csv_buffer, width_mode); csv_buffer.clear(); in_csv_table = False
                    else:
                        csv_buffer.append(raw_line)
                        continue

                if any(c in raw_line for c in BORDER_CHARS):
                    art_buffer.append(raw_line)
                    continue
                else:
                    if art_buffer: preview_html += get_html_art(art_buffer, calculated_art_indent, width_mode); art_buffer.clear()
                
                line_strip = convert_char_width(line_strip, width_mode)
                
                if line_strip in ["就　業　規　則", "就業規則", "賃　金　規　程", "賃金規程", "育児・介護休業規程", "退職金規程"]:
                    preview_html += f"<div style='text-align: center; font-weight: bold; margin-top: 15px; font-size: 1.2em;'>{line_strip}</div>"
                elif re.match(r'^第.*章', line_strip):
                    preview_html += f"<div style='text-align: center; font-weight: bold; margin-top: 15px;'>{line_strip}</div>"
                elif re.match(r'^第.*条', line_strip):
                    preview_html += f"<div style='font-weight: bold; margin-top: 10px;'>{line_strip}</div>"
                elif re.match(r'^([０-９0-9]+[．\.]|\([０-９0-9]+\))', line_strip):
                    preview_html += f"<div style='margin-left: {indent_kou}em; padding-left: {hanging_kou}em; text-indent: -{hanging_kou}em;'>{line_strip}</div>"
                elif re.match(r'^[①-⑳]', line_strip):
                    preview_html += f"<div style='margin-left: {indent_gou}em; padding-left: {hanging_gou}em; text-indent: -{hanging_gou}em;'>{line_strip}</div>"
                else:
                    output_text = line_strip
                    if add_body_space and output_text:
                        output_text = '　' + output_text
                    preview_html += f"<div style='margin-left: {indent_normal}em;'>{output_text}</div>"
            
            if block_buffer:
                if is_csv_block(block_buffer): preview_html += get_html_table(block_buffer, width_mode)
                else: preview_html += get_html_art(block_buffer, calculated_art_indent, width_mode)
            if csv_buffer: preview_html += get_html_table(csv_buffer, width_mode)
            if art_buffer: preview_html += get_html_art(art_buffer, calculated_art_indent, width_mode)
                
            preview_html += "</div>"
            st.markdown(preview_html, unsafe_allow_html=True)
        else:
            st.info("テキストを入力すると、ここにインデントのプレビューが表示されます。")

    if text_input:
        st.markdown("---")
        doc = create_word_doc(text_input, selected_font, width_mode, add_body_space, indent_normal, calculated_art_indent, indent_kou, hanging_kou, indent_gou, hanging_gou)
        bio = io.BytesIO()
        doc.save(bio)
        
        st.download_button(
            label="📥 設定を反映してWordファイルをダウンロード",
            data=bio.getvalue(),
            file_name="整形済み就業規則.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

if __name__ == "__main__":
    main()