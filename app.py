import streamlit as st
import docx
from docx.shared import Pt, Inches, RGBColor
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
import io
import re
import pandas as pd

# ==========================================================
# --- 全角/半角 変換用テーブル（英数字＋すべての記号を網羅） ---
# ==========================================================
HALF_CHARS = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ!\"#$%&'()*+,-./:;<=>?@[\\]^_`{|}~¥"
FULL_CHARS = "０１２３４５６７８９ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ！＂＃＄％＆’（）＊＋，－．／：；＜＝＞？＠［＼］＾＿｀｛｜｝～￥"

HALF2FULL = str.maketrans(HALF_CHARS, FULL_CHARS)
FULL2HALF = str.maketrans(FULL_CHARS, HALF_CHARS)

BORDER_CHARS = ['┌', '└', '│', '─', '├', '┤', '┬', '┴', '┼', '━', '┃', '┏', '┓', '┗', '┛', '┣', '┫', '┳', '┻', '╋', '┠', '┨', '┯', '┷', '┿', '┝', '┥', '┰', '┸', '╂']

# --- 階層ごとの正規表現パターン ---
RE_SHOU = r'^第[0-9１２３４５６７８９０一二三四五六七八九十百]+章'
RE_JOU  = r'^第[0-9１２３４５６７８９０一二三四五六七八九十百]+条'
RE_KOU  = r'^[０-９0-9]+[．\.]'
RE_GOU  = r'^[①-⑳㉑-㉟㊱-㊿]'
RE_LV5  = r'^[（\(][０-９0-9]+[）\)]'
RE_LV6  = r'^[（\(][ア-ンァ-ォ]+[）\)]'

def convert_char_width(text, mode):
    if mode == "半角に統一":
        return text.translate(FULL2HALF)
    elif mode == "全角に統一（推奨）":
        return text.translate(HALF2FULL)
    return text

def get_hierarchy_match(line, regex_pat, custom_markers_str, width_mode):
    m = re.match(regex_pat, line)
    if m:
        return m.group(0)
    if custom_markers_str:
        converted_custom = convert_char_width(custom_markers_str, width_mode)
        markers = [m.strip() for m in converted_custom.split(',') if m.strip()]
        for marker in markers:
            if line.startswith(marker):
                return marker
    return None

def strip_marker(text, marker_str):
    if not marker_str: return text
    if text.startswith(marker_str):
        return text[len(marker_str):].lstrip()
    return text

def set_font(run, font_name, size_pt=None, bold=False):
    run.font.name = font_name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
    run.font.color.rgb = RGBColor(0, 0, 0)
    if size_pt:
        run.font.size = Pt(size_pt)
    if bold:
        run.bold = True

def add_toc(doc):
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.snap_to_grid = False
    run = paragraph.add_run("目　次")
    set_font(run, 'MS ゴシック', size_pt=14, bold=True)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    p = doc.add_paragraph()
    p.paragraph_format.snap_to_grid = False
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

def apply_format_sync(p, font_name, size_pt=10.5, bold=False, base_ind=0.0, hanging_ind=0.0, align_center=False):
    pf = p.paragraph_format
    pf.snap_to_grid = False
    
    total_left = base_ind + hanging_ind
    pf.left_indent = Pt(10.5 * total_left)
    
    if hanging_ind != 0:
        pf.first_line_indent = Pt(-10.5 * hanging_ind)
    else:
        pf.first_line_indent = Pt(0)
    
    pf.space_after = Pt(0)
    pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
    
    if align_center:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    
    if not p.runs:
        return
    for run in p.runs:
        set_font(run, font_name, size_pt, bold)

def create_word_doc(text, selected_font, width_mode, add_space_shou, add_space_jou, add_space_kou, add_space_gou, add_space_lv5, add_space_lv6, indent_body_shou, indent_body_jou, indent_body_kou, indent_body_gou, indent_body_lv5, indent_body_lv6, art_indent_val, indent_kou, hanging_kou, indent_gou, hanging_gou, indent_lv5, hanging_lv5, indent_lv6, hanging_lv6, custom_kou, custom_gou, custom_lv5, custom_lv6, out_mode):
    doc = docx.Document('template.docx')
    if "アウトライン" in out_mode:
        set_update_fields(doc)
    
    style = doc.styles['Normal']
    font = style.font
    font.name = selected_font
    font.size = Pt(10.5)
    style._element.rPr.rFonts.set(qn('w:eastAsia'), selected_font)
    style.paragraph_format.space_after = Pt(0)

    toc_added = False
    art_buffer = []
    csv_buffer = []
    block_buffer = []
    in_code_block = False
    in_csv_table = False
    is_first_text_line = True 
    last_ctx = "none"
    
    def flush_art(lines_to_flush):
        if not lines_to_flush: return
        table = doc.add_table(rows=1, cols=1)
        table.style = 'Table Grid'
        tbl_pr = table._element.xpath('w:tblPr')
        if tbl_pr:
            tbl_ind = OxmlElement('w:tblInd')
            indent_twips = int(10.5 * art_indent_val * 20) 
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
        p.paragraph_format.left_indent = Pt(0) 
        
        for i, line in enumerate(lines_to_flush):
            if i > 0: p.add_run('\n')
            # アート内の文字も同様に指定モードで変換
            converted_line = convert_char_width(line, "全角に統一（推奨）")
            run = p.add_run(converted_line)
            for t in run._element.findall('.//w:t', namespaces=run._element.nsmap):
                t.set(qn('xml:space'), 'preserve')
            set_font(run, 'ＭＳ ゴシック', size_pt=8.5)
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
                p.paragraph_format.snap_to_grid = False
                p.paragraph_format.left_indent = Pt(10.5 * indent_body_jou)
                run = p.add_run(convert_char_width(row.strip(), width_mode))
                set_font(run, selected_font, size_pt=10.5)
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
                        p.paragraph_format.snap_to_grid = False
                        for run in p.runs:
                            # セル内のテキストを指定モードで変換
                            run.text = convert_char_width(run.text, width_mode)
                            set_font(run, selected_font, size_pt=9)
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
        
        is_heading = bool(re.match(r'^第[0-9１２３４５６７８９０一二三四五六七八九十百]+[章条]', line_strip)) or line_strip in ["就　業　規　則", "就業規則", "賃　金　規　程", "賃金規程", "育児・介護休業規程", "退職金規程"]
        
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
        
        m_shou = get_hierarchy_match(line_strip, RE_SHOU, None, width_mode)
        m_jou = get_hierarchy_match(line_strip, RE_JOU, None, width_mode)
        m_kou = get_hierarchy_match(line_strip, RE_KOU, custom_kou, width_mode)
        m_gou = get_hierarchy_match(line_strip, RE_GOU, custom_gou, width_mode)
        m_lv5 = get_hierarchy_match(line_strip, RE_LV5, custom_lv5, width_mode)
        m_lv6 = get_hierarchy_match(line_strip, RE_LV6, custom_lv6, width_mode)

        is_title = False
        if is_first_text_line:
            if not m_shou and not m_jou:
                is_title = True
            is_first_text_line = False
        elif line_strip in ["就　業　規　則", "就業規則", "賃　金　規　程", "賃金規程", "育児・介護休業規程", "退職金規程"]:
            is_title = True

        if (is_title or m_shou or m_jou) and not toc_added:
            add_toc(doc)
            toc_added = True

        if is_title:
            p = doc.add_paragraph()
            run = p.add_run(line_strip)
            apply_format_sync(p, selected_font, size_pt=13.5, bold=True, align_center=True)
            p.paragraph_format.space_after = Pt(24)
            last_ctx = "shou"

        elif m_shou:
            out_txt = strip_marker(line_strip, m_shou) if "アウトライン" in out_mode else line_strip
            p = doc.add_paragraph(out_txt, style='Heading 1' if "アウトライン" in out_mode else None)
            apply_format_sync(p, selected_font, size_pt=13, bold=True, align_center=True)
            p.paragraph_format.space_before = Pt(18)
            p.paragraph_format.space_after = Pt(12)
            last_ctx = "shou"
            
        elif m_jou:
            out_txt = strip_marker(line_strip, m_jou) if "アウトライン" in out_mode else line_strip
            p = doc.add_paragraph(out_txt, style='Heading 2' if "アウトライン" in out_mode else None)
            apply_format_sync(p, selected_font, size_pt=10.5, bold=True, base_ind=0.0)
            p.paragraph_format.space_before = Pt(6)
            
            pPr = p._element.get_or_add_pPr()
            pBdr = OxmlElement('w:pBdr')
            left = OxmlElement('w:left')
            left.set(qn('w:val'), 'single')
            left.set(qn('w:sz'), '24') 
            left.set(qn('w:space'), '4') 
            left.set(qn('w:color'), '333333')
            pBdr.append(left)
            pPr.append(pBdr)
            
            last_ctx = "jou"
        
        elif m_kou:
            out_txt = strip_marker(line_strip, m_kou) if "アウトライン" in out_mode else line_strip
            p = doc.add_paragraph(out_txt, style='Heading 3' if "アウトライン" in out_mode else None)
            apply_format_sync(p, selected_font, size_pt=10.5, bold=False, base_ind=indent_kou, hanging_ind=hanging_kou)
            last_ctx = "kou"

        elif m_gou:
            out_txt = strip_marker(line_strip, m_gou) if "アウトライン" in out_mode else line_strip
            p = doc.add_paragraph(out_txt, style='Heading 4' if "アウトライン" in out_mode else None)
            apply_format_sync(p, selected_font, size_pt=10.5, bold=False, base_ind=indent_gou, hanging_ind=hanging_gou)
            last_ctx = "gou"

        elif m_lv5:
            out_txt = strip_marker(line_strip, m_lv5) if "アウトライン" in out_mode else line_strip
            p = doc.add_paragraph(out_txt, style='Heading 5' if "アウトライン" in out_mode else None)
            apply_format_sync(p, selected_font, size_pt=10.5, bold=False, base_ind=indent_lv5, hanging_ind=hanging_lv5)
            last_ctx = "lv5"

        elif m_lv6:
            out_txt = strip_marker(line_strip, m_lv6) if "アウトライン" in out_mode else line_strip
            p = doc.add_paragraph(out_txt, style='Heading 6' if "アウトライン" in out_mode else None)
            apply_format_sync(p, selected_font, size_pt=10.5, bold=False, base_ind=indent_lv6, hanging_ind=hanging_lv6)
            last_ctx = "lv6"

        else:
            p = doc.add_paragraph()
            
            add_space = False
            if last_ctx in ["shou", "title", "none"]:
                ind = indent_body_shou
                add_space = add_space_shou
            elif last_ctx == "jou":
                ind = indent_body_jou
                add_space = add_space_jou
            elif last_ctx == "kou":
                ind = indent_body_kou
                add_space = add_space_kou
            elif last_ctx == "gou":
                ind = indent_body_gou
                add_space = add_space_gou
            elif last_ctx == "lv5":
                ind = indent_body_lv5
                add_space = add_space_lv5
            elif last_ctx == "lv6":
                ind = indent_body_lv6
                add_space = add_space_lv6
            
            output_text = line_strip
            if add_space and output_text:
                output_text = '　' + output_text
            run = p.add_run(output_text)
            
            apply_format_sync(p, selected_font, size_pt=10.5, bold=False, base_ind=ind, hanging_ind=0.0)
            
    if block_buffer: flush_code_block()
    if art_buffer: flush_art(art_buffer)
    if csv_buffer: flush_csv(csv_buffer)
        
    if not toc_added: add_toc(doc)
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
    html += f"<pre style='border: 1px solid #999; background-color: #fff; padding: 10px; font-family: \"MS Gothic\", \"ＭＳ ゴシック\", monospace; font-size: 13px; line-height: 1.2; margin: 0 0 10px; margin-left: {indent_art}em; white-space: pre; overflow-x: auto;'>{chr(10).join(converted_lines)}</pre>"
    return html

def reset_settings():
    st.session_state.font_choice = "MS 明朝"
    st.session_state.width_mode_choice = "全角に統一（推奨）"
    
    # --- スペース個別設定の初期化 ---
    st.session_state.add_space_shou = False
    st.session_state.add_space_jou = False
    st.session_state.add_space_kou = False
    st.session_state.add_space_gou = False
    st.session_state.add_space_lv5 = False
    st.session_state.add_space_lv6 = False
    
    # --- 本文インデントの初期値 ---
    st.session_state.in_body_shou = 0.0
    st.session_state.in_body_jou = 2.0
    st.session_state.in_body_kou = 4.1
    st.session_state.in_body_gou = 5.1
    st.session_state.in_body_lv5 = 7.1
    st.session_state.in_body_lv6 = 9.1
    
    # --- 階層の初期値 ---
    st.session_state.ik = 2.8                       
    st.session_state.hk = 1.3                       
    st.session_state.ig = 4.1                       
    st.session_state.hg = 1.0                       
    st.session_state.i5 = 4.1                       
    st.session_state.h5 = 3.0                       
    st.session_state.i6 = 6.1                       
    st.session_state.h6 = 3.0                       
    
    st.session_state.art_type = "本文（番号なし）の開始位置に合わせる"
    st.session_state.c_kou = ""
    st.session_state.c_gou = ""
    st.session_state.c_lv5 = ""
    st.session_state.c_lv6 = ""

# ==========================================================
# 🔒 顧問先 個別ログイン管理システム
# ==========================================================
CLIENT_CREDENTIALS = {
    "a_company": {"name": "株式会社A 様", "password": "passA"},
    "b_company": {"name": "株式会社B 様", "password": "passB"},
    # "c_company": {"name": "【解約済】株式会社C 様", "password": "passC"}, # ← このように # を付けるだけで遮断
    "test_user": {"name": "テストユーザー 様", "password": "1234"},
}

def login_screen():
    if "logged_in_client" not in st.session_state:
        st.session_state["logged_in_client"] = None

    if not st.session_state["logged_in_client"]:
        st.markdown("## 🔒 顧問先様専用 規程作成ツール")
        st.info("陶守社会保険労務士事務所から発行されたログインIDとパスワードを入力してください。")
        
        with st.form("login_form"):
            client_id = st.text_input("ログインID")
            client_pwd = st.text_input("パスワード", type="password")
            submitted = st.form_submit_button("ログイン")
            
            if submitted:
                if client_id in CLIENT_CREDENTIALS and CLIENT_CREDENTIALS[client_id]["password"] == client_pwd:
                    st.session_state["logged_in_client"] = CLIENT_CREDENTIALS[client_id]["name"]
                    st.rerun()
                else:
                    st.error("⚠️ IDまたはパスワードが間違っています。")
        return False
    return True

def main():
    st.set_page_config(page_title="就業規則 Word変換", layout="wide")
    
    # ▼ 全体公開（リードマグネット）として運用する場合は、以下の2行を # でコメントアウトしてください ▼
    if not login_screen():
        return
    st.sidebar.success(f"👤 ログイン中: **{st.session_state['logged_in_client']}**")
    # ▲ 全体公開する場合はここまでコメントアウト ▲

    st.title("📄 就業規則プロフェッショナル整形ツール")
    
    if "initialized" not in st.session_state:
        reset_settings()
        st.session_state.initialized = True

    st.sidebar.header("🚀 出力エンジン・書式設定")
    out_mode = st.sidebar.radio("モード選択", ["【推奨】アウトライン連動", "直接モード"])
    if "アウトライン" in out_mode:
        df_outline = pd.DataFrame({"Lv": ["1", "2", "3", "4", "5", "6"], "役割": ["章", "条", "項", "号", "(1)", "(ア)"], "スタイル": ["見出し1", "見出し2", "見出し3", "見出し4", "見出し5", "見出し6"]})
        st.sidebar.table(df_outline)

    st.sidebar.button("🔄 推奨設定（初期値）に戻す", on_click=reset_settings, use_container_width=True)
    st.sidebar.markdown("---")

    selected_font = st.sidebar.selectbox("基本フォント", options=["MS 明朝", "MS ゴシック", "Meiryo", "Yu Mincho", "Yu Gothic", "Arial"], key="font_choice")
    width_mode = st.sidebar.radio("英数字の表記統一", options=["全角に統一（推奨）", "半角に統一", "変換しない（元のまま）"], key="width_mode_choice")
    
    # ==========================================
    # UIの究極改修: 階層ごとの完全ブロック化
    # ==========================================
    st.sidebar.markdown("---")
    st.sidebar.write("**【第1階層】章（表題）の設定**")
    add_space_shou = st.sidebar.checkbox("本文の先頭に全角スペースを自動挿入", key="add_space_shou")
    indent_body_shou = st.sidebar.slider("章 直後の本文 字下げ", min_value=0.0, max_value=10.0, step=0.1, key="in_body_shou")
    
    st.sidebar.markdown("---")
    st.sidebar.write("**【第2階層】条の設定**")
    add_space_jou = st.sidebar.checkbox("本文の先頭に全角スペースを自動挿入", key="add_space_jou")
    indent_body_jou = st.sidebar.slider("条 直後の本文 字下げ", min_value=0.0, max_value=10.0, step=0.1, key="in_body_jou")

    st.sidebar.markdown("---")
    st.sidebar.write("**【第3階層】項（1. 等）の設定**")
    indent_kou = st.sidebar.slider("項 の字下げ", min_value=0.0, max_value=15.0, step=0.1, key="ik")
    hanging_kou = st.sidebar.slider("項 の突き出し幅", min_value=0.0, max_value=5.0, step=0.1, key="hk")
    add_space_kou = st.sidebar.checkbox("本文の先頭に全角スペースを自動挿入", key="add_space_kou")
    indent_body_kou = st.sidebar.slider("項 直後の本文 字下げ", min_value=0.0, max_value=15.0, step=0.1, key="in_body_kou")

    st.sidebar.markdown("---")
    st.sidebar.write("**【第4階層】号（① 等）の設定**")
    indent_gou = st.sidebar.slider("号 の字下げ", min_value=0.0, max_value=15.0, step=0.1, key="ig")
    hanging_gou = st.sidebar.slider("号 の突き出し幅", min_value=0.0, max_value=5.0, step=0.1, key="hg")
    add_space_gou = st.sidebar.checkbox("本文の先頭に全角スペースを自動挿入", key="add_space_gou")
    indent_body_gou = st.sidebar.slider("号 直後の本文 字下げ", min_value=0.0, max_value=15.0, step=0.1, key="in_body_gou")

    st.sidebar.markdown("---")
    st.sidebar.write("**【第5階層】(1) 等の設定**")
    indent_lv5 = st.sidebar.slider("第5階層 の字下げ", min_value=0.0, max_value=15.0, step=0.1, key="i5")
    hanging_lv5 = st.sidebar.slider("第5階層 の突き出し幅", min_value=0.0, max_value=5.0, step=0.1, key="h5")
    add_space_lv5 = st.sidebar.checkbox("本文の先頭に全角スペースを自動挿入", key="add_space_lv5")
    indent_body_lv5 = st.sidebar.slider("第5階層 直後の本文 字下げ", min_value=0.0, max_value=15.0, step=0.1, key="in_body_lv5")

    st.sidebar.markdown("---")
    st.sidebar.write("**【第6階層】(ア) 等の設定**")
    indent_lv6 = st.sidebar.slider("第6階層 の字下げ", min_value=0.0, max_value=15.0, step=0.1, key="i6")
    hanging_lv6 = st.sidebar.slider("第6階層 の突き出し幅", min_value=0.0, max_value=5.0, step=0.1, key="h6")
    add_space_lv6 = st.sidebar.checkbox("本文の先頭に全角スペースを自動挿入", key="add_space_lv6")
    indent_body_lv6 = st.sidebar.slider("第6階層 直後の本文 字下げ", min_value=0.0, max_value=15.0, step=0.1, key="in_body_lv6")

    st.sidebar.markdown("---")
    with st.sidebar.expander("⚙️ 特殊な記号を追加（上級者向け）"):
        st.write("自社の規定で特殊な記号を使っている場合のみ入力。（複数ある場合はカンマ区切り）")
        custom_kou = st.text_input("第3階層（項）に追加", placeholder="例: ◆, ●", key="c_kou")
        custom_gou = st.text_input("第4階層（号）に追加", placeholder="例: ・", key="c_gou")
        custom_lv5 = st.text_input("第5階層に追加", placeholder="例: a.", key="c_lv5")
        custom_lv6 = st.text_input("第6階層に追加", placeholder="例: (a)", key="c_lv6")

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
        calculated_art_indent = indent_body_jou
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
            is_first_text_line = True 
            last_p = "none"
            
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
                
                m_sh = get_hierarchy_match(line_strip, RE_SHOU, None, width_mode)
                m_jo = get_hierarchy_match(line_strip, RE_JOU, None, width_mode)
                m_ko = get_hierarchy_match(line_strip, RE_KOU, custom_kou, width_mode)
                m_go = get_hierarchy_match(line_strip, RE_GOU, custom_gou, width_mode)
                m_l5 = get_hierarchy_match(line_strip, RE_LV5, custom_lv5, width_mode)
                m_l6 = get_hierarchy_match(line_strip, RE_LV6, custom_lv6, width_mode)

                is_title = False
                if is_first_text_line:
                    if not m_sh and not m_jo:
                        is_title = True
                    is_first_text_line = False
                elif line_strip in ["就　業　規　則", "就業規則", "賃　金　規　程", "賃金規程", "育児・介護休業規程", "退職金規程"]:
                    is_title = True

                if is_title:
                    preview_html += f"<div style='text-align: center; font-weight: bold; margin-bottom: 24px; font-size: 1.3em;'>{line_strip}</div>"
                    last_p = "title"
                elif m_sh:
                    preview_html += f"<div style='text-align: center; font-weight: bold; margin-top: 15px;'>{line_strip}</div>"; last_p = "shou"
                elif m_jo:
                    preview_html += f"<div style='font-weight: bold; border-left: 4px solid #333; padding-left: 10px; margin-top: 10px;'>{line_strip}</div>"; last_p = "jou"
                elif m_ko:
                    preview_html += f"<div style='margin-left: {indent_kou}em; padding-left: {hanging_kou}em; text-indent: -{hanging_kou}em; color: red;'>{line_strip}</div>"; last_p = "kou"
                elif m_go:
                    preview_html += f"<div style='margin-left: {indent_gou}em; padding-left: {hanging_gou}em; text-indent: -{hanging_gou}em; color: red;'>{line_strip}</div>"; last_p = "gou"
                elif m_l5:
                    preview_html += f"<div style='margin-left: {indent_lv5}em; padding-left: {hanging_lv5}em; text-indent: -{hanging_lv5}em; color: red;'>{line_strip}</div>"; last_p = "lv5"
                elif m_l6:
                    preview_html += f"<div style='margin-left: {indent_lv6}em; padding-left: {hanging_lv6}em; text-indent: -{hanging_lv6}em; color: red;'>{line_strip}</div>"; last_p = "lv6"
                else:
                    add_space = False
                    if last_p in ["shou", "title", "none"]:
                        ind = indent_body_shou
                        add_space = add_space_shou
                    elif last_p == "jou":
                        ind = indent_body_jou
                        add_space = add_space_jou
                    elif last_p == "kou":
                        ind = indent_body_kou
                        add_space = add_space_kou
                    elif last_p == "gou":
                        ind = indent_body_gou
                        add_space = add_space_gou
                    elif last_p == "lv5":
                        ind = indent_body_lv5
                        add_space = add_space_lv5
                    elif last_p == "lv6":
                        ind = indent_body_lv6
                        add_space = add_space_lv6
                    
                    output_text = line_strip
                    if add_space and output_text:
                        output_text = '　' + output_text
                    
                    base_off = "14px" if last_p == "jou" else "0px"
                    preview_html += f"<div style='margin-left: calc({ind}em + {base_off}); color:#555;'>{output_text}</div>"
            
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
        doc = create_word_doc(text_input, selected_font, width_mode, add_space_shou, add_space_jou, add_space_kou, add_space_gou, add_space_lv5, add_space_lv6, indent_body_shou, indent_body_jou, indent_body_kou, indent_body_gou, indent_body_lv5, indent_body_lv6, calculated_art_indent, indent_kou, hanging_kou, indent_gou, hanging_gou, indent_lv5, hanging_lv5, indent_lv6, hanging_lv6, custom_kou, custom_gou, custom_lv5, custom_lv6, out_mode)
        bio = io.BytesIO()
        doc.save(bio)
        
        st.download_button(
            label=f"📥 {out_mode}設定を反映してWordファイルをダウンロード",
            data=bio.getvalue(),
            file_name="整形済み就業規則.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

if __name__ == "__main__":
    import sys
    import subprocess
    if "streamlit" not in sys.argv[0]:
        subprocess.run([sys.executable, "-m", "streamlit", "run", sys.argv[0]])
    else:
        main()
