import streamlit as st
import pandas as pd
import plotly.express as px
import psycopg2
import json
import io
import zipfile
import os
import re
import base64
import hashlib
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Flowable, Image
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

st.set_page_config(page_title="KTO Karatay YDY Portalı", page_icon="🎓", layout="wide")

def get_logo_b64():
    if os.path.exists("logo.png"):
        try:
            with open("logo.png", "rb") as f:
                return base64.b64encode(f.read()).decode()
        except Exception:
            return ""
    return ""

# ==========================================
# KURUMSAL ARAYÜZ (UI) & TEMA AYARLARI
# ==========================================
def kurumsal_tema_uygula():
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;700&family=Jost:wght@400;500;600&display=swap');

    html, body, [class*="css"] { font-family: 'DM Sans', sans-serif !important; }
    h1, h2, h3, h4, h5, h6 { font-family: 'Jost', sans-serif !important; color: #1A365D !important; }
    .stApp { background-color: #F4F6F9; }
    [data-testid="stSidebar"] { background-color: #1A365D; }
    [data-testid="stSidebar"] * { color: #FFFFFF !important; }

    /* Birincil ve İkincil Butonlar */
    button[kind="primary"] {
        background-color: #E84C0F !important;
        border-color: #E84C0F !important;
        color: white !important;
        font-weight: 600 !important;
        border-radius: 8px !important;
        transition: all 0.2s ease;
    }
    button[kind="primary"]:hover {
        background-color: #CC400A !important;
        border-color: #CC400A !important;
        transform: translateY(-1px);
    }
    button[kind="secondary"] {
        border-color: #1A365D !important;
        color: #1A365D !important;
        border-radius: 8px !important;
    }
    button[kind="secondary"]:hover {
        border-color: #E84C0F !important;
        color: #E84C0F !important;
    }

    /* Metrik ve Tablo (Tam ekran araç çubuğunu engellememesi için overflow: visible) */
    [data-testid="stMetricValue"] { color: #E84C0F !important; font-family: 'Jost', sans-serif !important; font-weight: 600; }
    [data-testid="stDataFrame"] {
        border-radius: 8px;
        box-shadow: 0 2px 8px rgba(26, 54, 93, 0.06);
    }

    /* Seçenek B: Modern Kart Giriş Ekranı */
    .login-card-header {
        background: linear-gradient(135deg, #1A365D 0%, #24477A 100%);
        padding: 28px 24px 22px 24px;
        border-radius: 14px 14px 0 0;
        text-align: center;
        border-bottom: 4px solid #E84C0F;
        margin-bottom: -1px;
    }
    .login-card-header img {
        max-height: 54px;
        margin-bottom: 12px;
    }
    .login-card-title {
        font-family: 'DM Sans', sans-serif;
        font-size: 20px;
        font-weight: 700;
        color: #FFFFFF;
        margin: 0;
        letter-spacing: 0.4px;
    }
    .login-card-sub {
        font-family: 'Jost', sans-serif;
        font-size: 13px;
        color: #CBD5E1;
        margin-top: 4px;
    }

    /* Seçenek C: Kompakt Profesyonel Hoca Paneli */
    .c-header-bar {
        background: linear-gradient(135deg, #1A365D 0%, #234575 100%);
        padding: 16px 22px;
        border-radius: 10px;
        border-bottom: 3px solid #E84C0F;
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 16px;
        box-shadow: 0 4px 12px rgba(26, 54, 93, 0.12);
    }
    .c-header-title {
        font-family: 'DM Sans', sans-serif;
        font-size: 18px;
        font-weight: 700;
        color: #FFFFFF;
        margin: 0;
    }
    .c-header-sub {
        font-family: 'Jost', sans-serif;
        font-size: 12.5px;
        color: #CBD5E1;
        margin-top: 2px;
    }
    .c-badge-open {
        background-color: rgba(22, 163, 74, 0.2);
        color: #86EFAC;
        border: 1px solid #22C55E;
        padding: 5px 12px;
        border-radius: 20px;
        font-size: 12px;
        font-weight: 600;
    }
    .c-badge-locked {
        background-color: rgba(220, 38, 38, 0.25);
        color: #FCA5A5;
        border: 1px solid #EF4444;
        padding: 5px 12px;
        border-radius: 20px;
        font-size: 12px;
        font-weight: 600;
    }
    .c-kpi-grid {
        display: grid;
        grid-template-columns: repeat(3, 1fr);
        gap: 12px;
        margin-bottom: 14px;
    }
    .c-kpi-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-top: 3px solid #1A365D;
        border-radius: 8px;
        padding: 12px 16px;
        box-shadow: 0 1px 4px rgba(0,0,0,0.04);
    }
    .c-kpi-card.accent {
        border-top-color: #E84C0F;
    }
    .c-kpi-lbl {
        font-size: 11px;
        font-weight: 700;
        color: #64748B;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .c-kpi-val {
        font-family: 'Jost', sans-serif;
        font-size: 24px;
        font-weight: 600;
        color: #1A365D;
        margin-top: 2px;
    }
    .c-kpi-card.accent .c-kpi-val {
        color: #E84C0F;
    }
    .sidebar-profile {
        background: rgba(255, 255, 255, 0.08);
        border: 1px solid rgba(255, 255, 255, 0.14);
        border-radius: 10px;
        padding: 12px 14px;
        margin-bottom: 14px;
    }
    .sidebar-role-badge {
        display: inline-block;
        background: #E84C0F;
        color: #FFFFFF;
        font-size: 11px;
        font-weight: 600;
        padding: 2px 9px;
        border-radius: 12px;
        margin-top: 5px;
    }
    </style>
    """, unsafe_allow_html=True)

kurumsal_tema_uygula()

def hash_sifre(sifre: str) -> str:
    if not sifre: return ""
    if len(sifre) == 64 and all(c in '0123456789abcdefABCDEF' for c in sifre): return sifre
    return hashlib.sha256((sifre + st.secrets["PASSWORD_SALT"]).encode('utf-8')).hexdigest()

def sifre_dogrula(girilen_sifre: str, hashli_sifre: str) -> bool:
    return hash_sifre(girilen_sifre) == hashli_sifre

def get_db_connection():
    return psycopg2.connect(st.secrets["SUPABASE_URI"])

@st.cache_resource
def check_and_update_schema():
    try:
        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute("ALTER TABLE sinav_tos ADD COLUMN IF NOT EXISTS b_grubu_karsiligi TEXT DEFAULT '';")
        cur.execute("ALTER TABLE sinav_ayarlari ADD COLUMN IF NOT EXISTS kitapcik_turu TEXT DEFAULT 'Çoklu Kitapçık (A ve B)';")
        conn.commit()
        cur.close(); conn.close()
    except Exception:
        pass
    return True

check_and_update_schema()

def get_db_backup_excel():
    conn = get_db_connection()
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        tablolar = ["sinavlar", "sinav_ayarlari", "sinav_tos", "kullanicilar", "ogrenciler", "hoca_atamalari", "notlar"]
        for t in tablolar:
            try:
                df = pd.read_sql_query(f"SELECT * FROM {t}", conn)
                df.to_excel(writer, sheet_name=t, index=False)
            except: pass
    conn.close()
    buffer.seek(0)
    return buffer

# ==========================================
# PDF VE EXCEL RAPORLAMA FONKSİYONLARI (SEÇENEK C)
# ==========================================
def tr_to_eng(text):
    if not isinstance(text, str): return str(text)
    tr_map = {'ı':'i', 'i':'i', 'I':'I', 'İ':'I', 'ğ':'g', 'Ğ':'G', 'ü':'u', 'Ü':'U', 'ş':'s', 'Ş':'S', 'ö':'o', 'Ö':'O', 'ç':'c', 'Ç':'C'}
    for k, v in tr_map.items(): text = text.replace(k, v)
    return text

FONT_HEADING = "Helvetica-Bold"
FONT_BODY = "Helvetica"
FONT_BODY_MED = "Helvetica-Bold"
USE_TTF = False

try:
    if os.path.exists("fonts/DMSans-Bold.ttf") and os.path.exists("fonts/Jost-Regular.ttf"):
        pdfmetrics.registerFont(TTFont("DMSans-Bold", "fonts/DMSans-Bold.ttf"))
        pdfmetrics.registerFont(TTFont("Jost-Regular", "fonts/Jost-Regular.ttf"))
        FONT_HEADING = "DMSans-Bold"
        FONT_BODY = "Jost-Regular"
        if os.path.exists("fonts/Jost-Medium.ttf"):
            pdfmetrics.registerFont(TTFont("Jost-Medium", "fonts/Jost-Medium.ttf"))
            FONT_BODY_MED = "Jost-Medium"
        else:
            FONT_BODY_MED = "DMSans-Bold"
        USE_TTF = True
except Exception:
    pass

def pdf_txt(text):
    s = str(text) if text is not None else ""
    return s if USE_TTF else tr_to_eng(s)

def safe_float(val):
    try:
        s = str(val).strip().replace(',', '.')
        if s.lower() in ['', 'nan', 'none', '<na>']: return 0.0
        return float(s)
    except Exception:
        return 0.0

def create_corporate_class_excel(sec_sinav, sec_sube, df_s, df_aktif, cols_to_show, beceri_listesi):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sınıf Raporu"
    ws.views.sheetView[0].showGridLines = False

    df_excel = df_s[cols_to_show].sort_values(by="Sınav Notu", ascending=False).copy()
    df_excel["Sınava Girdi"] = df_excel["Sınava Girdi"].map({True: "Katıldı", False: "Girmedi"})
    df_excel = df_excel.rename(columns={
        "Ogrenci_No": "Öğrenci No",
        "Ad_Soyad": "Ad Soyad",
        "Sınav Notu": "Toplam Not"
    })

    headers = list(df_excel.columns)
    num_cols = max(len(headers), 5)
    last_col_letter = get_column_letter(num_cols)

    fill_navy = PatternFill(start_color="1A365D", end_color="1A365D", fill_type="solid")
    fill_zebra = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    fill_kpi_lbl = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    fill_white = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    fill_green = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    fill_yellow = PatternFill(start_color="FEF9C3", end_color="FEF9C3", fill_type="solid")
    fill_red = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")

    font_uni = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    font_sub = Font(name="Calibri", size=10, color="CBD5E1")
    font_sec_title = Font(name="Calibri", size=11, bold=True, color="1A365D")
    font_th = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
    font_td = Font(name="Calibri", size=9, color="1E293B")
    font_td_bold = Font(name="Calibri", size=9, bold=True, color="1A365D")
    font_kpi_lbl = Font(name="Calibri", size=8, bold=True, color="64748B")
    font_kpi_orange = Font(name="Calibri", size=13, bold=True, color="E84C0F")
    font_kpi_navy = Font(name="Calibri", size=12, bold=True, color="1A365D")
    font_green = Font(name="Calibri", size=9, bold=True, color="166534")
    font_yellow = Font(name="Calibri", size=9, bold=True, color="854D0E")
    font_red = Font(name="Calibri", size=9, bold=True, color="991B1B")
    font_footer = Font(name="Calibri", size=8, italic=True, color="64748B")

    thin_gray = Side(border_style="thin", color="E2E8F0")
    med_orange = Side(border_style="medium", color="E84C0F")
    med_navy = Side(border_style="medium", color="1A365D")

    border_cell = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray)
    border_banner_bottom = Border(bottom=med_orange)
    border_kpi_top_orange = Border(top=med_orange, left=thin_gray, right=thin_gray)
    border_kpi_top_navy = Border(top=med_navy, left=thin_gray, right=thin_gray)
    border_kpi_bottom = Border(bottom=thin_gray, left=thin_gray, right=thin_gray)

    ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 20
    for r in [1, 2]:
        for c in range(1, num_cols + 1):
            cell = ws.cell(row=r, column=c)
            cell.fill = fill_navy
            if r == 2:
                cell.border = border_banner_bottom

    logo_added = False
    if os.path.exists("logo.png"):
        try:
            xl_img = XLImage("logo.png")
            ratio = xl_img.width / float(xl_img.height) if xl_img.height else 1.0
            xl_img.height = 46
            xl_img.width = int(46 * ratio)
            ws.add_image(xl_img, "A1")
            logo_added = True
        except Exception:
            logo_added = False

    start_col_text = "B" if logo_added else "A"
    ws.merge_cells(f"{start_col_text}1:{last_col_letter}1")
    ws.merge_cells(f"{start_col_text}2:{last_col_letter}2")

    sinav_temiz = str(sec_sinav).replace("_", " ").strip()
    c_uni = ws[f"{start_col_text}1"]
    c_uni.value = f"KTO KARATAY UNIVERSITY  |  {sinav_temiz} - {sec_sube} CLASS PERFORMANCE REPORT"
    c_uni.font = font_uni
    c_uni.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    c_sub = ws[f"{start_col_text}2"]
    c_sub.value = "Department of Foreign Languages | English Preparatory Program"
    c_sub.font = font_sub
    c_sub.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    ws.row_dimensions[3].height = 10
    ws.row_dimensions[4].height = 18
    ws.row_dimensions[5].height = 24

    kpi_items = [
        ("SINIF ORTALAMASI", f"{df_aktif['Sınav Notu'].mean():.2f}" if not df_aktif.empty else "0.00", True),
        ("MEDYAN", f"{df_aktif['Sınav Notu'].median():.2f}" if not df_aktif.empty else "0.00", False),
        ("EN YÜKSEK NOT", f"{df_aktif['Sınav Notu'].max():.2f}" if not df_aktif.empty else "0.00", False),
        ("EN DÜŞÜK NOT", f"{df_aktif['Sınav Notu'].min():.2f}" if not df_aktif.empty else "0.00", False),
        ("AKTİF ÖĞRENCİ", str(len(df_aktif)), False),
    ]
    for idx, (lbl, val, is_primary) in enumerate(kpi_items, start=1):
        c_lbl = ws.cell(row=4, column=idx, value=lbl)
        c_lbl.font = font_kpi_lbl
        c_lbl.fill = fill_kpi_lbl
        c_lbl.alignment = Alignment(horizontal="center", vertical="center")
        c_lbl.border = border_kpi_top_orange if is_primary else border_kpi_top_navy

        c_val = ws.cell(row=5, column=idx, value=val)
        c_val.font = font_kpi_orange if is_primary else font_kpi_navy
        c_val.fill = fill_white
        c_val.alignment = Alignment(horizontal="center", vertical="center")
        c_val.border = border_kpi_bottom

    ws.row_dimensions[6].height = 12
    ws.row_dimensions[7].height = 20
    ws.merge_cells(f"A7:{last_col_letter}7")
    c_t1 = ws["A7"]
    c_t1.value = "ÖĞRENCİ NOTLARI VE DETAY TABLOSU"
    c_t1.font = font_sec_title
    c_t1.alignment = Alignment(horizontal="left", vertical="center")

    header_row = 8
    ws.row_dimensions[header_row].height = 30
    for c_idx, col_name in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=c_idx, value=col_name)
        cell.font = font_th
        cell.fill = fill_navy
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border_cell

    curr_row = header_row + 1
    for r_i, (_, row) in enumerate(df_excel.iterrows()):
        ws.row_dimensions[curr_row].height = 20
        is_zebra = (r_i % 2 == 1)
        row_fill = fill_zebra if is_zebra else fill_white
        for c_idx, col_name in enumerate(headers, start=1):
            val = row[col_name]
            cell = ws.cell(row=curr_row, column=c_idx, value=val)
            cell.fill = row_fill
            cell.border = border_cell
            if col_name == "Ad Soyad":
                cell.font = font_td
                cell.alignment = Alignment(horizontal="left", vertical="center")
            elif col_name == "Toplam Not":
                cell.font = font_td_bold
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_name == "Sınava Girdi" and str(val) == "Girmedi":
                cell.font = font_red
                cell.fill = fill_red
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.font = font_td
                cell.alignment = Alignment(horizontal="center", vertical="center")
        curr_row += 1

    if beceri_listesi:
        curr_row += 1
        ws.row_dimensions[curr_row].height = 20
        ws.merge_cells(f"A{curr_row}:{last_col_letter}{curr_row}")
        c_t2 = ws[f"A{curr_row}"]
        c_t2.value = "BECERİ VE KAZANIM ANALİZİ"
        c_t2.font = font_sec_title
        c_t2.alignment = Alignment(horizontal="left", vertical="center")
        curr_row += 1

        ws.row_dimensions[curr_row].height = 24
        use_wide_merge = (num_cols >= 7)
        if use_wide_merge:
            ws.cell(row=curr_row, column=1, value="Bölüm")
            ws.merge_cells(start_row=curr_row, start_column=2, end_row=curr_row, end_column=4)
            ws.cell(row=curr_row, column=2, value="Ölçülen Kazanım")
            ws.cell(row=curr_row, column=5, value="Sınıf Ort.")
            ws.cell(row=curr_row, column=6, value="Max Puan")
            ws.cell(row=curr_row, column=7, value="Başarı (%)")
            for c_i in range(1, 8):
                hc = ws.cell(row=curr_row, column=c_i)
                hc.font = font_th
                hc.fill = fill_navy
                hc.alignment = Alignment(horizontal="center" if c_i != 2 else "left", vertical="center")
                hc.border = border_cell
        else:
            b_headers = ["Bölüm", "Ölçülen Kazanım", "Sınıf Ort.", "Max Puan", "Başarı (%)"]
            for c_i, bh in enumerate(b_headers, start=1):
                hc = ws.cell(row=curr_row, column=c_i, value=bh)
                hc.font = font_th
                hc.fill = fill_navy
                hc.alignment = Alignment(horizontal="center" if c_i != 2 else "left", vertical="center")
                hc.border = border_cell

        curr_row += 1
        for b_i, b in enumerate(beceri_listesi):
            ws.row_dimensions[curr_row].height = 22
            b_fill = fill_zebra if (b_i % 2 == 1) else fill_white
            yuzde = float(b.get('yuzde', 0))
            if yuzde >= 75:
                y_fill, y_font = fill_green, font_green
            elif yuzde >= 50:
                y_fill, y_font = fill_yellow, font_yellow
            else:
                y_fill, y_font = fill_red, font_red

            if use_wide_merge:
                ws.merge_cells(start_row=curr_row, start_column=2, end_row=curr_row, end_column=4)
                c1 = ws.cell(row=curr_row, column=1, value=b.get('isim', ''))
                c2 = ws.cell(row=curr_row, column=2, value=b.get('kazanim', ''))
                c5 = ws.cell(row=curr_row, column=5, value=round(float(b.get('ort', 0)), 2))
                c6 = ws.cell(row=curr_row, column=6, value=b.get('max', 0))
                c7 = ws.cell(row=curr_row, column=7, value=f"%{yuzde:.1f}")
                for c_i in range(1, 8):
                    cc = ws.cell(row=curr_row, column=c_i)
                    cc.fill = b_fill
                    cc.border = border_cell
                    cc.font = font_td
                c1.font = font_td_bold
                c1.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                c2.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                c5.alignment = Alignment(horizontal="center", vertical="center")
                c6.alignment = Alignment(horizontal="center", vertical="center")
                c7.alignment = Alignment(horizontal="center", vertical="center")
                c7.fill = y_fill
                c7.font = y_font
            else:
                vals = [b.get('isim', ''), b.get('kazanim', ''), round(float(b.get('ort', 0)), 2), b.get('max', 0), f"%{yuzde:.1f}"]
                for c_i, v in enumerate(vals, start=1):
                    cc = ws.cell(row=curr_row, column=c_i, value=v)
                    cc.fill = y_fill if c_i == 5 else b_fill
                    cc.font = y_font if c_i == 5 else (font_td_bold if c_i == 1 else font_td)
                    cc.border = border_cell
                    cc.alignment = Alignment(horizontal="left" if c_i in [1, 2] else "center", vertical="center", wrap_text=True)
            curr_row += 1

    curr_row += 1
    ws.merge_cells(f"A{curr_row}:{last_col_letter}{curr_row}")
    c_ft = ws[f"A{curr_row}"]
    c_ft.value = "SoFL | Testing and Evaluation Office — KTO Karatay University"
    c_ft.font = font_footer
    c_ft.alignment = Alignment(horizontal="left", vertical="center")

    col_widths = {1: 18, 2: 26, 3: 11, 4: 13, 5: 13}
    for c_i in range(1, num_cols + 1):
        ws.column_dimensions[get_column_letter(c_i)].width = col_widths.get(c_i, 14)

    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer

def karne_satirlari_hazirla(tos_df, df_aktif, ogr_row, tolerans=6):
    isaretler = [
        ("1. Okuma", "2. Okuma", "3. Okuma"),
        ("(1. Okuma)", "(2. Okuma)", "(3. Okuma)"),
        ("1st", "2nd", "3rd"),
        ("1. Hoca", "2. Hoca", "3. Hoca"),
        ("(1)", "(2)", "(3)"),
        ("Okuma 1", "Okuma 2", "Okuma 3"),
        ("Rater 1", "Rater 2", "Rater 3"),
        (" 1 ", " 2 ", " 3 "),
        (" 1", " 2", " 3"),
        ("1.", "2.", "3.")
    ]
    tos_map = {}
    ordered_cols = []
    for _, r in tos_df.iterrows():
        col_key = f"{r['ana_bolum']} - {r['alt_bolum']}"
        if col_key not in tos_map:
            tos_map[col_key] = {"max": float(r['max_puan']), "kazanim": str(r['kazanimlar'] or "")}
            ordered_cols.append(col_key)

    eslesen_gruplar = {}
    atlananlar = set()
    for c1 in ordered_cols:
        if c1 in atlananlar: continue
        for m1, m2, m3 in isaretler:
            if m1 in c1:
                c2 = c1.replace(m1, m2)
                c3 = c1.replace(m1, m3)
                if c2 in tos_map:
                    temiz_ad = re.sub(r'\s+', ' ', c1.replace(m1, "").replace("()", "")).strip()
                    temiz_ad = re.sub(r'\s*-\s*', ' - ', temiz_ad).strip(" -")
                    eslesen_gruplar[c1] = (c1, c2, c3 if c3 in tos_map else None, temiz_ad)
                    atlananlar.update([c1, c2])
                    if c3 in tos_map: atlananlar.add(c3)
                    break

    def cift_puan_hesapla(satir, c1, c2, c3):
        v1 = safe_float(satir.get(c1, 0))
        v2 = safe_float(satir.get(c2, 0))
        if c3 and c3 in satir and abs(v1 - v2) > tolerans:
            v3 = safe_float(satir.get(c3, 0))
            if v3 > 0 or str(satir.get(c3, "")).strip() == "0":
                d1, d2 = abs(v3 - v1), abs(v3 - v2)
                if d1 < d2: return (v3 + v1) / 2.0
                elif d2 < d1: return (v3 + v2) / 2.0
                else: return (v3 + max(v1, v2)) / 2.0
        return (v1 + v2) / 2.0

    satirlar = []
    for col in ordered_cols:
        if col in eslesen_gruplar:
            c1, c2, c3, temiz_ad = eslesen_gruplar[col]
            m_puan = tos_map[c1]["max"]
            kazanim = tos_map[c1]["kazanim"]
            ogr_puan = round(cift_puan_hesapla(ogr_row, c1, c2, c3), 2)
            if not df_aktif.empty:
                sinif_puanlari = [cift_puan_hesapla(r, c1, c2, c3) for _, r in df_aktif.iterrows()]
                sinif_ort = sum(sinif_puanlari) / len(sinif_puanlari)
            else:
                sinif_ort = 0.0
            yuzde = (ogr_puan / m_puan * 100) if m_puan > 0 else 0.0
            satirlar.append({
                "bolum": temiz_ad,
                "kazanim": kazanim,
                "puan": ogr_puan,
                "max": m_puan,
                "sinif_ort": round(sinif_ort, 1),
                "yuzde": min(max(yuzde, 0.0), 100.0)
            })
        elif col not in atlananlar:
            m_puan = tos_map[col]["max"]
            kazanim = tos_map[col]["kazanim"]
            ogr_puan = round(safe_float(ogr_row.get(col, 0)), 2)
            if not df_aktif.empty and col in df_aktif.columns:
                sinif_ort = pd.to_numeric(df_aktif[col].astype(str).str.replace(',', '.'), errors='coerce').fillna(0.0).mean()
            else:
                sinif_ort = 0.0
            yuzde = (ogr_puan / m_puan * 100) if m_puan > 0 else 0.0
            satirlar.append({
                "bolum": col,
                "kazanim": kazanim,
                "puan": ogr_puan,
                "max": m_puan,
                "sinif_ort": round(sinif_ort, 1),
                "yuzde": min(max(yuzde, 0.0), 100.0)
            })
    return satirlar

class ModernBarWithLabel(Flowable):
    def __init__(self, bar_width, height, percentage):
        Flowable.__init__(self)
        self.bar_width = bar_width
        self.height = height
        self.percentage = max(0.0, min(100.0, float(percentage)))
        self.width = bar_width + 28
        if self.percentage >= 75: self.color = colors.HexColor("#16A34A")
        elif self.percentage >= 50: self.color = colors.HexColor("#EAB308")
        else: self.color = colors.HexColor("#DC2626")

    def wrap(self, availWidth, availHeight):
        return (self.width, max(self.height, 10))

    def draw(self):
        y_off = 2
        self.canv.setFillColor(colors.HexColor("#E2E8F0"))
        self.canv.roundRect(0, y_off, self.bar_width, self.height, 3, fill=1, stroke=0)
        fill_w = (self.bar_width * self.percentage) / 100.0
        if fill_w > 0:
            self.canv.setFillColor(self.color)
            self.canv.roundRect(0, y_off, max(fill_w, 3), self.height, 3, fill=1, stroke=0)
        self.canv.setFont(FONT_BODY_MED, 7.5)
        self.canv.setFillColor(colors.HexColor("#334155"))
        self.canv.drawString(self.bar_width + 5, y_off + 0.5, f"{self.percentage:.0f}%")

def _draw_footer(canvas, doc):
    canvas.saveState()
    canvas.setStrokeColor(colors.HexColor("#CBD5E1"))
    canvas.setLineWidth(0.6)
    canvas.line(28, 32, A4[0] - 28, 32)
    canvas.setFont(FONT_BODY, 7.5)
    canvas.setFillColor(colors.HexColor("#64748B"))
    canvas.drawString(28, 20, pdf_txt("SoFL | Testing and Evaluation Office"))
    canvas.drawRightString(A4[0] - 28, 20, pdf_txt("KTO Karatay University"))
    canvas.restoreState()

def create_modern_pdf(rapor_verileri):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=28, leftMargin=28, topMargin=26, bottomMargin=42)
    elements = []

    st_banner_uni = ParagraphStyle('BannerUni', fontName=FONT_HEADING, fontSize=12.5, leading=15, textColor=colors.white)
    st_banner_dept = ParagraphStyle('BannerDept', fontName=FONT_BODY, fontSize=8.5, leading=11, textColor=colors.HexColor("#CBD5E1"))
    st_banner_exam = ParagraphStyle('BannerExam', fontName=FONT_HEADING, fontSize=12, leading=14, textColor=colors.white, alignment=2)
    st_banner_sub = ParagraphStyle('BannerSub', fontName=FONT_BODY_MED, fontSize=7.5, leading=10, textColor=colors.HexColor("#FDBA74"), alignment=2)
    st_stu_left = ParagraphStyle('StuLeft', fontName=FONT_BODY, fontSize=9, leading=12, textColor=colors.HexColor("#1E293B"))
    st_stu_right = ParagraphStyle('StuRight', fontName=FONT_BODY, fontSize=9, leading=12, textColor=colors.HexColor("#1E293B"), alignment=2)
    st_kpi_lbl = ParagraphStyle('KpiLbl', fontName=FONT_HEADING, fontSize=7, leading=9, textColor=colors.HexColor("#64748B"), alignment=1)
    st_kpi_val_orange = ParagraphStyle('KpiValO', fontName=FONT_HEADING, fontSize=13, leading=16, textColor=colors.HexColor("#E84C0F"), alignment=1)
    st_kpi_val_navy = ParagraphStyle('KpiValN', fontName=FONT_HEADING, fontSize=12, leading=16, textColor=colors.HexColor("#1A365D"), alignment=1)
    st_th = ParagraphStyle('TH', fontName=FONT_HEADING, fontSize=7.5, leading=9.5, textColor=colors.white, alignment=1)
    st_th_left = ParagraphStyle('THL', fontName=FONT_HEADING, fontSize=7.5, leading=9.5, textColor=colors.white, alignment=0)

    for idx, veri in enumerate(rapor_verileri):
        sinav_temiz = str(veri.get('sinav_kodu', '')).replace('_', ' ').strip()
        ogr_ad = pdf_txt(veri.get('ogr_ad', ''))
        ogr_no = pdf_txt(veri.get('ogr_no', ''))
        sube = pdf_txt(veri.get('sube', ''))
        kitapcik = pdf_txt(veri.get('kitapcik', 'A'))
        toplam_not = safe_float(veri.get('toplam_not', 0))
        sinif_genel_ort = safe_float(veri.get('sinif_genel_ort', 0))
        satirlar = veri.get('satirlar', [])

        sol_metin = [
            Paragraph("KTO KARATAY UNIVERSITY", st_banner_uni),
            Paragraph("Department of Foreign Languages | English Preparatory Program", st_banner_dept)
        ]
        sag_metin = [
            Paragraph(pdf_txt(sinav_temiz), st_banner_exam),
            Paragraph("INDIVIDUAL REPORT CARD", st_banner_sub)
        ]

        if os.path.exists("logo.png"):
            try:
                ir = ImageReader("logo.png")
                iw, ih = ir.getSize()
                logo_h = 34
                logo_w = min(logo_h * (iw / float(ih)), 105)
                logo_img = Image("logo.png", width=logo_w, height=logo_h)
                banner_data = [[logo_img, sol_metin, sag_metin]]
                banner_widths = [logo_w + 10, 539 - (logo_w + 10) - 145, 145]
            except Exception:
                banner_data = [[sol_metin, sag_metin]]
                banner_widths = [394, 145]
        else:
            banner_data = [[sol_metin, sag_metin]]
            banner_widths = [394, 145]

        banner_table = Table(banner_data, colWidths=banner_widths)
        banner_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor("#1A365D")),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
            ('LEFTPADDING', (0, 0), (0, -1), 12),
            ('RIGHTPADDING', (-1, 0), (-1, -1), 12),
            ('LINEBELOW', (0, 0), (-1, 0), 2.5, colors.HexColor("#E84C0F")),
        ]))
        elements.append(banner_table)
        elements.append(Spacer(1, 10))

        stu_table = Table(
            [[Paragraph(f"<font name='{FONT_HEADING}' color='#1A365D'>STUDENT:</font>  {ogr_ad}", st_stu_left),
              Paragraph(f"<font name='{FONT_HEADING}' color='#1A365D'>STUDENT ID:</font>  {ogr_no}", st_stu_right)]],
            colWidths=[359, 180]
        )
        stu_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor("#F1F5F9")),
            ('BOX', (0, 0), (-1, -1), 0.5, colors.HexColor("#E2E8F0")),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ('LEFTPADDING', (0, 0), (0, -1), 10),
            ('RIGHTPADDING', (-1, 0), (-1, -1), 10),
        ]))
        elements.append(stu_table)
        elements.append(Spacer(1, 8))

        def kpi_box(label, value_text, val_style, accent_hex):
            t = Table([[Paragraph(label, st_kpi_lbl)], [Paragraph(value_text, val_style)]], colWidths=[128])
            t.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, -1), colors.white),
                ('BOX', (0, 0), (-1, -1), 0.5, colors.HexColor("#E2E8F0")),
                ('LINEABOVE', (0, 0), (-1, 0), 2, colors.HexColor(accent_hex)),
                ('TOPPADDING', (0, 0), (-1, 0), 6),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 2),
                ('TOPPADDING', (0, 1), (-1, 1), 1),
                ('BOTTOMPADDING', (0, 1), (-1, 1), 6),
            ]))
            return t

        kpi_strip = Table([[
            kpi_box("TOTAL SCORE", f"{toplam_not:.2f}", st_kpi_val_orange, "#E84C0F"),
            "",
            kpi_box("CLASS AVERAGE", f"{sinif_genel_ort:.2f}", st_kpi_val_navy, "#1A365D"),
            "",
            kpi_box("CLASS", sube, st_kpi_val_navy, "#1A365D"),
            "",
            kpi_box("BOOKLET", kitapcik, st_kpi_val_navy, "#1A365D")
        ]], colWidths=[128, 9, 128, 9, 128, 9, 128])
        kpi_strip.setStyle(TableStyle([
            ('LEFTPADDING', (0, 0), (-1, -1), 0),
            ('RIGHTPADDING', (0, 0), (-1, -1), 0),
            ('TOPPADDING', (0, 0), (-1, -1), 0),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
        ]))
        elements.append(kpi_strip)
        elements.append(Spacer(1, 12))

        compact = len(satirlar) > 13
        body_fs = 7.5 if compact else 8.2
        pad_v = 3.5 if compact else 5.0

        st_sec = ParagraphStyle('Sec', fontName=FONT_BODY_MED, fontSize=body_fs, leading=body_fs + 2, textColor=colors.HexColor("#1E293B"))
        st_obj = ParagraphStyle('Obj', fontName=FONT_BODY, fontSize=body_fs, leading=body_fs + 2, textColor=colors.HexColor("#475569"))
        st_num = ParagraphStyle('Num', fontName=FONT_BODY_MED, fontSize=body_fs, leading=body_fs + 2, textColor=colors.HexColor("#0F172A"), alignment=1)
        st_avg = ParagraphStyle('Avg', fontName=FONT_BODY, fontSize=body_fs, leading=body_fs + 2, textColor=colors.HexColor("#64748B"), alignment=1)

        tablo_verisi = [[
            Paragraph("SECTION", st_th_left),
            Paragraph("LEARNING OBJECTIVES", st_th_left),
            Paragraph("SCORE", st_th),
            Paragraph("CLASS AVG", st_th),
            Paragraph("PERFORMANCE", st_th)
        ]]

        for s in satirlar:
            puan_str = f"{s['puan']:g} / {s['max']:g}"
            ort_str = f"{s['sinif_ort']:.1f}"
            bar = ModernBarWithLabel(48, 6.5, s['yuzde'])
            tablo_verisi.append([
                Paragraph(pdf_txt(s['bolum']), st_sec),
                Paragraph(pdf_txt(s['kazanim']), st_obj),
                Paragraph(puan_str, st_num),
                Paragraph(ort_str, st_avg),
                bar
            ])

        main_table = Table(tablo_verisi, colWidths=[108, 233, 58, 54, 86], repeatRows=1)
        t_style = [
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1A365D")),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('ALIGN', (2, 0), (-1, -1), 'CENTER'),
            ('TOPPADDING', (0, 0), (-1, 0), 6),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
            ('TOPPADDING', (0, 1), (-1, -1), pad_v),
            ('BOTTOMPADDING', (0, 1), (-1, -1), pad_v),
            ('LEFTPADDING', (0, 0), (-1, -1), 6),
            ('RIGHTPADDING', (0, 0), (-1, -1), 6),
            ('LINEBELOW', (0, 1), (-1, -1), 0.4, colors.HexColor("#E2E8F0")),
            ('LINEBELOW', (0, -1), (-1, -1), 1.0, colors.HexColor("#1A365D")),
        ]
        for i in range(1, len(tablo_verisi)):
            if i % 2 == 0:
                t_style.append(('BACKGROUND', (0, i), (-1, i), colors.HexColor("#F8FAFC")))
        main_table.setStyle(TableStyle(t_style))
        elements.append(main_table)

        if idx < len(rapor_verileri) - 1:
            elements.append(PageBreak())

    doc.build(elements, onFirstPage=_draw_footer, onLaterPages=_draw_footer)
    buffer.seek(0)
    return buffer

# ==========================================
# GİRİŞ KAPISI VE OTURUM (SEÇENEK B)
# ==========================================
if 'giris_basarili' not in st.session_state:
    st.session_state.update({'giris_basarili': False, 'kullanici_adi': "", 'rol': "", 'isim': ""})

def giris_yap(k_adi, sifre):
    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute("SELECT sifre, rol, isim FROM kullanicilar WHERE kullanici_adi=%s", (k_adi,))
    kullanici = cur.fetchone()
    cur.close(); conn.close()
    if kullanici and sifre_dogrula(sifre, kullanici[0]):
        st.session_state.update({'giris_basarili': True, 'kullanici_adi': k_adi, 'rol': kullanici[1], 'isim': kullanici[2]})
        st.rerun()
    else: st.error("Kullanıcı adı veya şifre hatalı!")

def cikis_yap():
    st.session_state.update({'giris_basarili': False, 'kullanici_adi': "", 'rol': "", 'isim': ""})
    st.rerun()

logo_b64 = get_logo_b64()

if not st.session_state['giris_basarili']:
    st.markdown("<div style='height: 30px;'></div>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1.2, 1.6, 1.2])
    with col2:
        logo_html = f"<img src='data:image/png;base64,{logo_b64}' alt='KTO Logo'/>" if logo_b64 else ""
        st.markdown(f"""
        <div class="login-card-header">
            {logo_html}
            <div class="login-card-title">KTO KARATAY UNIVERSITY</div>
            <div class="login-card-sub">Department of Foreign Languages | Exam &amp; Analytics Portal</div>
        </div>
        """, unsafe_allow_html=True)
        with st.form("giris_formu"):
            k_adi = st.text_input("Kullanıcı Adı", placeholder="Kullanıcı adınızı girin")
            sifre = st.text_input("Şifre", type="password", placeholder="••••••••")
            st.markdown("<div style='height: 6px;'></div>", unsafe_allow_html=True)
            if st.form_submit_button("Sisteme Giriş Yap", type="primary", use_container_width=True):
                giris_yap(k_adi, sifre)
        st.markdown("<p style='text-align:center; color:#64748B; font-size:12px; margin-top:14px;'>SoFL | Testing and Evaluation Office</p>", unsafe_allow_html=True)
else:
    if logo_b64:
        st.sidebar.markdown(
            f"<div style='text-align:center; padding: 6px 0 14px 0; border-bottom: 1px solid rgba(255,255,255,0.12); margin-bottom:14px;'>"
            f"<img src='data:image/png;base64,{logo_b64}' style='max-height:46px;'/></div>",
            unsafe_allow_html=True
        )
    st.sidebar.markdown(
        f"""
        <div class="sidebar-profile">
            <div style="font-size:15px; font-weight:700;">{st.session_state['isim']}</div>
            <span class="sidebar-role-badge">{st.session_state['rol']}</span>
        </div>
        """,
        unsafe_allow_html=True
    )
    if st.sidebar.button("🚪 Güvenli Çıkış", use_container_width=True): cikis_yap()
    
    with st.sidebar.expander("🔑 Şifremi Değiştir"):
        eski_sifre = st.text_input("Mevcut Şifre", type="password", key="eski_s")
        yeni_sifre_1 = st.text_input("Yeni Şifre", type="password", key="yeni_s1")
        if st.button("Şifreyi Güncelle", key="btn_sifre_degis"):
            if eski_sifre and yeni_sifre_1:
                conn = get_db_connection()
                cur = conn.cursor()
                cur.execute("SELECT sifre FROM kullanicilar WHERE kullanici_adi=%s", (st.session_state['kullanici_adi'],))
                mevcut_hash = cur.fetchone()[0]
                if sifre_dogrula(eski_sifre, mevcut_hash):
                    cur.execute("UPDATE kullanicilar SET sifre=%s WHERE kullanici_adi=%s", (hash_sifre(yeni_sifre_1), st.session_state['kullanici_adi']))
                    conn.commit()
                    st.success("Şifreniz güncellendi!")
                else: st.error("Mevcut şifreniz yanlış!")
                cur.close(); conn.close()

    if st.session_state['rol'] in ["Admin", "Testing"]:
        st.title("🎓 KTO Karatay YDY Öğrenme Analitiği Sistemi")
        if st.session_state['rol'] == "Admin":
            tabs = st.tabs(["📝 Sınav Tanımla", "🔒 Sınav & Veri Yönetimi", "📋 Öğrenci & Atamalar", "👥 Kullanıcılar", "📊 Analitik Raporlar"])
            tab_sinav, tab_yonetim, tab_atamalar, tab_kullanici, tab_rapor = tabs[0], tabs[1], tabs[2], tabs[3], tabs[4]
        else:
            tabs = st.tabs(["📝 Sınav Tanımla", "🔒 Sınav & Veri Yönetimi", "📊 Analitik Raporlar"])
            tab_sinav, tab_yonetim, tab_rapor = tabs[0], tabs[1], tabs[2]
        
        # ----------------------------------------------------
        # 1. SINAV TANIMLA
        # ----------------------------------------------------
        with tab_sinav:
            st.subheader("Standart Sınav Kodu Oluşturucu")
            c1, c2, c3, c4 = st.columns(4)
            with c1: secilen_modul = st.selectbox("Modül", ["M1", "M2", "M3", "M4", "YazOkulu"])
            with c2: secilen_kur = st.selectbox("Kur Seviyesi", ["A1", "A2", "B1", "B2", "PF"])
            with c3: secilen_tur = st.selectbox("Sınav Türü", ["Quiz_1", "Quiz_2", "Midterm", "Final"])
            with c4: tolerans_baremi = st.number_input("Çift Okuma Tolerans Farkı", min_value=0, value=6)
            
            st.markdown("---")
            kitapcik_turu = st.radio("Sınav Formatı (Kitapçık Yapısı)", ["Tek Kitapçık (Standart)", "Çoklu Kitapçık (A ve B)"], horizontal=True)
            sinav_kodu = f"{secilen_modul}_{secilen_kur}_{secilen_tur}"
            
            st.subheader("Gelişmiş Belirtke Tablosu (ToS) İnşaatçısı")
            tos_file = st.file_uploader("📥 Excel'den Hazır ToS Yükle", type=["xlsx", "xls"])
            if tos_file:
                try:
                    df_up_tos = pd.read_excel(tos_file)
                    st.session_state['taslak_df'] = df_up_tos
                    st.success("ToS Excel'den aktarıldı!")
                except: st.error("Format uyumsuz.")

            ANA_BOLUMLER = ["Listening I", "Listening II", "Reading I", "Reading II", "Grammar", "Functional Language", "Use of English", "Vocabulary", "Speaking", "Speaking 1st", "Speaking 2nd", "Speaking 3rd", "Writing", "Writing (1. Okuma)", "Writing (2. Okuma)", "Writing (3. Okuma)"]
            
            if 'taslak_df' not in st.session_state:
                st.session_state['taslak_df'] = pd.DataFrame({
                    "Ana Bölüm": ["Grammar"], 
                    "Alt Bölüm (A Kitapçığı Yeri)": ["Part A"], 
                    "B Kitapçığındaki Yeri (İsteğe Bağlı)": [""],
                    "Ölçülen Kazanımlar": ["Present Simple"], 
                    "Max Puan": ["15,0"]
                })
            
            col_config_tos = {
                "Ana Bölüm": st.column_config.SelectboxColumn("Ana Bölüm", options=ANA_BOLUMLER, required=True),
                "Alt Bölüm (A Kitapçığı Yeri)": st.column_config.TextColumn("Bölüm Adı / Yeri", required=True),
                "Ölçülen Kazanımlar": st.column_config.TextColumn("Ölçülen Kazanımlar"),
                "Max Puan": st.column_config.TextColumn("Max Puan")
            }
            
            if "Çoklu" in kitapcik_turu: col_config_tos["B Kitapçığındaki Yeri (İsteğe Bağlı)"] = st.column_config.TextColumn("B Kitapçığındaki Yeri")
            else: col_config_tos["B Kitapçığındaki Yeri (İsteğe Bağlı)"] = None 

            duzenlenen_rubrik = st.data_editor(st.session_state['taslak_df'], num_rows="dynamic", column_config=col_config_tos, use_container_width=True)
            
            if st.button("🚀 Sınavı Onayla & Öğretim Elemanlarına Aç", type="primary"):
                conn = get_db_connection()
                cur = conn.cursor()
                cur.execute("SELECT sinav_kodu FROM sinavlar WHERE sinav_kodu=%s", (sinav_kodu,))
                if cur.fetchone(): st.warning("Bu sınav kodu zaten açık!")
                else:
                    cur.execute("INSERT INTO sinavlar (sinav_kodu, kilitli) VALUES (%s, FALSE)", (sinav_kodu,))
                    cur.execute("INSERT INTO sinav_ayarlari (sinav_kodu, tolerans, kitapcik_turu) VALUES (%s, %s, %s)", (sinav_kodu, tolerans_baremi, kitapcik_turu))
                    
                    tos_data = []
                    for _, row in duzenlenen_rubrik.iterrows():
                        saf_max = float(str(row.get('Max Puan', '0')).replace(',', '.'))
                        b_grubu = str(row.get('B Kitapçığındaki Yeri (İsteğe Bağlı)', '')).strip() if "Çoklu" in kitapcik_turu else ""
                        tos_data.append((sinav_kodu, row['Ana Bölüm'], row['Alt Bölüm (A Kitapçığı Yeri)'], row['Ölçülen Kazanımlar'], saf_max, b_grubu))
                    
                    if tos_data:
                        cur.executemany("INSERT INTO sinav_tos (sinav_kodu, ana_bolum, alt_bolum, kazanimlar, max_puan, b_grubu_karsiligi) VALUES (%s, %s, %s, %s, %s, %s)", tos_data)
                    conn.commit()
                    st.success(f"{sinav_kodu} ({kitapcik_turu}) başarıyla oluşturuldu!")
                cur.close(); conn.close()

        # ----------------------------------------------------
        # 2. SINAV & VERİ YÖNETİMİ
        # ----------------------------------------------------
        with tab_yonetim:
            conn = get_db_connection()
            df_sinav_durum = pd.read_sql_query("SELECT sinav_kodu, kilitli FROM sinavlar", conn)
            
            if not df_sinav_durum.empty:
                st.markdown("### 🔒 Sınav Kilitleme (Finalization)")
                col_kilit1, col_kilit2 = st.columns(2)
                with col_kilit1:
                    islem_yapilacak_sinav = st.selectbox("Kilidi Değiştirilecek Sınav", df_sinav_durum['sinav_kodu'].tolist())
                    mevcut_durum = df_sinav_durum[df_sinav_durum['sinav_kodu'] == islem_yapilacak_sinav]['kilitli'].iloc[0]
                    st.markdown(f"**Mevcut Durum:** {'🔒 KİLİTLİ (Not girilemez)' if mevcut_durum else '🟢 AÇIK (Not girilebilir)'}")
                with col_kilit2:
                    st.write(""); st.write("")
                    if mevcut_durum and st.button("🔓 Sınav Kilidini Aç"):
                        conn.cursor().execute("UPDATE sinavlar SET kilitli=FALSE WHERE sinav_kodu=%s", (islem_yapilacak_sinav,))
                        conn.commit(); st.rerun()
                    elif not mevcut_durum and st.button("🔒 Sınavı Kilitle", type="primary"):
                        conn.cursor().execute("UPDATE sinavlar SET kilitli=TRUE WHERE sinav_kodu=%s", (islem_yapilacak_sinav,))
                        conn.commit(); st.rerun()
                st.divider()
                
                st.markdown("### ✏️ Sınav Düzenle (Ayarlar ve ToS)")
                duzenlenecek_sinav = st.selectbox("Düzenlenecek Sınavı Seçin", df_sinav_durum['sinav_kodu'].tolist(), key="edit_exam_select")
                cur = conn.cursor()
                cur.execute("SELECT tolerans, kitapcik_turu FROM sinav_ayarlari WHERE sinav_kodu=%s", (duzenlenecek_sinav,))
                ayarlar_res = cur.fetchone()
                
                mevcut_tolerans = int(ayarlar_res[0]) if (ayarlar_res and ayarlar_res[0] is not None) else 6
                mevcut_kturu = ayarlar_res[1] if (ayarlar_res and ayarlar_res[1] is not None) else "Çoklu Kitapçık (A ve B)"
                
                col_e1, col_e2 = st.columns(2)
                with col_e1: yeni_tolerans = st.number_input("Güncel Çift Okuma Tolerans Farkı", min_value=0, value=mevcut_tolerans)
                with col_e2: yeni_kturu = st.radio("Sınav Formatını Değiştir", ["Tek Kitapçık (Standart)", "Çoklu Kitapçık (A ve B)"], index=0 if "Tek" in mevcut_kturu else 1)
                
                df_mevcut_tos = pd.read_sql_query('SELECT ana_bolum AS "Ana Bölüm", alt_bolum AS "Alt Bölüm (A Kitapçığı Yeri)", b_grubu_karsiligi AS "B Kitapçığındaki Yeri (İsteğe Bağlı)", kazanimlar AS "Ölçülen Kazanımlar", max_puan AS "Max Puan" FROM sinav_tos WHERE sinav_kodu=%s', conn, params=(duzenlenecek_sinav,))
                df_mevcut_tos["Max Puan"] = df_mevcut_tos["Max Puan"].astype(str).str.replace(".", ",")
                
                edit_col_config = {
                    "Ana Bölüm": st.column_config.SelectboxColumn("Ana Bölüm", options=ANA_BOLUMLER, required=True),
                    "Alt Bölüm (A Kitapçığı Yeri)": st.column_config.TextColumn("Bölüm Adı / Yeri", required=True),
                    "Max Puan": st.column_config.TextColumn("Max Puan")
                }
                
                if "Çoklu" in yeni_kturu: edit_col_config["B Kitapçığındaki Yeri (İsteğe Bağlı)"] = st.column_config.TextColumn("B Kitapçığındaki Yeri")
                else: edit_col_config["B Kitapçığındaki Yeri (İsteğe Bağlı)"] = None

                st.markdown("**Güncel Belirtke Tablosu (ToS):**")
                guncel_tos = st.data_editor(df_mevcut_tos, num_rows="dynamic", column_config=edit_col_config, use_container_width=True, key="edit_tos_editor")
                
                if st.button("💾 Değişiklikleri Kaydet (Sınavı Güncelle)", type="primary"):
                    cur.execute("UPDATE sinav_ayarlari SET tolerans=%s, kitapcik_turu=%s WHERE sinav_kodu=%s", (yeni_tolerans, yeni_kturu, duzenlenecek_sinav))
                    cur.execute("DELETE FROM sinav_tos WHERE sinav_kodu=%s", (duzenlenecek_sinav,))
                    
                    tos_data = []
                    for _, row in guncel_tos.iterrows():
                        saf_max = float(str(row.get('Max Puan', '0')).replace(',', '.'))
                        b_grubu = str(row.get('B Kitapçığındaki Yeri (İsteğe Bağlı)', '')).strip() if "Çoklu" in yeni_kturu else ""
                        tos_data.append((duzenlenecek_sinav, row['Ana Bölüm'], row['Alt Bölüm (A Kitapçığı Yeri)'], row['Ölçülen Kazanımlar'], saf_max, b_grubu))
                        
                    if tos_data:
                        cur.executemany("INSERT INTO sinav_tos (sinav_kodu, ana_bolum, alt_bolum, kazanimlar, max_puan, b_grubu_karsiligi) VALUES (%s, %s, %s, %s, %s, %s)", tos_data)
                    conn.commit(); st.success(f"{duzenlenecek_sinav} başarıyla güncellendi!")
                st.divider()
                
                st.markdown("### 🗑️ Sınav Silme (Çift Güvenlikli Kalkan)")
                silinecek_sinav = st.selectbox("Silinecek Sınavı Seçin", df_sinav_durum['sinav_kodu'].tolist(), key="sil_sec")
                silme_onay = st.checkbox("⚠️ Bu sınavı ve öğrencilere ait tüm not kayıtlarını kalıcı olarak silmeyi onaylıyorum.")
                silme_sifre = st.text_input("İşlemi onaylamak için giriş şifrenizi tekrar girin:", type="password")
                
                if st.button("🗑️ Seçili Sınavı Tamamen Sil"):
                    if not silme_onay:
                        st.error("Lütfen yukarıdaki onay kutusunu işaretleyin.")
                    else:
                        cur.execute("SELECT sifre FROM kullanicilar WHERE kullanici_adi=%s", (st.session_state['kullanici_adi'],))
                        admin_hash = cur.fetchone()[0]
                        if sifre_dogrula(silme_sifre, admin_hash):
                            cur.execute("DELETE FROM sinavlar WHERE sinav_kodu=%s", (silinecek_sinav,))
                            conn.commit()
                            st.success(f"{silinecek_sinav} sınavı ve tüm verileri sistemden kalıcı olarak silindi!")
                            st.rerun()
                        else:
                            st.error("Şifre hatalı! Güvenlik nedeniyle silme işlemi reddedildi.")
            else: st.info("Sistemde kayıtlı sınav yok.")
            
            if st.session_state['rol'] == "Admin":
                st.divider()
                st.markdown("### 💾 Felaket Kurtarma Yedeği")
                st.download_button(label="📦 Tüm Veritabanını Excel Olarak İndir", data=get_db_backup_excel(), file_name="YDYO_Tam_Sistem_Yedegi.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
            conn.close()

        # ----------------------------------------------------
        # 3. KULLANICILAR & ATAMALAR 
        # ----------------------------------------------------
        if st.session_state['rol'] == "Admin":
            with tab_kullanici:
                st.subheader("👥 Kullanıcı Yönetimi")
                conn = get_db_connection()
                df_kullanicilar = pd.read_sql_query("SELECT kullanici_adi, rol, isim FROM kullanicilar", conn)
                
                with st.expander("➕ Yeni Kullanıcı Ekle", expanded=False):
                    c_k1, c_k2 = st.columns(2)
                    with c_k1: yeni_k_adi = st.text_input("Kullanıcı Adı (ID)", key="new_id"); yeni_sifre = st.text_input("Şifre", type="password", key="new_pw")
                    with c_k2: yeni_rol = st.selectbox("Rol", ["Öğretim Elemanı", "Testing", "Admin"], key="new_role"); yeni_isim = st.text_input("İsim (Örn: Dr. Ahmet)", key="new_name")
                    if st.button("Yeni Kullanıcıyı Kaydet", type="primary"):
                        if yeni_k_adi and yeni_sifre and yeni_isim:
                            cur = conn.cursor()
                            cur.execute("SELECT kullanici_adi FROM kullanicilar WHERE kullanici_adi=%s", (yeni_k_adi,))
                            if cur.fetchone(): st.error("Bu ID zaten mevcut!")
                            else:
                                cur.execute("INSERT INTO kullanicilar (kullanici_adi, sifre, rol, isim) VALUES (%s, %s, %s, %s)", (yeni_k_adi, hash_sifre(yeni_sifre), yeni_rol, yeni_isim))
                                conn.commit(); st.success(f"{yeni_k_adi} eklendi!"); st.rerun()
                        else: st.warning("Tüm alanları doldurunuz.")
                
                with st.expander("✏️ Mevcut Kullanıcıyı Düzenle", expanded=False):
                    if not df_kullanicilar.empty:
                        duzenlenecek_kullanici = st.selectbox("Düzenlenecek Kullanıcı", df_kullanicilar['kullanici_adi'].tolist())
                        mevcut_bilgi = df_kullanicilar[df_kullanicilar['kullanici_adi'] == duzenlenecek_kullanici].iloc[0]
                        
                        c_e1, c_e2 = st.columns(2)
                        with c_e1: 
                            edit_id = st.text_input("Kullanıcı Adı (ID)", value=mevcut_bilgi['kullanici_adi'], key=f"e_id_{duzenlenecek_kullanici}")
                            edit_pw = st.text_input("Yeni Şifre (Değişmeyecekse BOŞ Bırakın)", type="password", key=f"e_pw_{duzenlenecek_kullanici}")
                        with c_e2:
                            roller = ["Öğretim Elemanı", "Testing", "Admin"]
                            edit_rol = st.selectbox("Rol", roller, index=roller.index(mevcut_bilgi['rol']) if mevcut_bilgi['rol'] in roller else 0, key=f"e_rol_{duzenlenecek_kullanici}")
                            edit_isim = st.text_input("İsim", value=mevcut_bilgi['isim'], key=f"e_isim_{duzenlenecek_kullanici}")
                        
                        if st.button("Bilgileri Güncelle"):
                            cur = conn.cursor()
                            if edit_id != duzenlenecek_kullanici:
                                cur.execute("SELECT kullanici_adi FROM kullanicilar WHERE kullanici_adi=%s", (edit_id,))
                                if cur.fetchone(): st.error("Bu ID zaten kullanımda!"); st.stop()
                                cur.execute("UPDATE hoca_atamalari SET kullanici_adi=%s WHERE kullanici_adi=%s", (edit_id, duzenlenecek_kullanici))
                                cur.execute("UPDATE kullanicilar SET kullanici_adi=%s, rol=%s, isim=%s WHERE kullanici_adi=%s", (edit_id, edit_rol, edit_isim, duzenlenecek_kullanici))
                            else:
                                cur.execute("UPDATE kullanicilar SET rol=%s, isim=%s WHERE kullanici_adi=%s", (edit_rol, edit_isim, duzenlenecek_kullanici))
                            if edit_pw: cur.execute("UPDATE kullanicilar SET sifre=%s WHERE kullanici_adi=%s", (hash_sifre(edit_pw), edit_id))
                            conn.commit(); st.success("Kullanıcı güncellendi!"); st.rerun()
                
                st.markdown("**Sistemdeki Kullanıcılar**")
                st.dataframe(df_kullanicilar, use_container_width=True, hide_index=True)
                sil_k_adi = st.selectbox("Silinecek Kullanıcı", df_kullanicilar['kullanici_adi'].tolist() if not df_kullanicilar.empty else [])
                if st.button("Kullanıcıyı Sil") and sil_k_adi:
                    if sil_k_adi == st.session_state['kullanici_adi']: st.error("Kendinizi silemezsiniz!")
                    else: conn.cursor().execute("DELETE FROM kullanicilar WHERE kullanici_adi=%s", (sil_k_adi,)); conn.commit(); st.rerun()
                conn.close()
                
            with tab_atamalar:
                conn = get_db_connection()
                st.markdown("**Sınıf - Hoca Ataması Yap**")
                
                df_hocalar = pd.read_sql_query("SELECT kullanici_adi, isim FROM kullanicilar WHERE rol != 'Admin'", conn)
                hoca_dict = dict(zip(df_hocalar.kullanici_adi, df_hocalar.isim)) if not df_hocalar.empty else {}
                
                with st.form("yeni_atama_formu"):
                    c_f1, c_f2, c_f3 = st.columns(3)
                    yeni_a_modul = c_f1.selectbox("Modül", ["M1", "M2", "M3", "M4", "YazOkulu", "PF"])
                    yeni_a_sube = c_f2.text_input("Şube (Örn: C14, A21)")
                    if hoca_dict:
                        yeni_a_hoca = c_f3.selectbox("Öğretim Elemanı", list(hoca_dict.keys()), format_func=lambda x: hoca_dict[x])
                    else:
                        yeni_a_hoca = c_f3.selectbox("Öğretim Elemanı", ["Kayıtlı Hoca Yok"])
                    
                    if st.form_submit_button("Atamayı Ekle"):
                        if yeni_a_sube and hoca_dict:
                            cur = conn.cursor()
                            cur.execute("INSERT INTO hoca_atamalari (modul, sube, kullanici_adi) VALUES (%s, %s, %s)", (yeni_a_modul, yeni_a_sube.upper().strip(), yeni_a_hoca))
                            conn.commit()
                            st.success("Atama başarıyla eklendi!")
                            st.rerun()
                        else:
                            st.warning("Lütfen şube adını giriniz ve hoca seçtiğinizden emin olunuz.")
                
                st.markdown("**Mevcut Sınıf Atamaları (Silmek için butona basınız)**")
                mevcut_atamalar = pd.read_sql_query("SELECT modul, sube, kullanici_adi FROM hoca_atamalari ORDER BY modul, sube", conn)
                if not mevcut_atamalar.empty:
                    for _, row in mevcut_atamalar.iterrows():
                        c_a1, c_a2 = st.columns([4, 1])
                        c_a1.write(f"🔹 **{row['modul']} - {row['sube']}** : {hoca_dict.get(row['kullanici_adi'], row['kullanici_adi'])}")
                        if c_a2.button("Atamayı Sil", key=f"del_ata_{row['modul']}_{row['sube']}_{row['kullanici_adi']}"):
                            conn.cursor().execute("DELETE FROM hoca_atamalari WHERE modul=%s AND sube=%s AND kullanici_adi=%s", (row['modul'], row['sube'], row['kullanici_adi']))
                            conn.commit()
                            st.rerun()
                else: st.info("Henüz yapılmış bir sınıf ataması yok.")
                
                st.divider()
                st.markdown("**Öğrenci Kütüğü Yükleme Terminali**")
                duzenlenen_ogr = st.data_editor(pd.read_sql_query("SELECT ogrenci_no, ad_soyad, modul, sube FROM ogrenciler", conn), num_rows="dynamic", use_container_width=True)
                
                if st.button("Öğrenci Kütüğünü Kaydet"):
                    cur = conn.cursor(); cur.execute("DELETE FROM ogrenciler")
                    ogr_data = [(r['ogrenci_no'], r['ad_soyad'], r['modul'], r['sube']) for _, r in duzenlenen_ogr.iterrows()]
                    if ogr_data:
                        cur.executemany("INSERT INTO ogrenciler (ogrenci_no, ad_soyad, modul, sube) VALUES (%s, %s, %s, %s)", ogr_data)
                    conn.commit(); st.success("Kütük güncellendi!")
                
                conn.close()
                
        # ----------------------------------------------------
        # 4. GELİŞMİŞ ANALİTİK VE KARNE RAPORLARI
        # ----------------------------------------------------
        with tab_rapor:
            st.markdown("### 🔬 Gelişmiş Ölçme & Değerlendirme Analitiği")
            rapor_kapsami = st.radio("Analiz Türü:", [
                "🏆 Dönem Sonu Konsolide Başarı", 
                "🎯 Madde Ayırt Edicilik (Item Discrimination)", 
                "⚖️ Puanlayıcı Güvenirliği (Inter-Rater Reliability)",
                "👤 Bireysel Karneler",
                "🏫 Sınıf / Şube Raporu (Excel Formatı)",
                "📡 Canlı Sınav Takip Radarı",
                "📥 OBS Excel Çıktısı"
            ], horizontal=True)
            st.divider()
            
            conn = get_db_connection()
            cur = conn.cursor()
            
            if rapor_kapsami == "📡 Canlı Sınav Takip Radarı":
                cur.execute("SELECT sinav_kodu FROM sinavlar")
                sinavlar = [r[0] for r in cur.fetchall()]
                if not sinavlar: st.warning("Kayıtlı sınav yok.")
                else:
                    sec_sinav = st.selectbox("Takip Edilecek Sınavı Seçin", sinavlar)
                    ilgili_modul = sec_sinav.split("_")[0]
                    
                    cur.execute("SELECT h.sube, h.kullanici_adi, k.isim FROM hoca_atamalari h LEFT JOIN kullanicilar k ON h.kullanici_adi = k.kullanici_adi WHERE h.modul=%s", (ilgili_modul,))
                    hoca_atamalari = cur.fetchall()
                    
                    if not hoca_atamalari: st.info("Bu modül için atanmış sınıf/hoca bulunmuyor.")
                    else:
                        st.markdown(f"#### 📡 {sec_sinav} - Veri Giriş Durumları")
                        for sube, k_adi, hoca_isim in sorted(hoca_atamalari, key=lambda x: x[0]):
                            cur.execute("SELECT COUNT(*) FROM ogrenciler WHERE modul=%s AND sube=%s", (ilgili_modul, sube))
                            toplam_ogrenci = cur.fetchone()[0]
                            
                            cur.execute("SELECT veri_json FROM notlar WHERE sinav_kodu=%s AND sube=%s", (sec_sinav, sube))
                            not_verisi = cur.fetchone()
                            
                            okunan_ogrenci = 0
                            if not_verisi:
                                df_not = pd.read_json(io.StringIO(not_verisi[0]))
                                if "Sınav Notu" in df_not.columns:
                                    okunan_ogrenci = len(df_not[df_not["Sınav Notu"] > 0])
                            
                            if toplam_ogrenci > 0: yuzde = min(int((okunan_ogrenci / toplam_ogrenci) * 100), 100)
                            else: yuzde = 0
                                
                            col_text, col_bar = st.columns([1, 3])
                            col_text.write(f"**{sube}** ({hoca_isim or k_adi})")
                            col_bar.progress(yuzde, text=f"{okunan_ogrenci} / {toplam_ogrenci} Öğrenci Okundu (%{yuzde})")
                            st.write("")

            elif rapor_kapsami == "📥 OBS Excel Çıktısı":
                cur.execute("SELECT sinav_kodu FROM sinavlar")
                sinavlar = [r[0] for r in cur.fetchall()]
                if not sinavlar: st.warning("Kayıtlı sınav yok.")
                else:
                    sec_sinav = st.selectbox("OBS'ye Aktarılacak Sınavı Seçin", sinavlar)
                    if st.button("🚀 OBS Formatında Excel Üret", type="primary"):
                        cur.execute("SELECT sube, veri_json FROM notlar WHERE sinav_kodu=%s", (sec_sinav,))
                        not_verileri = cur.fetchall()
                        if not not_verileri: st.error("Bu sınava ait girilmiş hiçbir not bulunamadı!")
                        else:
                            obs_listesi = []
                            for sube, veri_json in not_verileri:
                                df_temp = pd.read_json(io.StringIO(veri_json))
                                if "Ogrenci_No" in df_temp.columns and "Sınav Notu" in df_temp.columns:
                                    for _, row in df_temp.iterrows():
                                        ogr_no = str(row["Ogrenci_No"]).strip()
                                        notu = row["Sınav Notu"]
                                        if ogr_no and ogr_no.lower() != "nan":
                                            obs_listesi.append({"Öğrenci No": ogr_no, "Notu": notu})
                                            
                            if obs_listesi:
                                df_obs = pd.DataFrame(obs_listesi).sort_values(by="Öğrenci No")
                                buffer = io.BytesIO()
                                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                                    df_obs.to_excel(writer, index=False, sheet_name="OBS_Aktarim")
                                buffer.seek(0)
                                st.success("OBS formatı hazır!")
                                st.download_button("📥 Excel Dosyasını İndir", data=buffer, file_name=f"OBS_{sec_sinav}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                            else:
                                st.warning("Excel formatına uygun öğrenci/not kaydı bulunamadı.")
            
            elif rapor_kapsami == "🏆 Dönem Sonu Konsolide Başarı":
                cur.execute("SELECT sinav_kodu FROM sinavlar")
                acik_sinavlar = [r[0] for r in cur.fetchall()]
                if not acik_sinavlar: st.warning("Sistemde sınav yok.")
                else:
                    secilen_sinavlar_konsolide = st.multiselect("Hesaplamaya Dahil Edilecek Sınavları Seçin", acik_sinavlar)
                    if secilen_sinavlar_konsolide:
                        c_weights = st.columns(len(secilen_sinavlar_konsolide))
                        agirliklar = {}; toplam_agirlik = 0
                        for i, sk in enumerate(secilen_sinavlar_konsolide):
                            with c_weights[i]:
                                w = st.number_input(f"{sk} Ağırlığı (%)", min_value=0, max_value=100, value=100//len(secilen_sinavlar_konsolide))
                                agirliklar[sk] = w; toplam_agirlik += w
                        if toplam_agirlik != 100: st.error(f"⚠️ Toplam ağırlık 100 olmalıdır! (Şu an: {toplam_agirlik})")
                        else:
                            gecme_notu = st.number_input("Geçme Barajı", min_value=0, max_value=100, value=65)
                            if st.button("Hesapla", type="primary"):
                                cur.execute("SELECT sinav_kodu, sube, veri_json FROM notlar WHERE sinav_kodu = ANY(%s)", (secilen_sinavlar_konsolide,))
                                ogr_dict = {}
                                for s_kodu, s_sube, v_json in cur.fetchall():
                                    df_temp = pd.read_json(io.StringIO(v_json))
                                    if "Sınav Notu" in df_temp.columns:
                                        for _, r in df_temp.iterrows():
                                            o_no = str(r.get("Ogrenci_No", ""))
                                            if not o_no: continue
                                            if o_no not in ogr_dict: ogr_dict[o_no] = {"Öğrenci No": o_no, "Ad Soyad": r.get("Ad_Soyad", ""), "Şube": s_sube}
                                            ogr_dict[o_no][s_kodu] = r["Sınav Notu"]
                                if not ogr_dict: st.warning("Not bulunamadı.")
                                else:
                                    sonuclar = []
                                    for o_no, data in ogr_dict.items():
                                        nihai_puan = sum(data.get(sk, 0) * (agirliklar[sk] / 100) for sk in secilen_sinavlar_konsolide)
                                        for sk in secilen_sinavlar_konsolide:
                                            if sk not in data: data[sk] = 0
                                        data["Nihai Puan"] = round(nihai_puan, 2)
                                        data["Durum"] = "BAŞARILI" if nihai_puan >= gecme_notu else "BAŞARISIZ"
                                        sonuclar.append(data)
                                    df_sonuc = pd.DataFrame(sonuclar).sort_values(by="Nihai Puan", ascending=False)
                                    try: styled_df = df_sonuc.style.map(lambda x: 'background-color: #D5F5E3' if x=='BAŞARILI' else 'background-color: #FADBD8' if x=='BAŞARISIZ' else '', subset=['Durum'])
                                    except: styled_df = df_sonuc.style.applymap(lambda x: 'background-color: #D5F5E3' if x=='BAŞARILI' else 'background-color: #FADBD8' if x=='BAŞARISIZ' else '', subset=['Durum'])
                                    st.dataframe(styled_df, use_container_width=True, hide_index=True)
            
            elif rapor_kapsami == "🎯 Madde Ayırt Edicilik (Item Discrimination)":
                cur.execute("SELECT sinav_kodu FROM sinavlar")
                sinavlar = [r[0] for r in cur.fetchall()]
                if not sinavlar: st.warning("Sınav bulunamadı.")
                else:
                    secilen_sinav = st.selectbox("Analiz Edilecek Sınav", sinavlar)
                    if st.button("Analizi Başlat", type="primary"):
                        df_tos = pd.read_sql_query("SELECT ana_bolum, alt_bolum, max_puan FROM sinav_tos WHERE sinav_kodu=%s", conn, params=(secilen_sinav,))
                        max_puanlar = {f"{r['ana_bolum']} - {r['alt_bolum']}": float(r['max_puan']) for _, r in df_tos.iterrows()}
                        cur.execute("SELECT veri_json FROM notlar WHERE sinav_kodu=%s", (secilen_sinav,))
                        tum_ogrenciler = []
                        for j in cur.fetchall(): tum_ogrenciler.extend(pd.read_json(io.StringIO(j[0])).to_dict('records'))
                        df_hepsi = pd.DataFrame(tum_ogrenciler)
                        if "Sınava Girdi" in df_hepsi.columns: df_hepsi = df_hepsi[df_hepsi["Sınava Girdi"] == True]
                        if "Sınav Notu" not in df_hepsi.columns or len(df_hepsi) < 10: st.warning("Aktif katılımcı verisi yetersiz (En az 10 öğrenci).")
                        else:
                            df_hepsi = df_hepsi.sort_values(by="Sınav Notu", ascending=False)
                            n_27 = max(1, int(len(df_hepsi) * 0.27))
                            df_ust, df_alt = df_hepsi.head(n_27), df_hepsi.tail(n_27)
                            sonuclar = []
                            for bolum, max_p in max_puanlar.items():
                                if bolum in df_hepsi.columns:
                                    ust_ortalama = df_ust[bolum].mean()
                                    alt_ortalama = df_alt[bolum].mean()
                                    d_index = (ust_ortalama / max_p if max_p > 0 else 0) - (alt_ortalama / max_p if max_p > 0 else 0)
                                    durum = "Çok İyi Madde" if d_index >= 0.40 else "İyi Madde" if d_index >= 0.30 else "Zayıf Madde" if d_index >= 0.20 else "Çöpe Atılmalı"
                                    sonuclar.append({"Bölüm": bolum, "Max": max_p, "Üst %27": round(ust_ortalama, 2), "Alt %27": round(alt_ortalama, 2), "D İndeksi": round(d_index, 3), "Yorum": durum})
                            df_s = pd.DataFrame(sonuclar)
                            if not df_s.empty:
                                def renk_ver(val):
                                    if not isinstance(val, str): return ''
                                    if "Çok İyi" in val: return 'color: green; font-weight: bold;'
                                    if "Çöpe" in val: return 'background-color: #FADBD8; color: black;'
                                    return ''
                                try: styler = df_s.style.map(renk_ver, subset=['Yorum'])
                                except AttributeError: styler = df_s.style.applymap(renk_ver, subset=['Yorum'])
                                st.dataframe(styler, use_container_width=True)

            elif rapor_kapsami == "⚖️ Puanlayıcı Güvenirliği (Inter-Rater Reliability)":
                cur.execute("SELECT sinav_kodu FROM sinavlar")
                sinavlar = [r[0] for r in cur.fetchall()]
                if not sinavlar: st.warning("Sınav bulunamadı.")
                else:
                    secilen_sinav_ir = st.selectbox("Sınav Seç", sinavlar)
                    if st.button("Güvenirlik Analizi Çıkar", type="primary"):
                        cur.execute("SELECT veri_json FROM notlar WHERE sinav_kodu=%s", (secilen_sinav_ir,))
                        tum_ogrenciler = []
                        for j in cur.fetchall(): tum_ogrenciler.extend(pd.read_json(io.StringIO(j[0])).to_dict('records'))
                        df_ir = pd.DataFrame(tum_ogrenciler)
                        if "Sınava Girdi" in df_ir.columns: df_ir = df_ir[df_ir["Sınava Girdi"] == True]
                        okuma_1_kolonlari = [c for c in df_ir.columns if "(1. Okuma)" in c]
                        if not okuma_1_kolonlari: st.warning("'(1. Okuma)' etiketi taşıyan sütun bulunamadı.")
                        else:
                            for ok1_col in okuma_1_kolonlari:
                                ok2_col = ok1_col.replace("(1. Okuma)", "(2. Okuma)")
                                if ok2_col in df_ir.columns:
                                    st.markdown(f"#### 📊 Beceri: {ok1_col.replace(' - (1. Okuma)', '')}")
                                    df_cift = df_ir[['Ad_Soyad', ok1_col, ok2_col]].copy().dropna()
                                    df_cift['Puan Farkı'] = abs(df_cift[ok1_col] - df_cift[ok2_col])
                                    if len(df_cift) > 2:
                                        r_degeri = df_cift[ok1_col].corr(df_cift[ok2_col])
                                        c1, c2, c3 = st.columns(3)
                                        c1.metric("Korelasyon (r)", f"{r_degeri:.2f}")
                                        c2.metric("1. Okuma Ort.", f"{df_cift[ok1_col].mean():.2f}")
                                        c3.metric("2. Okuma Ort.", f"{df_cift[ok2_col].mean():.2f}")
                                        st.write("**⚠️ Yüksek Sapma (Fark > 10):**")
                                        st.dataframe(df_cift[df_cift['Puan Farkı'] > 10].sort_values(by="Puan Farkı", ascending=False), use_container_width=True)
                                        st.divider()
            
            elif rapor_kapsami == "👤 Bireysel Karneler":
                cur.execute("SELECT sinav_kodu FROM sinavlar")
                sinavlar = [r[0] for r in cur.fetchall()]
                if not sinavlar: st.warning("Kayıtlı sınav yok.")
                else:
                    sec_sinav = st.selectbox("Sınav Seçin", sinavlar)
                    cur.execute("SELECT sube FROM notlar WHERE sinav_kodu=%s", (sec_sinav,))
                    subeler = [r[0] for r in cur.fetchall()]
                    if not subeler: st.warning("Bu sınava ait girilmiş not yok.")
                    else:
                        sec_sube = st.selectbox("Şube Seçin", subeler)
                        
                        cur.execute("SELECT tolerans FROM sinav_ayarlari WHERE sinav_kodu=%s", (sec_sinav,))
                        tol_res = cur.fetchone()
                        sinav_tol = int(tol_res[0]) if (tol_res and tol_res[0] is not None) else 6
                        
                        cur.execute("SELECT ana_bolum, alt_bolum, kazanimlar, max_puan FROM sinav_tos WHERE sinav_kodu=%s", (sec_sinav,))
                        tos_df = pd.DataFrame(cur.fetchall(), columns=["ana_bolum", "alt_bolum", "kazanimlar", "max_puan"])
                        
                        cur.execute("SELECT veri_json FROM notlar WHERE sinav_kodu=%s AND sube=%s", (sec_sinav, sec_sube))
                        df_k = pd.read_json(io.StringIO(cur.fetchone()[0]))
                        
                        if not df_k.empty:
                            if "Sınava Girdi" not in df_k.columns: df_k["Sınava Girdi"] = True
                            df_k_aktif = df_k[df_k["Sınava Girdi"] == True].copy()
                            sinif_genel_ort = pd.to_numeric(df_k_aktif["Sınav Notu"], errors='coerce').fillna(0.0).mean() if not df_k_aktif.empty else 0.0
                            
                            sec_ogr = st.selectbox("Öğrenci Seçin", df_k['Ad_Soyad'].tolist())
                            ogr_veri = df_k[df_k['Ad_Soyad'] == sec_ogr].iloc[0]
                            
                            bkl = ogr_veri.get('Kitapçık', 'A')
                            st.markdown(f"#### 🎓 Öğrenci: {sec_ogr} (Toplam: {ogr_veri.get('Sınav Notu', 0)} | Kitapçık: {bkl})")
                            
                            katildi_mi = bool(ogr_veri.get('Sınava Girdi', True))
                            if not katildi_mi:
                                st.error("⚠️ Bu öğrenci sınava katılmamıştır. (Absent)")
                            else:
                                satirlar = karne_satirlari_hazirla(tos_df, df_k_aktif, ogr_veri, sinav_tol)
                                for s in satirlar:
                                    st.write(f"**{s['bolum']}** | Kazanım: *{s['kazanim']}* ({s['puan']:g} / {s['max']:g} | Sınıf Ort: {s['sinif_ort']:.1f})")
                                    st.progress(min(int(s['yuzde']), 100))
                                
                                if st.button("📄 PDF Karne Oluştur", type="primary"):
                                    r_veri = [{
                                        'sinav_kodu': sec_sinav,
                                        'ogr_no': str(ogr_veri.get('Ogrenci_No', '')).strip(),
                                        'ogr_ad': sec_ogr,
                                        'sube': sec_sube,
                                        'kitapcik': bkl,
                                        'toplam_not': ogr_veri.get('Sınav Notu', 0),
                                        'sinif_genel_ort': sinif_genel_ort,
                                        'satirlar': satirlar
                                    }]
                                    pdf_buf = create_modern_pdf(r_veri)
                                    if pdf_buf:
                                        st.success("PDF başarıyla oluşturuldu!")
                                        st.download_button("📥 PDF Olarak İndir", data=pdf_buf, file_name=f"{tr_to_eng(sec_ogr).replace(' ', '_')}_ReportCard.pdf", mime="application/pdf")
            
            elif rapor_kapsami == "🏫 Sınıf / Şube Raporu (Excel Formatı)":
                cur.execute("SELECT sinav_kodu FROM sinavlar")
                sinavlar = [r[0] for r in cur.fetchall()]
                if not sinavlar: st.warning("Kayıtlı sınav yok.")
                else:
                    sec_sinav = st.selectbox("Sınav Seçin", sinavlar, key="sr_sinav")
                    cur.execute("SELECT sube FROM notlar WHERE sinav_kodu=%s", (sec_sinav,))
                    subeler = [r[0] for r in cur.fetchall()]
                    if not subeler: st.warning("Bu sınava ait not yok.")
                    else:
                        sec_sube = st.selectbox("Şube Seçin", subeler, key="sr_sube")
                        
                        cur.execute("SELECT ana_bolum, alt_bolum, kazanimlar, max_puan FROM sinav_tos WHERE sinav_kodu=%s", (sec_sinav,))
                        tos_df = pd.DataFrame(cur.fetchall(), columns=["ana_bolum", "alt_bolum", "kazanimlar", "max_puan"])
                        tos_dict = {f"{row['ana_bolum']} - {row['alt_bolum']}": {"max": float(row['max_puan']), "kazanim": row['kazanimlar']} for _, row in tos_df.iterrows()}
                        
                        cur.execute("SELECT veri_json FROM notlar WHERE sinav_kodu=%s AND sube=%s", (sec_sinav, sec_sube))
                        df_s = pd.read_json(io.StringIO(cur.fetchone()[0]))
                        
                        if "Sınava Girdi" not in df_s.columns: df_s["Sınava Girdi"] = True
                        if "Kitapçık" not in df_s.columns: df_s["Kitapçık"] = "A"
                        df_aktif = df_s[df_s["Sınava Girdi"] == True].copy()
                        
                        for c in tos_dict.keys():
                            if c in df_aktif.columns:
                                df_aktif[c] = pd.to_numeric(df_aktif[c].astype(str).str.replace(',', '.'), errors='coerce').fillna(0.0)
                        
                        st.markdown(f"### {sec_sube} - Genel Sınıf İstatistikleri")
                        c1, c2, c3, c4 = st.columns(4)
                        c1.metric("Sınıf Ortalaması", f"{df_aktif['Sınav Notu'].mean():.2f}" if not df_aktif.empty else "0.00")
                        c2.metric("Medyan", f"{df_aktif['Sınav Notu'].median():.2f}" if not df_aktif.empty else "0.00")
                        c3.metric("En Yüksek", f"{df_aktif['Sınav Notu'].max():.2f}" if not df_aktif.empty else "0.00")
                        c4.metric("En Düşük", f"{df_aktif['Sınav Notu'].min():.2f}" if not df_aktif.empty else "0.00")
                        
                        st.markdown("#### Beceri Bazlı Sınıf Ortalamaları")
                        beceri_listesi_pdf = []
                        for col, veri_dict in tos_dict.items():
                            if col in df_aktif.columns:
                                ort = df_aktif[col].mean() if not df_aktif.empty else 0
                                m_puan = veri_dict['max']
                                yuzde = (ort / m_puan * 100) if m_puan > 0 else 0
                                beceri_listesi_pdf.append({"isim": col, "kazanim": veri_dict['kazanim'], "ort": ort, "max": m_puan, "yuzde": yuzde})
                                st.write(f"**{col}** | *{veri_dict['kazanim']}* (Sınıf Ort: {ort:.1f} / {m_puan})")
                                st.progress(min(int(yuzde), 100))
                                
                        st.markdown("#### Öğrenci Detay Tablosu")
                        cols_to_show = ["Ogrenci_No", "Ad_Soyad", "Kitapçık", "Sınava Girdi", "Sınav Notu"] + [c for c in tos_dict.keys() if c in df_s.columns]
                        st.dataframe(df_s[cols_to_show].sort_values(by="Sınav Notu", ascending=False), use_container_width=True, hide_index=True)
                        
                        c_btn1, c_btn2 = st.columns(2)
                        with c_btn1:
                            if st.button("📊 Sınıf Raporunu Excel Olarak İndir (Tek Sayfa - Kurumsal)", type="primary"):
                                buffer = create_corporate_class_excel(sec_sinav, sec_sube, df_s, df_aktif, cols_to_show, beceri_listesi_pdf)
                                st.success("Kurumsal tek sayfalık Excel sınıf raporu hazırlandı!")
                                st.download_button("📥 Excel Raporunu İndir", data=buffer, file_name=f"{sec_sube}_{sec_sinav}_Sinif_Raporu.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                        
                        with c_btn2:
                            if st.button("📦 Tüm Sınıfın Karnelerini Oluştur (Ayrı PDF - ZIP)"):
                                if df_aktif.empty:
                                    st.warning("Sınava giren öğrenci bulunmuyor.")
                                else:
                                    cur.execute("SELECT tolerans FROM sinav_ayarlari WHERE sinav_kodu=%s", (sec_sinav,))
                                    tol_res = cur.fetchone()
                                    sinav_tol = int(tol_res[0]) if (tol_res and tol_res[0] is not None) else 6
                                    sinif_genel_ort = pd.to_numeric(df_aktif["Sınav Notu"], errors='coerce').fillna(0.0).mean() if not df_aktif.empty else 0.0
                                    
                                    zip_buffer = io.BytesIO()
                                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                                        for _, row in df_aktif.sort_values(by="Ad_Soyad").iterrows():
                                            ogr_no = str(row.get('Ogrenci_No', '')).strip()
                                            ogr_ad = str(row.get('Ad_Soyad', '')).strip()
                                            toplam = row.get('Sınav Notu', 0)
                                            bkl = row.get('Kitapçık', 'A')
                                            satirlar = karne_satirlari_hazirla(tos_df, df_aktif, row, sinav_tol)
                                            
                                            tek_veri = [{
                                                'sinav_kodu': sec_sinav,
                                                'ogr_no': ogr_no,
                                                'ogr_ad': ogr_ad,
                                                'sube': sec_sube,
                                                'kitapcik': bkl,
                                                'toplam_not': toplam,
                                                'sinif_genel_ort': sinif_genel_ort,
                                                'satirlar': satirlar
                                            }]
                                            pdf_buf = create_modern_pdf(tek_veri)
                                            temiz_ad = tr_to_eng(ogr_ad).replace(" ", "_")
                                            dosya_adi = f"{ogr_no}_{temiz_ad}_ReportCard.pdf"
                                            zf.writestr(dosya_adi, pdf_buf.getvalue())
                                            
                                    zip_buffer.seek(0)
                                    st.success(f"{len(df_aktif)} öğrencinin karnesi ayrı PDF'ler halinde paketlendi!")
                                    st.download_button(
                                        "📥 Tüm Karneleri İndir (ZIP)",
                                        data=zip_buffer,
                                        file_name=f"{sec_sube}_{sec_sinav}_Karneler.zip",
                                        mime="application/zip"
                                    )

            cur.close(); conn.close()

    # ----------------------------------------------------
    # 5. HOCA VERİ GİRİŞ PANELİ (SEÇENEK C + PINNED SÜTUN)
    # ----------------------------------------------------
    elif st.session_state['rol'] == "Öğretim Elemanı":
        conn = get_db_connection()
        cur = conn.cursor()
        
        cur.execute("SELECT modul FROM hoca_atamalari WHERE kullanici_adi=%s", (st.session_state['kullanici_adi'],))
        hocanin_modulleri = set([r[0] for r in cur.fetchall()])
        
        cur.execute("SELECT sinav_kodu, kilitli FROM sinavlar")
        sinav_data = cur.fetchall()
        if not sinav_data: st.warning("Sistemde açık sınav bulunmuyor."); st.stop()
        
        acik_sinavlar = [row[0] for row in sinav_data if row[0].split("_")[0] in hocanin_modulleri]
        if not acik_sinavlar: 
            st.error("⚠️ Size tanımlanmış herhangi bir modüle ait aktif sınav bulunmamaktadır.")
            st.stop()
            
        sinav_kilit_durumlari = {row[0]: row[1] for row in sinav_data}
        
        col_sel1, col_sel2 = st.columns([2, 1])
        with col_sel1:
            secilen_islem_sinavi = st.selectbox("Sorumlu Olduğunuz Sınav", acik_sinavlar)
        is_locked = sinav_kilit_durumlari.get(secilen_islem_sinavi, False)
        
        ilgili_modul = secilen_islem_sinavi.split("_")[0]
        cur.execute("SELECT sube FROM hoca_atamalari WHERE modul=%s AND kullanici_adi=%s", (ilgili_modul, st.session_state['kullanici_adi']))
        hoca_subeleri = [row[0] for row in cur.fetchall()]
        if not hoca_subeleri: st.error("Bu sınavın modülünde size atanmış bir şube bulunmuyor."); st.stop()
        
        with col_sel2:
            secilen_sube = st.selectbox("Şube", hoca_subeleri)

        # Seçenek C: Kompakt Profesyonel Üst Bant
        sinav_temiz_baslik = secilen_islem_sinavi.replace("_", " ")
        durum_badge = '<span class="c-badge-locked">🔒 KİLİTLİ (Salt Okunur)</span>' if is_locked else '<span class="c-badge-open">● Not Girişine Açık</span>'
        st.markdown(f"""
        <div class="c-header-bar">
            <div>
                <div class="c-header-title">{sinav_temiz_baslik}  —  {secilen_sube} Şubesi</div>
                <div class="c-header-sub">Department of Foreign Languages | Hızlı Veri Giriş Terminali</div>
            </div>
            <div>{durum_badge}</div>
        </div>
        """, unsafe_allow_html=True)
            
        cur.execute("SELECT tolerans, kitapcik_turu FROM sinav_ayarlari WHERE sinav_kodu=%s", (secilen_islem_sinavi,))
        ayarlar_res = cur.fetchone()
        tolerans = int(ayarlar_res[0]) if (ayarlar_res and ayarlar_res[0] is not None) else 6
        kitapcik_turu = ayarlar_res[1] if (ayarlar_res and ayarlar_res[1] is not None) else "Çoklu Kitapçık (A ve B)"
        is_multi_booklet = "Çoklu" in kitapcik_turu
        
        try:
            rubrik_df = pd.read_sql_query("SELECT ana_bolum, alt_bolum, max_puan, b_grubu_karsiligi FROM sinav_tos WHERE sinav_kodu=%s", conn, params=(secilen_islem_sinavi,))
        except:
            rubrik_df = pd.read_sql_query("SELECT ana_bolum, alt_bolum, max_puan FROM sinav_tos WHERE sinav_kodu=%s", conn, params=(secilen_islem_sinavi,))
            rubrik_df['b_grubu_karsiligi'] = ""
            
        max_puan_sozlugu = {f"{r['ana_bolum']} - {r['alt_bolum']}": float(r['max_puan']) for _, r in rubrik_df.iterrows()}
        
        isaretler = [("1. Okuma", "2. Okuma", "3. Okuma"), ("(1. Okuma)", "(2. Okuma)", "(3. Okuma)"), ("1st", "2nd", "3rd"), ("1. Hoca", "2. Hoca", "3. Hoca"), ("(1)", "(2)", "(3)"), ("Okuma 1", "Okuma 2", "Okuma 3"), ("Rater 1", "Rater 2", "Rater 3"), (" 1 ", " 2 ", " 3 "), (" 1", " 2", " 3"), ("1.", "2.", "3.")]
        
        cift_okumalar = []
        islenenler = set()
        
        for c1 in max_puan_sozlugu.keys():
            if c1 in islenenler: continue
            c2_match, c3_match = None, None
            for m1, m2, m3 in isaretler:
                if m1 in c1:
                    hedef_c2 = c1.replace(m1, m2)
                    hedef_c3 = c1.replace(m1, m3)
                    if hedef_c2 in max_puan_sozlugu:
                        c2_match = hedef_c2
                        if hedef_c3 in max_puan_sozlugu: c3_match = hedef_c3
                        break
            if c2_match:
                cift_okumalar.append((c1, c2_match, c3_match))
                islenenler.update([c1, c2_match])
                if c3_match: islenenler.add(c3_match)
                
        ort_mesaj = None
        if cift_okumalar:
            eslesen_isimler = []
            for c1, c2, c3 in cift_okumalar:
                isim = f"**{c1.split('-')[-1].strip()}** & **{c2.split('-')[-1].strip()}**"
                if c3: isim += f" *(Gerekirse {c3.split('-')[-1].strip()})*"
                eslesen_isimler.append(isim)
            ort_mesaj = "🤖 **Sistem Zekası Aktif:** Şu okumalar otomatik olarak eşleştirildi ve ortalamaları alınacaktır: " + " | ".join(eslesen_isimler)

        safe_b_map = {}
        reverse_b_map = {}
        for _, row in rubrik_df.iterrows():
            a_name = f"{row['ana_bolum']} - {row['alt_bolum']}"
            b_karsi_raw = str(row.get('b_grubu_karsiligi', '')).strip()
            
            if b_karsi_raw.lower() in ['nan', 'none', '']:
                safe_b_map[a_name] = a_name
                reverse_b_map[a_name] = a_name
            elif is_multi_booklet:
                b_name = f"{row['ana_bolum']} - {b_karsi_raw}"
                safe_b_map[a_name] = b_name
                reverse_b_map[b_name] = a_name
            else:
                safe_b_map[a_name] = a_name
                reverse_b_map[a_name] = a_name

        cur.execute("SELECT veri_json FROM notlar WHERE sinav_kodu=%s AND sube=%s", (secilen_islem_sinavi, secilen_sube))
        kaydedilmis_veri = cur.fetchone()
        
        current_tos_cols = list(dict.fromkeys([f"{row['ana_bolum']} - {row['alt_bolum']}" for _, row in rubrik_df.iterrows()]))
        base_cols = ["Ogrenci_No", "Ad_Soyad", "Sınava Girdi", "Kitapçık", "Sınav Notu"]
        
        if kaydedilmis_veri: 
            df_giris = pd.read_json(io.StringIO(kaydedilmis_veri[0]))
            df_giris.columns = [str(c).strip() for c in df_giris.columns]
            df_giris = df_giris.loc[:, ~df_giris.columns.duplicated()]
            
            if "Sınava Girdi" not in df_giris.columns: df_giris.insert(2, "Sınava Girdi", True)
            if "Kitapçık" not in df_giris.columns: df_giris.insert(3, "Kitapçık", "A")
            
            old_cols = [c for c in df_giris.columns if c not in current_tos_cols and c not in base_cols]
            if old_cols:
                for nc in current_tos_cols:
                    if nc not in df_giris.columns: df_giris[nc] = ""
                old_speak = sorted([c for c in old_cols if 'speaking' in c.lower() and c.lower() != 'nan'])
                new_speak = sorted([c for c in current_tos_cols if 'speaking' in c.lower()])
                if len(old_speak) == len(new_speak) and len(old_speak) > 0:
                    for oc, nc in zip(old_speak, new_speak):
                        df_giris[nc] = df_giris[oc].fillna('')
                df_giris.drop(columns=old_cols, inplace=True, errors='ignore')
        else:
            gercek_ogrenciler = pd.read_sql_query("SELECT ogrenci_no AS \"Ogrenci_No\", ad_soyad AS \"Ad_Soyad\" FROM ogrenciler WHERE modul=%s AND sube=%s", conn, params=(ilgili_modul, secilen_sube))
            df_giris = gercek_ogrenciler.copy() if not gercek_ogrenciler.empty else pd.DataFrame(columns=["Ogrenci_No", "Ad_Soyad"])
            df_giris.insert(2, "Sınava Girdi", True)
            df_giris.insert(3, "Kitapçık", "A")
            df_giris["Sınav Notu"] = 0.0 
        
        for nc in current_tos_cols:
            if nc not in df_giris.columns:
                df_giris[nc] = ""
            else:
                df_giris[nc] = df_giris[nc].apply(lambda x: str(x).replace(".", ",") if pd.notnull(x) else "")
        
        skill_cols_in_order = current_tos_cols
        
        total_ogr = len(df_giris)
        if total_ogr > 0:
            df_giris["Sınav Notu"] = pd.to_numeric(df_giris["Sınav Notu"], errors='coerce').fillna(0.0)
            graded_ogr = int((df_giris['Sınav Notu'] > 0).sum())
            tamamlanma_yuzdesi = min(int((graded_ogr / total_ogr) * 100), 100)
            sube_ort_str = f"{df_giris[df_giris['Sınav Notu'] > 0]['Sınav Notu'].mean():.1f}" if graded_ogr > 0 else "0.0"
            max_not_str = f"{df_giris['Sınav Notu'].max():.1f}" if graded_ogr > 0 else "0.0"
            
            st.markdown(f"""
            <div class="c-kpi-grid">
                <div class="c-kpi-card accent">
                    <div class="c-kpi-lbl">Şube Ortalaması</div>
                    <div class="c-kpi-val">{sube_ort_str}</div>
                </div>
                <div class="c-kpi-card">
                    <div class="c-kpi-lbl">En Yüksek Not</div>
                    <div class="c-kpi-val">{max_not_str}</div>
                </div>
                <div class="c-kpi-card">
                    <div class="c-kpi-lbl">Notlandırılan / Kayıtlı</div>
                    <div class="c-kpi-val">{graded_ogr} / {total_ogr}</div>
                </div>
            </div>
            """, unsafe_allow_html=True)
            st.progress(tamamlanma_yuzdesi, text=f"Sınıf Notlandırma İlerlemesi: %{tamamlanma_yuzdesi}")

        if is_multi_booklet:
            tab_yoklama, tab_notlar = st.tabs(["📍 1. Aşama: Yoklama ve Kitapçık Seçimi", "✍️ 2. Aşama: Not Girişi (A/B Görünümlü)"])
            
            with tab_yoklama:
                st.info("Lütfen 2. Aşamaya geçmeden önce sınıftaki öğrencilerin yoklamasını alın ve dağıtılan kitapçık (A/B/C/D) türlerini belirleyip KAYDET butonuna basın.")
                with st.form("form_yoklama"):
                    col_conf_yok = {
                        "Ogrenci_No": st.column_config.TextColumn("Öğrenci No", disabled=True),
                        "Ad_Soyad": st.column_config.TextColumn("Ad Soyad", disabled=True, pinned=True),
                        "Sınava Girdi": st.column_config.CheckboxColumn("Sınava Girdi"),
                        "Kitapçık": st.column_config.SelectboxColumn("Kitapçık", options=["A", "B", "C", "D"])
                    }
                    
                    y_df = st.data_editor(df_giris, column_order=["Ogrenci_No", "Ad_Soyad", "Sınava Girdi", "Kitapçık"], column_config=col_conf_yok, use_container_width=True, disabled=is_locked)
                    
                    if st.form_submit_button("💾 1. Aşamayı Kaydet (Yoklama ve Kitapçıklar)", type="primary", disabled=is_locked):
                        df_to_save = df_giris.copy()
                        df_to_save["Sınava Girdi"] = y_df["Sınava Girdi"]
                        df_to_save["Kitapçık"] = y_df["Kitapçık"]
                        cur.execute("""INSERT INTO notlar (sinav_kodu, sube, veri_json) VALUES (%s, %s, %s) ON CONFLICT (sinav_kodu, sube) DO UPDATE SET veri_json = EXCLUDED.veri_json;""", (secilen_islem_sinavi, secilen_sube, df_to_save.to_json(orient="records")))
                        conn.commit()
                        st.success("Kitapçık eşleştirmeleri kaydedildi! Artık yukarıdan 2. Aşama (Not Girişi) sekmesine geçebilirsiniz.")
                        st.rerun()

            with tab_notlar:
                gorunum = st.radio("Not Giriş Modu (Excel'deki sıralamanıza göre seçin):", ["👥 Tüm Sınıf (Karma Görünüm)", "📘 Sadece A Kitapçığı Alanlar", "📕 Sadece B Kitapçığı Alanlar"], horizontal=True)
                df_view = df_giris.copy()
                
                if "A Kitapçığı" in gorunum:
                    df_view = df_view[df_view["Kitapçık"] == "A"].copy()
                    df_view = df_view[base_cols + skill_cols_in_order]
                    
                elif "B Kitapçığı" in gorunum:
                    df_view = df_view[df_view["Kitapçık"] == "B"].copy()
                    active_safe_map = {k: v for k, v in safe_b_map.items() if k in df_view.columns}
                    df_view = df_view.rename(columns=active_safe_map)
                    skill_cols = [c for c in df_view.columns if c not in base_cols]
                    df_view = df_view[base_cols + skill_cols]
                
                if df_view.empty:
                    st.warning("⚠️ Seçtiğiniz filtreye uygun öğrenci bulunamadı. Lütfen önce 1. Aşamadan öğrencilere o kitapçığı atayıp atamadığınızı kontrol edin.")
                else:
                    if ort_mesaj: st.info(ort_mesaj)
                    
                    col_conf_not = {
                        "Ogrenci_No": st.column_config.TextColumn("Öğrenci No", disabled=True),
                        "Ad_Soyad": st.column_config.TextColumn("Ad Soyad", disabled=True, pinned=True),
                        "Sınava Girdi": st.column_config.CheckboxColumn("Sınava Girdi", disabled=True),
                        "Kitapçık": st.column_config.TextColumn("Kitapçık", disabled=True),
                        "Sınav Notu": st.column_config.NumberColumn("Sınav Notu", disabled=True)
                    }
                    
                    for c in df_view.columns:
                        if c not in base_cols: col_conf_not[c] = st.column_config.TextColumn(c)
                    
                    with st.form("hoca_veri_giris_formu"):
                        girilen_notlar_df = st.data_editor(df_view, num_rows="dynamic", use_container_width=True, column_config=col_conf_not, disabled=is_locked)
                        kaydet_btn = st.form_submit_button("💾 Notları Buluta Kaydet", type="primary", disabled=is_locked)
                        
                    if kaydet_btn and not is_locked:
                        hatalar = []
                        
                        if "B Kitapçığı" in gorunum:
                            girilen_notlar_df = girilen_notlar_df.rename(columns=reverse_b_map)
                            
                        df_to_save = df_giris.copy()
                        df_to_save.update(girilen_notlar_df)
                        
                        processed_cols = set()
                        for c1, c2, c3 in cift_okumalar:
                            processed_cols.update([c1, c2])
                            if c3: processed_cols.add(c3)
                        
                        for col in max_puan_sozlugu.keys():
                            df_to_save[col] = pd.to_numeric(df_to_save[col].astype(str).str.replace(',', '.').replace(['', 'nan', 'None', '<NA>'], '0'), errors='coerce').fillna(0.0)
                        
                        df_to_save["Sınav Notu"] = 0.0
                        
                        for idx, row in df_to_save.iterrows():
                            if row.get("Sınava Girdi", True):
                                ogrenci_toplam = 0.0
                                
                                for c, m_puan in max_puan_sozlugu.items():
                                    val = df_to_save.at[idx, c]
                                    if val > m_puan: hatalar.append(f"**{row.get('Ogrenci_No', '')}**: '{c.split('-')[-1].strip()}' notu ({val}), max ({m_puan}) aşıyor!")
                                    
                                for c1, c2, c3 in cift_okumalar:
                                    v1 = df_to_save.at[idx, c1]
                                    v2 = df_to_save.at[idx, c2]
                                    
                                    has_v3 = False
                                    v3 = 0.0
                                    if c3:
                                        v3_raw = str(girilen_notlar_df.at[idx, c3]).strip() if c3 in girilen_notlar_df.columns else ""
                                        has_v3 = v3_raw not in ["", "nan", "None", "<NA>"]
                                        v3 = df_to_save.at[idx, c3]
                                    
                                    if abs(v1 - v2) > tolerans:
                                        if not has_v3 and c3 is not None:
                                            hatalar.append(f"**{row.get('Ogrenci_No', '')}**: {c1.split('-')[-1].strip()} okumalar arası fark ({abs(v1-v2)}), {tolerans} puanlık toleransı aşıyor! 3. Okuma girilmelidir.")
                                        elif has_v3:
                                            diff1 = abs(v3 - v1)
                                            diff2 = abs(v3 - v2)
                                            if diff1 < diff2: ogrenci_toplam += (v3 + v1) / 2.0
                                            elif diff2 < diff1: ogrenci_toplam += (v3 + v2) / 2.0
                                            else: ogrenci_toplam += (v3 + max(v1, v2)) / 2.0
                                        else:
                                            ogrenci_toplam += (v1 + v2) / 2.0
                                    else:
                                        ogrenci_toplam += (v1 + v2) / 2.0
                                        
                                for c in max_puan_sozlugu.keys():
                                    if c not in processed_cols: ogrenci_toplam += df_to_save.at[idx, c]
                                        
                                df_to_save.at[idx, "Sınav Notu"] = round(ogrenci_toplam, 2)
                            else:
                                for col in max_puan_sozlugu.keys(): df_to_save.at[idx, col] = 0.0
                                df_to_save.at[idx, "Sınav Notu"] = 0.0
                                
                        if hatalar: 
                            st.error("🛑 SİSTEM KAYDI DURDURDU: KURALLAR İHLAL EDİLDİ VEYA MAX PUAN AŞILDI!")
                            for h in hatalar: st.warning(h)
                            st.stop()
                            
                        cur.execute("""INSERT INTO notlar (sinav_kodu, sube, veri_json) VALUES (%s, %s, %s) ON CONFLICT (sinav_kodu, sube) DO UPDATE SET veri_json = EXCLUDED.veri_json;""", (secilen_islem_sinavi, secilen_sube, df_to_save.to_json(orient="records")))
                        conn.commit(); cur.close(); conn.close()
                        st.success("Notlar bulut veritabanına başarıyla kaydedildi!")
                        st.rerun()

        # TEK KİTAPÇIK İÇİN DÜMDÜZ GİRİŞ EKRANI
        else:
            base_cols = ["Ogrenci_No", "Ad_Soyad", "Sınava Girdi", "Sınav Notu"]
            df_view = df_giris[base_cols + skill_cols_in_order].copy()
            
            if ort_mesaj: st.info(ort_mesaj)
            
            col_conf_not = {
                "Ogrenci_No": st.column_config.TextColumn("Öğrenci No", disabled=True),
                "Ad_Soyad": st.column_config.TextColumn("Ad Soyad", disabled=True, pinned=True),
                "Sınava Girdi": st.column_config.CheckboxColumn("Sınava Girdi"),
                "Sınav Notu": st.column_config.NumberColumn("Sınav Notu", disabled=True)
            }
            
            for c in df_view.columns:
                if c not in base_cols: col_conf_not[c] = st.column_config.TextColumn(c)
            
            st.info("Bu sınav **Tek Kitapçık** olarak yapılandırılmıştır. Tüm notları doğrudan aşağıdaki tabloya girebilirsiniz.")
            
            with st.form("hoca_veri_giris_formu_tek"):
                girilen_notlar_df = st.data_editor(df_view, num_rows="dynamic", use_container_width=True, column_config=col_conf_not, disabled=is_locked)
                kaydet_btn = st.form_submit_button("💾 Notları Buluta Kaydet", type="primary", disabled=is_locked)
                
            if kaydet_btn and not is_locked:
                hatalar = []
                
                df_to_save = df_giris.copy()
                df_to_save.update(girilen_notlar_df)
                
                processed_cols = set()
                for c1, c2, c3 in cift_okumalar:
                    processed_cols.update([c1, c2])
                    if c3: processed_cols.add(c3)
                
                for col in max_puan_sozlugu.keys():
                    df_to_save[col] = pd.to_numeric(df_to_save[col].astype(str).str.replace(',', '.').replace(['', 'nan', 'None', '<NA>'], '0'), errors='coerce').fillna(0.0)
                
                df_to_save["Sınav Notu"] = 0.0
                
                for idx, row in df_to_save.iterrows():
                    if row.get("Sınava Girdi", True):
                        ogrenci_toplam = 0.0
                        
                        for c, m_puan in max_puan_sozlugu.items():
                            val = df_to_save.at[idx, c]
                            if val > m_puan: hatalar.append(f"**{row.get('Ogrenci_No', '')}**: '{c.split('-')[-1].strip()}' notu ({val}), max ({m_puan}) aşıyor!")
                            
                        for c1, c2, c3 in cift_okumalar:
                            v1 = df_to_save.at[idx, c1]
                            v2 = df_to_save.at[idx, c2]
                            
                            has_v3 = False
                            v3 = 0.0
                            if c3:
                                v3_raw = str(girilen_notlar_df.at[idx, c3]).strip() if c3 in girilen_notlar_df.columns else ""
                                has_v3 = v3_raw not in ["", "nan", "None", "<NA>"]
                                v3 = df_to_save.at[idx, c3]
                            
                            if abs(v1 - v2) > tolerans:
                                if not has_v3 and c3 is not None:
                                    hatalar.append(f"**{row.get('Ogrenci_No', '')}**: {c1.split('-')[-1].strip()} okumalar arası fark ({abs(v1-v2)}), {tolerans} puanlık toleransı aşıyor! 3. Okuma girilmelidir.")
                                elif has_v3:
                                    diff1 = abs(v3 - v1)
                                    diff2 = abs(v3 - v2)
                                    if diff1 < diff2: ogrenci_toplam += (v3 + v1) / 2.0
                                    elif diff2 < diff1: ogrenci_toplam += (v3 + v2) / 2.0
                                    else: ogrenci_toplam += (v3 + max(v1, v2)) / 2.0
                                else:
                                    ogrenci_toplam += (v1 + v2) / 2.0
                            else:
                                ogrenci_toplam += (v1 + v2) / 2.0
                                
                        for c in max_puan_sozlugu.keys():
                            if c not in processed_cols: ogrenci_toplam += df_to_save.at[idx, c]
                                
                        df_to_save.at[idx, "Sınav Notu"] = round(ogrenci_toplam, 2)
                    else:
                        for col in max_puan_sozlugu.keys(): df_to_save.at[idx, col] = 0.0
                        df_to_save.at[idx, "Sınav Notu"] = 0.0
                        
                if hatalar: 
                    st.error("🛑 SİSTEM KAYDI DURDURDU: KURALLAR İHLAL EDİLDİ VEYA MAX PUAN AŞILDI!")
                    for h in hatalar: st.warning(h)
                    st.stop()
                    
                cur.execute("""INSERT INTO notlar (sinav_kodu, sube, veri_json) VALUES (%s, %s, %s) ON CONFLICT (sinav_kodu, sube) DO UPDATE SET veri_json = EXCLUDED.veri_json;""", (secilen_islem_sinavi, secilen_sube, df_to_save.to_json(orient="records")))
                conn.commit(); cur.close(); conn.close()
                st.success("Notlar bulut veritabanına başarıyla kaydedildi!")
                st.rerun()
