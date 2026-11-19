import io
import os
import re
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Flowable, Image
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib import colors

def tr_to_eng(text):
    if not isinstance(text, str):
        return str(text)
    tr_map = {'ı':'i', 'i':'i', 'I':'I', 'İ':'I', 'ğ':'g', 'Ğ':'G', 'ü':'u', 'Ü':'U', 'ş':'s', 'Ş':'S', 'ö':'o', 'Ö':'O', 'ç':'c', 'Ç':'C'}
    for k, v in tr_map.items():
        text = text.replace(k, v)
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
        if s.lower() in ['', 'nan', 'none', '<na>']:
            return 0.0
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
        if c1 in atlananlar:
            continue
        for m1, m2, m3 in isaretler:
            if m1 in c1:
                c2 = c1.replace(m1, m2)
                c3 = c1.replace(m1, m3)
                if c2 in tos_map:
                    temiz_ad = re.sub(r'\s+', ' ', c1.replace(m1, "").replace("()", "")).strip()
                    temiz_ad = re.sub(r'\s*-\s*', ' - ', temiz_ad).strip(" -")
                    eslesen_gruplar[c1] = (c1, c2, c3 if c3 in tos_map else None, temiz_ad)
                    atlananlar.update([c1, c2])
                    if c3 in tos_map:
                        atlananlar.add(c3)
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
