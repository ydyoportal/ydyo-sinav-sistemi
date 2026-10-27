import streamlit as st
import pandas as pd
import plotly.express as px
import psycopg2
import json
import io
import hashlib
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Flowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

st.set_page_config(page_title="KTO Karatay YDY Portalı", page_icon="🔐", layout="wide")

SALT = "KTO_YDYO_SECURE_SALT_2026"

def hash_sifre(sifre: str) -> str:
    if not sifre: return ""
    if len(sifre) == 64 and all(c in '0123456789abcdefABCDEF' for c in sifre): return sifre
    return hashlib.sha256((sifre + SALT).encode('utf-8')).hexdigest()

def sifre_dogrula(girilen_sifre: str, hashli_sifre: str) -> bool:
    return hash_sifre(girilen_sifre) == hashli_sifre

def get_db_connection():
    if "SUPABASE_URI" in st.secrets: uri = st.secrets["SUPABASE_URI"]
    else: uri = "postgresql://postgres.atyewikzyuvembeikirx:X9R%24t9vVu7ctF%24%21@aws-0-ap-southeast-2.pooler.supabase.com:5432/postgres"
    return psycopg2.connect(uri)

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
# PDF RAPORLAMA FONKSİYONLARI 
# ==========================================
def tr_to_eng(text):
    if not isinstance(text, str): return str(text)
    tr_map = {'ı':'i', 'i':'i', 'I':'I', 'İ':'I', 'ğ':'g', 'Ğ':'G', 'ü':'u', 'Ü':'U', 'ş':'s', 'Ş':'S', 'ö':'o', 'Ö':'O', 'ç':'c', 'Ç':'C'}
    for k, v in tr_map.items(): text = text.replace(k, v)
    return text

class ModernBar(Flowable):
    def __init__(self, width, height, percentage):
        Flowable.__init__(self)
        self.width = width; self.height = height; self.percentage = max(0, min(100, percentage))
        if self.percentage >= 75: self.color = colors.HexColor("#27AE60")
        elif self.percentage >= 50: self.color = colors.HexColor("#F1C40F")
        else: self.color = colors.HexColor("#E74C3C")
    
    # LAYOUT ERROR ÇÖZÜMÜ İÇİN EKLENEN WRAP FONKSİYONU
    def wrap(self, availWidth, availHeight):
        return (self.width, self.height)
        
    def draw(self):
        self.canv.setFillColor(colors.HexColor("#EAECEE"))
        self.canv.roundRect(0, 0, self.width, self.height, 3, fill=1, stroke=0)
        fill_width = (self.width * self.percentage) / 100
        if fill_width > 0:
            self.canv.setFillColor(self.color)
            self.canv.roundRect(0, 0, fill_width, self.height, 3, fill=1, stroke=0)

styles = getSampleStyleSheet()
header_style = ParagraphStyle('Header', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=14, textColor=colors.HexColor("#1A5276"), alignment=1, spaceAfter=5)
sub_header_style = ParagraphStyle('SubHeader', parent=styles['Normal'], fontName='Helvetica', fontSize=10, textColor=colors.grey, alignment=1, spaceAfter=20)
title_style = ParagraphStyle('ModernTitle', parent=styles['Title'], fontName='Helvetica-Bold', fontSize=16, textColor=colors.HexColor("#2C3E50"), spaceAfter=10)
info_style = ParagraphStyle('Info', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, textColor=colors.HexColor("#34495E"), spaceAfter=15)
cell_style = ParagraphStyle('Cell', parent=styles['Normal'], fontName='Helvetica', fontSize=9, alignment=1)
cell_style_left = ParagraphStyle('CellLeft', parent=styles['Normal'], fontName='Helvetica', fontSize=9, alignment=0)
header_cell_style = ParagraphStyle('HeaderCell', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, textColor=colors.whitesmoke, alignment=1)

def build_pdf_doc(elements, is_landscape=False):
    buffer = io.BytesIO()
    page_format = landscape(A4) if is_landscape else A4
    doc = SimpleDocTemplate(buffer, pagesize=page_format, rightMargin=30, leftMargin=30, topMargin=40, bottomMargin=30)
    doc.build(elements)
    buffer.seek(0)
    return buffer

def create_landscape_class_pdf(veri):
    elements = []
    elements.append(Paragraph("KTO KARATAY UNIVERSITY", header_style))
    elements.append(Paragraph("School of Foreign Languages | Testing Office", sub_header_style))
    elements.append(Paragraph(tr_to_eng(veri['baslik']), title_style))
    elements.append(Paragraph(tr_to_eng(veri['alt_baslik']), info_style))
    
    kpi_df = veri['df_kpi']
    kpi_data = [[Paragraph("Metric", header_cell_style), Paragraph("Value", header_cell_style)]]
    for _, row in kpi_df.iterrows(): kpi_data.append([Paragraph(tr_to_eng(str(row['Metric'])), cell_style_left), Paragraph(str(row['Value']), cell_style)])
    kpi_table = Table(kpi_data, colWidths=[120, 100])
    kpi_table.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor("#2C3E50")), ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('BOTTOMPADDING', (0,0), (-1,-1), 8), ('TOPPADDING', (0,0), (-1,-1), 8), ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey), ('BACKGROUND', (0,1), (-1,-1), colors.HexColor("#F8F9F9"))]))
    
    skills_data = [[Paragraph("Section / Part", header_cell_style), Paragraph("Objectives", header_cell_style), Paragraph("Avg", header_cell_style), Paragraph("Max", header_cell_style), Paragraph("Performance", header_cell_style)]]
    for b_dict in veri['beceri_listesi']:
        bar = ModernBar(130, 10, b_dict['yuzde'])
        skills_data.append([Paragraph(tr_to_eng(b_dict['isim']), cell_style_left), Paragraph(tr_to_eng(b_dict.get('kazanim', '')), cell_style_left), Paragraph(f"{b_dict['ort']:.1f}", cell_style), Paragraph(str(b_dict['max']), cell_style), bar])
    
    # SÜTUN GENİŞLİKLERİ KÜÇÜLTÜLEREK SAYFAYA SIĞDIRILDI (TOPLAM 510)
    skill_table = Table(skills_data, colWidths=[110, 180, 40, 40, 140])
    s_style = [('BACKGROUND', (0,0), (-1,0), colors.HexColor("#2C3E50")), ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey)]
    for i in range(1, len(skills_data)):
        if i % 2 == 0: s_style.append(('BACKGROUND', (0,i), (-1,i), colors.HexColor("#F8F9F9")))
    skill_table.setStyle(TableStyle(s_style))
    
    # TAŞMA HATASI (LAYOUT ERROR) OLMAMASI İÇİN İÇ BOŞLUKLAR (PADDING) SIFIRLANDI
    kapsayici_tablo = Table([[kpi_table, skill_table]], colWidths=[240, 542])
    kapsayici_tablo.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0)
    ]))
    elements.extend([kapsayici_tablo, Spacer(1, 20)])
    
    df_liste = veri['df_liste'].copy()
    elements.append(Paragraph("<b>Detailed Student List & Scores</b>", styles['Heading4']))
    elements.append(Spacer(1, 5))
    col_count = len(df_liste.columns)
    base_widths = {'Student ID': 70, 'Name Surname': 140, 'Status': 60, 'Total Score': 60}
    kalan_genislik = 782 - sum(base_widths.values())
    skill_col_width = kalan_genislik / max(1, col_count - 4)
    col_widths = [base_widths.get(c, skill_col_width) for c in df_liste.columns]
    
    tablo_verisi = [[Paragraph(c, header_cell_style) for c in df_liste.columns]]
    for _, row in df_liste.iterrows():
        row_data = []
        for c_name in df_liste.columns:
            c_text = tr_to_eng(str(row[c_name]))
            c_style = cell_style_left if c_name == 'Name Surname' else cell_style
            row_data.append(Paragraph(c_text, c_style))
        tablo_verisi.append(row_data)
    l_table = Table(tablo_verisi, colWidths=col_widths, repeatRows=1)
    l_style = [('BACKGROUND', (0,0), (-1,0), colors.HexColor("#2C3E50")), ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey), ('BOTTOMPADDING', (0,0), (-1,0), 8), ('TOPPADDING', (0,0), (-1,0), 8)]
    for i in range(1, len(tablo_verisi)):
        if i % 2 == 0: l_style.append(('BACKGROUND', (0,i), (-1,i), colors.HexColor("#F8F9F9")))
    l_table.setStyle(TableStyle(l_style))
    elements.append(l_table)
    return build_pdf_doc(elements, is_landscape=True)

def create_modern_pdf(rapor_verileri):
    elements = []
    for idx, veri in enumerate(rapor_verileri):
        elements.append(Paragraph("KTO KARATAY UNIVERSITY", header_style))
        elements.append(Paragraph("School of Foreign Languages | Testing Office", sub_header_style))
        elements.append(Paragraph(tr_to_eng(veri['baslik']), title_style))
        if veri.get('alt_baslik'): elements.append(Paragraph(tr_to_eng(veri['alt_baslik']), info_style))
        
        if veri.get('beceri_grafik_verisi'):
            elements.append(Paragraph("<b>Skill Profile</b>", styles['Heading4']))
            elements.append(Spacer(1, 5))
            b_data = []
            for b_name, b_val in veri['beceri_grafik_verisi'].items():
                bar = ModernBar(120, 10, b_val)
                b_data.append([Paragraph(tr_to_eng(b_name), cell_style_left), bar, Paragraph(f"%{b_val:.1f}", cell_style)])
            if b_data:
                b_table = Table(b_data, colWidths=[120, 130, 50])
                b_table.setStyle(TableStyle([('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('BOTTOMPADDING', (0,0), (-1,-1), 4), ('TOPPADDING', (0,0), (-1,-1), 4)]))
                elements.extend([b_table, Spacer(1, 15)])
                
        df_listesi = veri.get('df_list', [])
        for tablo_baslik, df_tablo in df_listesi:
            if not df_tablo.empty:
                if tablo_baslik:
                    elements.append(Paragraph(f"<b>{tr_to_eng(tablo_baslik)}</b>", styles['Heading4']))
                    elements.append(Spacer(1, 5))
                df_temp = df_tablo.copy()
                df_temp.columns = [tr_to_eng(c) for c in df_temp.columns]
                cols_count = len(df_temp.columns)
                if cols_count == 4: col_widths = [100, 295, 70, 70]
                else: col_widths = [535 / max(1, cols_count)] * cols_count
                tablo_verisi = [[Paragraph(c, header_cell_style) for c in df_temp.columns]]
                for _, row in df_temp.iterrows():
                    row_data = []
                    for col_idx, cell_val in enumerate(row):
                        c_style = cell_style_left if cols_count == 4 and col_idx in [0, 1] else cell_style
                        row_data.append(Paragraph(tr_to_eng(str(cell_val)), c_style))
                    tablo_verisi.append(row_data)
                tablo = Table(tablo_verisi, colWidths=col_widths, repeatRows=1)
                t_style = [('BACKGROUND', (0,0), (-1,0), colors.HexColor("#283747")), ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey)]
                for i in range(1, len(tablo_verisi)):
                    if i % 2 == 0: t_style.append(('BACKGROUND', (0,i), (-1,i), colors.HexColor("#F8F9F9")))
                tablo.setStyle(TableStyle(t_style))
                elements.extend([tablo, Spacer(1, 20)])
        if idx < len(rapor_verileri) - 1: elements.append(PageBreak())
    return build_pdf_doc(elements, is_landscape=False)

# ==========================================
# GİRİŞ KAPISI VE OTURUM
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

if not st.session_state['giris_basarili']:
    st.title("🔐 KTO Karatay YDY - Merkezi Giriş")
    with st.form("giris_formu"):
        k_adi = st.text_input("Kullanıcı Adı")
        sifre = st.text_input("Şifre", type="password")
        if st.form_submit_button("Sisteme Giriş Yap", type="primary"): giris_yap(k_adi, sifre)
else:
    st.sidebar.title(f"👤 {st.session_state['isim']}")
    st.sidebar.info(f"Yetki: {st.session_state['rol']}")
    if st.sidebar.button("🚪 Güvenli Çıkış"): cikis_yap()
    
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
        
    st.title("🎓 KTO Karatay YDY Öğrenme Analitiği Sistemi")

    if st.session_state['rol'] in ["Admin", "Testing"]:
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
                st.session_state['taslak_df'] = pd.DataFrame({"Ana Bölüm": ["Grammar"], "Alt Bölüm (Part)": ["Part A"], "Ölçülen Kazanımlar": ["Present Simple"], "Max Puan": [15.0]})
            
            duzenlenen_rubrik = st.data_editor(
                st.session_state['taslak_df'], 
                num_rows="dynamic", 
                column_config={
                    "Ana Bölüm": st.column_config.SelectboxColumn("Ana Bölüm", options=ANA_BOLUMLER, required=True),
                    "Max Puan": st.column_config.NumberColumn("Max Puan", step=0.5, format="%.2f")
                }, 
                use_container_width=True
            )
            
            if st.button("🚀 Sınavı Onayla & Öğretim Elemanlarına Aç", type="primary"):
                conn = get_db_connection()
                cur = conn.cursor()
                cur.execute("SELECT sinav_kodu FROM sinavlar WHERE sinav_kodu=%s", (sinav_kodu,))
                if cur.fetchone(): st.warning("Bu sınav kodu zaten açık!")
                else:
                    cur.execute("INSERT INTO sinavlar (sinav_kodu, kilitli) VALUES (%s, FALSE)", (sinav_kodu,))
                    cur.execute("INSERT INTO sinav_ayarlari (sinav_kodu, tolerans) VALUES (%s, %s)", (sinav_kodu, tolerans_baremi))
                    for _, row in duzenlenen_rubrik.iterrows():
                        cur.execute("INSERT INTO sinav_tos (sinav_kodu, ana_bolum, alt_bolum, kazanimlar, max_puan) VALUES (%s, %s, %s, %s, %s)", (sinav_kodu, row['Ana Bölüm'], row['Alt Bölüm (Part)'], row['Ölçülen Kazanımlar'], float(row['Max Puan'])))
                    conn.commit()
                    st.success("Sınav başarıyla buluta kaydedildi!")
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
                cur.execute("SELECT tolerans FROM sinav_ayarlari WHERE sinav_kodu=%s", (duzenlenecek_sinav,))
                tol_res = cur.fetchone()
                mevcut_tolerans = int(tol_res[0]) if (tol_res and tol_res[0] is not None) else 6
                yeni_tolerans = st.number_input("Güncel Çift Okuma Tolerans Farkı", min_value=0, value=mevcut_tolerans)
                df_mevcut_tos = pd.read_sql_query("SELECT ana_bolum AS \"Ana Bölüm\", alt_bolum AS \"Alt Bölüm (Part)\", kazanimlar AS \"Ölçülen Kazanımlar\", max_puan AS \"Max Puan\" FROM sinav_tos WHERE sinav_kodu=%s", conn, params=(duzenlenecek_sinav,))
                
                st.markdown("**Güncel Belirtke Tablosu (ToS):**")
                guncel_tos = st.data_editor(df_mevcut_tos, num_rows="dynamic", column_config={"Ana Bölüm": st.column_config.SelectboxColumn("Ana Bölüm", options=ANA_BOLUMLER, required=True), "Max Puan": st.column_config.NumberColumn("Max Puan", step=0.5, format="%.2f")}, use_container_width=True, key="edit_tos_editor")
                if st.button("💾 Değişiklikleri Kaydet (Sınavı Güncelle)", type="primary"):
                    cur.execute("UPDATE sinav_ayarlari SET tolerans=%s WHERE sinav_kodu=%s", (yeni_tolerans, duzenlenecek_sinav))
                    cur.execute("DELETE FROM sinav_tos WHERE sinav_kodu=%s", (duzenlenecek_sinav,))
                    for _, row in guncel_tos.iterrows():
                        cur.execute("INSERT INTO sinav_tos (sinav_kodu, ana_bolum, alt_bolum, kazanimlar, max_puan) VALUES (%s, %s, %s, %s, %s)", (duzenlenecek_sinav, row['Ana Bölüm'], row['Alt Bölüm (Part)'], row['Ölçülen Kazanımlar'], float(row['Max Puan'])))
                    conn.commit(); st.success(f"{duzenlenecek_sinav} başarıyla güncellendi!")
                st.divider()
                
                st.markdown("### 🗑️ Sınav Silme")
                silinecek_sinav = st.selectbox("Silinecek Sınavı Seçin", df_sinav_durum['sinav_kodu'].tolist(), key="sil_sec")
                if st.button("🗑️ Seçili Sınavı Tamamen Sil"):
                    conn.cursor().execute("DELETE FROM sinavlar WHERE sinav_kodu=%s", (silinecek_sinav,)); conn.commit(); st.success("Sınav silindi!"); st.rerun()
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
                st.info("Aşağıdaki formu kullanarak hangi hocanın hangi modülde ve şubede ders verdiğini atayabilirsiniz. Öğretmenler kendi panellerine girdiklerinde SADECE burada atandıkları modülün sınavlarını görebileceklerdir.")
                
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
                    for _, r in duzenlenen_ogr.iterrows(): cur.execute("INSERT INTO ogrenciler (ogrenci_no, ad_soyad, modul, sube) VALUES (%s, %s, %s, %s)", (r['ogrenci_no'], r['ad_soyad'], r['modul'], r['sube']))
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
                "🏫 Sınıf / Şube Raporu"
            ], horizontal=True)
            st.divider()
            
            conn = get_db_connection()
            cur = conn.cursor()
            
            if rapor_kapsami == "🏆 Dönem Sonu Konsolide Başarı":
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
            
            # 4. BİREYSEL KARNELER
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
                        
                        cur.execute("SELECT ana_bolum, alt_bolum, kazanimlar, max_puan FROM sinav_tos WHERE sinav_kodu=%s", (sec_sinav,))
                        tos_df = pd.DataFrame(cur.fetchall(), columns=["ana_bolum", "alt_bolum", "kazanimlar", "max_puan"])
                        tos_dict = {f"{row['ana_bolum']} - {row['alt_bolum']}": {"max": float(row['max_puan']), "kazanim": row['kazanimlar']} for _, row in tos_df.iterrows()}
                        
                        cur.execute("SELECT veri_json FROM notlar WHERE sinav_kodu=%s AND sube=%s", (sec_sinav, sec_sube))
                        df_k = pd.read_json(io.StringIO(cur.fetchone()[0]))
                        
                        if not df_k.empty:
                            sec_ogr = st.selectbox("Öğrenci Seçin", df_k['Ad_Soyad'].tolist())
                            ogr_veri = df_k[df_k['Ad_Soyad'] == sec_ogr].iloc[0]
                            st.markdown(f"#### 🎓 Öğrenci: {sec_ogr} (Toplam: {ogr_veri.get('Sınav Notu', 0)})")
                            
                            katildi_mi = bool(ogr_veri.get('Sınava Girdi', True))
                            if not katildi_mi:
                                st.error("⚠️ Bu öğrenci sınava katılmamıştır. (Absent)")
                            else:
                                b_yuzdeleri = {}
                                for col, veri_dict in tos_dict.items():
                                    if col in ogr_veri:
                                        alinan = float(ogr_veri[col])
                                        m_puan = veri_dict["max"]
                                        yuzde = (alinan / m_puan * 100) if m_puan > 0 else 0
                                        b_yuzdeleri[col] = yuzde
                                        st.write(f"**{col}** | Kazanım: *{veri_dict['kazanim']}* ({alinan} / {m_puan})")
                                        st.progress(min(int(yuzde), 100))
                                
                                if st.button("📄 PDF Karne Oluştur", type="primary"):
                                    df_pdf = pd.DataFrame({
                                        "Section / Skill": list(tos_dict.keys()), 
                                        "Objectives": [tos_dict[c]['kazanim'] for c in tos_dict.keys()],
                                        "Score Achieved": [ogr_veri.get(c, 0) for c in tos_dict.keys()],
                                        "Max Score": [tos_dict[c]['max'] for c in tos_dict.keys()]
                                    })
                                    r_veri = [{'baslik': f"{sec_sinav} - Individual Report Card", 'alt_baslik': f"Student: {sec_ogr} | Class: {sec_sube} | Total Score: {ogr_veri.get('Sınav Notu', 0)}", 'df_list': [("Skill & Objectives Details", df_pdf)], 'beceri_grafik_verisi': b_yuzdeleri}]
                                    pdf_buf = create_modern_pdf(r_veri)
                                    if pdf_buf:
                                        st.success("PDF başarıyla oluşturuldu!")
                                        st.download_button("📥 PDF Olarak İndir", data=pdf_buf, file_name=f"{sec_ogr}_ReportCard.pdf", mime="application/pdf")
            
            # 5. SINIF / ŞUBE RAPORU
            elif rapor_kapsami == "🏫 Sınıf / Şube Raporu":
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
                        df_aktif = df_s[df_s["Sınava Girdi"] == True]
                        
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
                        cols_to_show = ["Ogrenci_No", "Ad_Soyad", "Sınava Girdi", "Sınav Notu"] + [c for c in tos_dict.keys() if c in df_s.columns]
                        st.dataframe(df_s[cols_to_show].sort_values(by="Sınav Notu", ascending=False), use_container_width=True, hide_index=True)
                        
                        c_btn1, c_btn2 = st.columns(2)
                        with c_btn1:
                            if st.button("📊 Sınıf Raporu PDF Oluştur (Yatay)", type="primary"):
                                kpi_df = pd.DataFrame({"Metric": ["Class Average", "Median", "Highest Score", "Lowest Score", "Active Students"], "Value": [f"{df_aktif['Sınav Notu'].mean():.2f}" if not df_aktif.empty else "0", f"{df_aktif['Sınav Notu'].median():.2f}" if not df_aktif.empty else "0", f"{df_aktif['Sınav Notu'].max():.2f}" if not df_aktif.empty else "0", f"{df_aktif['Sınav Notu'].min():.2f}" if not df_aktif.empty else "0", str(len(df_aktif))]})
                                df_pdf_liste = df_s[cols_to_show].copy()
                                df_pdf_liste["Sınava Girdi"] = df_pdf_liste["Sınava Girdi"].map({True: "Attended", False: "Absent"})
                                df_pdf_liste = df_pdf_liste.rename(columns={"Ogrenci_No": "Student ID", "Ad_Soyad": "Name Surname", "Sınava Girdi": "Status", "Sınav Notu": "Total Score"})
                                veri = {'baslik': f"{sec_sinav} - Class Performance Report", 'alt_baslik': f"Class: {sec_sube}", 'df_kpi': kpi_df, 'beceri_listesi': beceri_listesi_pdf, 'df_liste': df_pdf_liste}
                                pdf_buf = create_landscape_class_pdf(veri)
                                if pdf_buf:
                                    st.success("Sınıf raporu oluşturuldu!")
                                    st.download_button("📥 Yatay PDF İndir", data=pdf_buf, file_name=f"{sec_sube}_ClassReport.pdf", mime="application/pdf")
                        
                        with c_btn2:
                            if st.button("📦 Tüm Sınıfın Toplu Karnelerini Oluştur"):
                                toplu_veriler = []
                                for _, row in df_aktif.sort_values(by="Ad_Soyad").iterrows(): 
                                    ogr_no = row.get('Ogrenci_No', '')
                                    ogr_ad = row.get('Ad_Soyad', '')
                                    toplam = row.get('Sınav Notu', 0)
                                    b_yuzdeleri = {}
                                    for c, v_dict in tos_dict.items():
                                        mp = v_dict['max']
                                        if c in row: b_yuzdeleri[c] = (row[c]/mp*100) if mp>0 else 0
                                    df_pdf = pd.DataFrame({"Section / Skill": list(tos_dict.keys()), "Objectives": [tos_dict[c]['kazanim'] for c in tos_dict.keys()], "Score Achieved": [row.get(c, 0) for c in tos_dict.keys()], "Max Score": [tos_dict[c]['max'] for c in tos_dict.keys()]})
                                    toplu_veriler.append({'baslik': f"{sec_sinav} - Individual Report Card", 'alt_baslik': f"Student: {ogr_no} - {ogr_ad} | Total Score: {toplam}", 'df_list': [("Skill & Objectives Details", df_pdf)], 'beceri_grafik_verisi': b_yuzdeleri})
                                if toplu_veriler:
                                    pdf_buf = create_modern_pdf(toplu_veriler)
                                    if pdf_buf:
                                        st.success("Toplu karneler oluşturuldu!")
                                        st.download_button("📥 Tüm Karneleri İndir (Tek PDF)", data=pdf_buf, file_name=f"{sec_sube}_AllReportCards.pdf", mime="application/pdf")
                                else: st.warning("Sınava giren öğrenci bulunmuyor.")

            cur.close(); conn.close()

    # ----------------------------------------------------
    # 5. HOCA VERİ GİRİŞ PANELİ
    # ----------------------------------------------------
    elif st.session_state['rol'] == "Öğretim Elemanı":
        st.subheader("📝 Hızlı Veri Giriş Terminali")
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
            st.info("Eğer bir sınıfta derse giriyorsanız, lütfen Testing Ofisi (Admin) ile iletişime geçerek atamanızın yapılmasını isteyiniz.")
            st.stop()
            
        sinav_kilit_durumlari = {row[0]: row[1] for row in sinav_data}
        secilen_islem_sinavi = st.selectbox("Sorumlu Olduğunuz Sınavlar", acik_sinavlar)
        is_locked = sinav_kilit_durumlari.get(secilen_islem_sinavi, False)
        
        if is_locked:
            st.error("🔒 BU SINAV TESTING OFİSİ TARAFINDAN KİLİTLENMİŞTİR!")
            st.warning("Not girişi kapatılmıştır. Yalnızca görüntüleme yapabilirsiniz.")
        
        ilgili_modul = secilen_islem_sinavi.split("_")[0]
        cur.execute("SELECT sube FROM hoca_atamalari WHERE modul=%s AND kullanici_adi=%s", (ilgili_modul, st.session_state['kullanici_adi']))
        hoca_subeleri = [row[0] for row in cur.fetchall()]
        if not hoca_subeleri: st.error("Bu sınavın modülünde size atanmış bir şube bulunmuyor."); st.stop()
        
        secilen_sube = st.selectbox("Sorumlu Olduğunuz Şubeler", hoca_subeleri)
        rubrik_df = pd.read_sql_query("SELECT * FROM sinav_tos WHERE sinav_kodu=%s", conn, params=(secilen_islem_sinavi,))
        
        st.markdown(f"### 📊 {secilen_sube} - Not Tablosu")
        cur.execute("SELECT veri_json FROM notlar WHERE sinav_kodu=%s AND sube=%s", (secilen_islem_sinavi, secilen_sube))
        kaydedilmis_veri = cur.fetchone()
        
        if kaydedilmis_veri: 
            df_giris = pd.read_json(io.StringIO(kaydedilmis_veri[0]))
            if "Sınava Girdi" not in df_giris.columns: df_giris.insert(2, "Sınava Girdi", True)
        else:
            gercek_ogrenciler = pd.read_sql_query("SELECT ogrenci_no AS \"Ogrenci_No\", ad_soyad AS \"Ad_Soyad\" FROM ogrenciler WHERE modul=%s AND sube=%s", conn, params=(ilgili_modul, secilen_sube))
            df_giris = gercek_ogrenciler.copy() if not gercek_ogrenciler.empty else pd.DataFrame(columns=["Ogrenci_No", "Ad_Soyad"])
            df_giris.insert(2, "Sınava Girdi", True)
            for _, row in rubrik_df.iterrows(): df_giris[f"{row['ana_bolum']} - {row['alt_bolum']}"] = 0.0 
            df_giris["Sınav Notu"] = 0.0 
            
        col_conf = {"Sınav Notu": st.column_config.NumberColumn("Sınav Notu", disabled=True), "Sınava Girdi": st.column_config.CheckboxColumn("Sınava Girdi", default=True)}
        girilen_notlar_df = st.data_editor(df_giris, num_rows="dynamic", use_container_width=True, column_config=col_conf, disabled=is_locked)
        
        if not is_locked and st.button("💾 Notları Buluta Kaydet", type="primary"):
            hatalar = []
            max_puan_sozlugu = {f"{r['ana_bolum']} - {r['alt_bolum']}": float(r['max_puan']) for _, r in rubrik_df.iterrows()}
            for idx, row in girilen_notlar_df.iterrows():
                if row.get("Sınava Girdi", True):
                    for col, m_puan in max_puan_sozlugu.items():
                        val = pd.to_numeric(row.get(col), errors='coerce')
                        if pd.notnull(val) and val > m_puan: hatalar.append(f"**{row.get('Ogrenci_No', '')}**: '{col}' notu ({val}), max ({m_puan}) aşıyor!")
            if hatalar: st.error("🛑 HATA!"); [st.warning(h) for h in hatalar]; st.stop()
            
            toplam_notlar = []
            for idx, row in girilen_notlar_df.iterrows():
                if row.get("Sınava Girdi", True): ogrenci_toplam = sum([pd.to_numeric(row.get(col), errors='coerce') for col in max_puan_sozlugu.keys() if pd.notnull(pd.to_numeric(row.get(col), errors='coerce'))])
                else: ogrenci_toplam = 0.0
                toplam_notlar.append(round(ogrenci_toplam, 2))
                
            girilen_notlar_df["Sınav Notu"] = toplam_notlar
            cur.execute("""INSERT INTO notlar (sinav_kodu, sube, veri_json) VALUES (%s, %s, %s) ON CONFLICT (sinav_kodu, sube) DO UPDATE SET veri_json = EXCLUDED.veri_json;""", (secilen_islem_sinavi, secilen_sube, girilen_notlar_df.to_json(orient="records")))
            conn.commit(); cur.close(); conn.close()
            st.success("Notlar bulut veritabanına başarıyla kaydedildi!"); st.rerun()
