import streamlit as st
import pandas as pd
import plotly.express as px
import psycopg2
import json
import io
import hashlib
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Flowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

st.set_page_config(page_title="KTO Karatay YDY Portalı", page_icon="🎓", layout="wide")

# ==========================================
# KURUMSAL ARAYÜZ (UI) & TEMA AYARLARI
# ==========================================
def kurumsal_tema_uygula():
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;700&family=Jost:wght@400;500;600&display=swap');

    html, body, [class*="css"] { font-family: 'DM Sans', sans-serif !important; }
    h1, h2, h3, h4, h5, h6 { font-family: 'Jost', sans-serif !important; color: #1A365D !important; }
    .stApp { background-color: #F8F9FA; }
    [data-testid="stSidebar"] { background-color: #1A365D; }
    [data-testid="stSidebar"] * { color: #FFFFFF !important; }
    button[kind="primary"] { background-color: #E84C0F !important; border-color: #E84C0F !important; color: white !important; font-weight: 600 !important; border-radius: 6px !important; transition: all 0.3s ease; }
    button[kind="primary"]:hover { background-color: #CC400A !important; border-color: #CC400A !important; transform: translateY(-2px); }
    button[kind="secondary"] { border-color: #1A365D !important; color: #1A365D !important; border-radius: 6px !important; }
    button[kind="secondary"]:hover { border-color: #E84C0F !important; color: #E84C0F !important; }
    [data-testid="stMetricValue"] { color: #E84C0F !important; font-family: 'Jost', sans-serif !important; font-weight: 600; }
    [data-testid="stDataFrame"] { border-radius: 8px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.05); }
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
# BİREYSEL KARNE (DİKEY PDF) FONKSİYONLARI 
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
header_style = ParagraphStyle('Header', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=14, textColor=colors.HexColor("#1A365D"), alignment=1, spaceAfter=5)
sub_header_style = ParagraphStyle('SubHeader', parent=styles['Normal'], fontName='Helvetica', fontSize=10, textColor=colors.grey, alignment=1, spaceAfter=20)
title_style = ParagraphStyle('ModernTitle', parent=styles['Title'], fontName='Helvetica-Bold', fontSize=16, textColor=colors.HexColor("#2C3E50"), spaceAfter=10)
info_style = ParagraphStyle('Info', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, textColor=colors.HexColor("#E84C0F"), spaceAfter=15)
cell_style = ParagraphStyle('Cell', parent=styles['Normal'], fontName='Helvetica', fontSize=8, alignment=1, wordWrap='CJK')
cell_style_left = ParagraphStyle('CellLeft', parent=styles['Normal'], fontName='Helvetica', fontSize=8, alignment=0, wordWrap='CJK')
header_cell_style = ParagraphStyle('HeaderCell', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=9, textColor=colors.whitesmoke, alignment=1, wordWrap='CJK')

def build_pdf_doc(elements):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=40, bottomMargin=30)
    doc.build(elements)
    buffer.seek(0)
    return buffer

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
                t_style = [('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1A365D")), ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('VALIGN', (0,0), (-1,-1), 'MIDDLE'), ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey)]
                for i in range(1, len(tablo_verisi)):
                    if i % 2 == 0: t_style.append(('BACKGROUND', (0,i), (-1,i), colors.HexColor("#F8F9F9")))
                tablo.setStyle(TableStyle(t_style))
                elements.extend([tablo, Spacer(1, 20)])
        if idx < len(rapor_verileri) - 1: elements.append(PageBreak())
    return build_pdf_doc(elements)

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
    st.markdown("<h1 style='text-align: center;'>KTO Karatay YDY - Merkezi Giriş</h1>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        with st.form("giris_formu"):
            k_adi = st.text_input("Kullanıcı Adı")
            sifre = st.text_input("Şifre", type="password")
            if st.form_submit_button("Sisteme Giriş Yap", type="primary", use_container_width=True): giris_yap(k_adi, sifre)
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
                        
                        cur.execute("SELECT ana_bolum, alt_bolum, kazanimlar, max_puan FROM sinav_tos WHERE sinav_kodu=%s", (sec_sinav,))
                        tos_df = pd.DataFrame(cur.fetchall(), columns=["ana_bolum", "alt_bolum", "kazanimlar", "max_puan"])
                        tos_dict = {f"{row['ana_bolum']} - {row['alt_bolum']}": {"max": float(row['max_puan']), "kazanim": row['kazanimlar']} for _, row in tos_df.iterrows()}
                        
                        cur.execute("SELECT veri_json FROM notlar WHERE sinav_kodu=%s AND sube=%s", (sec_sinav, sec_sube))
                        df_k = pd.read_json(io.StringIO(cur.fetchone()[0]))
                        
                        if not df_k.empty:
                            sec_ogr = st.selectbox("Öğrenci Seçin", df_k['Ad_Soyad'].tolist())
                            ogr_veri = df_k[df_k['Ad_Soyad'] == sec_ogr].iloc[0]
                            
                            bkl = ogr_veri.get('Kitapçık', 'A')
                            st.markdown(f"#### 🎓 Öğrenci: {sec_ogr} (Toplam: {ogr_veri.get('Sınav Notu', 0)} | Kitapçık: {bkl})")
                            
                            katildi_mi = bool(ogr_veri.get('Sınava Girdi', True))
                            if not katildi_mi:
                                st.error("⚠️ Bu öğrenci sınava katılmamıştır. (Absent)")
                            else:
                                b_yuzdeleri = {}
                                for col, veri_dict in tos_dict.items():
                                    if col in ogr_veri:
                                        try: alinan = float(str(ogr_veri[col]).replace(',', '.'))
                                        except: alinan = 0.0
                                            
                                        m_puan = veri_dict["max"]
                                        yuzde = (alinan / m_puan * 100) if m_puan > 0 else 0
                                        b_yuzdeleri[col] = yuzde
                                        st.write(f"**{col}** | Kazanım: *{veri_dict['kazanim']}* ({alinan} / {m_puan})")
                                        st.progress(min(int(yuzde), 100))
                                
                                if st.button("📄 PDF Karne Oluştur", type="primary"):
                                    safe_scores = []
                                    for c in tos_dict.keys():
                                        try: safe_scores.append(float(str(ogr_veri.get(c, 0)).replace(',', '.')))
                                        except: safe_scores.append(0.0)
                                        
                                    df_pdf = pd.DataFrame({
                                        "Section / Skill": list(tos_dict.keys()), 
                                        "Objectives": [tos_dict[c]['kazanim'] for c in tos_dict.keys()],
                                        "Score Achieved": safe_scores,
                                        "Max Score": [tos_dict[c]['max'] for c in tos_dict.keys()]
                                    })
                                    r_veri = [{'baslik': f"{sec_sinav} - Individual Report Card", 'alt_baslik': f"Student: {sec_ogr} | Class: {sec_sube} | Booklet: {bkl} | Total Score: {ogr_veri.get('Sınav Notu', 0)}", 'df_list': [("Skill & Objectives Details", df_pdf)], 'beceri_grafik_verisi': b_yuzdeleri}]
                                    pdf_buf = create_modern_pdf(r_veri)
                                    if pdf_buf:
                                        st.success("PDF başarıyla oluşturuldu!")
                                        st.download_button("📥 PDF Olarak İndir", data=pdf_buf, file_name=f"{sec_ogr}_ReportCard.pdf", mime="application/pdf")
            
            # GÜNCELLEME: Çökmeyen Sınıf Raporu (Sadece Excel Dışa Aktarımı)
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
                            if st.button("📊 Sınıf Raporunu Excel Olarak İndir (Tavsiye Edilen)", type="primary"):
                                buffer = io.BytesIO()
                                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                                    # Sayfa 1: Öğrenci Notları
                                    df_excel_liste = df_s[cols_to_show].copy()
                                    df_excel_liste["Sınava Girdi"] = df_excel_liste["Sınava Girdi"].map({True: "Katıldı", False: "Girmedi"})
                                    df_excel_liste = df_excel_liste.rename(columns={"Ogrenci_No": "Öğrenci No", "Ad_Soyad": "Ad Soyad", "Sınav Notu": "Toplam Not"})
                                    df_excel_liste.to_excel(writer, sheet_name="Öğrenci Notları", index=False)
                                    
                                    # Sayfa 2: İstatistikler
                                    kpi_df = pd.DataFrame({
                                        "Metrik": ["Sınıf Ortalaması", "Medyan", "En Yüksek Not", "En Düşük Not", "Aktif Öğrenci Sayısı"], 
                                        "Değer": [
                                            f"{df_aktif['Sınav Notu'].mean():.2f}" if not df_aktif.empty else "0", 
                                            f"{df_aktif['Sınav Notu'].median():.2f}" if not df_aktif.empty else "0", 
                                            f"{df_aktif['Sınav Notu'].max():.2f}" if not df_aktif.empty else "0", 
                                            f"{df_aktif['Sınav Notu'].min():.2f}" if not df_aktif.empty else "0", 
                                            str(len(df_aktif))
                                        ]
                                    })
                                    kpi_df.to_excel(writer, sheet_name="İstatistikler", index=False)
                                    
                                    # Sayfa 3: Beceriler
                                    beceri_df = pd.DataFrame(beceri_listesi_pdf)
                                    if not beceri_df.empty:
                                        beceri_df = beceri_df.rename(columns={"isim": "Bölüm", "kazanim": "Kazanım", "ort": "Sınıf Ortalaması", "max": "Max Puan", "yuzde": "Başarı Yüzdesi (%)"})
                                        beceri_df.to_excel(writer, sheet_name="Beceri Analizi", index=False)
                                        
                                buffer.seek(0)
                                st.success("Sınıf raporu Excel olarak hazırlandı! Hiçbir veri kaybı olmadan tüm geniş sütunları inceleyebilirsiniz.")
                                st.download_button("📥 Excel Raporunu İndir", data=buffer, file_name=f"{sec_sube}_Sinif_Raporu.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                        
                        with c_btn2:
                            if st.button("📦 Tüm Sınıfın Toplu Karnelerini Oluştur (PDF)"):
                                toplu_veriler = []
                                for _, row in df_aktif.sort_values(by="Ad_Soyad").iterrows(): 
                                    ogr_no = row.get('Ogrenci_No', '')
                                    ogr_ad = row.get('Ad_Soyad', '')
                                    toplam = row.get('Sınav Notu', 0)
                                    bkl = row.get('Kitapçık', 'A')
                                    b_yuzdeleri = {}
                                    for c, v_dict in tos_dict.items():
                                        mp = v_dict['max']
                                        if c in row: b_yuzdeleri[c] = (row[c]/mp*100) if mp>0 else 0
                                    
                                    safe_scores = []
                                    for c in tos_dict.keys():
                                        try: safe_scores.append(float(str(row.get(c, 0)).replace(',', '.')))
                                        except: safe_scores.append(0.0)
                                        
                                    df_pdf = pd.DataFrame({
                                        "Section / Skill": list(tos_dict.keys()), 
                                        "Objectives": [tos_dict[c]['kazanim'] for c in tos_dict.keys()], 
                                        "Score Achieved": safe_scores, 
                                        "Max Score": [tos_dict[c]['max'] for c in tos_dict.keys()]
                                    })
                                    toplu_veriler.append({'baslik': f"{sec_sinav} - Individual Report Card", 'alt_baslik': f"Student: {ogr_no} - {ogr_ad} | Booklet: {bkl} | Total Score: {toplam}", 'df_list': [("Skill & Objectives Details", df_pdf)], 'beceri_grafik_verisi': b_yuzdeleri})
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
            
        cur.execute("SELECT tolerans, kitapcik_turu FROM sinav_ayarlari WHERE sinav_kodu=%s", (secilen_islem_sinavi,))
        ayarlar_res = cur.fetchone()
        tolerans = int(ayarlar_res[0]) if (ayarlar_res and ayarlar_res[0] is not None) else 6
        kitapcik_turu = ayarlar_res[1] if (ayarlar_res and ayarlar_res[1] is not None) else "Çoklu Kitapçık (A ve B)"
        is_multi_booklet = "Çoklu" in kitapcik_turu
        
        ilgili_modul = secilen_islem_sinavi.split("_")[0]
        cur.execute("SELECT sube FROM hoca_atamalari WHERE modul=%s AND kullanici_adi=%s", (ilgili_modul, st.session_state['kullanici_adi']))
        hoca_subeleri = [row[0] for row in cur.fetchall()]
        if not hoca_subeleri: st.error("Bu sınavın modülünde size atanmış bir şube bulunmuyor."); st.stop()
        
        secilen_sube = st.selectbox("Sorumlu Olduğunuz Şubeler", hoca_subeleri)
        
        try:
            rubrik_df = pd.read_sql_query("SELECT ana_bolum, alt_bolum, max_puan, b_grubu_karsiligi FROM sinav_tos WHERE sinav_kodu=%s", conn, params=(secilen_islem_sinavi,))
        except:
            rubrik_df = pd.read_sql_query("SELECT ana_bolum, alt_bolum, max_puan FROM sinav_tos WHERE sinav_kodu=%s", conn, params=(secilen_islem_sinavi,))
            rubrik_df['b_grubu_karsiligi'] = ""
            
        b_map = {}
        reverse_b_map = {}
        for _, row in rubrik_df.iterrows():
            a_name = f"{row['ana_bolum']} - {row['alt_bolum']}"
            b_karsi = str(row.get('b_grubu_karsiligi', '')).strip()
            
            if b_karsi and is_multi_booklet:
                b_name = f"{row['ana_bolum']} - {b_karsi}"
                b_map[a_name] = b_name
                reverse_b_map[b_name] = a_name
            else:
                b_map[a_name] = a_name
                reverse_b_map[a_name] = a_name

        st.markdown(f"### 📊 {secilen_sube} - Veri Giriş Paneli")
        cur.execute("SELECT veri_json FROM notlar WHERE sinav_kodu=%s AND sube=%s", (secilen_islem_sinavi, secilen_sube))
        kaydedilmis_veri = cur.fetchone()
        
        if kaydedilmis_veri: 
            df_giris = pd.read_json(io.StringIO(kaydedilmis_veri[0]))
            if "Sınava Girdi" not in df_giris.columns: df_giris.insert(2, "Sınava Girdi", True)
            if "Kitapçık" not in df_giris.columns: df_giris.insert(3, "Kitapçık", "A")
        else:
            gercek_ogrenciler = pd.read_sql_query("SELECT ogrenci_no AS \"Ogrenci_No\", ad_soyad AS \"Ad_Soyad\" FROM ogrenciler WHERE modul=%s AND sube=%s", conn, params=(ilgili_modul, secilen_sube))
            df_giris = gercek_ogrenciler.copy() if not gercek_ogrenciler.empty else pd.DataFrame(columns=["Ogrenci_No", "Ad_Soyad"])
            df_giris.insert(2, "Sınava Girdi", True)
            df_giris.insert(3, "Kitapçık", "A")
            df_giris["Sınav Notu"] = 0.0 
        
        for _, row in rubrik_df.iterrows():
            col_name = f"{row['ana_bolum']} - {row['alt_bolum']}"
            if col_name not in df_giris.columns:
                df_giris[col_name] = ""
            else:
                df_giris[col_name] = df_giris[col_name].apply(lambda x: str(x).replace(".", ",") if pd.notnull(x) else "")
        
        skill_cols_in_order = [f"{row['ana_bolum']} - {row['alt_bolum']}" for _, row in rubrik_df.iterrows()]
        
        total_ogr = len(df_giris)
        if total_ogr > 0:
            df_giris["Sınav Notu"] = pd.to_numeric(df_giris["Sınav Notu"], errors='coerce').fillna(0.0)
            graded_ogr = (df_giris['Sınav Notu'] > 0).sum()
            tamamlanma_yuzdesi = min(int((graded_ogr / total_ogr) * 100), 100)
            st.progress(tamamlanma_yuzdesi, text=f"Sınıf Notlandırma İlerlemesi: %{tamamlanma_yuzdesi} ({graded_ogr} / {total_ogr} Öğrenci)")
            
            mc1, mc2, mc3 = st.columns(3)
            mc1.metric("Şube Ortalaması", f"{df_giris[df_giris['Sınav Notu'] > 0]['Sınav Notu'].mean():.1f}" if graded_ogr > 0 else "0")
            mc2.metric("En Yüksek Not", f"{df_giris['Sınav Notu'].max():.1f}" if graded_ogr > 0 else "0")
            mc3.metric("Kayıtlı Öğrenci", f"{total_ogr}")
            st.markdown("---")

        if is_multi_booklet:
            tab_yoklama, tab_notlar = st.tabs(["📍 1. Aşama: Yoklama ve Kitapçık Seçimi", "✍️ 2. Aşama: Not Girişi (A/B Görünümlü)"])
            
            with tab_yoklama:
                st.info("Lütfen 2. Aşamaya geçmeden önce sınıftaki öğrencilerin yoklamasını alın ve dağıtılan kitapçık (A/B/C/D) türlerini belirleyip KAYDET butonuna basın.")
                with st.form("form_yoklama"):
                    col_conf_yok = {
                        "Ogrenci_No": st.column_config.TextColumn("Öğrenci No", disabled=True),
                        "Ad_Soyad": st.column_config.TextColumn("Ad Soyad", disabled=True),
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
                    base_cols = ["Ogrenci_No", "Ad_Soyad", "Sınava Girdi", "Kitapçık", "Sınav Notu"]
                    df_view = df_view[base_cols + skill_cols_in_order]
                    
                elif "B Kitapçığı" in gorunum:
                    df_view = df_view[df_view["Kitapçık"] == "B"].copy()
                    df_view = df_view.rename(columns=b_map)
                    base_cols = ["Ogrenci_No", "Ad_Soyad", "Sınava Girdi", "Kitapçık", "Sınav Notu"]
                    skill_cols = [c for c in df_view.columns if c not in base_cols]
                    df_view = df_view[base_cols + skill_cols]
                
                if df_view.empty:
                    st.warning("⚠️ Seçtiğiniz filtreye uygun öğrenci bulunamadı. Lütfen önce 1. Aşamadan öğrencilere o kitapçığı atayıp atamadığınızı kontrol edin.")
                else:
                    col_conf_not = {
                        "Ogrenci_No": st.column_config.TextColumn("Öğrenci No", disabled=True),
                        "Ad_Soyad": st.column_config.TextColumn("Ad Soyad", disabled=True),
                        "Sınava Girdi": st.column_config.CheckboxColumn("Sınava Girdi", disabled=True),
                        "Kitapçık": st.column_config.TextColumn("Kitapçık", disabled=True),
                        "Sınav Notu": st.column_config.NumberColumn("Sınav Notu", disabled=True)
                    }
                    
                    for c in df_view.columns:
                        if c not in ["Ogrenci_No", "Ad_Soyad", "Sınava Girdi", "Kitapçık", "Sınav Notu"]:
                            col_conf_not[c] = st.column_config.TextColumn(c)
                    
                    with st.form("hoca_veri_giris_formu"):
                        girilen_notlar_df = st.data_editor(df_view, num_rows="dynamic", use_container_width=True, column_config=col_conf_not, disabled=is_locked)
                        kaydet_btn = st.form_submit_button("💾 Notları Buluta Kaydet", type="primary", disabled=is_locked)
                        
                    if kaydet_btn and not is_locked:
                        hatalar = []
                        max_puan_sozlugu = {f"{r['ana_bolum']} - {r['alt_bolum']}": float(r['max_puan']) for _, r in rubrik_df.iterrows()}
                        
                        if "B Kitapçığı" in gorunum:
                            girilen_notlar_df = girilen_notlar_df.rename(columns=reverse_b_map)
                            
                        df_to_save = df_giris.copy()
                        df_to_save.update(girilen_notlar_df)
                        
                        cift_okumalar = []
                        for c1 in max_puan_sozlugu.keys():
                            if "(1. Okuma)" in c1 or "1st" in c1:
                                ana_base = c1.split(" - ")[0].replace("(1. Okuma)", "").replace("1st", "").strip()
                                c2, c3 = None, None
                                for k in max_puan_sozlugu.keys():
                                    if "(2. Okuma)" in k or "2nd" in k:
                                        if k.split(" - ")[0].replace("(2. Okuma)", "").replace("2nd", "").strip() == ana_base: c2 = k
                                    if "(3. Okuma)" in k or "3rd" in k:
                                        if k.split(" - ")[0].replace("(3. Okuma)", "").replace("3rd", "").strip() == ana_base: c3 = k
                                if c2: cift_okumalar.append((c1, c2, c3))
                                
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
                                    if c1 in max_puan_sozlugu and c2 in max_puan_sozlugu:
                                        v1 = df_to_save.at[idx, c1]
                                        v2 = df_to_save.at[idx, c2]
                                        
                                        v3_raw = str(girilen_notlar_df.at[idx, c3]).strip() if c3 in max_puan_sozlugu else ""
                                        has_v3 = v3_raw not in ["", "nan", "None", "<NA>"]
                                        v3 = df_to_save.at[idx, c3] if c3 in max_puan_sozlugu else 0.0
                                        
                                        if abs(v1 - v2) > tolerans:
                                            if not has_v3 and c3 in max_puan_sozlugu:
                                                hatalar.append(f"**{row.get('Ogrenci_No', '')}**: {c1.split('(')[0].strip()} okumalar arası fark ({abs(v1-v2)}), {tolerans} puanlık toleransı aşıyor! 3. Okuma girilmelidir.")
                                            elif has_v3:
                                                diff1 = abs(v3 - v1)
                                                diff2 = abs(v3 - v2)
                                                if diff1 < diff2: ogrenci_toplam += (v3 + v1) / 2.0
                                                elif diff2 < diff1: ogrenci_toplam += (v3 + v2) / 2.0
                                                else: ogrenci_toplam += (v3 + max(v1, v2)) / 2.0
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
            
            col_conf_not = {
                "Ogrenci_No": st.column_config.TextColumn("Öğrenci No", disabled=True),
                "Ad_Soyad": st.column_config.TextColumn("Ad Soyad", disabled=True),
                "Sınava Girdi": st.column_config.CheckboxColumn("Sınava Girdi"),
                "Sınav Notu": st.column_config.NumberColumn("Sınav Notu", disabled=True)
            }
            
            for c in df_view.columns:
                if c not in ["Ogrenci_No", "Ad_Soyad", "Sınava Girdi", "Sınav Notu"]:
                    col_conf_not[c] = st.column_config.TextColumn(c)
            
            st.info("Bu sınav **Tek Kitapçık** olarak yapılandırılmıştır. Tüm notları doğrudan aşağıdaki tabloya girebilirsiniz.")
            
            with st.form("hoca_veri_giris_formu_tek"):
                girilen_notlar_df = st.data_editor(df_view, num_rows="dynamic", use_container_width=True, column_config=col_conf_not, disabled=is_locked)
                kaydet_btn = st.form_submit_button("💾 Notları Buluta Kaydet", type="primary", disabled=is_locked)
                
            if kaydet_btn and not is_locked:
                hatalar = []
                max_puan_sozlugu = {f"{r['ana_bolum']} - {r['alt_bolum']}": float(r['max_puan']) for _, r in rubrik_df.iterrows()}
                
                df_to_save = df_giris.copy()
                df_to_save.update(girilen_notlar_df)
                
                cift_okumalar = []
                for c1 in max_puan_sozlugu.keys():
                    if "(1. Okuma)" in c1 or "1st" in c1:
                        ana_base = c1.split(" - ")[0].replace("(1. Okuma)", "").replace("1st", "").strip()
                        c2, c3 = None, None
                        for k in max_puan_sozlugu.keys():
                            if "(2. Okuma)" in k or "2nd" in k:
                                if k.split(" - ")[0].replace("(2. Okuma)", "").replace("2nd", "").strip() == ana_base: c2 = k
                            if "(3. Okuma)" in k or "3rd" in k:
                                if k.split(" - ")[0].replace("(3. Okuma)", "").replace("3rd", "").strip() == ana_base: c3 = k
                        if c2: cift_okumalar.append((c1, c2, c3))
                        
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
                            if c1 in max_puan_sozlugu and c2 in max_puan_sozlugu:
                                v1 = df_to_save.at[idx, c1]
                                v2 = df_to_save.at[idx, c2]
                                
                                has_v3 = False
                                v3 = 0.0
                                if c3:
                                    v3_raw = str(girilen_notlar_df.at[idx, c3]).strip() if c3 in girilen_notlar_df.columns else ""
                                    has_v3 = v3_raw not in ["", "nan", "None", "<NA>"]
                                    v3 = df_to_save.at[idx, c3]
                                
                                if abs(v1 - v2) > tolerans:
                                    if not has_v3 and c3 in max_puan_sozlugu:
                                        hatalar.append(f"**{row.get('Ogrenci_No', '')}**: {c1.split('(')[0].strip()} okumalar arası fark ({abs(v1-v2)}), {tolerans} puanlık toleransı aşıyor! 3. Okuma girilmelidir.")
                                    elif has_v3:
                                        diff1 = abs(v3 - v1)
                                        diff2 = abs(v3 - v2)
                                        if diff1 < diff2: ogrenci_toplam += (v3 + v1) / 2.0
                                        elif diff2 < diff1: ogrenci_toplam += (v3 + v2) / 2.0
                                        else: ogrenci_toplam += (v3 + max(v1, v2)) / 2.0
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
