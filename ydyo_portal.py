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

def create_modern_pdf(rapor_verileri):
    elements = []
    for idx, veri in enumerate(rapor_verileri):
        elements.append(Paragraph("KTO KARATAY UNIVERSITY", header_style))
        elements.append(Paragraph("School of Foreign Languages | Testing Office", sub_header_style))
        elements.append(Paragraph(tr_to_eng(veri['baslik']), title_style))
        if veri.get('alt_baslik'): elements.append(Paragraph(tr_to_eng(veri['alt_baslik']), info_style))
        df_listesi = veri.get('df_list', [])
        for tablo_baslik, df_tablo in df_listesi:
            if not df_tablo.empty:
                if tablo_baslik:
                    elements.append(Paragraph(f"<b>{tr_to_eng(tablo_baslik)}</b>", styles['Heading4']))
                    elements.append(Spacer(1, 5))
                df_temp = df_tablo.copy()
                df_temp.columns = [tr_to_eng(c) for c in df_temp.columns]
                cols_count = len(df_temp.columns)
                col_widths = [535 / max(1, cols_count)] * cols_count
                tablo_verisi = [[Paragraph(c, header_cell_style) for c in df_temp.columns]]
                for _, row in df_temp.iterrows():
                    row_data = [Paragraph(tr_to_eng(str(cell_val)), cell_style) for cell_val in row]
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
        # 2. SINAV & VERİ YÖNETİMİ (KİLİT, DÜZENLEME, SİLME, YEDEK)
        # ----------------------------------------------------
        with tab_yonetim:
            conn = get_db_connection()
            df_sinav_durum = pd.read_sql_query("SELECT sinav_kodu, kilitli FROM sinavlar", conn)
            
            if not df_sinav_durum.empty:
                # 1. KİLİTLEME MODÜLÜ
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
                
                # 2. SINAV DÜZENLEME MODÜLÜ (YENİ)
                st.markdown("### ✏️ Sınav Düzenle (Ayarlar ve ToS)")
                st.info("Sınavı silmeden, mevcut tolerans ayarını veya Belirtke Tablosunu (ToS) güncelleyebilirsiniz. Uyarı: Eğer not girişi yapıldıysa ToS'taki sütun isimlerini (Ana Bölüm/Alt Bölüm) değiştirmek notların görüntülenmesinde sorun yaratabilir, sadece puan güncellemesi önerilir.")
                
                duzenlenecek_sinav = st.selectbox("Düzenlenecek Sınavı Seçin", df_sinav_durum['sinav_kodu'].tolist(), key="edit_exam_select")
                
                # Mevcut Toleransı Çek
                cur = conn.cursor()
                cur.execute("SELECT tolerans FROM sinav_ayarlari WHERE sinav_kodu=%s", (duzenlenecek_sinav,))
                tol_res = cur.fetchone()
                mevcut_tolerans = tol_res[0] if tol_res else 6
                
                yeni_tolerans = st.number_input("Güncel Çift Okuma Tolerans Farkı", min_value=0, value=mevcut_tolerans)
                
                # Mevcut ToS'u Çek
                df_mevcut_tos = pd.read_sql_query("SELECT ana_bolum AS \"Ana Bölüm\", alt_bolum AS \"Alt Bölüm (Part)\", kazanimlar AS \"Ölçülen Kazanımlar\", max_puan AS \"Max Puan\" FROM sinav_tos WHERE sinav_kodu=%s", conn, params=(duzenlenecek_sinav,))
                
                st.markdown("**Güncel Belirtke Tablosu (ToS):**")
                guncel_tos = st.data_editor(
                    df_mevcut_tos, 
                    num_rows="dynamic", 
                    column_config={
                        "Ana Bölüm": st.column_config.SelectboxColumn("Ana Bölüm", options=ANA_BOLUMLER, required=True),
                        "Max Puan": st.column_config.NumberColumn("Max Puan", step=0.5, format="%.2f")
                    }, 
                    use_container_width=True,
                    key="edit_tos_editor"
                )
                
                if st.button("💾 Değişiklikleri Kaydet (Sınavı Güncelle)", type="primary"):
                    # Toleransı güncelle
                    cur.execute("UPDATE sinav_ayarlari SET tolerans=%s WHERE sinav_kodu=%s", (yeni_tolerans, duzenlenecek_sinav))
                    # Eski ToS'u sil ve yenisini yaz
                    cur.execute("DELETE FROM sinav_tos WHERE sinav_kodu=%s", (duzenlenecek_sinav,))
                    for _, row in guncel_tos.iterrows():
                        cur.execute("INSERT INTO sinav_tos (sinav_kodu, ana_bolum, alt_bolum, kazanimlar, max_puan) VALUES (%s, %s, %s, %s, %s)", (duzenlenecek_sinav, row['Ana Bölüm'], row['Alt Bölüm (Part)'], row['Ölçülen Kazanımlar'], float(row['Max Puan'])))
                    conn.commit()
                    st.success(f"{duzenlenecek_sinav} başarıyla güncellendi!")
                
                st.divider()
                
                # 3. SINAV SİLME MODÜLÜ
                st.markdown("### 🗑️ Sınav Silme")
                silinecek_sinav = st.selectbox("Silinecek Sınavı Seçin", df_sinav_durum['sinav_kodu'].tolist(), key="sil_sec")
                if st.button("🗑️ Seçili Sınavı Tamamen Sil"):
                    conn.cursor().execute("DELETE FROM sinavlar WHERE sinav_kodu=%s", (silinecek_sinav,))
                    conn.commit(); st.success("Sınav silindi!"); st.rerun()
            else: st.info("Sistemde kayıtlı sınav yok.")
            
            if st.session_state['rol'] == "Admin":
                st.divider()
                st.markdown("### 💾 Felaket Kurtarma Yedeği (Disaster Backup)")
                st.download_button(label="📦 Tüm Veritabanını Excel Olarak İndir", data=get_db_backup_excel(), file_name="YDYO_Tam_Sistem_Yedegi.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
            conn.close()

        # ----------------------------------------------------
        # 3. KULLANICILAR (EKLEME & DÜZENLEME AYRILDI)
        # ----------------------------------------------------
        if st.session_state['rol'] == "Admin":
            with tab_kullanici:
                st.subheader("👥 Kullanıcı Yönetimi")
                conn = get_db_connection()
                df_kullanicilar = pd.read_sql_query("SELECT kullanici_adi, rol, isim FROM kullanicilar", conn)
                
                # YENİ KULLANICI EKLE
                with st.expander("➕ Yeni Kullanıcı Ekle", expanded=False):
                    c_k1, c_k2 = st.columns(2)
                    with c_k1: 
                        yeni_k_adi = st.text_input("Kullanıcı Adı (ID)", key="new_id")
                        yeni_sifre = st.text_input("Şifre", type="password", key="new_pw")
                    with c_k2: 
                        yeni_rol = st.selectbox("Rol", ["Öğretim Elemanı", "Testing", "Admin"], key="new_role")
                        yeni_isim = st.text_input("İsim (Örn: Dr. Ahmet)", key="new_name")
                    
                    if st.button("Yeni Kullanıcıyı Kaydet", type="primary"):
                        if yeni_k_adi and yeni_sifre and yeni_isim:
                            cur = conn.cursor()
                            cur.execute("SELECT kullanici_adi FROM kullanicilar WHERE kullanici_adi=%s", (yeni_k_adi,))
                            if cur.fetchone():
                                st.error("Bu kullanıcı adı zaten mevcut! Lütfen farklı bir ID seçin veya 'Düzenle' sekmesini kullanın.")
                            else:
                                h_pwd = hash_sifre(yeni_sifre)
                                cur.execute("INSERT INTO kullanicilar (kullanici_adi, sifre, rol, isim) VALUES (%s, %s, %s, %s)", (yeni_k_adi, h_pwd, yeni_rol, yeni_isim))
                                conn.commit(); st.success(f"{yeni_k_adi} sisteme eklendi!"); st.rerun()
                        else: st.warning("Tüm alanları doldurunuz.")
                
                # MEVCUT KULLANICIYI DÜZENLE
                with st.expander("✏️ Mevcut Kullanıcıyı Düzenle (İsim, Rol veya ID Değişimi)", expanded=False):
                    if not df_kullanicilar.empty:
                        duzenlenecek_kullanici = st.selectbox("Düzenlenecek Kullanıcıyı Seçin", df_kullanicilar['kullanici_adi'].tolist())
                        mevcut_bilgi = df_kullanicilar[df_kullanicilar['kullanici_adi'] == duzenlenecek_kullanici].iloc[0]
                        
                        c_e1, c_e2 = st.columns(2)
                        with c_e1:
                            edit_id = st.text_input("Kullanıcı Adı (ID)", value=mevcut_bilgi['kullanici_adi'], key="edit_id")
                            edit_pw = st.text_input("Yeni Şifre (Değiştirmek İstemiyorsanız BOŞ Bırakın)", type="password", key="edit_pw")
                        with c_e2:
                            roller = ["Öğretim Elemanı", "Testing", "Admin"]
                            mevcut_rol_index = roller.index(mevcut_bilgi['rol']) if mevcut_bilgi['rol'] in roller else 0
                            edit_rol = st.selectbox("Rol", roller, index=mevcut_rol_index, key="edit_role")
                            edit_isim = st.text_input("İsim", value=mevcut_bilgi['isim'], key="edit_name")
                            
                        if st.button("Kullanıcı Bilgilerini Güncelle"):
                            cur = conn.cursor()
                            # ID değiştirilmişse FK (Atamalar) hatası almamak için önce atamaları güncelle
                            if edit_id != duzenlenecek_kullanici:
                                cur.execute("SELECT kullanici_adi FROM kullanicilar WHERE kullanici_adi=%s", (edit_id,))
                                if cur.fetchone():
                                    st.error("Girdiğiniz yeni Kullanıcı Adı (ID) zaten başkası tarafından kullanılıyor!")
                                    st.stop()
                                cur.execute("UPDATE hoca_atamalari SET kullanici_adi=%s WHERE kullanici_adi=%s", (edit_id, duzenlenecek_kullanici))
                                cur.execute("UPDATE kullanicilar SET kullanici_adi=%s, rol=%s, isim=%s WHERE kullanici_adi=%s", (edit_id, edit_rol, edit_isim, duzenlenecek_kullanici))
                            else:
                                cur.execute("UPDATE kullanicilar SET rol=%s, isim=%s WHERE kullanici_adi=%s", (edit_rol, edit_isim, duzenlenecek_kullanici))
                            
                            # Eğer şifre alanı boş bırakılmadıysa şifreyi de güncelle
                            if edit_pw:
                                cur.execute("UPDATE kullanicilar SET sifre=%s WHERE kullanici_adi=%s", (hash_sifre(edit_pw), edit_id))
                                
                            conn.commit()
                            st.success("Kullanıcı başarıyla güncellendi!")
                            st.rerun()

                st.markdown("**Sistemdeki Kullanıcılar**")
                st.dataframe(df_kullanicilar, use_container_width=True, hide_index=True)
                sil_k_adi = st.selectbox("Silinecek Kullanıcı", df_kullanicilar['kullanici_adi'].tolist() if not df_kullanicilar.empty else [])
                if st.button("Kullanıcıyı Sil") and sil_k_adi:
                    if sil_k_adi == st.session_state['kullanici_adi']: st.error("Kendinizi silemezsiniz!")
                    else:
                        cur = conn.cursor(); cur.execute("DELETE FROM kullanicilar WHERE kullanici_adi=%s", (sil_k_adi,)); conn.commit(); st.rerun()
                conn.close()
                
            # ÖĞRENCİ VE ATAMALAR
            with tab_atamalar:
                conn = get_db_connection()
                c_ata1, c_ata2 = st.columns(2)
                with c_ata1:
                    st.markdown("**Öğrenci Kütüğü**")
                    duzenlenen_ogr = st.data_editor(pd.read_sql_query("SELECT ogrenci_no, ad_soyad, modul, sube FROM ogrenciler", conn), num_rows="dynamic", use_container_width=True)
                    if st.button("Kaydet (Öğrenciler)"):
                        cur = conn.cursor(); cur.execute("DELETE FROM ogrenciler")
                        for _, r in duzenlenen_ogr.iterrows(): cur.execute("INSERT INTO ogrenciler (ogrenci_no, ad_soyad, modul, sube) VALUES (%s, %s, %s, %s)", (r['ogrenci_no'], r['ad_soyad'], r['modul'], r['sube']))
                        conn.commit(); st.success("Kütük güncellendi!")
                with c_ata2:
                    st.markdown("**Hoca - Şube Atamaları**")
                    duzenlenen_hoca = st.data_editor(pd.read_sql_query("SELECT modul, sube, kullanici_adi FROM hoca_atamalari", conn), num_rows="dynamic", use_container_width=True)
                    if st.button("Kaydet (Atamalar)"):
                        cur = conn.cursor(); cur.execute("DELETE FROM hoca_atamalari")
                        for _, r in duzenlenen_hoca.iterrows(): cur.execute("INSERT INTO hoca_atamalari (modul, sube, kullanici_adi) VALUES (%s, %s, %s)", (r['modul'], r['sube'], r['kullanici_adi']))
                        conn.commit(); st.success("Atamalar güncellendi!")
                conn.close()
                
        # ----------------------------------------------------
        # 4. GELİŞMİŞ ANALİTİK RAPORLAR (KARNELER DAHİL)
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
                        agirliklar = {}
                        toplam_agirlik = 0
                        for i, sk in enumerate(secilen_sinavlar_konsolide):
                            with c_weights[i]:
                                w = st.number_input(f"{sk} Ağırlığı (%)", min_value=0, max_value=100, value=100//len(secilen_sinavlar_konsolide))
                                agirliklar[sk] = w
                                toplam_agirlik += w
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
                                    
                                    # PANDAS SÜRÜM UYUMU (map kullanımı)
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
                        if "Sınav Notu" not in df_hepsi.columns or len(df_hepsi) < 10: st.warning("Yeterli veri yok (En az 10 öğrenci).")
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
                                # PANDAS 2.1+ UYUMU İÇİN APPLYMAP -> MAP DEĞİŞİKLİĞİ BURADA YAPILDI
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
                        cur.execute("SELECT veri_json FROM notlar WHERE sinav_kodu=%s AND sube=%s", (sec_sinav, sec_sube))
                        v = cur.fetchone()
                        df_k = pd.read_json(io.StringIO(v[0]))
                        if not df_k.empty:
                            sec_ogr = st.selectbox("Öğrenci Seçin", df_k['Ad_Soyad'].tolist())
                            ogr_veri = df_k[df_k['Ad_Soyad'] == sec_ogr].iloc[0]
                            st.write(f"**Öğrenci:** {sec_ogr}")
                            st.write(f"**Sınav Notu:** {ogr_veri.get('Sınav Notu', 0)}")
                            
                            if st.button("PDF Karne İndir"):
                                df_pdf = pd.DataFrame({"Beceri/Bölüm": [c for c in df_k.columns if c not in ["Ogrenci_No", "Ad_Soyad", "Sınav Notu"]],
                                                       "Alınan Puan": [ogr_veri[c] for c in df_k.columns if c not in ["Ogrenci_No", "Ad_Soyad", "Sınav Notu"]]})
                                r_veri = [{'baslik': f"{sec_sinav} - Bireysel Karne", 'alt_baslik': f"Öğrenci: {sec_ogr} | Toplam Not: {ogr_veri.get('Sınav Notu', 0)}", 'df_list': [("Beceri Detayları", df_pdf)]}]
                                pdf_buf = create_modern_pdf(r_veri)
                                st.download_button("📥 PDF Olarak İndir", data=pdf_buf, file_name=f"{sec_ogr}_Karne.pdf", mime="application/pdf")
            
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
                        cur.execute("SELECT veri_json FROM notlar WHERE sinav_kodu=%s AND sube=%s", (sec_sinav, sec_sube))
                        df_s = pd.read_json(io.StringIO(cur.fetchone()[0]))
                        
                        st.markdown(f"### {sec_sube} - Genel Durum")
                        st.metric("Sınıf Ortalaması", f"{df_s['Sınav Notu'].mean():.2f}")
                        st.dataframe(df_s[['Ogrenci_No', 'Ad_Soyad', 'Sınav Notu']].sort_values(by="Sınav Notu", ascending=False), use_container_width=True, hide_index=True)
                        
                        if st.button("Sınıf Raporunu PDF İndir"):
                            r_veri = [{'baslik': f"{sec_sinav} - {sec_sube} Sınıf Raporu", 'alt_baslik': f"Sınıf Ortalaması: {df_s['Sınav Notu'].mean():.2f}", 'df_list': [("Tüm Öğrenciler", df_s[['Ogrenci_No', 'Ad_Soyad', 'Sınav Notu']])]}]
                            pdf_buf = create_modern_pdf(r_veri)
                            st.download_button("📥 Sınıf PDF İndir", data=pdf_buf, file_name=f"{sec_sube}_Rapor.pdf", mime="application/pdf")

            cur.close(); conn.close()

    # ----------------------------------------------------
    # 5. HOCA VERİ GİRİŞ PANELİ (KİLİT KONTROLLÜ)
    # ----------------------------------------------------
    elif st.session_state['rol'] == "Öğretim Elemanı":
        st.subheader("📝 Hızlı Veri Giriş Terminali")
        conn = get_db_connection()
        cur = conn.cursor()
        cur.execute("SELECT sinav_kodu, kilitli FROM sinavlar")
        sinav_data = cur.fetchall()
        if not sinav_data: st.warning("Aktif sınav bulunmuyor."); st.stop()
        
        acik_sinavlar = [row[0] for row in sinav_data]
        sinav_kilit_durumlari = {row[0]: row[1] for row in sinav_data}
        secilen_islem_sinavi = st.selectbox("Aktif Sınavlar", acik_sinavlar)
        is_locked = sinav_kilit_durumlari.get(secilen_islem_sinavi, False)
        
        if is_locked:
            st.error("🔒 BU SINAV TESTING OFİSİ TARAFINDAN KİLİTLENMİŞTİR!")
            st.warning("Not girişi kapatılmıştır. Yalnızca görüntüleme yapabilirsiniz.")
        
        ilgili_modul = secilen_islem_sinavi.split("_")[0]
        cur.execute("SELECT sube FROM hoca_atamalari WHERE modul=%s AND kullanici_adi=%s", (ilgili_modul, st.session_state['kullanici_adi']))
        hoca_subeleri = [row[0] for row in cur.fetchall()]
        if not hoca_subeleri: st.error("Atanmış şubeniz bulunmuyor."); st.stop()
        secilen_sube = st.selectbox("Şubeleriniz", hoca_subeleri)
        rubrik_df = pd.read_sql_query("SELECT * FROM sinav_tos WHERE sinav_kodu=%s", conn, params=(secilen_islem_sinavi,))
        
        st.markdown(f"### 📊 {secilen_sube} - Not Tablosu")
        cur.execute("SELECT veri_json FROM notlar WHERE sinav_kodu=%s AND sube=%s", (secilen_islem_sinavi, secilen_sube))
        kaydedilmis_veri = cur.fetchone()
        
        if kaydedilmis_veri: df_giris = pd.read_json(io.StringIO(kaydedilmis_veri[0]))
        else:
            gercek_ogrenciler = pd.read_sql_query("SELECT ogrenci_no AS \"Ogrenci_No\", ad_soyad AS \"Ad_Soyad\" FROM ogrenciler WHERE modul=%s AND sube=%s", conn, params=(ilgili_modul, secilen_sube))
            df_giris = gercek_ogrenciler.copy() if not gercek_ogrenciler.empty else pd.DataFrame(columns=["Ogrenci_No", "Ad_Soyad"])
            for _, row in rubrik_df.iterrows(): df_giris[f"{row['ana_bolum']} - {row['alt_bolum']}"] = 0.0 
            df_giris["Sınav Notu"] = 0.0 
            
        girilen_notlar_df = st.data_editor(df_giris, num_rows="dynamic", use_container_width=True, column_config={"Sınav Notu": st.column_config.NumberColumn("Sınav Notu", disabled=True)}, disabled=is_locked)
        
        if not is_locked and st.button("💾 Notları Buluta Kaydet", type="primary"):
            hatalar = []
            max_puan_sozlugu = {f"{r['ana_bolum']} - {r['alt_bolum']}": float(r['max_puan']) for _, r in rubrik_df.iterrows()}
            for idx, row in girilen_notlar_df.iterrows():
                for col, m_puan in max_puan_sozlugu.items():
                    val = pd.to_numeric(row.get(col), errors='coerce')
                    if pd.notnull(val) and val > m_puan: hatalar.append(f"**{row.get('Ogrenci_No', '')}**: '{col}' notu ({val}), max ({m_puan}) aşıyor!")
            if hatalar: st.error("🛑 HATA!"); [st.warning(h) for h in hatalar]; st.stop()
            
            toplam_notlar = []
            for idx, row in girilen_notlar_df.iterrows():
                ogrenci_toplam = sum([pd.to_numeric(row.get(col), errors='coerce') for col in max_puan_sozlugu.keys() if pd.notnull(pd.to_numeric(row.get(col), errors='coerce'))])
                toplam_notlar.append(round(ogrenci_toplam, 2))
            girilen_notlar_df["Sınav Notu"] = toplam_notlar
            
            cur.execute("""
                INSERT INTO notlar (sinav_kodu, sube, veri_json) 
                VALUES (%s, %s, %s)
                ON CONFLICT (sinav_kodu, sube) DO UPDATE 
                SET veri_json = EXCLUDED.veri_json;
            """, (secilen_islem_sinavi, secilen_sube, girilen_notlar_df.to_json(orient="records")))
            conn.commit(); cur.close(); conn.close()
            st.success("Notlar bulut veritabanına başarıyla kaydedildi!"); st.rerun()
