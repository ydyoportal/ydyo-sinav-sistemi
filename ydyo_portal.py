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
            df = pd.read_sql_query(f"SELECT * FROM {t}", conn)
            df.to_excel(writer, sheet_name=t, index=False)
    conn.close()
    buffer.seek(0)
    return buffer

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
        if st.button("Şifreyi Güncelle"):
            if eski_sifre and yeni_sifre_1:
                conn = get_db_connection()
                cur = conn.cursor()
                cur.execute("SELECT sifre FROM kullanicilar WHERE kullanici_adi=%s", (st.session_state['kullanici_adi'],))
                if sifre_dogrula(eski_sifre, cur.fetchone()[0]):
                    cur.execute("UPDATE kullanicilar SET sifre=%s WHERE kullanici_adi=%s", (hash_sifre(yeni_sifre_1), st.session_state['kullanici_adi']))
                    conn.commit()
                    st.success("Şifreniz güncellendi!")
                else: st.error("Mevcut şifreniz yanlış!")
                cur.close(); conn.close()
        
    st.title("🎓 KTO Karatay YDY Öğrenme Analitiği Sistemi")

    if st.session_state['rol'] in ["Admin", "Testing"]:
        tabs = st.tabs(["📝 Sınav Tanımla", "🔒 Sınav & Veri Yönetimi", "📋 Öğrenci & Atamalar", "👥 Kullanıcılar", "📊 Analitik Raporlar"]) if st.session_state['rol'] == "Admin" else st.tabs(["📝 Sınav Tanımla", "🔒 Sınav & Veri Yönetimi", "📊 Analitik Raporlar"])
        tab_sinav = tabs[0]
        tab_yonetim = tabs[1]
        
        # 1. Sınav Tanımla Modülü (Öncekiyle Aynı)
        with tab_sinav:
            st.subheader("Sınav & ToS Oluşturucu")
            # ... (Bu kısımlar önceki yapınızla aynı kalır, yer kaplamaması için kodun odak noktası raporlara alındı)
            st.info("Sınav tanımlama alanı aktiftir.")

        # 2. Sınav Yönetimi (Kilit & Yedek)
        with tab_yonetim:
            st.markdown("### 🔒 Sınav Kilitleme")
            conn = get_db_connection()
            df_sinav = pd.read_sql_query("SELECT sinav_kodu, kilitli FROM sinavlar", conn)
            if not df_sinav.empty:
                sil_sinav = st.selectbox("İşlem Yapılacak Sınav", df_sinav['sinav_kodu'].tolist())
                mevcut_durum = df_sinav[df_sinav['sinav_kodu'] == sil_sinav]['kilitli'].iloc[0]
                st.write(f"**Durum:** {'🔒 KİLİTLİ' if mevcut_durum else '🟢 AÇIK'}")
                if mevcut_durum and st.button("🔓 Kilidi Aç"): 
                    conn.cursor().execute("UPDATE sinavlar SET kilitli=FALSE WHERE sinav_kodu=%s", (sil_sinav,)); conn.commit(); st.rerun()
                elif not mevcut_durum and st.button("🔒 Kilitle", type="primary"):
                    conn.cursor().execute("UPDATE sinavlar SET kilitli=TRUE WHERE sinav_kodu=%s", (sil_sinav,)); conn.commit(); st.rerun()
            if st.session_state['rol'] == "Admin":
                st.divider()
                st.download_button("📦 Tam Veritabanını İndir (Excel)", data=get_db_backup_excel(), file_name="YDYO_Tam_Yedek.xlsx")
            conn.close()

        # Atamalar ve Kullanıcılar (Sadece Admin)
        if st.session_state['rol'] == "Admin":
            with tabs[2]: st.info("Atama modülü aktif.")
            with tabs[3]: st.info("Kullanıcı yönetimi aktif.")

        # 3. YENİ PAKET: GELİŞMİŞ ANALİTİK RAPORLAR (TESTING)
        tab_rapor = tabs[4] if st.session_state['rol'] == "Admin" else tabs[2]
        with tab_rapor:
            st.markdown("### 🔬 Gelişmiş Ölçme & Değerlendirme Analitiği")
            rapor_kapsami = st.radio("Analiz Türü:", [
                "🏆 Dönem Sonu Konsolide Başarı", 
                "🎯 Madde Ayırt Edicilik (Item Discrimination)", 
                "⚖️ Puanlayıcı Güvenirliği (Inter-Rater Reliability)"
            ], horizontal=True)
            st.divider()
            
            conn = get_db_connection()
            cur = conn.cursor()
            
            # -- MADDE AYIRT EDİCİLİK ANALİZİ --
            if rapor_kapsami == "🎯 Madde Ayırt Edicilik (Item Discrimination)":
                st.info("Bu analiz, sınavdaki her bir sorunun/bölümün başarılı öğrencilerle (Üst %27) başarısız öğrencileri (Alt %27) ne kadar iyi ayırabildiğini gösterir.")
                cur.execute("SELECT sinav_kodu FROM sinavlar")
                sinavlar = [r[0] for r in cur.fetchall()]
                if not sinavlar: st.warning("Sınav bulunamadı.")
                else:
                    secilen_sinav = st.selectbox("Analiz Edilecek Sınav", sinavlar)
                    if st.button("Analizi Başlat", type="primary"):
                        # Sınavın ToS verisini (max puanlar) çek
                        df_tos = pd.read_sql_query("SELECT ana_bolum, alt_bolum, max_puan FROM sinav_tos WHERE sinav_kodu=%s", conn, params=(secilen_sinav,))
                        max_puanlar = {f"{r['ana_bolum']} - {r['alt_bolum']}": float(r['max_puan']) for _, r in df_tos.iterrows()}
                        
                        # Öğrenci notlarını çek ve birleştir
                        cur.execute("SELECT veri_json FROM notlar WHERE sinav_kodu=%s", (secilen_sinav,))
                        json_verileri = cur.fetchall()
                        tum_ogrenciler = []
                        for j in json_verileri:
                            df_temp = pd.read_json(io.StringIO(j[0]))
                            tum_ogrenciler.extend(df_temp.to_dict('records'))
                        
                        df_hepsi = pd.DataFrame(tum_ogrenciler)
                        if "Sınav Notu" not in df_hepsi.columns or len(df_hepsi) < 10:
                            st.warning("Sağlıklı bir %27 analizi için yeterli öğrenci verisi yok (En az 10 öğrenci önerilir).")
                        else:
                            # Üst ve Alt Grupları Belirle
                            df_hepsi = df_hepsi.sort_values(by="Sınav Notu", ascending=False)
                            n_27 = max(1, int(len(df_hepsi) * 0.27))
                            df_ust = df_hepsi.head(n_27)
                            df_alt = df_hepsi.tail(n_27)
                            
                            sonuclar = []
                            for bolum, max_p in max_puanlar.items():
                                if bolum in df_hepsi.columns:
                                    ust_ortalama = df_ust[bolum].mean()
                                    alt_ortalama = df_alt[bolum].mean()
                                    
                                    # Puanları 100'lük (Yüzdelik) formata çevirerek oran bul
                                    p_ust = ust_ortalama / max_p if max_p > 0 else 0
                                    p_alt = alt_ortalama / max_p if max_p > 0 else 0
                                    d_index = p_ust - p_alt
                                    
                                    # Pedagojik Yorumlama
                                    if d_index >= 0.40: durum, renk = "Çok İyi Madde", "#27AE60"
                                    elif d_index >= 0.30: durum, renk = "İyi Madde (Geliştirilebilir)", "#F1C40F"
                                    elif d_index >= 0.20: durum, renk = "Zayıf Madde (Düzeltilmeli)", "#E67E22"
                                    else: durum, renk = "Çöpe Atılmalı", "#E74C3C"
                                    
                                    sonuclar.append({
                                        "Bölüm": bolum,
                                        "Max Puan": max_p,
                                        "Üst %27 Ort.": round(ust_ortalama, 2),
                                        "Alt %27 Ort.": round(alt_ortalama, 2),
                                        "Ayırt Edicilik (D)": round(d_index, 3),
                                        "Durum Yorumu": durum
                                    })
                            
                            df_sonuc = pd.DataFrame(sonuclar)
                            def color_d(val):
                                if "Çok İyi" in val: return 'color: #27AE60; font-weight: bold;'
                                elif "Çöpe" in val: return 'background-color: #FADBD8; color: black;'
                                elif "Zayıf" in val: return 'color: #E67E22;'
                                return ''
                            
                            st.markdown(f"**Toplam Öğrenci:** {len(df_hepsi)} | **Grup Büyüklüğü (%27):** {n_27} Öğrenci")
                            st.dataframe(df_sonuc.style.map(color_d, subset=['Durum Yorumu']), use_container_width=True)

            # -- PUANLAYICI GÜVENİRLİĞİ (INTER-RATER) --
            elif rapor_kapsami == "⚖️ Puanlayıcı Güvenirliği (Inter-Rater Reliability)":
                st.info("Writing ve Speaking gibi çift okuma gerektiren becerilerde 1. ve 2. okuma arasındaki uyumu (Korelasyon) ve sapmaları analiz eder.")
                cur.execute("SELECT sinav_kodu FROM sinavlar")
                sinavlar = [r[0] for r in cur.fetchall()]
                if not sinavlar: st.warning("Sınav bulunamadı.")
                else:
                    secilen_sinav_ir = st.selectbox("Sınav Seç", sinavlar, key="ir_sinav")
                    
                    if st.button("Güvenirlik Analizi Çıkar", type="primary"):
                        cur.execute("SELECT veri_json FROM notlar WHERE sinav_kodu=%s", (secilen_sinav_ir,))
                        json_verileri = cur.fetchall()
                        tum_ogrenciler = []
                        for j in json_verileri:
                            tum_ogrenciler.extend(pd.read_json(io.StringIO(j[0])).to_dict('records'))
                        
                        df_ir = pd.DataFrame(tum_ogrenciler)
                        # "1. Okuma" ve "2. Okuma" içeren sütun gruplarını bul
                        kolonlar = df_ir.columns.tolist()
                        okuma_1_kolonlari = [c for c in kolonlar if "(1. Okuma)" in c]
                        
                        if not okuma_1_kolonlari:
                            st.warning("Bu sınavda '(1. Okuma)' etiketi taşıyan bir değerlendirme sütunu bulunamadı.")
                        else:
                            for ok1_col in okuma_1_kolonlari:
                                ok2_col = ok1_col.replace("(1. Okuma)", "(2. Okuma)")
                                beceri_adi = ok1_col.replace(" - (1. Okuma)", "").replace("(1. Okuma)", "")
                                
                                if ok2_col in kolonlar:
                                    st.markdown(f"#### 📊 Analiz Edilen Beceri: {beceri_adi}")
                                    df_cift = df_ir[['Ad_Soyad', ok1_col, ok2_col]].copy().dropna()
                                    df_cift['Puan Farkı (Mutlak)'] = abs(df_cift[ok1_col] - df_cift[ok2_col])
                                    
                                    if len(df_cift) > 2:
                                        # Pearson Korelasyonu (r)
                                        r_degeri = df_cift[ok1_col].corr(df_cift[ok2_col])
                                        r_durum = "Yüksek Uyum 🟢" if r_degeri >= 0.70 else "Kabul Edilebilir 🟡" if r_degeri >= 0.50 else "Zayıf Uyum 🔴 (Gözden Geçirilmeli)"
                                        
                                        # Ortalama Eğilimleri (Lenient vs Harsh)
                                        ort_1 = df_cift[ok1_col].mean()
                                        ort_2 = df_cift[ok2_col].mean()
                                        
                                        c_m1, c_m2, c_m3 = st.columns(3)
                                        c_m1.metric("Korelasyon (r)", f"{r_degeri:.2f}", r_durum)
                                        c_m2.metric("1. Okuma Sınıf Ortalaması", f"{ort_1:.2f}")
                                        c_m3.metric("2. Okuma Sınıf Ortalaması", f"{ort_2:.2f}")
                                        
                                        st.write("**⚠️ Yüksek Sapma Gösteren Kâğıtlar (Fark > 10 Puan):**")
                                        sapmalar = df_cift[df_cift['Puan Farkı (Mutlak)'] > 10].sort_values(by="Puan Farkı (Mutlak)", ascending=False)
                                        if not sapmalar.empty:
                                            st.dataframe(sapmalar, use_container_width=True)
                                        else: st.success("10 Puan üzerinde sapma olan kağıt bulunmamaktadır.")
                                        st.divider()

            elif rapor_kapsami == "🏆 Dönem Sonu Konsolide Başarı":
                st.info("Konsolide modülü aktif.")
            
            cur.close(); conn.close()

    # 4. Hoca Veri Giriş Paneli (Kilit Kontrollü)
    elif st.session_state['rol'] == "Öğretim Elemanı":
        st.subheader("📝 Hızlı Veri Giriş Terminali")
        st.info("Hoca modülü aktif (Önceki kodunuz ile aynı).")
