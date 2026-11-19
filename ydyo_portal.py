import streamlit as st
import pandas as pd
from db_utils import (
    get_logo_b64,
    hash_sifre,
    sifre_dogrula,
    get_db_connection,
    check_and_update_schema,
    get_db_backup_excel,
)
from admin_analitik import render_analitik_raporlar
from hoca_paneli import render_hoca_paneli

st.set_page_config(page_title="KTO Karatay YDY Portalı", page_icon="🎓", layout="wide")

def kurumsal_tema_uygula():
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;700&family=Jost:wght@400;500;600&display=swap');

    html, body, [class*="css"] { font-family: 'DM Sans', sans-serif !important; }
    h1, h2, h3, h4, h5, h6 { font-family: 'Jost', sans-serif !important; color: #1A365D !important; }
    .stApp { background-color: #F4F6F9; }
    [data-testid="stSidebar"] { background-color: #1A365D; }
    [data-testid="stSidebar"] * { color: #FFFFFF !important; }

    /* Sol Menü Buton & Input Düzeltmesi */
    [data-testid="stSidebar"] button {
        background-color: rgba(255, 255, 255, 0.12) !important;
        border: 1px solid rgba(255, 255, 255, 0.28) !important;
        color: #FFFFFF !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        transition: all 0.2s ease;
    }
    [data-testid="stSidebar"] button:hover {
        background-color: #E84C0F !important;
        border-color: #E84C0F !important;
        color: #FFFFFF !important;
    }
    [data-testid="stSidebar"] input {
        color: #1E293B !important;
    }

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
    .login-card-header img { max-height: 54px; margin-bottom: 12px; }
    .login-card-title { font-family: 'DM Sans', sans-serif; font-size: 20px; font-weight: 700; color: #FFFFFF; margin: 0; letter-spacing: 0.4px; }
    .login-card-sub { font-family: 'Jost', sans-serif; font-size: 13px; color: #CBD5E1; margin-top: 4px; }

    /* Kompakt Profesyonel Üst Bant (Admin Opsiyon A & Hoca Paneli C) */
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
    .c-header-title { font-family: 'DM Sans', sans-serif; font-size: 18px; font-weight: 700; color: #FFFFFF; margin: 0; }
    .c-header-sub { font-family: 'Jost', sans-serif; font-size: 12.5px; color: #CBD5E1; margin-top: 2px; }
    .c-badge-open { background-color: rgba(22, 163, 74, 0.2); color: #86EFAC; border: 1px solid #22C55E; padding: 5px 12px; border-radius: 20px; font-size: 12px; font-weight: 600; }
    .c-badge-locked { background-color: rgba(220, 38, 38, 0.25); color: #FCA5A5; border: 1px solid #EF4444; padding: 5px 12px; border-radius: 20px; font-size: 12px; font-weight: 600; }
    .c-badge-role { background-color: rgba(232, 76, 15, 0.22); color: #FDBA74; border: 1px solid #E84C0F; padding: 5px 12px; border-radius: 20px; font-size: 12px; font-weight: 600; }

    /* Kurumsal KPI Kartları */
    .c-kpi-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; margin-bottom: 14px; }
    .c-kpi-card { background: #FFFFFF; border: 1px solid #E2E8F0; border-top: 3px solid #1A365D; border-radius: 8px; padding: 12px 16px; box-shadow: 0 1px 4px rgba(0,0,0,0.04); }
    .c-kpi-card.accent { border-top-color: #E84C0F; }
    .c-kpi-lbl { font-size: 11px; font-weight: 700; color: #64748B; text-transform: uppercase; letter-spacing: 0.5px; }
    .c-kpi-val { font-family: 'Jost', sans-serif; font-size: 22px; font-weight: 600; color: #1A365D; margin-top: 2px; }
    .c-kpi-card.accent .c-kpi-val { color: #E84C0F; }

    /* Analitik Bölüm Başlıkları */
    .analytics-section-title { font-family: 'Jost', sans-serif; font-size: 19px; font-weight: 600; color: #1A365D; margin-bottom: 10px; }
    .analytics-sub-title { font-family: 'DM Sans', sans-serif; font-size: 14px; font-weight: 700; color: #1A365D; margin: 14px 0 8px 0; border-left: 3px solid #E84C0F; padding-left: 8px; }

    .sidebar-profile { background: rgba(255, 255, 255, 0.08); border: 1px solid rgba(255, 255, 255, 0.14); border-radius: 10px; padding: 12px 14px; margin-bottom: 14px; }
    .sidebar-role-badge { display: inline-block; background: #E84C0F; color: #FFFFFF; font-size: 11px; font-weight: 600; padding: 2px 9px; border-radius: 12px; margin-top: 5px; }
    </style>
    """, unsafe_allow_html=True)

kurumsal_tema_uygula()
check_and_update_schema()

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
    else:
        st.error("Kullanıcı adı veya şifre hatalı!")

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
    if st.sidebar.button("🚪 Güvenli Çıkış", use_container_width=True):
        cikis_yap()

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
                else:
                    st.error("Mevcut şifreniz yanlış!")
                cur.close(); conn.close()

    if st.session_state['rol'] in ["Admin", "Testing"]:
        st.markdown(f"""
        <div class="c-header-bar">
            <div>
                <div class="c-header-title">KTO Karatay YDY Öğrenme Analitiği Sistemi</div>
                <div class="c-header-sub">Department of Foreign Languages | Testing and Evaluation Office</div>
            </div>
            <div><span class="c-badge-role">{st.session_state['rol']} Paneli</span></div>
        </div>
        """, unsafe_allow_html=True)

        if st.session_state['rol'] == "Admin":
            tabs = st.tabs(["📝 Sınav Tanımla", "🔒 Sınav & Veri Yönetimi", "📋 Öğrenci & Atamalar", "👥 Kullanıcılar", "📊 Analitik Raporlar"])
            tab_sinav, tab_yonetim, tab_atamalar, tab_kullanici, tab_rapor = tabs[0], tabs[1], tabs[2], tabs[3], tabs[4]
        else:
            tabs = st.tabs(["📝 Sınav Tanımla", "🔒 Sınav & Veri Yönetimi", "📊 Analitik Raporlar"])
            tab_sinav, tab_yonetim, tab_rapor = tabs[0], tabs[1], tabs[2]

        ANA_BOLUMLER = ["Listening I", "Listening II", "Reading I", "Reading II", "Grammar", "Functional Language", "Use of English", "Vocabulary", "Speaking", "Speaking 1st", "Speaking 2nd", "Speaking 3rd", "Writing", "Writing (1. Okuma)", "Writing (2. Okuma)", "Writing (3. Okuma)"]

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
                except Exception:
                    st.error("Format uyumsuz.")

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
            if "Çoklu" in kitapcik_turu:
                col_config_tos["B Kitapçığındaki Yeri (İsteğe Bağlı)"] = st.column_config.TextColumn("B Kitapçığındaki Yeri")
            else:
                col_config_tos["B Kitapçığındaki Yeri (İsteğe Bağlı)"] = None

            duzenlenen_rubrik = st.data_editor(st.session_state['taslak_df'], num_rows="dynamic", column_config=col_config_tos, use_container_width=True)

            if st.button("🚀 Sınavı Onayla & Öğretim Elemanlarına Aç", type="primary"):
                conn = get_db_connection()
                cur = conn.cursor()
                cur.execute("SELECT sinav_kodu FROM sinavlar WHERE sinav_kodu=%s", (sinav_kodu,))
                if cur.fetchone():
                    st.warning("Bu sınav kodu zaten açık!")
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
                if "Çoklu" in yeni_kturu:
                    edit_col_config["B Kitapçığındaki Yeri (İsteğe Bağlı)"] = st.column_config.TextColumn("B Kitapçığındaki Yeri")
                else:
                    edit_col_config["B Kitapçığındaki Yeri (İsteğe Bağlı)"] = None

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
            else:
                st.info("Sistemde kayıtlı sınav yok.")

            if st.session_state['rol'] == "Admin":
                st.divider()
                st.markdown("### 💾 Felaket Kurtarma Yedeği")
                st.download_button(label="📦 Tüm Veritabanını Excel Olarak İndir", data=get_db_backup_excel(), file_name="YDYO_Tam_Sistem_Yedegi.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
            conn.close()

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
                else:
                    st.info("Henüz yapılmış bir sınıf ataması yok.")

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

        with tab_rapor:
            render_analitik_raporlar()

    elif st.session_state['rol'] == "Öğretim Elemanı":
        render_hoca_paneli()
