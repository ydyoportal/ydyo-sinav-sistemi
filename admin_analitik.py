import streamlit as st
import pandas as pd
import io
import zipfile
from db_utils import get_db_connection
from rapor_motoru import (
    tr_to_eng,
    karne_satirlari_hazirla,
    create_modern_pdf,
    create_corporate_class_excel,
)

def _kpi_cards_html(items):
    cols = len(items)
    cards = "".join(
        f'<div class="{"c-kpi-card accent" if is_accent else "c-kpi-card"}">'
        f'<div class="c-kpi-lbl">{lbl}</div>'
        f'<div class="c-kpi-val">{val}</div>'
        f'</div>'
        for lbl, val, is_accent in items
    )
    return f'<div class="c-kpi-grid" style="grid-template-columns: repeat({cols}, 1fr);">{cards}</div>'
def render_analitik_raporlar():
    st.markdown('<div class="analytics-section-title">🔬 Gelişmiş Ölçme &amp; Değerlendirme Analitiği</div>', unsafe_allow_html=True)
    rapor_kapsami = st.radio("Analiz Türü:", [
        "🏆 Dönem Sonu Konsolide Başarı",
        "🎯 Madde Ayırt Edicilik (Item Discrimination)",
        "⚖️ Puanlayıcı Güvenirliği (Inter-Rater Reliability)",
        "👤 Bireysel Karneler",
        "🏫 Sınıf / Şube Raporu (Excel Formatı)",
        "📡 Canlı Sınav Takip Radarı",
        "📥 OBS Excel Çıktısı"
    ], horizontal=True, label_visibility="collapsed")
    st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)

    conn = get_db_connection()
    cur = conn.cursor()

    if rapor_kapsami == "📡 Canlı Sınav Takip Radarı":
        cur.execute("SELECT sinav_kodu FROM sinavlar ORDER BY sinav_kodu")
        sinavlar = [r[0] for r in cur.fetchall()]
        if not sinavlar:
            st.warning("Kayıtlı sınav yok.")
        else:
            with st.container(border=True):
                sec_sinav = st.selectbox("Takip Edilecek Sınavı Seçin", sinavlar)
            ilgili_modul = sec_sinav.split("_")[0]

            cur.execute("SELECT h.sube, h.kullanici_adi, k.isim FROM hoca_atamalari h LEFT JOIN kullanicilar k ON h.kullanici_adi = k.kullanici_adi WHERE h.modul=%s", (ilgili_modul,))
            hoca_atamalari = cur.fetchall()

            if not hoca_atamalari:
                st.info("Bu modül için atanmış sınıf/hoca bulunmuyor.")
            else:
                st.markdown(f'<div class="analytics-sub-title">📡 {sec_sinav.replace("_", " ")} — Veri Giriş Durumları</div>', unsafe_allow_html=True)
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

                    yuzde = min(int((okunan_ogrenci / toplam_ogrenci) * 100), 100) if toplam_ogrenci > 0 else 0
                    with st.container(border=True):
                        col_text, col_bar = st.columns([1.2, 3])
                        col_text.markdown(f"**{sube}**  \n<span style='color:#64748B; font-size:12px;'>{hoca_isim or k_adi}</span>", unsafe_allow_html=True)
                        col_bar.progress(yuzde, text=f"{okunan_ogrenci} / {toplam_ogrenci} Öğrenci Okundu (%{yuzde})")

    elif rapor_kapsami == "📥 OBS Excel Çıktısı":
        cur.execute("SELECT sinav_kodu FROM sinavlar ORDER BY sinav_kodu")
        sinavlar = [r[0] for r in cur.fetchall()]
        if not sinavlar:
            st.warning("Kayıtlı sınav yok.")
        else:
            with st.container(border=True):
                sec_sinav = st.selectbox("OBS'ye Aktarılacak Sınavı Seçin", sinavlar)
                if st.button("🚀 OBS Formatında Excel Üret", type="primary"):
                    cur.execute("SELECT sube, veri_json FROM notlar WHERE sinav_kodu=%s", (sec_sinav,))
                    not_verileri = cur.fetchall()
                    if not not_verileri:
                        st.error("Bu sınava ait girilmiş hiçbir not bulunamadı!")
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
        cur.execute("SELECT sinav_kodu FROM sinavlar ORDER BY sinav_kodu")
        acik_sinavlar = [r[0] for r in cur.fetchall()]
        if not acik_sinavlar:
            st.warning("Sistemde sınav yok.")
        else:
            with st.container(border=True):
                secilen_sinavlar_konsolide = st.multiselect("Hesaplamaya Dahil Edilecek Sınavları Seçin", acik_sinavlar)
                if secilen_sinavlar_konsolide:
                    c_weights = st.columns(len(secilen_sinavlar_konsolide))
                    agirliklar = {}
                    toplam_agirlik = 0
                    for i, sk in enumerate(secilen_sinavlar_konsolide):
                        with c_weights[i]:
                            w = st.number_input(f"{sk} Ağırlığı (%)", min_value=0, max_value=100, value=100 // len(secilen_sinavlar_konsolide))
                            agirliklar[sk] = w
                            toplam_agirlik += w
                    if toplam_agirlik != 100:
                        st.error(f"⚠️ Toplam ağırlık 100 olmalıdır! (Şu an: {toplam_agirlik})")
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
                                        if not o_no:
                                            continue
                                        if o_no not in ogr_dict:
                                            ogr_dict[o_no] = {"Öğrenci No": o_no, "Ad Soyad": r.get("Ad_Soyad", ""), "Şube": s_sube}
                                        ogr_dict[o_no][s_kodu] = r["Sınav Notu"]
                            if not ogr_dict:
                                st.warning("Not bulunamadı.")
                            else:
                                sonuclar = []
                                for o_no, data in ogr_dict.items():
                                    nihai_puan = sum(data.get(sk, 0) * (agirliklar[sk] / 100) for sk in secilen_sinavlar_konsolide)
                                    for sk in secilen_sinavlar_konsolide:
                                        if sk not in data:
                                            data[sk] = 0
                                    data["Nihai Puan"] = round(nihai_puan, 2)
                                    data["Durum"] = "BAŞARILI" if nihai_puan >= gecme_notu else "BAŞARISIZ"
                                    sonuclar.append(data)
                                df_sonuc = pd.DataFrame(sonuclar).sort_values(by="Nihai Puan", ascending=False)
                                try:
                                    styled_df = df_sonuc.style.map(lambda x: 'background-color: #D5F5E3' if x == 'BAŞARILI' else 'background-color: #FADBD8' if x == 'BAŞARISIZ' else '', subset=['Durum'])
                                except Exception:
                                    styled_df = df_sonuc.style.applymap(lambda x: 'background-color: #D5F5E3' if x == 'BAŞARILI' else 'background-color: #FADBD8' if x == 'BAŞARISIZ' else '', subset=['Durum'])
                                st.dataframe(styled_df, use_container_width=True, hide_index=True)

    elif rapor_kapsami == "🎯 Madde Ayırt Edicilik (Item Discrimination)":
        cur.execute("SELECT DISTINCT sube FROM notlar ORDER BY sube")
        mevcut_subeler = [r[0] for r in cur.fetchall()]
        if not mevcut_subeler:
            st.warning("Not girişi yapılmış şube bulunamadı.")
        else:
            with st.container(border=True):
                c_f1, c_f2 = st.columns(2)
                with c_f1:
                    sec_sube_md = st.selectbox("Şube Seçin", ["Tüm Şubeler"] + mevcut_subeler, key="md_sube")
                if sec_sube_md == "Tüm Şubeler":
                    cur.execute("SELECT DISTINCT sinav_kodu FROM notlar ORDER BY sinav_kodu")
                else:
                    cur.execute("SELECT DISTINCT sinav_kodu FROM notlar WHERE sube=%s ORDER BY sinav_kodu", (sec_sube_md,))
                sinavlar = [r[0] for r in cur.fetchall()]
                if not sinavlar:
                    st.warning("Bu şube için kayıtlı sınav notu bulunamadı.")
                else:
                    with c_f2:
                        secilen_sinav = st.selectbox("Analiz Edilecek Sınav", sinavlar, key="md_sinav")
                    baslat = st.button("Analizi Başlat", type="primary")

            if sinavlar and baslat:
                df_tos = pd.read_sql_query("SELECT ana_bolum, alt_bolum, max_puan FROM sinav_tos WHERE sinav_kodu=%s", conn, params=(secilen_sinav,))
                max_puanlar = {f"{r['ana_bolum']} - {r['alt_bolum']}": float(r['max_puan']) for _, r in df_tos.iterrows()}
                if sec_sube_md == "Tüm Şubeler":
                    cur.execute("SELECT veri_json FROM notlar WHERE sinav_kodu=%s", (secilen_sinav,))
                else:
                    cur.execute("SELECT veri_json FROM notlar WHERE sinav_kodu=%s AND sube=%s", (secilen_sinav, sec_sube_md))
                tum_ogrenciler = []
                for j in cur.fetchall():
                    tum_ogrenciler.extend(pd.read_json(io.StringIO(j[0])).to_dict('records'))
                df_hepsi = pd.DataFrame(tum_ogrenciler)
                if "Sınava Girdi" in df_hepsi.columns:
                    df_hepsi = df_hepsi[df_hepsi["Sınava Girdi"] == True]
                if "Sınav Notu" not in df_hepsi.columns or len(df_hepsi) < 10:
                    st.warning("Aktif katılımcı verisi yetersiz (En az 10 öğrenci).")
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
                            durum = "Çok İyi Madde" if d_index >= 0.40 else "İyi Madde" if d_index >= 0.30 else "Zayıf Madde" if d_index >= 0.20 else "Revize Edilmeli"
                            sonuclar.append({"Bölüm": bolum, "Max": max_p, "Üst %27": round(ust_ortalama, 2), "Alt %27": round(alt_ortalama, 2), "D İndeksi": round(d_index, 3), "Yorum": durum})
                    df_s = pd.DataFrame(sonuclar)
                    if not df_s.empty:
                        def renk_ver(val):
                            if not isinstance(val, str): return ''
                            if "Çok İyi" in val: return 'color: #166534; font-weight: bold;'
                            if "Revize" in val: return 'background-color: #FEE2E2; color: #991B1B; font-weight: bold;'
                            return ''
                        try:
                            styler = df_s.style.map(renk_ver, subset=['Yorum'])
                        except AttributeError:
                            styler = df_s.style.applymap(renk_ver, subset=['Yorum'])
                        st.dataframe(styler, use_container_width=True, hide_index=True)

    elif rapor_kapsami == "⚖️ Puanlayıcı Güvenirliği (Inter-Rater Reliability)":
        cur.execute("SELECT DISTINCT sube FROM notlar ORDER BY sube")
        mevcut_subeler = [r[0] for r in cur.fetchall()]
        if not mevcut_subeler:
            st.warning("Not girişi yapılmış şube bulunamadı.")
        else:
            with st.container(border=True):
                c_f1, c_f2 = st.columns(2)
                with c_f1:
                    sec_sube_ir = st.selectbox("Şube Seçin", ["Tüm Şubeler"] + mevcut_subeler, key="ir_sube")
                if sec_sube_ir == "Tüm Şubeler":
                    cur.execute("SELECT DISTINCT sinav_kodu FROM notlar ORDER BY sinav_kodu")
                else:
                    cur.execute("SELECT DISTINCT sinav_kodu FROM notlar WHERE sube=%s ORDER BY sinav_kodu", (sec_sube_ir,))
                sinavlar = [r[0] for r in cur.fetchall()]
                if not sinavlar:
                    st.warning("Bu şube için kayıtlı sınav notu bulunamadı.")
                else:
                    with c_f2:
                        secilen_sinav_ir = st.selectbox("Sınav Seç", sinavlar, key="ir_sinav")
                    baslat_ir = st.button("Güvenirlik Analizi Çıkar", type="primary")

            if sinavlar and baslat_ir:
                if sec_sube_ir == "Tüm Şubeler":
                    cur.execute("SELECT veri_json FROM notlar WHERE sinav_kodu=%s", (secilen_sinav_ir,))
                else:
                    cur.execute("SELECT veri_json FROM notlar WHERE sinav_kodu=%s AND sube=%s", (secilen_sinav_ir, sec_sube_ir))
                tum_ogrenciler = []
                for j in cur.fetchall():
                    tum_ogrenciler.extend(pd.read_json(io.StringIO(j[0])).to_dict('records'))
                df_ir = pd.DataFrame(tum_ogrenciler)
                if "Sınava Girdi" in df_ir.columns:
                    df_ir = df_ir[df_ir["Sınava Girdi"] == True]
                okuma_1_kolonlari = [c for c in df_ir.columns if "(1. Okuma)" in c]
                if not okuma_1_kolonlari:
                    st.warning("'(1. Okuma)' etiketi taşıyan sütun bulunamadı.")
                else:
                    for ok1_col in okuma_1_kolonlari:
                        ok2_col = ok1_col.replace("(1. Okuma)", "(2. Okuma)")
                        if ok2_col in df_ir.columns:
                            st.markdown(f'<div class="analytics-sub-title">📊 Beceri: {ok1_col.replace(" - (1. Okuma)", "")}</div>', unsafe_allow_html=True)
                            df_cift = df_ir[['Ad_Soyad', ok1_col, ok2_col]].copy().dropna()
                            df_cift['Puan Farkı'] = abs(df_cift[ok1_col] - df_cift[ok2_col])
                            if len(df_cift) > 2:
                                r_degeri = df_cift[ok1_col].corr(df_cift[ok2_col])
                                st.markdown(_kpi_cards_html([
                                    ("Korelasyon (r)", f"{r_degeri:.2f}", True),
                                    ("1. Okuma Ort.", f"{df_cift[ok1_col].mean():.2f}", False),
                                    ("2. Okuma Ort.", f"{df_cift[ok2_col].mean():.2f}", False),
                                ]), unsafe_allow_html=True)
                                st.write("**⚠️ Yüksek Sapma (Fark > 10):**")
                                st.dataframe(df_cift[df_cift['Puan Farkı'] > 10].sort_values(by="Puan Farkı", ascending=False), use_container_width=True, hide_index=True)
                                st.divider()

    elif rapor_kapsami == "👤 Bireysel Karneler":
        cur.execute("SELECT DISTINCT sube FROM notlar ORDER BY sube")
        subeler = [r[0] for r in cur.fetchall()]
        if not subeler:
            st.warning("Not girişi yapılmış şube yok.")
        else:
            with st.container(border=True):
                c_f1, c_f2 = st.columns(2)
                with c_f1:
                    sec_sube = st.selectbox("Şube Seçin", subeler, key="bk_sube")
                cur.execute("SELECT DISTINCT sinav_kodu FROM notlar WHERE sube=%s ORDER BY sinav_kodu", (sec_sube,))
                sinavlar = [r[0] for r in cur.fetchall()]
                if not sinavlar:
                    st.warning("Bu şubeye ait girilmiş sınav notu yok.")
                else:
                    with c_f2:
                        sec_sinav = st.selectbox("Sınav Seçin", sinavlar, key="bk_sinav")

            if sinavlar:
                cur.execute("SELECT tolerans FROM sinav_ayarlari WHERE sinav_kodu=%s", (sec_sinav,))
                tol_res = cur.fetchone()
                sinav_tol = int(tol_res[0]) if (tol_res and tol_res[0] is not None) else 6

                cur.execute("SELECT ana_bolum, alt_bolum, kazanimlar, max_puan FROM sinav_tos WHERE sinav_kodu=%s", (sec_sinav,))
                tos_df = pd.DataFrame(cur.fetchall(), columns=["ana_bolum", "alt_bolum", "kazanimlar", "max_puan"])

                cur.execute("SELECT veri_json FROM notlar WHERE sinav_kodu=%s AND sube=%s", (sec_sinav, sec_sube))
                df_k = pd.read_json(io.StringIO(cur.fetchone()[0]))

                if not df_k.empty:
                    if "Sınava Girdi" not in df_k.columns:
                        df_k["Sınava Girdi"] = True
                    df_k_aktif = df_k[df_k["Sınava Girdi"] == True].copy()
                    sinif_genel_ort = pd.to_numeric(df_k_aktif["Sınav Notu"], errors='coerce').fillna(0.0).mean() if not df_k_aktif.empty else 0.0

                    sec_ogr = st.selectbox("Öğrenci Seçin", df_k['Ad_Soyad'].tolist())
                    ogr_veri = df_k[df_k['Ad_Soyad'] == sec_ogr].iloc[0]

                    bkl = ogr_veri.get('Kitapçık', 'A')
                    st.markdown(_kpi_cards_html([
                        ("Öğrenci", str(sec_ogr), False),
                        ("Toplam Puan", f"{ogr_veri.get('Sınav Notu', 0)}", True),
                        ("Sınıf Ortalaması", f"{sinif_genel_ort:.2f}", False),
                        ("Kitapçık", str(bkl), False),
                    ]), unsafe_allow_html=True)

                    katildi_mi = bool(ogr_veri.get('Sınava Girdi', True))
                    if not katildi_mi:
                        st.error("⚠️ Bu öğrenci sınava katılmamıştır. (Absent)")
                    else:
                        satirlar = karne_satirlari_hazirla(tos_df, df_k_aktif, ogr_veri, sinav_tol)
                        with st.container(border=True):
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
        cur.execute("SELECT DISTINCT sube FROM notlar ORDER BY sube")
        subeler = [r[0] for r in cur.fetchall()]
        if not subeler:
            st.warning("Not girişi yapılmış şube yok.")
        else:
            with st.container(border=True):
                c_f1, c_f2 = st.columns(2)
                with c_f1:
                    sec_sube = st.selectbox("Şube Seçin", subeler, key="sr_sube")
                cur.execute("SELECT DISTINCT sinav_kodu FROM notlar WHERE sube=%s ORDER BY sinav_kodu", (sec_sube,))
                sinavlar = [r[0] for r in cur.fetchall()]
                if not sinavlar:
                    st.warning("Bu şubeye ait girilmiş sınav notu yok.")
                else:
                    with c_f2:
                        sec_sinav = st.selectbox("Sınav Seçin", sinavlar, key="sr_sinav")

            if sinavlar:
                cur.execute("SELECT ana_bolum, alt_bolum, kazanimlar, max_puan FROM sinav_tos WHERE sinav_kodu=%s", (sec_sinav,))
                tos_df = pd.DataFrame(cur.fetchall(), columns=["ana_bolum", "alt_bolum", "kazanimlar", "max_puan"])
                tos_dict = {f"{row['ana_bolum']} - {row['alt_bolum']}": {"max": float(row['max_puan']), "kazanim": row['kazanimlar']} for _, row in tos_df.iterrows()}

                cur.execute("SELECT veri_json FROM notlar WHERE sinav_kodu=%s AND sube=%s", (sec_sinav, sec_sube))
                df_s = pd.read_json(io.StringIO(cur.fetchone()[0]))

                if "Sınava Girdi" not in df_s.columns:
                    df_s["Sınava Girdi"] = True
                if "Kitapçık" not in df_s.columns:
                    df_s["Kitapçık"] = "A"
                df_aktif = df_s[df_s["Sınava Girdi"] == True].copy()

                for c in tos_dict.keys():
                    if c in df_aktif.columns:
                        df_aktif[c] = pd.to_numeric(df_aktif[c].astype(str).str.replace(',', '.'), errors='coerce').fillna(0.0)

                st.markdown(f'<div class="analytics-sub-title">📊 {sec_sube} — Genel Sınıf İstatistikleri</div>', unsafe_allow_html=True)
                st.markdown(_kpi_cards_html([
                    ("Sınıf Ortalaması", f"{df_aktif['Sınav Notu'].mean():.2f}" if not df_aktif.empty else "0.00", True),
                    ("Medyan", f"{df_aktif['Sınav Notu'].median():.2f}" if not df_aktif.empty else "0.00", False),
                    ("En Yüksek Not", f"{df_aktif['Sınav Notu'].max():.2f}" if not df_aktif.empty else "0.00", False),
                    ("En Düşük Not", f"{df_aktif['Sınav Notu'].min():.2f}" if not df_aktif.empty else "0.00", False),
                ]), unsafe_allow_html=True)

                st.markdown('<div class="analytics-sub-title">🎯 Beceri Bazlı Sınıf Ortalamaları</div>', unsafe_allow_html=True)
                beceri_listesi_pdf = []
                with st.container(border=True):
                    for col, veri_dict in tos_dict.items():
                        if col in df_aktif.columns:
                            ort = df_aktif[col].mean() if not df_aktif.empty else 0
                            m_puan = veri_dict['max']
                            yuzde = (ort / m_puan * 100) if m_puan > 0 else 0
                            beceri_listesi_pdf.append({"isim": col, "kazanim": veri_dict['kazanim'], "ort": ort, "max": m_puan, "yuzde": yuzde})
                            st.write(f"**{col}** | *{veri_dict['kazanim']}* (Sınıf Ort: {ort:.1f} / {m_puan})")
                            st.progress(min(int(yuzde), 100))

                st.markdown('<div class="analytics-sub-title">📋 Öğrenci Detay Tablosu</div>', unsafe_allow_html=True)
                cols_to_show = ["Ogrenci_No", "Ad_Soyad", "Kitapçık", "Sınava Girdi", "Sınav Notu"] + [c for c in tos_dict.keys() if c in df_s.columns]
                st.dataframe(df_s[cols_to_show].sort_values(by="Sınav Notu", ascending=False), use_container_width=True, hide_index=True)

                c_btn1, c_btn2 = st.columns(2)
                with c_btn1:
                    if st.button("📊 Sınıf Raporunu Excel Olarak İndir (Tek Sayfa - Kurumsal)", type="primary", use_container_width=True):
                        buffer = create_corporate_class_excel(sec_sinav, sec_sube, df_s, df_aktif, cols_to_show, beceri_listesi_pdf)
                        st.success("Kurumsal tek sayfalık Excel sınıf raporu hazırlandı!")
                        st.download_button("📥 Excel Raporunu İndir", data=buffer, file_name=f"{sec_sube}_{sec_sinav}_Sinif_Raporu.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

                with c_btn2:
                    if st.button("📦 Tüm Sınıfın Karnelerini Oluştur (Ayrı PDF - ZIP)", use_container_width=True):
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

    cur.close()
    conn.close()
