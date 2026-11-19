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


def _analitik_css_uygula():
    st.markdown(
        "<style>"
        "div[data-testid='stRadio'] > div[role='radiogroup'] {"
        "display: flex; flex-wrap: wrap; gap: 8px; background: #FFFFFF;"
        "padding: 10px 12px; border-radius: 10px; border: 1px solid #E2E8F0;"
        "box-shadow: 0 2px 6px rgba(26, 54, 93, 0.04); margin-bottom: 6px; }"
        "div[data-testid='stRadio'] > div[role='radiogroup'] label {"
        "background: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px;"
        "padding: 6px 12px; margin: 0 !important; cursor: pointer;"
        "transition: all 0.18s ease; font-size: 13px; font-weight: 500; color: #334155 !important; }"
        "div[data-testid='stRadio'] > div[role='radiogroup'] label:hover {"
        "border-color: #1A365D; background: #EFF6FF; color: #1A365D !important; }"
        "div[data-testid='stRadio'] > div[role='radiogroup'] label:has(input:checked) {"
        "background: #1A365D !important; border-color: #1A365D !important;"
        "box-shadow: 0 2px 6px rgba(26, 54, 93, 0.18); }"
        "div[data-testid='stRadio'] > div[role='radiogroup'] label:has(input:checked) * {"
        "color: #FFFFFF !important; font-weight: 600 !important; }"
        "div[data-testid='stRadio'] > div[role='radiogroup'] label > div:first-child {"
        "display: none !important; }"
        ".an-banner { background: #FFFFFF; border: 1px solid #E2E8F0;"
        "border-left: 4px solid #E84C0F; border-radius: 10px; padding: 14px 18px;"
        "margin: 10px 0 14px 0; display: flex; justify-content: space-between; align-items: center;"
        "box-shadow: 0 2px 6px rgba(26, 54, 93, 0.04); }"
        ".an-banner-title { font-family: 'DM Sans', sans-serif; font-size: 15.5px;"
        "font-weight: 700; color: #1A365D; margin: 0; }"
        ".an-banner-sub { font-family: 'Jost', sans-serif; font-size: 12.5px;"
        "color: #64748B; margin-top: 2px; }"
        ".an-pill { background: #EFF6FF; color: #1A365D; border: 1px solid #BFDBFE;"
        "padding: 4px 11px; border-radius: 20px; font-size: 11.5px; font-weight: 600; }"
        ".an-skill-row { padding: 8px 2px; border-bottom: 1px solid #F1F5F9; }"
        ".an-skill-row:last-child { border-bottom: none; }"
        ".an-skill-top { display: flex; justify-content: space-between; align-items: baseline;"
        "margin-bottom: 5px; gap: 12px; }"
        ".an-skill-name { font-family: 'DM Sans', sans-serif; font-size: 13px;"
        "font-weight: 700; color: #1E293B; }"
        ".an-skill-obj { font-family: 'Jost', sans-serif; font-size: 12px; color: #64748B;"
        "font-weight: 400; margin-left: 6px; }"
        ".an-skill-stat { font-family: 'Jost', sans-serif; font-size: 12.5px;"
        "font-weight: 600; color: #1A365D; white-space: nowrap; }"
        ".an-bar-bg { width: 100%; height: 8px; background: #E2E8F0; border-radius: 99px; overflow: hidden; }"
        ".an-bar-fill { height: 100%; border-radius: 99px; transition: width 0.3s ease; }"
        ".an-radar-card { background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 10px;"
        "padding: 12px 16px; margin-bottom: 10px; box-shadow: 0 1px 4px rgba(0,0,0,0.03); }"
        ".an-radar-top { display: flex; justify-content: space-between; align-items: center; margin-bottom: 7px; }"
        ".an-status-ok { background: #DCFCE7; color: #166534; padding: 2px 10px; border-radius: 12px; font-size: 11px; font-weight: 700; }"
        ".an-status-mid { background: #FEF9C3; color: #854D0E; padding: 2px 10px; border-radius: 12px; font-size: 11px; font-weight: 700; }"
        ".an-status-low { background: #FEE2E2; color: #991B1B; padding: 2px 10px; border-radius: 12px; font-size: 11px; font-weight: 700; }"
        "</style>",
        unsafe_allow_html=True,
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


def _banner_html(title, subtitle, badge_text=""):
    badge = f'<span class="an-pill">{badge_text}</span>' if badge_text else ""
    return (
        f'<div class="an-banner">'
        f'<div><div class="an-banner-title">{title}</div>'
        f'<div class="an-banner-sub">{subtitle}</div></div>'
        f'<div>{badge}</div>'
        f'</div>'
    )


def _bar_color(yuzde):
    if yuzde >= 75:
        return "#16A34A"
    if yuzde >= 50:
        return "#EAB308"
    return "#DC2626"


def _skill_rows_html(rows):
    parts = []
    for name, obj, stat_text, yuzde in rows:
        pct = max(0.0, min(100.0, float(yuzde)))
        color = _bar_color(pct)
        obj_html = f'<span class="an-skill-obj">— {obj}</span>' if obj else ""
        parts.append(
            f'<div class="an-skill-row">'
            f'<div class="an-skill-top">'
            f'<div><span class="an-skill-name">{name}</span>{obj_html}</div>'
            f'<div class="an-skill-stat">{stat_text} ({pct:.0f}%)</div>'
            f'</div>'
            f'<div class="an-bar-bg"><div class="an-bar-fill" style="width:{pct:.1f}%; background:{color};"></div></div>'
            f'</div>'
        )
    return "".join(parts)


def render_analitik_raporlar():
    _analitik_css_uygula()
    st.markdown('<div class="analytics-section-title">🔬 Gelişmiş Ölçme &amp; Değerlendirme Analitiği</div>', unsafe_allow_html=True)

    rapor_kapsami = st.radio(
        "Analiz Türü:",
        [
            "🏫 Sınıf / Şube Raporu",
            "👤 Bireysel Karneler",
            "🎯 Madde Ayırt Edicilik",
            "⚖️ Puanlayıcı Güvenirliği",
            "🏆 Dönem Sonu Konsolide",
            "📡 Canlı Sınav Radarı",
            "📥 OBS Excel Çıktısı",
        ],
        horizontal=True,
        label_visibility="collapsed",
    )

    conn = get_db_connection()
    cur = conn.cursor()

    # ----------------------------------------------------
    # 1. CANLI SINAV TAKİP RADARI
    # ----------------------------------------------------
    if rapor_kapsami == "📡 Canlı Sınav Radarı":
        cur.execute("SELECT sinav_kodu FROM sinavlar ORDER BY sinav_kodu")
        sinavlar = [r[0] for r in cur.fetchall()]
        if not sinavlar:
            st.warning("Kayıtlı sınav yok.")
        else:
            with st.container(border=True):
                sec_sinav = st.selectbox("Takip Edilecek Sınavı Seçin", sinavlar)
            ilgili_modul = sec_sinav.split("_")[0]

            cur.execute(
                "SELECT h.sube, h.kullanici_adi, k.isim FROM hoca_atamalari h "
                "LEFT JOIN kullanicilar k ON h.kullanici_adi = k.kullanici_adi WHERE h.modul=%s",
                (ilgili_modul,),
            )
            hoca_atamalari = cur.fetchall()

            if not hoca_atamalari:
                st.info("Bu modül için atanmış sınıf/hoca bulunmuyor.")
            else:
                st.markdown(
                    _banner_html(
                        f"{sec_sinav.replace('_', ' ')} — Canlı Veri Giriş Takibi",
                        "Modüle bağlı tüm şubelerin anlık notlandırma ilerlemesi",
                        f"{len(hoca_atamalari)} Şube",
                    ),
                    unsafe_allow_html=True,
                )
                tamamlanan_sube = 0
                toplam_ogr_all = 0
                okunan_ogr_all = 0
                radar_kartlari = []

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
                    if yuzde >= 100 and toplam_ogrenci > 0:
                        tamamlanan_sube += 1
                    toplam_ogr_all += toplam_ogrenci
                    okunan_ogr_all += okunan_ogrenci

                    if yuzde >= 100:
                        st_badge = '<span class="an-status-ok">Tamamlandı</span>'
                    elif yuzde > 0:
                        st_badge = '<span class="an-status-mid">Devam Ediyor</span>'
                    else:
                        st_badge = '<span class="an-status-low">Başlanmadı</span>'

                    bar_col = _bar_color(yuzde)
                    radar_kartlari.append(
                        f'<div class="an-radar-card">'
                        f'<div class="an-radar-top">'
                        f'<div><strong style="color:#1A365D; font-size:14px;">{sube}</strong> '
                        f'<span style="color:#64748B; font-size:12.5px;">· {hoca_isim or k_adi}</span></div>'
                        f'<div style="display:flex; align-items:center; gap:10px;">'
                        f'<span style="font-size:12.5px; font-weight:600; color:#334155;">{okunan_ogrenci} / {toplam_ogrenci} (%{yuzde})</span>'
                        f'{st_badge}</div>'
                        f'</div>'
                        f'<div class="an-bar-bg"><div class="an-bar-fill" style="width:{yuzde}%; background:{bar_col};"></div></div>'
                        f'</div>'
                    )

                genel_yuzde = (okunan_ogr_all / toplam_ogr_all * 100) if toplam_ogr_all > 0 else 0.0
                st.markdown(
                    _kpi_cards_html([
                        ("Genel İlerleme", f"%{genel_yuzde:.1f}", True),
                        ("Okunan / Toplam Öğrenci", f"{okunan_ogr_all} / {toplam_ogr_all}", False),
                        ("Tamamlanan Şube", f"{tamamlanan_sube} / {len(hoca_atamalari)}", False),
                    ]),
                    unsafe_allow_html=True,
                )
                st.markdown("".join(radar_kartlari), unsafe_allow_html=True)

    # ----------------------------------------------------
    # 2. OBS EXCEL ÇIKTISI
    # ----------------------------------------------------
    elif rapor_kapsami == "📥 OBS Excel Çıktısı":
        cur.execute("SELECT sinav_kodu FROM sinavlar ORDER BY sinav_kodu")
        sinavlar = [r[0] for r in cur.fetchall()]
        if not sinavlar:
            st.warning("Kayıtlı sınav yok.")
        else:
            st.markdown(
                _banner_html(
                    "OBS Toplu Not Aktarım Çıktısı",
                    "Seçili sınavın tüm şubelerdeki öğrenci notlarını tek listede OBS formatında dışa aktarır",
                    "Excel (.xlsx)",
                ),
                unsafe_allow_html=True,
            )
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
                            st.success(f"{len(df_obs)} öğrenci için OBS formatı hazır!")
                            st.download_button("📥 Excel Dosyasını İndir", data=buffer, file_name=f"OBS_{sec_sinav}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                        else:
                            st.warning("Excel formatına uygun öğrenci/not kaydı bulunamadı.")

    # ----------------------------------------------------
    # 3. DÖNEM SONU KONSOLİDE BAŞARI
    # ----------------------------------------------------
    elif rapor_kapsami == "🏆 Dönem Sonu Konsolide":
        cur.execute("SELECT sinav_kodu FROM sinavlar ORDER BY sinav_kodu")
        acik_sinavlar = [r[0] for r in cur.fetchall()]
        if not acik_sinavlar:
            st.warning("Sistemde sınav yok.")
        else:
            st.markdown(
                _banner_html(
                    "Dönem Sonu Konsolide Başarı Hesaplayıcı",
                    "Birden fazla sınavı ağırlıklı ortalama ile birleştirip geçme/kalma durumlarını analiz eder",
                    "Çoklu Sınav",
                ),
                unsafe_allow_html=True,
            )
            with st.container(border=True):
                secilen_sinavlar_konsolide = st.multiselect("Hesaplamaya Dahil Edilecek Sınavları Seçin", acik_sinavlar)
                hesapla_btn = False
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
                        hesapla_btn = st.button("Hesapla", type="primary")

            if secilen_sinavlar_konsolide and hesapla_btn:
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
                    basarili_say = int((df_sonuc["Durum"] == "BAŞARILI").sum())
                    basarisiz_say = int((df_sonuc["Durum"] == "BAŞARISIZ").sum())
                    basari_orani = (basarili_say / len(df_sonuc) * 100) if len(df_sonuc) > 0 else 0.0
                    st.markdown(
                        _kpi_cards_html([
                            ("Başarı Oranı", f"%{basari_orani:.1f}", True),
                            ("Başarılı Öğrenci", str(basarili_say), False),
                            ("Başarısız Öğrenci", str(basarisiz_say), False),
                            ("Genel Ortalama", f"{df_sonuc['Nihai Puan'].mean():.2f}", False),
                        ]),
                        unsafe_allow_html=True,
                    )
                    try:
                        styled_df = df_sonuc.style.map(lambda x: 'background-color: #DCFCE7; color: #166534; font-weight: bold;' if x == 'BAŞARILI' else 'background-color: #FEE2E2; color: #991B1B; font-weight: bold;' if x == 'BAŞARISIZ' else '', subset=['Durum'])
                    except Exception:
                        styled_df = df_sonuc.style.applymap(lambda x: 'background-color: #DCFCE7; color: #166534; font-weight: bold;' if x == 'BAŞARILI' else 'background-color: #FEE2E2; color: #991B1B; font-weight: bold;' if x == 'BAŞARISIZ' else '', subset=['Durum'])
                    st.dataframe(styled_df, use_container_width=True, hide_index=True)

    # ----------------------------------------------------
    # 4. MADDE AYIRT EDİCİLİK (ŞUBE -> SINAV)
    # ----------------------------------------------------
    elif rapor_kapsami == "🎯 Madde Ayırt Edicilik":
        cur.execute("SELECT DISTINCT sube FROM notlar ORDER BY sube")
        mevcut_subeler = [r[0] for r in cur.fetchall()]
        if not mevcut_subeler:
            st.warning("Not girişi yapılmış şube bulunamadı.")
        else:
            with st.container(border=True):
                c_f1, c_f2 = st.columns(2)
                with c_f1:
                    sec_sube_md = st.selectbox("1. Şube Seçin", ["Tüm Şubeler"] + mevcut_subeler, key="md_sube")
                if sec_sube_md == "Tüm Şubeler":
                    cur.execute("SELECT DISTINCT sinav_kodu FROM notlar ORDER BY sinav_kodu")
                else:
                    cur.execute("SELECT DISTINCT sinav_kodu FROM notlar WHERE sube=%s ORDER BY sinav_kodu", (sec_sube_md,))
                sinavlar = [r[0] for r in cur.fetchall()]
                baslat = False
                if not sinavlar:
                    st.warning("Bu şube için kayıtlı sınav notu bulunamadı.")
                else:
                    with c_f2:
                        secilen_sinav = st.selectbox("2. Analiz Edilecek Sınav", sinavlar, key="md_sinav")
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
                        iyi_say = int(df_s["Yorum"].isin(["Çok İyi Madde", "İyi Madde"]).sum())
                        revize_say = int((df_s["Yorum"] == "Revize Edilmeli").sum())
                        st.markdown(
                            _banner_html(
                                f"{secilen_sinav.replace('_', ' ')} — Madde Ayırt Edicilik Analizi",
                                f"Kapsam: {sec_sube_md} · Üst ve Alt %27 grup karşılaştırması (n={len(df_hepsi)})",
                                f"{len(df_s)} Bölüm",
                            ),
                            unsafe_allow_html=True,
                        )
                        st.markdown(
                            _kpi_cards_html([
                                ("Ortalama D İndeksi", f"{df_s['D İndeksi'].mean():.3f}", True),
                                ("Çok İyi / İyi Madde", str(iyi_say), False),
                                ("Revize Edilmeli", str(revize_say), False),
                                ("Aktif Katılımcı", str(len(df_hepsi)), False),
                            ]),
                            unsafe_allow_html=True,
                        )

                        def renk_ver(val):
                            if not isinstance(val, str): return ''
                            if "Çok İyi" in val: return 'background-color: #DCFCE7; color: #166534; font-weight: bold;'
                            if "İyi Madde" in val: return 'color: #166534; font-weight: bold;'
                            if "Zayıf" in val: return 'background-color: #FEF9C3; color: #854D0E; font-weight: bold;'
                            if "Revize" in val: return 'background-color: #FEE2E2; color: #991B1B; font-weight: bold;'
                            return ''

                        try:
                            styler = df_s.style.map(renk_ver, subset=['Yorum'])
                        except AttributeError:
                            styler = df_s.style.applymap(renk_ver, subset=['Yorum'])
                        st.dataframe(styler, use_container_width=True, hide_index=True)

    # ----------------------------------------------------
    # 5. PUANLAYICI GÜVENİRLİĞİ (ŞUBE -> SINAV)
    # ----------------------------------------------------
    elif rapor_kapsami == "⚖️ Puanlayıcı Güvenirliği":
        cur.execute("SELECT DISTINCT sube FROM notlar ORDER BY sube")
        mevcut_subeler = [r[0] for r in cur.fetchall()]
        if not mevcut_subeler:
            st.warning("Not girişi yapılmış şube bulunamadı.")
        else:
            with st.container(border=True):
                c_f1, c_f2 = st.columns(2)
                with c_f1:
                    sec_sube_ir = st.selectbox("1. Şube Seçin", ["Tüm Şubeler"] + mevcut_subeler, key="ir_sube")
                if sec_sube_ir == "Tüm Şubeler":
                    cur.execute("SELECT DISTINCT sinav_kodu FROM notlar ORDER BY sinav_kodu")
                else:
                    cur.execute("SELECT DISTINCT sinav_kodu FROM notlar WHERE sube=%s ORDER BY sinav_kodu", (sec_sube_ir,))
                sinavlar = [r[0] for r in cur.fetchall()]
                baslat_ir = False
                if not sinavlar:
                    st.warning("Bu şube için kayıtlı sınav notu bulunamadı.")
                else:
                    with c_f2:
                        secilen_sinav_ir = st.selectbox("2. Sınav Seçin", sinavlar, key="ir_sinav")
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
                    st.markdown(
                        _banner_html(
                            f"{secilen_sinav_ir.replace('_', ' ')} — Puanlayıcı Güvenirliği",
                            f"Kapsam: {sec_sube_ir} · Çift okuma korelasyon ve sapma analizi",
                            "Inter-Rater",
                        ),
                        unsafe_allow_html=True,
                    )
                    for ok1_col in okuma_1_kolonlari:
                        ok2_col = ok1_col.replace("(1. Okuma)", "(2. Okuma)")
                        if ok2_col in df_ir.columns:
                            st.markdown(f'<div class="analytics-sub-title">📊 Beceri: {ok1_col.replace(" - (1. Okuma)", "")}</div>', unsafe_allow_html=True)
                            df_cift = df_ir[['Ad_Soyad', ok1_col, ok2_col]].copy().dropna()
                            df_cift['Puan Farkı'] = abs(df_cift[ok1_col] - df_cift[ok2_col])
                            if len(df_cift) > 2:
                                r_degeri = df_cift[ok1_col].corr(df_cift[ok2_col])
                                st.markdown(
                                    _kpi_cards_html([
                                        ("Korelasyon (r)", f"{r_degeri:.2f}", True),
                                        ("1. Okuma Ort.", f"{df_cift[ok1_col].mean():.2f}", False),
                                        ("2. Okuma Ort.", f"{df_cift[ok2_col].mean():.2f}", False),
                                    ]),
                                    unsafe_allow_html=True,
                                )
                                st.write("**⚠️ Yüksek Sapma (Fark > 10):**")
                                st.dataframe(df_cift[df_cift['Puan Farkı'] > 10].sort_values(by="Puan Farkı", ascending=False), use_container_width=True, hide_index=True)

    # ----------------------------------------------------
    # 6. BİREYSEL KARNELER (ŞUBE -> SINAV)
    # ----------------------------------------------------
    elif rapor_kapsami == "👤 Bireysel Karneler":
        cur.execute("SELECT DISTINCT sube FROM notlar ORDER BY sube")
        subeler = [r[0] for r in cur.fetchall()]
        if not subeler:
            st.warning("Not girişi yapılmış şube yok.")
        else:
            with st.container(border=True):
                c_f1, c_f2 = st.columns(2)
                with c_f1:
                    sec_sube = st.selectbox("1. Şube Seçin", subeler, key="bk_sube")
                cur.execute("SELECT DISTINCT sinav_kodu FROM notlar WHERE sube=%s ORDER BY sinav_kodu", (sec_sube,))
                sinavlar = [r[0] for r in cur.fetchall()]
                if not sinavlar:
                    st.warning("Bu şubeye ait girilmiş sınav notu yok.")
                else:
                    with c_f2:
                        sec_sinav = st.selectbox("2. Sınav Seçin", sinavlar, key="bk_sinav")

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

                    with st.container(border=True):
                        sec_ogr = st.selectbox("3. Öğrenci Seçin", df_k['Ad_Soyad'].tolist())
                    ogr_veri = df_k[df_k['Ad_Soyad'] == sec_ogr].iloc[0]

                    bkl = ogr_veri.get('Kitapçık', 'A')
                    st.markdown(
                        _banner_html(
                            f"{sec_ogr} — Bireysel Performans Karnesi",
                            f"{sec_sinav.replace('_', ' ')} · Şube: {sec_sube} · Öğrenci No: {ogr_veri.get('Ogrenci_No', '')}",
                            f"Kitapçık {bkl}",
                        ),
                        unsafe_allow_html=True,
                    )
                    st.markdown(
                        _kpi_cards_html([
                            ("Toplam Puan", f"{ogr_veri.get('Sınav Notu', 0)}", True),
                            ("Sınıf Ortalaması", f"{sinif_genel_ort:.2f}", False),
                            ("Şube", str(sec_sube), False),
                            ("Kitapçık", str(bkl), False),
                        ]),
                        unsafe_allow_html=True,
                    )

                    katildi_mi = bool(ogr_veri.get('Sınava Girdi', True))
                    if not katildi_mi:
                        st.error("⚠️ Bu öğrenci sınava katılmamıştır. (Absent)")
                    else:
                        satirlar = karne_satirlari_hazirla(tos_df, df_k_aktif, ogr_veri, sinav_tol)
                        st.markdown('<div class="analytics-sub-title">🎯 Beceri ve Kazanım Profili</div>', unsafe_allow_html=True)
                        with st.container(border=True):
                            skill_data = [
                                (s['bolum'], s['kazanim'], f"{s['puan']:g} / {s['max']:g} · Ort: {s['sinif_ort']:.1f}", s['yuzde'])
                                for s in satirlar
                            ]
                            st.markdown(_skill_rows_html(skill_data), unsafe_allow_html=True)

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

    # ----------------------------------------------------
    # 7. SINIF / ŞUBE RAPORU (ŞUBE -> SINAV)
    # ----------------------------------------------------
    elif rapor_kapsami == "🏫 Sınıf / Şube Raporu":
        cur.execute("SELECT DISTINCT sube FROM notlar ORDER BY sube")
        subeler = [r[0] for r in cur.fetchall()]
        if not subeler:
            st.warning("Not girişi yapılmış şube yok.")
        else:
            with st.container(border=True):
                c_f1, c_f2 = st.columns(2)
                with c_f1:
                    sec_sube = st.selectbox("1. Şube Seçin", subeler, key="sr_sube")
                cur.execute("SELECT DISTINCT sinav_kodu FROM notlar WHERE sube=%s ORDER BY sinav_kodu", (sec_sube,))
                sinavlar = [r[0] for r in cur.fetchall()]
                if not sinavlar:
                    st.warning("Bu şubeye ait girilmiş sınav notu yok.")
                else:
                    with c_f2:
                        sec_sinav = st.selectbox("2. Sınav Seçin", sinavlar, key="sr_sinav")

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

                st.markdown(
                    _banner_html(
                        f"{sec_sube} Şubesi — {sec_sinav.replace('_', ' ')} Performans Raporu",
                        "Genel sınıf istatistikleri, beceri kazanım analizi ve toplu dışa aktarım merkezi",
                        f"{len(df_aktif)} / {len(df_s)} Katılım",
                    ),
                    unsafe_allow_html=True,
                )

                st.markdown(
                    _kpi_cards_html([
                        ("Sınıf Ortalaması", f"{df_aktif['Sınav Notu'].mean():.2f}" if not df_aktif.empty else "0.00", True),
                        ("Medyan", f"{df_aktif['Sınav Notu'].median():.2f}" if not df_aktif.empty else "0.00", False),
                        ("En Yüksek Not", f"{df_aktif['Sınav Notu'].max():.2f}" if not df_aktif.empty else "0.00", False),
                        ("En Düşük Not", f"{df_aktif['Sınav Notu'].min():.2f}" if not df_aktif.empty else "0.00", False),
                    ]),
                    unsafe_allow_html=True,
                )

                st.markdown('<div class="analytics-sub-title">🎯 Beceri Bazlı Sınıf Ortalamaları</div>', unsafe_allow_html=True)
                beceri_listesi_pdf = []
                skill_rows = []
                for col, veri_dict in tos_dict.items():
                    if col in df_aktif.columns:
                        ort = df_aktif[col].mean() if not df_aktif.empty else 0
                        m_puan = veri_dict['max']
                        yuzde = (ort / m_puan * 100) if m_puan > 0 else 0
                        beceri_listesi_pdf.append({"isim": col, "kazanim": veri_dict['kazanim'], "ort": ort, "max": m_puan, "yuzde": yuzde})
                        skill_rows.append((col, veri_dict['kazanim'], f"Ort: {ort:.1f} / {m_puan:g}", yuzde))

                with st.container(border=True):
                    if skill_rows:
                        st.markdown(_skill_rows_html(skill_rows), unsafe_allow_html=True)
                    else:
                        st.info("Beceri verisi bulunamadı.")

                st.markdown('<div class="analytics-sub-title">📋 Öğrenci Detay Tablosu</div>', unsafe_allow_html=True)
                cols_to_show = ["Ogrenci_No", "Ad_Soyad", "Kitapçık", "Sınava Girdi", "Sınav Notu"] + [c for c in tos_dict.keys() if c in df_s.columns]
                st.dataframe(df_s[cols_to_show].sort_values(by="Sınav Notu", ascending=False), use_container_width=True, hide_index=True)

                st.markdown('<div class="analytics-sub-title">📤 Rapor ve Karne Dışa Aktarım</div>', unsafe_allow_html=True)
                with st.container(border=True):
                    c_btn1, c_btn2 = st.columns(2)
                    with c_btn1:
                        if st.button("📊 Sınıf Raporunu Excel Olarak İndir (Tek Sayfa - Kurumsal)", type="primary", use_container_width=True):
                            buffer = create_corporate_class_excel(sec_sinav, sec_sube, df_s, df_aktif, cols_to_show, beceri_listesi_pdf)
                            st.success("Kurumsal tek sayfalık Excel sınıf raporu hazırlandı!")
                            st.download_button("📥 Excel Raporunu İndir", data=buffer, file_name=f"{sec_sube}_{sec_sinav}_Sinif_Raporu.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

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
                                    mime="application/zip",
                                    use_container_width=True,
                                )

    cur.close()
    conn.close()
