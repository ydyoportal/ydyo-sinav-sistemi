import streamlit as st
import pandas as pd
import io
from db_utils import get_db_connection, filtrele_sube_sinavlari

def render_hoca_paneli():
    conn = get_db_connection()
    cur = conn.cursor()

    cur.execute("SELECT DISTINCT modul, sube FROM hoca_atamalari WHERE kullanici_adi=%s ORDER BY modul, sube", (st.session_state['kullanici_adi'],))
    hoca_atamalari = cur.fetchall()
    if not hoca_atamalari:
        st.error("⚠️ Size tanımlanmış herhangi bir şube ataması bulunmamaktadır.")
        cur.close(); conn.close()
        st.stop()

    cur.execute("SELECT sinav_kodu, kilitli FROM sinavlar ORDER BY sinav_kodu")
    sinav_data = cur.fetchall()
    if not sinav_data:
        st.warning("Sistemde açık sınav bulunmuyor.")
        cur.close(); conn.close()
        st.stop()

    tum_sinav_kodlari = [row[0] for row in sinav_data]
    sinav_kilit_durumlari = {row[0]: row[1] for row in sinav_data}

    sube_secenekleri = [f"{m} - {s}" for m, s in hoca_atamalari]
    with st.container(border=True):
        col_sel1, col_sel2 = st.columns([1, 2])
        with col_sel1:
            secilen_sube_etiket = st.selectbox("Sorumlu Olduğunuz Şube", sube_secenekleri)
        ilgili_modul, secilen_sube = secilen_sube_etiket.split(" - ", 1)

        sube_sinavlari = filtrele_sube_sinavlari(tum_sinav_kodlari, ilgili_modul, secilen_sube)
        if not sube_sinavlari:
            st.warning(f"⚠️ **{secilen_sube}** şubesine ({ilgili_modul}) tanımlı aktif bir sınav bulunmamaktadır.")
            cur.close(); conn.close()
            st.stop()

        with col_sel2:
            secilen_islem_sinavi = st.selectbox("Şubeye Ait Sınav", sube_sinavlari)
    is_locked = sinav_kilit_durumlari.get(secilen_islem_sinavi, False)

    sinav_temiz_baslik = secilen_islem_sinavi.replace("_", " ")
    durum_badge = '<span class="c-badge-locked">🔒 KİLİTLİ</span>' if is_locked else '<span class="c-badge-open">● Not Girişine Açık</span>'
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
    except Exception:
        rubrik_df = pd.read_sql_query("SELECT ana_bolum, alt_bolum, max_puan FROM sinav_tos WHERE sinav_kodu=%s", conn, params=(secilen_islem_sinavi,))
        rubrik_df['b_grubu_karsiligi'] = ""

    max_puan_sozlugu = {f"{r['ana_bolum']} - {r['alt_bolum']}": float(r['max_puan']) for _, r in rubrik_df.iterrows()}

    isaretler = [("1. Okuma", "2. Okuma", "3. Okuma"), ("(1. Okuma)", "(2. Okuma)", "(3. Okuma)"), ("1st", "2nd", "3rd"), ("1. Hoca", "2. Hoca", "3. Hoca"), ("(1)", "(2)", "(3)"), ("Okuma 1", "Okuma 2", "Okuma 3"), ("Rater 1", "Rater 2", "Rater 3"), (" 1 ", " 2 ", " 3 "), (" 1", " 2", " 3"), ("1.", "2.", "3.")]

    cift_okumalar = []
    islenenler = set()

    for c1 in max_puan_sozlugu.keys():
        if c1 in islenenler:
            continue
        c2_match, c3_match = None, None
        for m1, m2, m3 in isaretler:
            if m1 in c1:
                hedef_c2 = c1.replace(m1, m2)
                hedef_c3 = c1.replace(m1, m3)
                if hedef_c2 in max_puan_sozlugu:
                    c2_match = hedef_c2
                    if hedef_c3 in max_puan_sozlugu:
                        c3_match = hedef_c3
                    break
        if c2_match:
            cift_okumalar.append((c1, c2_match, c3_match))
            islenenler.update([c1, c2_match])
            if c3_match:
                islenenler.add(c3_match)

    ort_mesaj = None
    if cift_okumalar:
        eslesen_isimler = []
        for c1, c2, c3 in cift_okumalar:
            isim = f"**{c1.split('-')[-1].strip()}** & **{c2.split('-')[-1].strip()}**"
            if c3:
                isim += f" *(Gerekirse {c3.split('-')[-1].strip()})*"
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
        gercek_ogrenciler = pd.read_sql_query('SELECT ogrenci_no AS "Ogrenci_No", ad_soyad AS "Ad_Soyad" FROM ogrenciler WHERE modul=%s AND sube=%s', conn, params=(ilgili_modul, secilen_sube))
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

    def _hesapla_ve_kaydet(girilen_notlar_df):
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
                    if val > m_puan:
                        hatalar.append(f"**{row.get('Ogrenci_No', '')}**: '{c.split('-')[-1].strip()}' notu ({val}), max ({m_puan}) aşıyor!")

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
                    if c not in processed_cols:
                        ogrenci_toplam += df_to_save.at[idx, c]

                df_to_save.at[idx, "Sınav Notu"] = round(ogrenci_toplam, 2)
            else:
                for col in max_puan_sozlugu.keys():
                    df_to_save.at[idx, col] = 0.0
                df_to_save.at[idx, "Sınav Notu"] = 0.0

        if hatalar:
            st.error("🛑 SİSTEM KAYDI DURDURDU: KURALLAR İHLAL EDİLDİ VEYA MAX PUAN AŞILDI!")
            for h in hatalar:
                st.warning(h)
            st.stop()

        cur.execute("""INSERT INTO notlar (sinav_kodu, sube, veri_json) VALUES (%s, %s, %s) ON CONFLICT (sinav_kodu, sube) DO UPDATE SET veri_json = EXCLUDED.veri_json;""", (secilen_islem_sinavi, secilen_sube, df_to_save.to_json(orient="records")))
        conn.commit()
        cur.close()
        conn.close()
        st.success("Notlar bulut veritabanına başarıyla kaydedildi!")
        st.rerun()

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
                if ort_mesaj:
                    st.info(ort_mesaj)
                col_conf_not = {
                    "Ogrenci_No": st.column_config.TextColumn("Öğrenci No", disabled=True),
                    "Ad_Soyad": st.column_config.TextColumn("Ad Soyad", disabled=True, pinned=True),
                    "Sınava Girdi": st.column_config.CheckboxColumn("Sınava Girdi", disabled=True),
                    "Kitapçık": st.column_config.TextColumn("Kitapçık", disabled=True),
                    "Sınav Notu": st.column_config.NumberColumn("Sınav Notu", disabled=True)
                }
                for c in df_view.columns:
                    if c not in base_cols:
                        col_conf_not[c] = st.column_config.TextColumn(c)

                with st.form("hoca_veri_giris_formu"):
                    girilen_notlar_df = st.data_editor(df_view, num_rows="dynamic", use_container_width=True, column_config=col_conf_not, disabled=is_locked)
                    kaydet_btn = st.form_submit_button("💾 Notları Buluta Kaydet", type="primary", disabled=is_locked)

                if kaydet_btn and not is_locked:
                    if "B Kitapçığı" in gorunum:
                        girilen_notlar_df = girilen_notlar_df.rename(columns=reverse_b_map)
                    _hesapla_ve_kaydet(girilen_notlar_df)
    else:
        base_cols_tek = ["Ogrenci_No", "Ad_Soyad", "Sınava Girdi", "Sınav Notu"]
        df_view = df_giris[base_cols_tek + skill_cols_in_order].copy()

        if ort_mesaj:
            st.info(ort_mesaj)

        col_conf_not = {
            "Ogrenci_No": st.column_config.TextColumn("Öğrenci No", disabled=True),
            "Ad_Soyad": st.column_config.TextColumn("Ad Soyad", disabled=True, pinned=True),
            "Sınava Girdi": st.column_config.CheckboxColumn("Sınava Girdi"),
            "Sınav Notu": st.column_config.NumberColumn("Sınav Notu", disabled=True)
        }
        for c in df_view.columns:
            if c not in base_cols_tek:
                col_conf_not[c] = st.column_config.TextColumn(c)

        st.info("Bu sınav **Tek Kitapçık** olarak yapılandırılmıştır. Tüm notları doğrudan aşağıdaki tabloya girebilirsiniz.")

        with st.form("hoca_veri_giris_formu_tek"):
            girilen_notlar_df = st.data_editor(df_view, num_rows="dynamic", use_container_width=True, column_config=col_conf_not, disabled=is_locked)
            kaydet_btn = st.form_submit_button("💾 Notları Buluta Kaydet", type="primary", disabled=is_locked)

        if kaydet_btn and not is_locked:
            _hesapla_ve_kaydet(girilen_notlar_df)

    cur.close()
    conn.close()
