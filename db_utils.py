import streamlit as st
import pandas as pd
import psycopg2
import io
import os
import base64
import hashlib

def get_logo_b64():
    if os.path.exists("logo.png"):
        try:
            with open("logo.png", "rb") as f:
                return base64.b64encode(f.read()).decode()
        except Exception:
            return ""
    return ""

def filtrele_sube_sinavlari(tum_sinavlar, modul, sube):
    modul_sinavlari = [s for s in tum_sinavlar if s.split("_")[0] == modul]
    sube_up = str(sube).upper().strip()
    tespit_kur = None
    for kur in ["A1", "A2", "B1", "B2", "PF"]:
        if sube_up.startswith(kur):
            tespit_kur = kur
            break
    if tespit_kur:
        kur_sinavlari = [s for s in modul_sinavlari if len(s.split("_")) > 1 and s.split("_")[1].upper() == tespit_kur]
        if kur_sinavlari:
            return kur_sinavlari
    return modul_sinavlari

def hash_sifre(sifre: str) -> str:
    if not sifre:
        return ""
    if len(sifre) == 64 and all(c in '0123456789abcdefABCDEF' for c in sifre):
        return sifre
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
        cur.close()
        conn.close()
    except Exception:
        pass
    return True

def get_db_backup_excel():
    conn = get_db_connection()
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        tablolar = ["sinavlar", "sinav_ayarlari", "sinav_tos", "kullanicilar", "ogrenciler", "hoca_atamalari", "notlar"]
        for t in tablolar:
            try:
                df = pd.read_sql_query(f"SELECT * FROM {t}", conn)
                df.to_excel(writer, sheet_name=t, index=False)
            except Exception:
                pass
    conn.close()
    buffer.seek(0)
    return buffer
