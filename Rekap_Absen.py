import os
import re
import io
import pandas as pd
import streamlit as st
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import gspread
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="Sistem Rekap & Database Absensi", layout="wide")
st.title("📊 Sistem Rekap & Database Absensi Online")

# ==============================================================================
# KONFIGURASI GOOGLE SHEETS MASTER DATABASE
# ==============================================================================
CREDENTIALS_FILE = "credentials.json"
DEFAULT_GSHEET_URL = "https://docs.google.com/spreadsheets/d/1zgvGZjd5hfV3fUqH9Ns9HQLLtttDzVJVvF8GTNWIwvI/edit?usp=sharing"

BULAN_ID = {
    1: "Januari", 2: "Februari", 3: "Maret", 4: "April",
    5: "Mei", 6: "Juni", 7: "Juli", 8: "Agustus",
    9: "September", 10: "Oktober", 11: "November", 12: "Desember"
}

BULAN_MAP_STR = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "mei": 5, "may": 5,
    "jun": 6, "jul": 7, "agu": 8, "agust": 8, "aug": 8,
    "sep": 9, "sept": 9, "okt": 10, "oct": 10, "nov": 11, "des": 12, "dec": 12
}


# ==============================================================================
# KONEKSI GOOGLE SHEETS (GSPREAD)
# ==============================================================================
def get_gsheet_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    if "gcp_service_account" in st.secrets:
        creds = Credentials.from_service_account_info(
            st.secrets["gcp_service_account"], scopes=scopes
        )
    else:
        creds = Credentials.from_service_account_file(
            CREDENTIALS_FILE, scopes=scopes
        )
    return gspread.authorize(creds)


@st.cache_data(ttl=60, show_spinner=False)
def ambil_daftar_tab_gsheet(gsheet_url):
    if not gsheet_url.strip():
        return []
    try:
        gc = get_gsheet_client()
        sh = gc.open_by_url(gsheet_url.strip())
        return [ws.title for ws in sh.worksheets()]
    except Exception:
        return []


def simpan_ke_gsheet(gsheet_url, nama_tab, df_rekap, mode_simpan="Timpa Semua (Gelondongan)"):
    """Menyimpan data ke Google Sheets: bisa timpa seluruh tab atau tambah/update per pegawai"""
    gc = get_gsheet_client()
    sh = gc.open_by_url(gsheet_url)

    header = [
        "No", "NIP", "Nama", "Total Masuk", "Efeketif Kerja",
        "Terlambat (Menit)", "Total Tidak Absen Pulang",
        "Presentase Kehadiran", "Presentase Jam Kerja"
    ]

    existing_titles = [ws.title for ws in sh.worksheets()]
    if nama_tab in existing_titles:
        ws = sh.worksheet(nama_tab)
        if mode_simpan == "Tambah / Update Per Nama (Sisipan)":
            existing_vals = ws.get_all_values()
            if len(existing_vals) > 1:
                old_hdr = existing_vals[0]
                df_old = pd.DataFrame(existing_vals[1:], columns=old_hdr)
                # Pastikan kolom standar tersedia meski dari tab format lama (6 atau 8 kolom)
                for col_name, default_val in [
                    ("Terlambat (Menit)", 0),
                    ("Total Tidak Absen Pulang", 0),
                    ("Presentase Kehadiran", 0.0),
                    ("Presentase Jam Kerja", 1.0),
                ]:
                    if col_name not in df_old.columns:
                        if col_name == "Presentase Kehadiran" and "Presentase" in df_old.columns:
                            df_old[col_name] = df_old["Presentase"]
                        else:
                            df_old[col_name] = default_val

                # Bersihkan jika ada nilai terformat % pada kolom angka bulat
                for num_col in ["Total Masuk", "Efeketif Kerja", "Terlambat (Menit)", "Total Tidak Absen Pulang"]:
                    def _clean_int(v):
                        s = str(v).strip()
                        if "%" in s:
                            return int(round(float(s.replace("%", "").strip()) / 100.0))
                        return int(round(float(s))) if s else 0
                    df_old[num_col] = df_old[num_col].apply(_clean_int)

                df_old["NIP"] = df_old["NIP"].astype(str).str.strip()
                new_nips = set(df_rekap["NIP"].astype(str).str.strip())
                df_old_keep = df_old[~df_old["NIP"].isin(new_nips)][header[1:]]
                df_new_part = df_rekap[header[1:]].copy()
                df_gabung_gs = pd.concat([df_old_keep, df_new_part], ignore_index=True)
            else:
                df_gabung_gs = df_rekap[header[1:]].copy()
        else:
            df_gabung_gs = df_rekap[header[1:]].copy()
        ws.clear()
    else:
        ws = sh.add_worksheet(title=nama_tab, rows=str(len(df_rekap) + 20), cols="10")
        df_gabung_gs = df_rekap[header[1:]].copy()

    rows_to_write = [header]
    for idx, r in df_gabung_gs.reset_index(drop=True).iterrows():
        def _clean_pct(val):
            s = str(val).replace("%", "").strip()
            v = float(s) if s else 0.0
            return round(v if v <= 1.0 else v / 100.0, 6)

        rows_to_write.append([
            int(idx + 1),
            str(r["NIP"]).strip(),
            str(r["Nama"]).strip(),
            int(float(r["Total Masuk"])),
            int(float(r["Efeketif Kerja"])),
            int(round(float(r["Terlambat (Menit)"]))),
            int(round(float(r.get("Total Tidak Absen Pulang", 0)))),
            _clean_pct(r["Presentase Kehadiran"]),
            _clean_pct(r["Presentase Jam Kerja"]),
        ])

    ws.update(range_name="A1", values=rows_to_write, value_input_option="USER_ENTERED")

    ws.format("D2:G" + str(len(rows_to_write)), {
        "numberFormat": {"type": "NUMBER", "pattern": "0"}
    })
    ws.format("H2:I" + str(len(rows_to_write)), {
        "numberFormat": {"type": "PERCENT", "pattern": "0.00%"}
    })
    ws.format("A1:I1", {
        "textFormat": {"bold": True},
        "horizontalAlignment": "CENTER"
    })
    return sh.title, len(rows_to_write) - 1


# ==============================================================================
# PENGOLAHAN DATA MENTAH (GABUNG MULTI-FILE & MULTI-TAB -> FORMAT SHEET 1)
# ==============================================================================
def deteksi_periode_otomatis(df_gabungan, fallback_name=""):
    """Mendeteksi Bulan & Tahun otomatis dari kolom Tanggal Masuk"""
    if "Tanggal Masuk" in df_gabungan.columns:
        tgl_series = pd.to_datetime(df_gabungan["Tanggal Masuk"], errors="coerce").dropna()
        if not tgl_series.empty:
            mode_periode = tgl_series.dt.to_period("M").mode()[0]
            return f"{BULAN_ID[mode_periode.month]} {mode_periode.year}"
    return fallback_name


def proses_data_mentah(df_raw, efektif_kerja=22, filter_periode=None):
    """
    Mengolah gabungan log mentah absensi menjadi rekap per NIP:
    - Menghitung Total Masuk, Terlambat (Menit), Total Tidak Absen Pulang,
      Presentase Kehadiran, dan Presentase Jam Kerja: ((HK*480) - Telat) / (HK*480).
    """
    df = df_raw.copy()
    df.columns = [str(c).strip() for c in df.columns]

    kolom_wajib = ["No Induk", "Nama", "Tanggal Masuk", "Scan Masuk", "Scan Keluar"]
    kurang = [k for k in kolom_wajib if k not in df.columns]
    if kurang:
        raise ValueError(f"Kolom wajib tidak ditemukan: {', '.join(kurang)}")

    df["NIP"] = df["No Induk"].astype(str).str.strip()
    df["Nama"] = df["Nama"].astype(str).str.strip()
    dt_series = pd.to_datetime(df["Tanggal Masuk"], errors="coerce")
    df["Tgl"] = dt_series.dt.strftime("%Y-%m-%d")
    df["Periode_Str"] = dt_series.apply(
        lambda d: f"{BULAN_ID[d.month]} {d.year}" if pd.notna(d) else ""
    )

    df = df[
        df["NIP"].notna()
        & (df["NIP"] != "")
        & (df["NIP"].str.lower() != "nan")
        & (df["NIP"].str.lower() != "no induk")
        & df["Tgl"].notna()
    ].copy()

    # Jika file per-nama berisi banyak bulan (seperti export riwayat 1 orang), filter sesuai periode target
    if filter_periode and filter_periode.strip() in df["Periode_Str"].unique():
        df = df[df["Periode_Str"] == filter_periode.strip()].copy()

    sm_str = df["Scan Masuk"].astype(str).str.strip()
    sk_str = df["Scan Keluar"].astype(str).str.strip()

    df["sm_ok"] = (sm_str != "-") & (sm_str != "") & (sm_str.str.lower() != "nan")
    df["sk_ok"] = (sk_str != "-") & (sk_str != "") & (sk_str.str.lower() != "nan")

    tgl_count = df.groupby(["NIP", "Tgl"])["sm_ok"].transform("count")
    lengkap = df["sm_ok"] & df["sk_ok"]
    first_lengkap = lengkap & (~df[lengkap].duplicated(subset=["NIP", "Tgl"])).reindex(df.index, fill_value=False)
    lupa_tunggal = df["sm_ok"] & (~df["sk_ok"]) & (tgl_count == 1)

    df["Hitung_Hadir"] = (first_lengkap | lupa_tunggal).astype(int)
    df["Tidak_Pulang"] = ((df["Hitung_Hadir"] == 1) & (~df["sk_ok"])).astype(int)

    if "Keterlambatan (Menit)" in df.columns:
        telat_num = pd.to_numeric(df["Keterlambatan (Menit)"], errors="coerce").fillna(0)
        df["Telat_Valid"] = telat_num.where(df["Hitung_Hadir"] == 1, 0)
    else:
        df["Telat_Valid"] = 0

    rekap = (
        df.groupby("NIP", as_index=False, sort=False)
        .agg(
            Nama=("Nama", "first"),
            Total_Masuk=("Hitung_Hadir", "sum"),
            Telat_Menit=("Telat_Valid", "sum"),
            Tidak_Pulang=("Tidak_Pulang", "sum"),
        )
        .rename(columns={
            "Total_Masuk": "Total Masuk",
            "Telat_Menit": "Terlambat (Menit)",
            "Tidak_Pulang": "Total Tidak Absen Pulang",
        })
    )
    rekap.insert(0, "No", range(1, len(rekap) + 1))
    rekap["Efeketif Kerja"] = int(efektif_kerja)
    rekap["Terlambat (Menit)"] = rekap["Terlambat (Menit)"].round(0).astype(int)
    rekap["Total Tidak Absen Pulang"] = rekap["Total Tidak Absen Pulang"].astype(int)

    total_menit_bulan = float(efektif_kerja) * 480.0
    rekap["Presentase Kehadiran"] = (
        rekap["Total Masuk"] / float(efektif_kerja) if efektif_kerja > 0 else 0.0
    )
    rekap["Presentase Jam Kerja"] = (
        ((total_menit_bulan - rekap["Terlambat (Menit)"]) / total_menit_bulan).clip(lower=0.0)
        if total_menit_bulan > 0 else 0.0
    )
    return rekap


def export_format_sheet1(df_rekap):
    """Mengekspor hasil rekap ke Excel lengkap dengan rumus Presentase Jam Kerja"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    font_default = Font(name="Calibri", size=11, bold=False)
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")

    headers = [
        "No", "NIP", "Nama", "Total Masuk", "Efeketif Kerja",
        "Terlambat (Menit)", "Total Tidak Absen Pulang",
        "Presentase Kehadiran", "Presentase Jam Kerja"
    ]
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = font_default
        cell.alignment = align_center

    for r_idx, row in df_rekap.reset_index(drop=True).iterrows():
        excel_row = r_idx + 2
        vals = [
            int(r_idx + 1),
            str(row["NIP"]),
            str(row["Nama"]),
            int(row["Total Masuk"]),
            int(row["Efeketif Kerja"]),
            int(row["Terlambat (Menit)"]),
            int(row["Total Tidak Absen Pulang"]),
            f"=D{excel_row}/E{excel_row}",
            f"=((E{excel_row}*480)-F{excel_row})/(E{excel_row}*480)",
        ]
        for c_idx, val in enumerate(vals, 1):
            cell = ws.cell(row=excel_row, column=c_idx, value=val)
            cell.font = font_default
            if c_idx == 3:
                cell.alignment = align_left
            elif c_idx in (8, 9):
                cell.number_format = "0.00%"

    lebar_kolom = {1: 6, 2: 16, 3: 36, 4: 14, 5: 15, 6: 18, 7: 24, 8: 21, 9: 22}
    for col_idx, width in lebar_kolom.items():
        ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = width

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# ==============================================================================
# PENCARIAN RIWAYAT DARI GOOGLE SHEETS ONLINE
# ==============================================================================
def ekstrak_periode_sort(nama_tab):
    nama_lower = nama_tab.lower()
    bulan_ditemukan = ""
    angka_bulan = 0
    for k, no_b in BULAN_MAP_STR.items():
        if k in nama_lower:
            bulan_ditemukan = BULAN_ID[no_b]
            angka_bulan = no_b
            break

    match_tahun = re.search(r"(\d{2,4})\b", nama_tab)
    tahun_ditemukan = ""
    angka_tahun = 0
    if match_tahun:
        thn = match_tahun.group(1)
        if len(thn) == 2:
            tahun_ditemukan = f"20{thn}"
            angka_tahun = int(f"20{thn}")
        elif len(thn) == 4:
            tahun_ditemukan = thn
            angka_tahun = int(thn)

    tampilan_bulan = f"{bulan_ditemukan} {tahun_ditemukan}".strip() if bulan_ditemukan else nama_tab
    return tampilan_bulan, angka_tahun, angka_bulan


def format_excel_custom(df):
    """Membuat file Excel berformat rapi dengan header kuning untuk riwayat per pegawai"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Rekap Absen"

    font_title = Font(name="Calibri", size=14, bold=True, underline="single")
    font_header = Font(name="Calibri", size=11, bold=True)
    font_body = Font(name="Calibri", size=11)
    fill_header = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    thin_border = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000"),
    )

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    ws.merge_cells("C2:J2")
    ws["C2"] = "REKAP ABSENSI PEGAWAI"
    ws["C2"].font = font_title
    ws["C2"].alignment = align_center

    headers = [
        "NO", "NIP", "Nama", "Bulan", "Total Masuk", "Efektif Hari Kerja",
        "Terlambat (Menit)", "Total Tidak Absen Pulang",
        "Presentase Kehadiran", "Presentase Jam Kerja"
    ]
    start_row = 5

    for col_num, header_title in enumerate(headers, 1):
        cell = ws.cell(row=start_row, column=col_num, value=header_title)
        cell.font = font_header
        cell.fill = fill_header
        cell.border = thin_border
        cell.alignment = align_center

    for row_idx, row_data in df.reset_index(drop=True).iterrows():
        current_row = start_row + 1 + row_idx
        no_val = row_idx + 1
        nip_val = str(row_data["NIP"])
        nama_val = str(row_data["Nama"])
        bulan_val = str(row_data["Bulan"])

        hadir_val = float(row_data.get("Total Masuk", row_data.get("Hadir", 0)))
        total_val = float(row_data.get("Efektif Hari Kerja", row_data.get("Total Hari", 0)))
        telat_val = float(row_data.get("Terlambat (Menit)", 0))
        tdk_pulang_val = float(row_data.get("Total Tidak Absen Pulang", 0))

        # Kolom E = Total Masuk, Kolom F = Efektif Hari Kerja, Kolom G = Terlambat (Menit)
        formula_hadir = f"=E{current_row}/F{current_row}"
        formula_jam = f"=((F{current_row}*480)-G{current_row})/(F{current_row}*480)"

        row_values = [
            no_val, nip_val, nama_val, bulan_val, hadir_val,
            total_val, telat_val, tdk_pulang_val, formula_hadir, formula_jam
        ]
        for col_idx, val in enumerate(row_values, 1):
            cell = ws.cell(row=current_row, column=col_idx, value=val)
            cell.font = font_body
            cell.border = thin_border
            if col_idx in [1, 4]:
                cell.alignment = align_center
            elif col_idx in [2, 3]:
                cell.alignment = align_left if col_idx == 3 else align_center
            elif col_idx in [5, 6, 7, 8]:
                cell.alignment = align_center
                cell.number_format = "#,##0"
            elif col_idx in [9, 10]:
                cell.alignment = align_right
                cell.number_format = "0.00%"

    for col in ws.columns:
        max_len = 0
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        for cell in col:
            if cell.row >= start_row and cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max(max_len + 4, 14)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# ==============================================================================
# SIDEBAR: KONFIGURASI LINK GOOGLE SHEETS MASTER (DIPAKAI OLEH TAB 1 & TAB 2)
# ==============================================================================
with st.sidebar:
    st.header("⚙️ Pengaturan Database")
    gsheet_url = st.text_input(
        "Link Google Sheets Master:",
        value=DEFAULT_GSHEET_URL,
        placeholder="https://docs.google.com/spreadsheets/d/xxxx/edit",
        help="Satu file Google Sheets ini akan menyimpan seluruh rekap bulanan (tiap bulan = 1 tab)."
    )
    if "gcp_service_account" in st.secrets:
        st.success("✅ Streamlit Secrets Aktif")
    elif os.path.exists(CREDENTIALS_FILE):
        st.success("✅ credentials.json terdeteksi")
    else:
        st.warning("⚠️ Secrets / credentials.json belum terkonfigurasi")

# ==============================================================================
# ANTARMUKA UTAMA (2 TAB)
# ==============================================================================
tab1, tab2 = st.tabs([
    "📤 Upload Data Mentah -> Auto Google Sheets & Excel",
    "🔍 Cari Rekap Per NIP / Nama (Database Online)"
])

# ------------------------------------------------------------------------------
# TAB 1: UPLOAD MULTI-FILE & MULTI-TAB -> SIMPAN KE GSHEET & DOWNLOAD SHEET1
# ------------------------------------------------------------------------------
with tab1:
    st.subheader("Olah & Gabung Data Mentah Mesin Absen (.xls / .xlsx)")
    st.caption("Mendukung penggabungan otomatis dari beberapa file sekaligus maupun beberapa tab (seperti tab Organik & PKWT).")

    col_up, col_hk = st.columns([3, 1])
    with col_up:
        uploaded_files = st.file_uploader(
            "Pilih 1 atau beberapa file mentah absensi:",
            type=["xls", "xlsx"],
            accept_multiple_files=True
        )
    with col_hk:
        efektif_kerja = st.number_input("Hari Efektif Kerja", min_value=1, max_value=31, value=22, step=1)

    if uploaded_files:
        semua_opsi_sheet = []
        peta_excel = {}

        for f in uploaded_files:
            xls_obj = pd.ExcelFile(f)
            peta_excel[f.name] = xls_obj
            for s_name in xls_obj.sheet_names:
                label_opsi = f"{f.name}  |  Tab: {s_name}"
                semua_opsi_sheet.append((label_opsi, f.name, s_name))

        default_pilihan = [
            item[0] for item in semua_opsi_sheet
            if item[2].strip().lower() != "sheet1"
        ]
        if not default_pilihan:
            default_pilihan = [item[0] for item in semua_opsi_sheet]

        pilihan_user = st.multiselect(
            "Tab yang akan digabung & dihitung:",
            options=[item[0] for item in semua_opsi_sheet],
            default=default_pilihan
        )

        if pilihan_user:
            try:
                kumpulan_df = []
                for label_pil in pilihan_user:
                    _, nama_f, nama_s = next(x for x in semua_opsi_sheet if x[0] == label_pil)
                    df_part = pd.read_excel(peta_excel[nama_f], sheet_name=nama_s)
                    kumpulan_df.append(df_part)

                df_gabungan = pd.concat(kumpulan_df, ignore_index=True)
                periode_terdeteksi = deteksi_periode_otomatis(df_gabungan, fallback_name="Rekap_Baru")

                # 1. Pilih mode simpan terlebih dahulu (Gelondongan vs Sisipan Per-Nama)
                mode_simpan_gsheet = st.radio(
                    "Mode Simpan ke Google Sheets:",
                    ["Timpa Semua (Gelondongan)", "Tambah / Update Per Nama (Sisipan)"],
                    index=0,
                    horizontal=True,
                    help="Pilih 'Tambah / Update Per Nama' untuk menyisipkan data pegawai susulan ke tab bulan yang sudah ada tanpa salah ketik nama tab."
                )

                col_per, col_fil = st.columns([1, 2])
                with col_per:
                    if mode_simpan_gsheet == "Tambah / Update Per Nama (Sisipan)":
                        daftar_tab_db = ambil_daftar_tab_gsheet(gsheet_url)
                        if daftar_tab_db:
                            default_idx = (
                                daftar_tab_db.index(periode_terdeteksi)
                                if periode_terdeteksi in daftar_tab_db
                                else 0
                            )
                            nama_tab_tujuan = st.selectbox(
                                "Pilih Tab Periode Tujuan (Database):",
                                options=daftar_tab_db,
                                index=default_idx,
                                help="Daftar tab otomatis diambil dari Google Sheets Master agar tidak salah ketik."
                            )
                        else:
                            st.warning("Belum ada tab di Google Sheets atau link belum terhubung.")
                            nama_tab_tujuan = st.text_input("Nama Tab Periode:", value=periode_terdeteksi)
                    else:
                        nama_tab_tujuan = st.text_input(
                            "Nama Tab Periode di Google Sheets:",
                            value=periode_terdeteksi,
                            help="Otomatis terdeteksi dari Tanggal Masuk. Bisa Anda ubah sebelum disimpan ke Google Sheets."
                        )

                with col_fil:
                    filter_mentah = st.text_input(
                        "Filter Cepat NIP / Nama (Opsional):",
                        placeholder="Kosongkan untuk memproses seluruh pegawai...",
                        key="filter_tab1"
                    )

                df_rekap_full = proses_data_mentah(
                    df_gabungan,
                    efektif_kerja=efektif_kerja,
                    filter_periode=nama_tab_tujuan
                )

                if filter_mentah.strip():
                    kw = filter_mentah.strip().upper()
                    df_tampil = df_rekap_full[
                        df_rekap_full["NIP"].str.upper().str.contains(kw, na=False)
                        | df_rekap_full["Nama"].str.upper().str.contains(kw, na=False)
                    ].copy()
                    df_tampil["No"] = range(1, len(df_tampil) + 1)
                else:
                    df_tampil = df_rekap_full

                st.success(
                    f"Berhasil menggabungkan **{len(pilihan_user)} tab** ({len(df_gabungan):,} baris log) "
                    f"menjadi **{len(df_tampil)} pegawai** untuk periode **{nama_tab_tujuan}**."
                )

                # Tampilkan preview di layar
                df_preview = df_tampil.copy()
                df_preview["Presentase Kehadiran"] = (df_preview["Presentase Kehadiran"] * 100).round(2).astype(str) + "%"
                df_preview["Presentase Jam Kerja"] = (df_preview["Presentase Jam Kerja"] * 100).round(2).astype(str) + "%"
                st.dataframe(df_preview, use_container_width=True, hide_index=True)

                col_btn1, col_btn2 = st.columns(2)
                with col_btn1:
                    if st.button(f"☁️ Simpan ke Google Sheets (Tab: {nama_tab_tujuan})", type="primary", use_container_width=True):
                        if not gsheet_url.strip():
                            st.error("Tempelkan Link Google Sheets Master di panel sebelah kiri terlebih dahulu.")
                        else:
                            with st.spinner(f"Menyimpan data ke tab '{nama_tab_tujuan}' di Google Sheets..."):
                                judul_gsheet, total_baris = simpan_ke_gsheet(
                                    gsheet_url.strip(),
                                    nama_tab_tujuan.strip(),
                                    df_tampil,
                                    mode_simpan=mode_simpan_gsheet
                                )
                                st.cache_data.clear()
                                st.success(f"Mantap! Data masuk ke tab **{nama_tab_tujuan}** di **{judul_gsheet}** (Total sekarang: **{total_baris} pegawai**).")

                with col_btn2:
                    file_sheet1_bytes = export_format_sheet1(df_tampil)
                    nama_unduh = f"Rekap_{nama_tab_tujuan.strip().replace(' ', '_')}.xlsx"
                    st.download_button(
                        label="📥 Unduh Hasil Format Sheet1 (.xlsx)",
                        data=file_sheet1_bytes,
                        file_name=nama_unduh,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )

            except Exception as e:
                st.error(f"Terjadi kesalahan saat mengolah data: {e}")

# ------------------------------------------------------------------------------
# TAB 2: PENCARIAN PER NIP / NAMA DARI SELURUH TAB DI GOOGLE SHEETS MASTER
# ------------------------------------------------------------------------------
with tab2:
    st.subheader("Cari Riwayat Absensi Pegawai Lintas Bulan (Google Sheets)")

    keyword_online = st.text_input(
        "Masukkan NIP atau Nama Pegawai:",
        placeholder="Contoh: 9422011HPL atau ABDULLOH",
        key="keyword_tab2"
    )

    if st.button("🔍 Cari Data di Google Sheets", type="primary"):
        if not gsheet_url.strip():
            st.warning("Tempelkan Link Google Sheets Master di panel sebelah kiri terlebih dahulu.")
        elif not keyword_online.strip():
            st.warning("Masukkan NIP atau Nama pegawai yang ingin dicari.")
        else:
            try:
                gc = get_gsheet_client()
                sh = gc.open_by_url(gsheet_url.strip())
                semua_tab = sh.worksheets()

                hasil_pencarian = []
                keyword_clean = keyword_online.strip().upper()

                with st.spinner(f"Memindai {len(semua_tab)} tab bulan di '{sh.title}'..."):
                    for ws in semua_tab:
                        records = ws.get_all_values()
                        for row in records[1:]:
                            if len(row) < 5:
                                continue
                            col0 = str(row[0]).strip()
                            col1 = str(row[1]).strip()
                            col2 = str(row[2]).strip() if len(row) > 2 else ""

                            # Format baru 9 kolom
                            if col0.isdigit() and len(row) >= 9:
                                nip_val, nama_val = col1, col2
                                hadir_val, total_val = row[3], row[4]
                                telat_val, tdk_pulang_val = row[5], row[6]
                                persen_hadir, persen_jam = row[7], row[8]
                            # Format 8 kolom
                            elif col0.isdigit() and len(row) >= 8:
                                nip_val, nama_val = col1, col2
                                hadir_val, total_val = row[3], row[4]
                                telat_val, tdk_pulang_val = row[5], "0"
                                persen_hadir, persen_jam = row[6], row[7]
                            # Format lama 6 kolom
                            elif col0.isdigit() and len(row) >= 6:
                                nip_val, nama_val = col1, col2
                                hadir_val, total_val = row[3], row[4]
                                telat_val, tdk_pulang_val = "0", "0"
                                persen_hadir, persen_jam = row[5], row[5]
                            else:
                                nip_val, nama_val = col0, col1
                                hadir_val, total_val = row[2], row[3]
                                telat_val, tdk_pulang_val = "0", "0"
                                persen_hadir, persen_jam = row[4], row[4]

                            if "%" in str(telat_val):
                                telat_val = str(int(round(float(str(telat_val).replace("%", "").strip()) / 100)))

                            if keyword_clean in nip_val.upper() or keyword_clean in nama_val.upper():
                                tampilan_b, thn_val, bln_val = ekstrak_periode_sort(ws.title)
                                hasil_pencarian.append({
                                    "NIP": nip_val,
                                    "Nama": nama_val,
                                    "Bulan": tampilan_b,
                                    "Total Masuk": hadir_val,
                                    "Efektif Hari Kerja": total_val,
                                    "Terlambat (Menit)": telat_val,
                                    "Total Tidak Absen Pulang": tdk_pulang_val,
                                    "Presentase Kehadiran": persen_hadir,
                                    "Presentase Jam Kerja": persen_jam,
                                    "_sort_tahun": thn_val,
                                    "_sort_bulan": bln_val,
                                })
                                break

                if hasil_pencarian:
                    df_hasil = pd.DataFrame(hasil_pencarian).sort_values(
                        by=["_sort_tahun", "_sort_bulan"], ascending=[True, True]
                    )
                    df_display = df_hasil.drop(columns=["_sort_tahun", "_sort_bulan"])

                    st.success(f"Ditemukan **{len(df_display)} bulan** riwayat presensi!")
                    st.dataframe(df_display, use_container_width=True, hide_index=True)

                    excel_data = format_excel_custom(df_display)
                    output_name = f"Rekap_Absensi_{keyword_online.strip().replace(' ', '_')}.xlsx"
                    st.download_button(
                        label="📥 Unduh Hasil Format Excel (.xlsx)",
                        data=excel_data,
                        file_name=output_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                else:
                    st.warning(f"Data '{keyword_online}' tidak ditemukan di tab mana pun pada Google Sheets.")
            except Exception as e:
                st.error(f"Gagal mengakses Google Sheets: {e}")
