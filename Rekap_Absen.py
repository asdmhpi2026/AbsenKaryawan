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
st.title("📝 Sistem Rekap & Database Absensi Online")

# ==============================================================================
# KONFIGURASI GOOGLE SHEETS MASTER DATABASE
# ==============================================================================
# 1. Simpan file JSON Service Account dari Google Cloud di folder yang sama dengan skrip ini
CREDENTIALS_FILE = "credentials.json"

# 2. Tempelkan URL Google Sheets Master Anda (cukup 1 file Google Sheets untuk semua bulan)
#    Pastikan email client_email di dalam credentials.json sudah di-Share sebagai "Editor"
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
            "credentials.json", scopes=scopes
        )
    return gspread.authorize(creds)

def simpan_ke_gsheet(gsheet_url, nama_tab, df_rekap):
    """Membuat tab baru (atau menimpa jika sudah ada) di Google Sheets dengan format Sheet1"""
    gc = get_gsheet_client()
    sh = gc.open_by_url(gsheet_url)

    # Cek apakah tab bulan tersebut sudah ada
    existing_titles = [ws.title for ws in sh.worksheets()]
    if nama_tab in existing_titles:
        ws = sh.worksheet(nama_tab)
        ws.clear()
    else:
        ws = sh.add_worksheet(title=nama_tab, rows=str(len(df_rekap) + 20), cols="10")

    # Siapkan data persis format Sheet1 (Presentase disimpan sebagai angka desimal / string persen)
    header = ["No", "NIP", "Nama", "Total Masuk", "Efeketif Kerja", "Presentase"]
    rows_to_write = [header]
    for idx, r in df_rekap.reset_index(drop=True).iterrows():
        rows_to_write.append([
            int(idx + 1),
            str(r["NIP"]),
            str(r["Nama"]),
            int(r["Total Masuk"]),
            int(r["Efeketif Kerja"]),
            round(float(r["Presentase"]), 6),
        ])

    ws.update(range_name="A1", values=rows_to_write, value_input_option="USER_ENTERED")

    # Format kolom F (Presentase) menjadi format persen (0%) di Google Sheets
    ws.format("F2:F" + str(len(rows_to_write)), {
        "numberFormat": {"type": "PERCENT", "pattern": "0%"}
    })
    ws.format("A1:F1", {
        "textFormat": {"bold": True},
        "horizontalAlignment": "CENTER"
    })
    return sh.title


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


def proses_data_mentah(df_raw, efektif_kerja=22):
    """
    Mengolah gabungan log mentah absensi menjadi rekap per NIP:
    1. Tanggal kembar dengan scan lengkap -> dihitung 1 (duplikatnya 0)
    2. Tanggal tunggal hanya Scan Masuk (Scan Keluar '-') -> dihitung 1
    3. Tanggal kembar yang semuanya tanpa Scan Keluar -> dihitung 0
    """
    df = df_raw.copy()
    df.columns = [str(c).strip() for c in df.columns]

    kolom_wajib = ["No Induk", "Nama", "Tanggal Masuk", "Scan Masuk", "Scan Keluar"]
    kurang = [k for k in kolom_wajib if k not in df.columns]
    if kurang:
        raise ValueError(f"Kolom wajib tidak ditemukan: {', '.join(kurang)}")

    df["NIP"] = df["No Induk"].astype(str).str.strip()
    df["Nama"] = df["Nama"].astype(str).str.strip()
    df["Tgl"] = pd.to_datetime(df["Tanggal Masuk"], errors="coerce").dt.strftime("%Y-%m-%d")

    # Buang baris kosong / header berulang jika ada
    df = df[
        df["NIP"].notna()
        & (df["NIP"] != "")
        & (df["NIP"].str.lower() != "nan")
        & (df["NIP"].str.lower() != "no induk")
        & df["Tgl"].notna()
    ].copy()

    sm_str = df["Scan Masuk"].astype(str).str.strip()
    sk_str = df["Scan Keluar"].astype(str).str.strip()

    df["sm_ok"] = (sm_str != "-") & (sm_str != "") & (sm_str.str.lower() != "nan")
    df["sk_ok"] = (sk_str != "-") & (sk_str != "") & (sk_str.str.lower() != "nan")

    # Hitung berapa kali (NIP, Tgl) muncul dalam data mentah gabungan
    tgl_count = df.groupby(["NIP", "Tgl"])["sm_ok"].transform("count")

    # Aturan 1: Scan Masuk & Keluar terisi -> ambil kemunculan pertama per (NIP, Tgl)
    lengkap = df["sm_ok"] & df["sk_ok"]
    first_lengkap = lengkap & (~df[lengkap].duplicated(subset=["NIP", "Tgl"])).reindex(df.index, fill_value=False)

    # Aturan 2: Hanya Scan Masuk terisi (Scan Keluar '-') DAN tanggal tunggal (tidak kembar)
    lupa_tunggal = df["sm_ok"] & (~df["sk_ok"]) & (tgl_count == 1)

    df["Hitung_Hadir"] = (first_lengkap | lupa_tunggal).astype(int)

    # Kelompokkan berdasarkan NIP (ambil Nama pertama agar konsisten jika ada perbedaan spasi)
    rekap = (
        df.groupby("NIP", as_index=False, sort=False)
        .agg(Nama=("Nama", "first"), Total_Masuk=("Hitung_Hadir", "sum"))
        .rename(columns={"Total_Masuk": "Total Masuk"})
    )
    rekap.insert(0, "No", range(1, len(rekap) + 1))
    rekap["Efeketif Kerja"] = int(efektif_kerja)
    rekap["Presentase"] = rekap["Total Masuk"] / float(efektif_kerja) if efektif_kerja > 0 else 0.0
    return rekap


def export_format_sheet1(df_rekap):
    """Mengekspor hasil rekap ke Excel persis seperti format Sheet1"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    font_default = Font(name="Calibri", size=11, bold=False)
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")

    headers = ["No", "NIP", "Nama", "Total Masuk", "Efeketif Kerja", "Presentase"]
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
            float(row["Presentase"]),
        ]
        for c_idx, val in enumerate(vals, 1):
            cell = ws.cell(row=excel_row, column=c_idx, value=val)
            cell.font = font_default
            if c_idx == 3:
                cell.alignment = align_left
            elif c_idx == 6:
                cell.number_format = "0%"

    lebar_kolom = {1: 6, 2: 16, 3: 38, 4: 14, 5: 15, 6: 13}
    for col_idx, width in lebar_kolom.items():
        col_letter = openpyxl.utils.get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = width

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

    ws.merge_cells("C2:G2")
    ws["C2"] = "REKAP ABSENSI PEGAWAI"
    ws["C2"].font = font_title
    ws["C2"].alignment = align_center

    headers = ["NO", "NIP", "Nama", "Bulan", "Hadir", "Total Hari", "Persentase"]
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

        try:
            hadir_val = float(row_data["Hadir"]) if pd.notna(row_data["Hadir"]) else 0
        except Exception:
            hadir_val = row_data["Hadir"]

        try:
            total_val = float(row_data["Total Hari"]) if pd.notna(row_data["Total Hari"]) else 0
        except Exception:
            total_val = row_data["Total Hari"]

        persen_val = 0
        try:
            raw_p = str(row_data["Persentase"]).replace("%", "").strip()
            v = float(raw_p)
            persen_val = v if v <= 1.0 else v / 100.0
        except Exception:
            persen_val = 0

        row_values = [no_val, nip_val, nama_val, bulan_val, hadir_val, total_val, persen_val]
        for col_idx, val in enumerate(row_values, 1):
            cell = ws.cell(row=current_row, column=col_idx, value=val)
            cell.font = font_body
            cell.border = thin_border
            if col_idx in [1, 4]:
                cell.alignment = align_center
            elif col_idx in [2, 3]:
                cell.alignment = align_left if col_idx == 3 else align_center
            elif col_idx in [5, 6]:
                cell.alignment = align_center
                if isinstance(val, (int, float)):
                    cell.number_format = "#,##0"
            elif col_idx == 7:
                cell.alignment = align_right
                cell.number_format = "0.0%" if persen_val == 1.0 or persen_val == 0 else "0.00%"

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
    if ("gcp_service_account" in st.secrets) or os.path.exists(CREDENTIALS_FILE):
        st.success("✅ Kredensial Google Sheets siap")
    else:
        st.warning("⚠️ credentials.json belum ada di folder")

# ==============================================================================
# ANTARMUKA UTAMA (2 TAB)
# ==============================================================================
tab1, tab2 = st.tabs([
    "📤 Upload Data Absen dari Simponik -> Google Sheets",
    "🔍 Cari Rekap Per NIP / Nama (Database Online)"
])

# ------------------------------------------------------------------------------
# TAB 1: UPLOAD MULTI-FILE & MULTI-TAB -> SIMPAN KE GSHEET & DOWNLOAD SHEET1
# ------------------------------------------------------------------------------
with tab1:
    st.subheader("Upload & Rekap Log Absensi Karyawan (.xls / .xlsx)")
    st.caption("Mendukung penggabungan otomatis dari beberapa file sekaligus maupun beberapa tab (seperti tab Organik & PKWT).")

    col_up, col_hk = st.columns([3, 1])
    with col_up:
        uploaded_files = st.file_uploader(
            "Upload File Absensi:",
            type=["xls", "xlsx"],
            accept_multiple_files=True
        )
    with col_hk:
        efektif_kerja = st.number_input("Hari Efektif Kerja", min_value=1, max_value=31, value=22, step=1)

    if uploaded_files:
        semua_opsi_sheet = []
        peta_excel = {}

        # Baca daftar tab dari semua file yang diunggah
        for f in uploaded_files:
            xls_obj = pd.ExcelFile(f)
            peta_excel[f.name] = xls_obj
            for s_name in xls_obj.sheet_names:
                # Abaikan tab rekap lama bernama "Sheet1" jika ada tab data mentah lain
                label_opsi = f"{f.name}  |  Tab: {s_name}"
                semua_opsi_sheet.append((label_opsi, f.name, s_name))

        # Otomatis pilih semua tab mentah (kecuali 'Sheet1' jika masih ada tab lain)
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

                col_per, col_fil = st.columns([1, 2])
                with col_per:
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

                df_rekap_full = proses_data_mentah(df_gabungan, efektif_kerja=efektif_kerja)

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
                df_preview["Presentase"] = (df_preview["Presentase"] * 100).round(0).astype(int).astype(str) + "%"
                st.dataframe(df_preview, use_container_width=True, hide_index=True)

                # Dua tombol aksi berdampingan: Simpan Otomatis ke Google Sheets & Download Excel Sheet1
                col_btn1, col_btn2 = st.columns(2)

                with col_btn1:
                    if st.button(f"☁️ Simpan Otomatis ke Google Sheets (Tab: {nama_tab_tujuan})", type="primary", use_container_width=True):
                        if not gsheet_url.strip():
                            st.error("Tempelkan Link Google Sheets Master di panel sebelah kiri terlebih dahulu.")
                        else:
                            with st.spinner(f"Menyimpan data ke tab '{nama_tab_tujuan}' di Google Sheets..."):
                                judul_gsheet = simpan_ke_gsheet(gsheet_url.strip(), nama_tab_tujuan.strip(), df_tampil)
                                st.success(f"Mantap! Data **{nama_tab_tujuan}** ({len(df_tampil)} pegawai) berhasil masuk ke Google Sheets **{judul_gsheet}**!")

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
    st.subheader("Cari Riwayat Absensi Pegawai")

    keyword_online = st.text_input(
        "Masukkan NIP atau Nama Pegawai:",
        placeholder="Contoh: 9422011HPL atau IQBAL",
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
                        for row in records[1:]:  # Lewati baris header
                            if len(row) < 5:
                                continue
                            col0 = str(row[0]).strip()
                            col1 = str(row[1]).strip()
                            col2 = str(row[2]).strip() if len(row) > 2 else ""

                            # Mendukung format Sheet1 6 kolom (No, NIP, Nama, Total Masuk, Efektif Kerja, Presentase)
                            if col0.isdigit() and len(row) >= 6:
                                nip_val, nama_val = col1, col2
                                hadir_val, total_val, persen_val = row[3], row[4], row[5]
                            else:
                                nip_val, nama_val = col0, col1
                                hadir_val, total_val, persen_val = row[2], row[3], row[4]

                            if keyword_clean in nip_val.upper() or keyword_clean in nama_val.upper():
                                tampilan_b, thn_val, bln_val = ekstrak_periode_sort(ws.title)
                                hasil_pencarian.append({
                                    "NIP": nip_val,
                                    "Nama": nama_val,
                                    "Bulan": tampilan_b,
                                    "Hadir": hadir_val,
                                    "Total Hari": total_val,
                                    "Persentase": persen_val,
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
