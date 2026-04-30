# command untuk run di terminal = streamlit run "d:\Axel\Byru\Monthly Report AE Automation\Monthly Report AE.py"

import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import re
import matplotlib.pyplot as plt
import io
import base64


st.set_page_config(page_title="Byru Monthly Report", layout="wide")

# ======================
# STYLE
# ======================
st.markdown("""
<style>

/* BACKGROUND HALUS BIRU */
body {
    background: linear-gradient(180deg, #f5f9fd 0%, #ffffff 100%);
}

/* TITLE */
.title {
    text-align:center;
    font-size:28px;
    font-weight:bold;
    color:#01120A;
}

/* KPI CARD */
.kpi {
    padding:12px;
    border-radius:10px;
    background:white;
    text-align:center;
    border-left:4px solid #26A9E0;
    box-shadow: 0 2px 6px rgba(0,0,0,0.05);
}

/* TEXTAREA */
.copy-box textarea {
    height:200px;
}

/* DIVIDER */
hr {
    border: none;
    height: 1px;
    background: #e5e7eb;
}

</style>
""", unsafe_allow_html=True)

# ======================
# UTIL
# ======================
def format_rp(val):
    try:
        val = int(float(val))
        return f"Rp {val:,}".replace(",", ".")
    except:
        return val
    
def format_tanggal(val):
    try:
        dt = pd.to_datetime(val)
        return dt.strftime("%#d %b %Y")  # contoh: 2 Jan 2025
    except:
        return val

def build_table(df):
    df = df.copy()

    # ======================
    # HAPUS KOLOM NO JIKA SUDAH ADA
    # ======================
    df.columns = [str(c).strip().upper() for c in df.columns]

    if "NO" in df.columns:
        df = df.drop(columns=["NO"])

    # ======================
    # TAMBAH NOMOR BARU
    # ======================
    df.insert(0, "No", range(1, len(df) + 1))

    html = "<table width='100%' style='border-collapse:collapse;font-size:12px;'>"

    # HEADER
    html += "<tr>"
    for col in df.columns:
        html += f"""
        <th style='border:1px solid #ddd;padding:6px;
        background:#26A9E0;color:white'>
        {col}
        </th>
        """
    html += "</tr>"

    # ROW
    for i, (_, row) in enumerate(df.iterrows()):
        bg = "#f9fbfd" if i % 2 == 0 else "white"

        html += f"<tr style='background:{bg}'>"
        for col in df.columns:
            html += f"""
            <td style='border:1px solid #eee;padding:6px;text-align:center'>
            {row[col]}
            </td>
            """
        html += "</tr>"

    html += "</table>"
    return html

def find_sheet(sheets, keyword):
    for s in sheets:
        if keyword.lower() in s.lower():
            return s
    return None

def extract_months(sheets):
    months = []
    for s in sheets:
        parts = s.split()
        for p in parts:
            if p.upper() in [
                "JAN","FEB","MAR","APR","MEI","JUN",
                "JUL","AGU","SEP","OKT","NOV","DES"
            ]:
                months.append(p.upper())
    return list(set(months))

# ======================
# HEADER
# ======================
st.markdown("<div class='title'>📊 Byru AE Monthly Report Generator</div>", unsafe_allow_html=True)

uploaded_file = st.file_uploader("Upload Excel", type=["xlsx"])

if uploaded_file:

    xls = pd.ExcelFile(uploaded_file)
    sheets = xls.sheet_names

    # ======================
    # AUTO DETECT SHEETS
    # ======================
    sheet_company = find_sheet(sheets, "NAMA PERUSAHAAN")
    sheet_status = find_sheet(sheets, "STATUS KLIEN")
    sheet_ref = find_sheet(sheets, "REFERALL")
    sheet_fee = find_sheet(sheets, "PERBANDINGAN")
    sheet_update = find_sheet(sheets, "UPDATE PERUSAHAAN")
    sheet_inv = find_sheet(sheets, "STATUS INV")

    if not all([sheet_company, sheet_status, sheet_ref, sheet_fee, sheet_update, sheet_inv]):
        st.error("❌ Struktur sheet tidak sesuai format Byru")
        st.stop()

    # ======================
    # LOAD DATA
    # ======================
    # df_company = pd.read_excel(xls, sheet_company, header=1)

    def extract_company_table(xls, sheet_name):
        df_raw = pd.read_excel(xls, sheet_name, header=None)

        data_rows = []
        start = False
        empty_count = 0

        col_nama = None
        col_aktif = None
        col_tidak_aktif = None
        col_awal_berlangganan = None
        col_akhir_berlangganan = None

        for i in range(len(df_raw)):
            row = df_raw.iloc[i].fillna("").astype(str).str.upper().tolist()
            row_str = " ".join(row)

            # ======================
            # DETECT START
            # ======================
            if "DATA NAMA PERUSAHAAN PERIODE" in row_str:
                start = True
                continue

            # ======================
            # DETECT HEADER (INI KUNCI FIX)
            # ======================
            if start and "NAMA PERUSAHAAN" in row:

                for idx, col in enumerate(row):
                    if "NAMA" in col:
                        col_nama = idx
                    elif "AKTIF" in col and "TIDAK" not in col:
                        col_aktif = idx
                    elif "TIDAK AKTIF" in col:
                        col_tidak_aktif = idx
                    elif "AWAL" in col:
                        col_awal_berlangganan = idx
                    elif "AKHIR" in col:
                        col_akhir_berlangganan = idx

                continue

            # ======================
            # READ DATA
            # ======================
            if start and col_nama is not None:

                row_values = df_raw.iloc[i]

                nama = row_values[col_nama]

                if pd.isna(nama) or str(nama).strip() == "":
                    empty_count += 1
                else:
                    empty_count = 0

                    user_aktif = row_values[col_aktif] if col_aktif is not None else 0
                    user_tidak_aktif = row_values[col_tidak_aktif] if col_tidak_aktif is not None else 0
                    awal_berlangganan = row_values[col_awal_berlangganan] if col_awal_berlangganan is not None else 0
                    akhir_berlangganan = row_values[col_akhir_berlangganan] if col_akhir_berlangganan is not None else 0

                    data = {
                        "NAMA PERUSAHAAN": nama,
                        "USER AKTIF": int(user_aktif) if pd.notna(user_aktif) else 0,
                        "USER TIDAK AKTIF": int(user_tidak_aktif) if pd.notna(user_tidak_aktif) else 0,
                        "AWAL BERLANGGANAN": awal_berlangganan if pd.notna(awal_berlangganan) else 0,
                        "AKHIR BERLANGGANAN": akhir_berlangganan if pd.notna(akhir_berlangganan) else 0
                    }

                    data_rows.append(data)

                # STOP jika 2 baris kosong
                if empty_count >= 2:
                    break

        df = pd.DataFrame(data_rows)

        return df
    
    df_company = extract_company_table(xls, sheet_company)
    if "AWAL BERLANGGANAN" in df_company.columns:
        df_company["AWAL BERLANGGANAN"] = df_company["AWAL BERLANGGANAN"].apply(format_tanggal)

    if "AKHIR BERLANGGANAN" in df_company.columns:
        df_company["AKHIR BERLANGGANAN"] = df_company["AKHIR BERLANGGANAN"].apply(format_tanggal)
    # STANDARDIZE
    # df_company.columns = df_company.columns.str.strip().str.upper()
    # df_company.columns = df_company.columns.map(lambda x: str(x).strip().upper())

    # def read_excel_dynamic_header(xls, sheet_name, target_column):
    #     df_raw = pd.read_excel(xls, sheet_name, header=None)

    #     for i in range(5):  # scan 5 baris pertama
    #         row = df_raw.iloc[i].astype(str).str.upper().tolist()
    #         if any(target_column in col for col in row):
    #             df = pd.read_excel(xls, sheet_name, header=i)
    #             return df

    #     return pd.read_excel(xls, sheet_name)  # fallback

    # df_company = read_excel_dynamic_header(xls, sheet_company, "USER AKTIF")
    # df_company.columns = df_company.columns.str.strip().str.upper()

    # df_status = pd.read_excel(xls, sheet_status, header=1)

    def extract_status_table(xls, sheet_name):
        df_raw = pd.read_excel(xls, sheet_name, header=None)

        data_rows = []
        start = False
        empty_count = 0

        col_map = {}

        for i in range(len(df_raw)):
            row = df_raw.iloc[i].fillna("").astype(str).str.upper().tolist()
            row_str = " ".join(row)

            # ======================
            # DETECT START
            # ======================
            if "STATUS KLIEN PERIODE" in row_str:
                start = True
                continue

            # ======================
            # DETECT HEADER
            # ======================
            if start and "NAMA" in row_str:
                for idx, col in enumerate(row):
                    if "NAMA" in col:
                        col_map["nama"] = idx
                    elif "STATUS" in col:
                        col_map["status"] = idx
                    elif "PERIODE" in col:
                        col_map["periode"] = idx
                continue

            # ======================
            # READ DATA
            # ======================
            if start and col_map:

                row_values = df_raw.iloc[i]

                nama = row_values[col_map.get("nama")]

                # STOP jika 1 baris kosong
                if pd.isna(nama) or str(nama).strip() == "":
                    empty_count += 1
                    if empty_count >= 1:
                        break
                else:
                    empty_count = 0

                    data = {
                        "NAMA": nama,
                        "STATUS": str(row_values[col_map.get("status")]).strip().upper()
                    }

                    # optional
                    if "periode" in col_map:
                        data["PERIODE"] = row_values[col_map.get("periode")]

                    data_rows.append(data)

        df = pd.DataFrame(data_rows)

        return df

    df_status = extract_status_table(xls, sheet_status)
    
    if "PERIODE" in df_status.columns:
        df_status["PERIODE"] = df_status["PERIODE"].apply(format_tanggal)

    df_ref = pd.read_excel(xls, sheet_ref, header=1)

    # df_fee = pd.read_excel(xls, sheet_fee)

    def extract_fee_comparison(xls, sheet_name):
        df_raw = pd.read_excel(xls, sheet_name, header=None)

        # ======================
        # PREVIOUS MONTH (F-H)
        # ======================
        prev_fee = df_raw.iloc[:, 5:8].copy()  # F=5
        prev_fee.columns = ["SUBSCRIBE FEE", "QTY", "IGNORE"]

        # ======================
        # CURRENT MONTH (J-L)
        # ======================
        curr_fee = df_raw.iloc[:, 9:12].copy()  # J=9
        curr_fee.columns = ["SUBSCRIBE FEE", "QTY", "IGNORE"]

        # ======================
        # CLEAN DATA
        # ======================
        def clean_df(df):
            df = df.copy()

            df["SUBSCRIBE FEE"] = df["SUBSCRIBE FEE"].astype(str).str.strip()
            df["QTY"] = pd.to_numeric(df["QTY"], errors="coerce")

            # buang baris kosong
            df = df[df["SUBSCRIBE FEE"] != "nan"]
            df = df.dropna(subset=["QTY"])

            return df[["SUBSCRIBE FEE", "QTY"]]

        prev_fee = clean_df(prev_fee)
        curr_fee = clean_df(curr_fee)

        return prev_fee, curr_fee

    prev_fee, curr_fee = extract_fee_comparison(xls, sheet_fee)


    def plot_fee_comparison(prev_fee, curr_fee):
        df_merge = pd.merge(
            prev_fee,
            curr_fee,
            on="SUBSCRIBE FEE",
            how="outer",
            suffixes=("_PREV", "_CURR")
        ).fillna(0)

        labels = df_merge["SUBSCRIBE FEE"]
        prev_vals = df_merge["QTY_PREV"]
        curr_vals = df_merge["QTY_CURR"]

        x = range(len(labels))

        fig, ax = plt.subplots(figsize=(8,4))

        # 🎨 Byru Color (Subtle)
        bars1 = ax.bar(x, prev_vals, width=0.4, color="#26A9E0", label="Previous Month")
        bars2 = ax.bar([i + 0.4 for i in x], curr_vals, width=0.4, color="#F6921E", label="Current Month")

        # LABEL ANGKA
        for bar in bars1:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2, height, int(height),
                    ha='center', va='bottom', fontsize=8)

        for bar in bars2:
            height = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2, height, int(height),
                    ha='center', va='bottom', fontsize=8)

        # AXIS
        ax.set_xticks([i + 0.2 for i in x])
        ax.set_xticklabels(labels, rotation=30)

        ax.set_xlabel("Subscribe Fee")
        ax.set_ylabel("Quantity")

        ax.set_title("Perbandingan Subscribe Fee (QTY)")

        # GRID SUBTLE
        ax.grid(axis='y', linestyle='--', alpha=0.2)

        ax.legend()

        plt.tight_layout()

        return fig
    

    def fig_to_base64(fig):
        buf = io.BytesIO()
        fig.savefig(buf, format="png", bbox_inches='tight')
        buf.seek(0)
        img_base64 = base64.b64encode(buf.read()).decode("utf-8")
        return img_base64
    # df_update = pd.read_excel(xls, sheet_update)

    def extract_update_company(xls, sheet_name):
        df_raw = pd.read_excel(xls, sheet_name, header=None)

        new_data = []
        close_data = []

        mode = None

        for i in range(len(df_raw)):
            row = df_raw.iloc[i].fillna("").astype(str).str.upper().tolist()

            # ======================
            # DETECT SECTION
            # ======================
            if any("NEW COMPANY" in cell for cell in row):
                mode = "NEW"
                continue

            if any("CLOSE COMPANY" in cell for cell in row):
                mode = "CLOSE"
                continue

            # ======================
            # SKIP HEADER
            # ======================
            if "NAMA PERUSAHAAN" in row:
                continue

            # ======================
            # AMBIL DATA
            # ======================
            if mode in ["NEW", "CLOSE"]:
                nama = df_raw.iloc[i, 0]
                qty = df_raw.iloc[i, 1]
                fee = df_raw.iloc[i, 2]

                if pd.notna(nama) and str(nama).strip() != "":
                    
                    if mode == "NEW":
                        data = {
                            "Nama Perusahaan Baru": nama,
                            "QTY": qty,
                            "Subscribe Fee": fee
                        }
                        new_data.append(data)

                    elif mode == "CLOSE":
                        data = {
                            "Nama Perusahaan Tutup": nama,
                            "QTY": qty,
                            "Subscribe Fee": fee
                        }
                        close_data.append(data)

        df_new = pd.DataFrame(new_data)
        df_close = pd.DataFrame(close_data)

        df_new.columns = df_new.columns.str.strip().str.upper()
        df_close.columns = df_close.columns.str.strip().str.upper()

        for col in df_new.columns:
            if "FEE" in col:
                df_new[col] = df_new[col].apply(format_rp)

        for col in df_close.columns:
            if "FEE" in col:
                df_close[col] = df_close[col].apply(format_rp)

        return df_new, df_close

    # df_inv = pd.read_excel(xls, sheet_inv, header=1)
    # import pandas as pd
    # import re

    def extract_number(val):
        if pd.isna(val):
            return None
        val = str(val).replace(".", "").replace(",", "")
        nums = re.findall(r"\d+", val)
        return int(nums[0]) if nums else None


    def find_column(df, keywords):
        for col in df.columns:
            col_str = str(col).upper()
            if all(k in col_str for k in keywords):
                return col
        return None


    def extract_invoice_data(xls, sheet_name):
        df_raw = pd.read_excel(xls, sheet_name, header=None)

        # ======================
        # DETECT HEADER OTOMATIS
        # ======================
        header_row = None
        for i in range(len(df_raw)):
            row = df_raw.iloc[i].fillna("").astype(str).str.upper().tolist()
            if any("PERUSAHAAN" in c for c in row) and any("STATUS" in c for c in row):
                header_row = i
                break

        if header_row is None:
            raise Exception("❌ Header tabel invoice tidak ditemukan")

        df = pd.read_excel(xls, sheet_name, header=header_row)

        # ======================
        # NORMALISASI KOLOM
        # ======================
        df.columns = [str(col).strip().upper() for col in df.columns]

        # ======================
        # FIND COLUMN (SAFE)
        # ======================
        col_nama = find_column(df, ["PERUSAHAAN"])
        col_jangka = find_column(df, ["JANGKA"])
        col_user = find_column(df, ["USER"])
        col_harga = find_column(df, ["HARGA"])
        col_total = find_column(df, ["INVOICE"])
        col_status = find_column(df, ["STATUS"])

        # VALIDASI
        if not all([col_nama, col_jangka, col_user, col_harga, col_total, col_status]):
            raise Exception(f"❌ Kolom tidak lengkap: {df.columns}")

        # ======================
        # CLEAN DATA
        # ======================
        df = df.dropna(subset=[col_nama])
        df = df[~df[col_nama].astype(str).str.upper().str.contains("TOTAL")]

        # ======================
        # SPLIT DATA (FIXED)
        # ======================
        df["STATUS_CLEAN"] = df[col_status].astype(str).str.upper()

        df_paid = df[
            (df["STATUS_CLEAN"].str.contains("LUNAS")) &
            (~df["STATUS_CLEAN"].str.contains("BELUM"))
        ]

        df_unpaid = df[
            df["STATUS_CLEAN"].str.contains("BELUM")
        ]

        # ======================
        # FORMAT
        # ======================S
        def format_df(d):
            df_out = pd.DataFrame({
                "Nama Perusahaan": d[col_nama],
                "Jangka": d[col_jangka],
                "Total User": pd.to_numeric(d[col_user], errors="coerce"),
                "Harga Per-User": pd.to_numeric(d[col_harga], errors="coerce"),
                "Total Invoice": pd.to_numeric(d[col_total], errors="coerce"),
            })

            # ======================
            # CONVERT KE INTEGER
            # ======================
            df_out["Total User"] = df_out["Total User"].fillna(0).astype(int)
            df_out["Harga Per-User"] = df_out["Harga Per-User"].fillna(0).astype(int)
            df_out["Total Invoice"] = df_out["Total Invoice"].fillna(0).astype(int)

            return df_out

        df_paid = format_df(df_paid)
        df_unpaid = format_df(df_unpaid)

        # ======================
        # TOTAL (FIXED)
        # ======================
        total_paid = pd.to_numeric(df_paid["Total Invoice"], errors="coerce").sum()
        total_unpaid = pd.to_numeric(df_unpaid["Total Invoice"], errors="coerce").sum()

        df_paid.columns = df_paid.columns.str.strip().str.upper()
        df_unpaid.columns = df_unpaid.columns.str.strip().str.upper()

        for df_ in [df_paid, df_unpaid]:
            for col in df_.columns:
                if "HARGA" in col or "TOTAL INVOICE" in col:
                    df_[col] = df_[col].apply(format_rp)

        return df_paid, df_unpaid, total_paid, total_unpaid


    # ======================
    # SELECT MONTH
    # ======================
    # months = sorted(df_fee["BULAN"].dropna().unique())



    # ======================
    # PROCESS
    # ======================
    df_company = df_company.sort_values(by="USER AKTIF", ascending=False)
    top10 = df_company.head(10)

    total_active = df_company["USER AKTIF"].sum()
    total_inactive = df_company["USER TIDAK AKTIF"].sum()

    trial = len(df_status[df_status["STATUS"] == "TRIAL"])
    stop_trial = len(df_status[df_status["STATUS"] == "STOP TRIAL"])
    waiting = len(df_status[df_status["STATUS"] == "MENUNGGU PEMBAYARAN"])

    df_ref["Periode"] = pd.to_datetime(df_ref["Periode"], errors='coerce')
    last_year = datetime.now().replace(year=datetime.now().year - 1)
    df_ref = df_ref[df_ref["Periode"] >= last_year]
    total_ref = df_ref["Komisi (10%)"].sum()
    if "Periode" in df_ref.columns:
        df_ref["Periode"] = df_ref["Periode"].apply(format_tanggal)
    
    df_ref.columns = df_ref.columns.str.strip().str.upper()

    for col in df_ref.columns:
        if "KOMISI" in col:
            df_ref[col] = df_ref[col].apply(format_rp)

    # current_fee = df_fee[df_fee["BULAN_FIX"] == selected_month]
    # prev_fee = df_fee[df_fee[bulan_col] == prev_month] if prev_month else pd.DataFrame()

    # new_company = df_update[df_update["STATUS"] == "BARU"]
    # closed_company = df_update[df_update["STATUS"] == "TUTUP"]
    new_company, closed_company = extract_update_company(xls, sheet_update)

    # paid = df_inv[df_inv["Status"] == "LUNAS"]
    # unpaid = df_inv[df_inv["Status"] != "LUNAS"]

    # total_paid = paid["TOTAL"].sum()
    # total_unpaid = unpaid["TOTAL"].sum()
    paid, unpaid, total_paid, total_unpaid = extract_invoice_data(xls, sheet_inv)

    # ======================
    # KPI
    # ======================
    col1, col2, col3 = st.columns(3)
    col1.metric("User Aktif", int(total_active))
    col2.metric("User Tidak Aktif", int(total_inactive))
    col3.metric("Referral", format_rp(total_ref))

    # ======================
    # CHART PERBANDINGAN FEE
    # ======================
    st.subheader("Perbandingan Subscribe Fee (QTY)")

    fig = plot_fee_comparison(prev_fee, curr_fee)
    st.pyplot(fig)

    chart_base64 = fig_to_base64(fig)
    # ======================
    # HTML GENERATOR
    # ======================
    html = f"""
    <div style="width:980px;margin:auto;font-family:Arial">

    <h2 style="
        text-align:center;
        background: linear-gradient(90deg,#26A9E0,#237FEA);
        color:white;
        padding:12px;
        border-radius:8px;
        border-bottom:3px solid #F6921E;
    ">
    
    Monthly Report {datetime.now().strftime('%B %Y')}
    </h2>

    <h2 style="color:#26A9E0;border-bottom:1px solid #e5e7eb;padding-bottom:4px;">
    Top 10 Perusahaan 
    </h2>
    {build_table(top10)}

    <p>Total Aktif: {total_active} | Tidak Aktif: {total_inactive}</p>

    <h2 style="color:#26A9E0;border-bottom:1px solid #e5e7eb;">Status Client</h2>
    {build_table(df_status)}
    <p>Trial: {trial} | Stop Trial: {stop_trial} | Waiting: {waiting}</p>

    <h2 style="color:#26A9E0;border-bottom:1px solid #e5e7eb;">Referral Fee (1 Tahun)</h2>
    {build_table(df_ref)}
    <p>Total: {format_rp(total_ref)}</p>

    <h2 style="color:#26A9E0;border-bottom:1px solid #e5e7eb;">Perbandingan Fee</h2>

    <img src="data:image/png;base64,{chart_base64}" 
    style="width:100%;max-width:700px;display:block;margin:auto;" />

    <h2 style="color:#26A9E0;border-bottom:1px solid #e5e7eb;">Update Perusahaan</h2>
    {build_table(new_company)}
    <br>
    {build_table(closed_company)}

    <h2 style="color:#26A9E0;border-bottom:1px solid #e5e7eb;">Status Invoice</h2>
    <p>Lunas: {format_rp(total_paid)} | Belum Lunas: {format_rp(total_unpaid)}</p>

    <h4>Lunas</h4>
    {build_table(paid)}

    <h4>Belum Lunas</h4>
    {build_table(unpaid)}

    </div>
    """

    # ======================
    # OUTPUT
    # ======================
    st.subheader("Preview")
    st.components.v1.html(html, height=800, scrolling=True)

    st.download_button("Download HTML", html, file_name="monthly_report.html")

    # ======================
    # COPY HTML
    # ======================
    st.subheader("Copy to Email")
    st.text_area("HTML Ready (Copy All)", html, height=250)