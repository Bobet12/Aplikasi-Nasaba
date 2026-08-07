import os
import pandas as pd
import streamlit as st
from fpdf import FPDF

st.set_page_config(page_title="Pencatatan Tabungan Nasabah", layout="wide")

st.title("📊 Aplikasi Pencatatan Tabungan Nasabah")
st.subheader("Hitung Bunga Majemuk & Cetak Laporan PDF")

# File Excel penyimpanan
EXCEL_FILE = "Buku_Catatan_Tabungan_Bunga_Majemuk.xlsx"

# Inisialisasi file Excel jika belum ada
if not os.path.exists(EXCEL_FILE):
    df_init = pd.DataFrame(
        columns=[
            "Tanggal",
            "Nama Nasabah",
            "Setoran Awal (Rp)",
            "Bunga Per Tahun (%)",
            "Tenor (Bulan)",
            "Total Akhir (Rp)",
        ]
    )
    df_init.to_excel(EXCEL_FILE, index=False)


# --- DIALOG KONFIRMASI HAPUS DATA ---
@st.dialog("⚠️ Konfirmasi Penghapusan Data")
def konfirmasi_hapus_dialog(data_dihapus):
    st.write(
        "Apakah Anda yakin ingin menghapus data nasabah berikut dari database?"
    )

    nama_list = data_dihapus["Nama Nasabah"].tolist()
    for nama_nasabah in nama_list:
        st.markdown(f"- **{nama_nasabah}**")

    st.warning(
        "Tindakan ini akan menghapus data dari file Excel secara permanen!"
    )

    col1, col2 = st.columns(2)
    with col1:
        if st.button("✅ Ya, Hapus Sekarang", type="primary"):
            # Baca data Excel saat ini
            df_curr = pd.read_excel(EXCEL_FILE)

            # Filter out data yang ingin dihapus
            df_updated = df_curr[
                ~df_curr["Nama Nasabah"].isin(nama_list)
            ].reset_index(drop=True)

            # Simpan kembali ke Excel
            df_updated.to_excel(EXCEL_FILE, index=False)

            st.toast("Data nasabah berhasil dihapus!", icon="🗑️")
            st.rerun()

    with col2:
        if st.button("❌ Batal"):
            st.rerun()


# Form Input Data Nasabah
with st.sidebar:
    st.header("📝 Input Data Nasabah")
    nama = st.text_input("Nama Nasabah")
    setoran = st.number_input(
        "Setoran Awal (Rp)", min_value=0.0, step=100000.0, format="%.2f"
    )
    bunga = st.number_input(
        "Suku Bunga Per Tahun (%)",
        min_value=0.0,
        max_value=100.0,
        value=12.0,
        step=0.5,
    )
    tenor = st.number_input(
        "Tenor (Bulan)", min_value=1, max_value=360, value=12, step=1
    )

    simpan_btn = st.button("💾 Hitung & Simpan Data")


# Fungsi Perhitungan Bunga Majemuk
def hitung_bunga_majemuk(p, r_annual, n_months):
    r_monthly = (r_annual / 100) / 12
    detail = []
    saldo = p
    for month in range(1, n_months + 1):
        bunga_bulan = saldo * r_monthly
        saldo += bunga_bulan
        detail.append(
            {"Bulan": month, "Bunga (Rp)": bunga_bulan, "Saldo (Rp)": saldo}
        )
    return saldo, pd.DataFrame(detail)


# Fungsi Generator PDF
def generate_pdf(
    nama_nasabah, setoran_awal, rate, tenor_bulan, total_akhir, df_detail
):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", "B", 16)
    pdf.cell(200, 10, txt="LAPORAN TABUNGAN NASABAH", ln=True, align="C")
    pdf.ln(5)

    pdf.set_font("Arial", size=11)
    pdf.cell(200, 8, txt=f"Nama Nasabah : {nama_nasabah}", ln=True)
    pdf.cell(200, 8, txt=f"Setoran Awal : Rp {setoran_awal:,.2f}", ln=True)
    pdf.cell(200, 8, txt=f"Bunga / Tahun : {rate}%", ln=True)
    pdf.cell(200, 8, txt=f"Tenor        : {tenor_bulan} Bulan", ln=True)
    pdf.cell(200, 8, txt=f"Total Akhir  : Rp {total_akhir:,.2f}", ln=True)
    pdf.ln(10)

    # Tabel Detail
    pdf.set_font("Arial", "B", 10)
    pdf.cell(30, 8, "Bulan", border=1, align="C")
    pdf.cell(80, 8, "Bunga Per Bulan (Rp)", border=1, align="C")
    pdf.cell(80, 8, "Saldo Akhir Bulan (Rp)", border=1, align="C")
    pdf.ln()

    pdf.set_font("Arial", size=10)
    for index, row in df_detail.iterrows():
        pdf.cell(30, 7, str(int(row["Bulan"])), border=1, align="C")
        pdf.cell(80, 7, f"Rp {row['Bunga (Rp)']:,.2f}", border=1, align="R")
        pdf.cell(80, 7, f"Rp {row['Saldo (Rp)']:,.2f}", border=1, align="R")
        pdf.ln()

    return pdf.output(dest="S").encode("latin-1")


# Skenario setelah tombol diklik
if simpan_btn:
    if nama.strip() == "":
        st.error("Silakan masukkan Nama Nasabah terlebih dahulu!")
    else:
        total_akhir, df_detail = hitung_bunga_majemuk(setoran, bunga, tenor)

        # Simpan ke Excel
        df_existing = pd.read_excel(EXCEL_FILE)
        new_data = pd.DataFrame(
            [
                {
                    "Tanggal": pd.Timestamp.now().strftime("%Y-%m-%d %H:%M"),
                    "Nama Nasabah": nama,
                    "Setoran Awal (Rp)": setoran,
                    "Bunga Per Tahun (%)": bunga,
                    "Tenor (Bulan)": tenor,
                    "Total Akhir (Rp)": total_akhir,
                }
            ]
        )
        df_updated = pd.concat([df_existing, new_data], ignore_index=True)
        df_updated.to_excel(EXCEL_FILE, index=False)

        st.success(f"Data untuk {nama} berhasil dihitung dan disimpan!")

        # Tampilkan Hasil
        col1, col2 = st.columns(2)
        col1.metric("Setoran Awal", f"Rp {setoran:,.2f}")
        col2.metric("Total Hasil Akhir", f"Rp {total_akhir:,.2f}")

        st.write("### Rincian Perkembangan Saldo Per Bulan")

        # Format tampilan Rupiah pada tabel detail
        st.dataframe(
            df_detail,
            column_config={
                "Bunga (Rp)": st.column_config.NumberColumn(
                    "Bunga (Rp)", format="Rp %,.2f"
                ),
                "Saldo (Rp)": st.column_config.NumberColumn(
                    "Saldo (Rp)", format="Rp %,.2f"
                ),
            },
            use_container_width=True,
        )

        # Download PDF
        pdf_bytes = generate_pdf(
            nama, setoran, bunga, tenor, total_akhir, df_detail
        )
        st.download_button(
            label="📄 Download Laporan PDF",
            data=pdf_bytes,
            file_name=f"Laporan_Tabungan_{nama}.pdf",
            mime="application/pdf",
        )

# --- RIWAYAT DATA NASABAH & KELOLA HAPUS ---
st.markdown("---")
st.write("### 📁 Riwayat Data Nasabah Tersimpan")

if os.path.exists(EXCEL_FILE):
    df_history = pd.read_excel(EXCEL_FILE)

    if not df_history.empty:
        # Fitur Pencarian
        kata_kunci = st.text_input(
            "🔍 Cari Nama Nasabah:",
            placeholder="Ketik nama nasabah di sini...",
        )

        df_tampil = df_history.copy()
        if kata_kunci:
            df_tampil = df_tampil[
                df_tampil["Nama Nasabah"].str.contains(
                    kata_kunci, case=False, na=False
                )
            ]

        st.caption(
            "💡 **Petunjuk Hapus:** Centang kolom **'Hapus?'** pada nasabah yang ingin dikeluarkan, lalu klik tombol **Hapus Nasabah Terpilih**."
        )

        # Sisipkan kolom centang (Pilih) secara temporer
        df_tampil.insert(0, "Pilih", False)

        # Tabel Interaktif dengan Format Rupiah & Checkbox
        edited_df = st.data_editor(
            df_tampil,
            column_config={
                "Pilih": st.column_config.CheckboxColumn(
                    "Hapus?",
                    help="Centang untuk memilih nasabah yang akan dihapus",
                    default=False,
                ),
                "Setoran Awal (Rp)": st.column_config.NumberColumn(
                    "Setoran Awal (Rp)", format="Rp %,.2f"
                ),
                "Total Akhir (Rp)": st.column_config.NumberColumn(
                    "Total Akhir (Rp)", format="Rp %,.2f"
                ),
            },
            disabled=[
                "Tanggal",
                "Nama Nasabah",
                "Setoran Awal (Rp)",
                "Bunga Per Tahun (%)",
                "Tenor (Bulan)",
                "Total Akhir (Rp)",
            ],
            hide_index=True,
            use_container_width=True,
        )

        # Tombol Pemicu Konfirmasi Hapus
        if st.button("🗑️ Hapus Nasabah Terpilih", type="primary"):
            data_dihapus = edited_df[edited_df["Pilih"] == True]

            if not data_dihapus.empty:
                konfirmasi_hapus_dialog(data_dihapus)
            else:
                st.warning(
                    "Silakan centang terlebih dahulu nama nasabah yang ingin dihapus pada tabel."
                )
    else:
        st.info("Belum ada data nasabah yang tersimpan.")
