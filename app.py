import io
import os
import re
import pandas as pd
import streamlit as st
from streamlit_option_menu import option_menu

# ==========================================
# FUNGSI HELPER / PENDUKUNG
# ==========================================


def remove_titles_and_clean(val):
  """Membersihkan gelar akademik/kehormatan dan merapikan format nama."""
  if pd.isna(val):
    return ""
  text = str(val)
  text = re.sub(r"[,;\/].*", "", text)
  titles_pattern = r"\b(S\.Kom|M\.Kom|S\.E|S\.H|S\.T|M\.T|S\.Pd|M\.Pd|Dr|Ir|H\.|Hj\.)\.?"
  text = re.sub(titles_pattern, "", text, flags=re.IGNORECASE)
  text = re.sub(r"[\.,]+$", "", text)
  text = text.strip()
  text = " ".join(text.split())
  return text.title()


def check_partial_words(str1, str2):
  """Pencocokan fleksibel berbasis kata antara dua string."""
  words1 = set(str1.split())
  words2 = set(str2.split())
  if not words1 or not words2:
    return False
  common = words1.intersection(words2)
  if len(words1) <= 2 or len(words2) <= 2:
    return len(common) >= min(len(words1), len(words2))
  else:
    return len(common) >= 2


def convert_df_to_excel(df, sheet_name="Sheet1"):
  """Mengonversi DataFrame menjadi byte stream Excel untuk tombol download."""
  output = io.BytesIO()
  with pd.ExcelWriter(output, engine="openpyxl") as writer:
    df.to_excel(writer, index=False, sheet_name=sheet_name)
  return output.getvalue()


def clean_nip_data(val):
  """Membersihkan NIP dari tanda petik atau karakter non-digit."""
  if pd.isna(val):
    return val
  text = str(val).replace("'", "").replace('"', "")
  return re.sub(r"\D", "", text)


# ==========================================
# KONFIGURASI HALAMAN STREAMLIT
# ==========================================
st.set_page_config(
    page_title="Aplikasi Absensi Karyawan", page_icon="📊", layout="wide"
)

# ==========================================
# NAVIGASI SIDEBAR
# ==========================================
with st.sidebar:
  menu = option_menu(
      menu_title="Main Menu",
      options=["Beranda (Cek Absen)", "Convert Data", "Author"],
      icons=["house-door", "arrow-repeat", "person-badge"],
      menu_icon="tv",
      default_index=0,
      styles={
          "container": {"padding": "0!important", "background-color": "transparent"},
          "icon": {"color": "white", "font-size": "18px"},
          "nav-link": {
              "font-size": "15px",
              "text-align": "left",
              "margin": "0px",
              "color": "white",
              "--hover-color": "rgba(255, 255, 255, 0.1)",
          },
          "nav-link-selected": {"background-color": "#ff4b4b"},
      },
  )

# ==========================================
# MENU 1: BERANDA (CEK ABSEN)
# ==========================================
if menu == "Beranda (Cek Absen)":
  st.title("👤 Daftar Karyawan Belum Absen")
  st.markdown(
      "Aplikasi membandingkan kehadiran berdasarkan **Nama** dengan pembersih"
      " gelar otomatis serta filter status kepegawaian."
  )
  st.markdown("---")

  col_up1, col_up2 = st.columns(2)
  with col_up1:
    master_file = st.file_uploader(
        "1. Upload File Master Karyawan",
        type=["xlsx", "xls"],
        key="master_file_input",
    )
  with col_up2:
    rekap_file = st.file_uploader(
        "2. Upload File Rekap Absen", type=["xlsx", "xls"], key="rekap_file_input"
    )

  if master_file and rekap_file:
    try:
      df_master = pd.read_excel(master_file)
      df_rekap = pd.read_excel(rekap_file)

      # Deteksi otomatis kolom status jika ada di file Master Excel
      status_cols = [
          col
          for col in df_master.columns
          if any(
              keyword in col.lower()
              for keyword in ["status", "keterangan", "keaktifan"]
          )
      ]

      st.markdown("---")
      st.subheader("⚙️ Pengaturan Filter Status Pegawai")

      df_master_filtered = df_master.copy()

      if status_cols:
        col_status = status_cols[0]
        st.info(
            f"ℹ️ Sistem mendeteksi kolom status kepegawaian pada: **{col_status}**"
        )

        # Checkbox interaktif untuk menyembunyikan pegawai Non Aktif
        abaikan_nonaktif = st.checkbox(
            "Kecualikan / Sembunyikan Pegawai dengan Status 'Non Aktif'",
            value=True,
        )

        if abaikan_nonaktif:
          df_master_filtered["_status_clean"] = (
              df_master_filtered[col_status]
              .astype(str)
              .str.lower()
              .str.strip()
          )
          # Menggunakan str.contains untuk menangkap teks seperti "Non Aktif (Tugas Belajar)"
          df_master_filtered = df_master_filtered[
              ~df_master_filtered["_status_clean"].str.contains(
                  "non aktif|non-aktif|nonaktif", na=False
              )
          ].copy()
          df_master_filtered = df_master_filtered.drop(
              columns=["_status_clean"]
          )

          jumlah_dilewati = len(df_master) - len(df_master_filtered)
          if jumlah_dilewati > 0:
            st.warning(
                f"⚠️ Sebanyak {jumlah_dilewati} pegawai dengan status 'Non"
                " Aktif' berhasil disembunyikan dari daftar absen."
            )
      else:
        st.warning(
            "⚠️ Tidak ditemukan kolom 'Status' di file Excel Master. Semua"
            " baris akan diproses."
        )

      st.markdown("---")
      st.subheader("⚙️ Pilih Kolom Nama untuk Pencocokan Data")
      master_cols = df_master_filtered.columns.tolist()
      rekap_cols = df_rekap.columns.tolist()

      cc3, cc4 = st.columns(2)
      with cc3:
        key_master = st.selectbox(
            "Kolom yang ingin dicocokkan di Master Karyawan (Nama):",
            master_cols,
            index=0 if "Nama" in master_cols else 0,
            key="key_master_select",
        )
      with cc4:
        key_rekap = st.selectbox(
            "Kolom yang ingin dicocokkan di Rekap Absen (Nama):",
            rekap_cols,
            index=0 if "Nama" in rekap_cols else 0,
            key="key_rekap_select",
        )

      if st.button("🚀 Proses Pengecekan Absen", key="btn_proses_absen"):
        with st.spinner("Sedang memproses dan mencocokkan data..."):
          df_processed = df_master_filtered.copy()
          df_processed["Nama_Bersih"] = df_processed[key_master].apply(
              remove_titles_and_clean
          )

          list_hadir_bersih = [
              remove_titles_and_clean(x).lower()
              for x in df_rekap[key_rekap].dropna().unique()
          ]
          df_processed["_key_clean"] = df_processed["Nama_Bersih"].str.lower()

          def cek_kehadiran(cleaned_master):
            if not cleaned_master:
              return False
            for nama_hadir in list_hadir_bersih:
              if (
                  cleaned_master in nama_hadir
                  or nama_hadir in cleaned_master
                  or check_partial_words(cleaned_master, nama_hadir)
              ):
                return True
            return False

          df_processed["Sudah_Absen"] = df_processed["_key_clean"].apply(
              cek_kehadiran
          )

          df_tidak_hadir = df_processed[
              df_processed["Sudah_Absen"] == False
          ].copy()
          df_tidak_hadir[key_master] = df_tidak_hadir["Nama_Bersih"]

          cols_to_drop = [
              col
              for col in ["Sudah_Absen", "Nama_Bersih", "_key_clean"]
              if col in df_tidak_hadir.columns
          ]
          df_tidak_hadir = df_tidak_hadir.drop(columns=cols_to_drop)

          st.session_state["df_tidak_hadir"] = df_tidak_hadir
          st.session_state["df_master_len"] = len(df_master_filtered)
          st.session_state["key_master_col"] = key_master
          st.success("✅ Data berhasil diproses!")

      if "df_tidak_hadir" in st.session_state:
        df_tidak_hadir = st.session_state["df_tidak_hadir"]

        st.markdown("---")
        st.header("🔍 Filter Berdasarkan Unit / Departemen")

        unit_cols = [
            col
            for col in df_master_filtered.columns
            if any(
                keyword in col.lower()
                for keyword in ["unit", "departemen", "bagian"]
            )
        ]

        if unit_cols:
          selected_unit = st.selectbox(
              "Pilih Unit / Departemen:",
              ["Semua Unit"]
              + list(df_tidak_hadir[unit_cols[0]].dropna().unique()),
              key="select_unit_filter",
          )
          df_tampil = (
              df_tidak_hadir
              if selected_unit == "Semua Unit"
              else df_tidak_hadir[df_tidak_hadir[unit_cols[0]] == selected_unit]
          )
        else:
          df_tampil = df_tidak_hadir

        st.metric(
            "Total Karyawan Belum Absen",
            f"{len(df_tampil)} Orang (dari"
            f" {st.session_state.get('df_master_len', 0)} Total Karyawan Aktif)",
        )

        st.markdown("### Daftar Karyawan yang Belum Absen (Tanpa Gelar):")
        if len(df_tampil) > 0:
          st.dataframe(df_tampil, use_container_width=True)

          excel_data = convert_df_to_excel(
              df_tampil, sheet_name="Belum_Absen"
          )
          st.download_button(
              label="📥 Download Daftar Belum Absen (Format Excel .xlsx)",
              data=excel_data,
              file_name="karyawan_belum_absen_tanpa_gelar.xlsx",
              mime=(
                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              ),
              key="download_btn_absen",
          )
        else:
          st.success("🎉 Luar biasa! Semua karyawan di unit ini sudah absen.")

    except Exception as e:
      st.error(f"Terjadi kesalahan saat memproses file: {e}")
  else:
    st.info(
        "👆 Silakan upload file **Master Karyawan** dan **Rekap Absen** di"
        " atas untuk memulai."
    )

# ==========================================
# MENU 2: CONVERT DATA
# ==========================================
elif menu == "Convert Data":
  st.title("🔄 Menu Konversi & Pilih Kolom Export")
  st.markdown("Gunakan menu ini untuk merapikan file Excel dan data NIP.")
  st.markdown("---")

  convert_file = st.file_uploader(
      "Upload File Excel yang ingin dikonversi:",
      type=["xlsx", "xls"],
      key="upload_conv",
  )

  if convert_file:
    try:
      df_conv = pd.read_excel(convert_file)
      st.success("File berhasil dimuat!")
      st.subheader("Preview Data Asli")
      st.dataframe(df_conv.head(5), use_container_width=True)

      opsi_konversi = st.selectbox(
          "Pilih tindakan konversi (Opsional):",
          [
              "-- Tanpa Konversi (Hanya Pilih Kolom) --",
              "Merapihkan angka pada NIP Absen (Hapus tanda petik/spasi)",
              "Hapus Gelar pada Kolom Nama",
          ],
      )

      kolom_target = None
      if opsi_konversi != "-- Tanpa Konversi (Hanya Pilih Kolom) --":
        label_text = (
            "Pilih kolom NIP yang ingin dirapikan:"
            if "NIP" in opsi_konversi
            else "Pilih kolom nama yang ingin dibersihkan gelarnya:"
        )
        kolom_target = st.selectbox(label_text, df_conv.columns.tolist())

      kolom_terpilih = st.multiselect(
          "Daftar Kolom:",
          df_conv.columns.tolist(),
          default=df_conv.columns.tolist(),
      )

      if st.button("Proses Data"):
        if not kolom_terpilih:
          st.warning("⚠️ Harap pilih minimal 1 kolom untuk di-export!")
        else:
          df_hasil = df_conv.copy()
          if (
              opsi_konversi != "-- Tanpa Konversi (Hanya Pilih Kolom) --"
              and kolom_target
          ):
            if "NIP" in opsi_konversi:
              df_hasil[kolom_target] = df_hasil[kolom_target].apply(
                  clean_nip_data
              )
            elif opsi_konversi == "Hapus Gelar pada Kolom Nama":
              df_hasil[kolom_target] = df_hasil[kolom_target].apply(
                  remove_titles_and_clean
              )

          st.session_state["df_final_export"] = df_hasil[kolom_terpilih]
          st.success("✅ Proses konversi dan pemilihan kolom berhasil!")

      if "df_final_export" in st.session_state:
        st.markdown("---")
        st.subheader("📋 Preview Hasil Akhir:")
        st.dataframe(
            st.session_state["df_final_export"].head(5), use_container_width=True
        )

        excel_data = convert_df_to_excel(
            st.session_state["df_final_export"], sheet_name="Export_Data"
        )

        st.download_button(
            label="📥 Download File Excel yang Dipilih (.xlsx)",
            data=excel_data,
            file_name="data_hasil_export.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )

    except Exception as e:
      st.error(f"Terjadi kesalahan: {e}")

# ==========================================
# MENU 3: AUTHOR
# ==========================================
elif menu == "Author":
  st.title("👨‍💻 Tentang Author")
  st.markdown("Informasi mengenai pengembang aplikasi ini.")
  st.markdown("---")

  st.info(
      "Aplikasi **Aplikasi Absensi Karyawan** ini dikembangkan secara mandiri"
      " untuk memudahkan proses pengecekan rekapitulasi absensi."
  )

  st.markdown("### 📌 Detail Author")
  st.markdown("- **Nama:** Irfan Bayu Seno")
  st.markdown("- **No. Telepon / WhatsApp:** 087775762410")
  st.markdown(
      "- **LinkedIn:**"
      " [Irfan Bayu"
      " Seno](https://www.linkedin.com/in/irfan-bayu-seno)"
  )