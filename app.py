import io
import re
import pandas as pd
import streamlit as st

# Konfigurasi halaman Streamlit
st.set_page_config(
    page_title="Aplikasi Absensi Karyawan", page_icon="📊", layout="wide"
)

# ==========================================
# NAVIGASI SIDEBAR
# ==========================================
st.sidebar.title("📌 Menu Navigasi")
menu = st.sidebar.radio(
    "Pilih Menu:", ["🏠 Beranda (Cek Absen)", "🔄 Convert Data"]
)

# ==========================================
# MENU 1: BERANDA (CEK ABSEN)
# ==========================================
if menu == "🏠 Beranda (Cek Absen)":
  st.title("👤 Daftar Karyawan Belum Absen")
  st.markdown(
      "Aplikasi membandingkan kehadiran berdasarkan **Nama** dengan pembersih"
      " gelar otomatis."
  )
  st.markdown("---")

  # Area Upload File di tengah (halaman utama)
  st.subheader("📁 Upload File Excel")
  col_up1, col_up2 = st.columns(2)
  with col_up1:
    master_file = st.file_uploader(
        "1. Upload File Master Karyawan", type=["xlsx", "xls"], key="master"
    )
  with col_up2:
    rekap_file = st.file_uploader(
        "2. Upload File Rekap Absen", type=["xlsx", "xls"], key="rekap"
    )

  if master_file and rekap_file:
    try:
      # Membaca file Excel
      df_master = pd.read_excel(master_file)
      df_rekap = pd.read_excel(rekap_file)

      st.success("File berhasil diunggah!")

      # Pengaturan Kolom Pencocokan Nama
      st.markdown("---")
      st.subheader("⚙️ Pilih Kolom Nama untuk Pencocokan Data")
      master_cols = df_master.columns.tolist()
      rekap_cols = df_rekap.columns.tolist()

      cc3, cc4 = st.columns(2)
      with cc3:
        key_master = st.selectbox(
            "Kolom yang ingin dicocokkan di Master Karyawan (Nama):",
            master_cols,
            index=0 if "Nama" in master_cols else 0,
        )
      with cc4:
        key_rekap = st.selectbox(
            "Kolom yang ingin dicocokkan di Rekap Absen (Nama):",
            rekap_cols,
            index=0 if "Nama" in rekap_cols else 0,
        )

      # FUNGSI PEMBERSIH GELAR & FORMAT NAMA YANG KUAT
      def remove_titles_and_clean(val):
        if pd.isna(val):
          return ""
        text = str(val)

        # 1. Hapus gelar di belakang koma, titik koma, atau garis miring (contoh: "Budi, S.Kom" atau "Budi/S.T")
        text = re.sub(r"[,;\/].*", "", text)

        # 2. Daftar singkatan gelar lengkap yang sering muncul di depan/belakang nama
        titles_pattern = r"\b(dr|drg|prof|ir|h|hj|dra|drs|se|sh|skom|mkom|mm|mpd|spd|si|st|mt|mh|msc|bsc|ba|ama|ak|S\.Kom|M\.Kom|S\.E|S\.H|S\.T|M\.T|S\.Pd|M\.Pd|Dr|Ir|H\.|Hj\.)\.?"
        text = re.sub(titles_pattern, "", text, flags=re.IGNORECASE)

        # 3. Bersihkan sisa titik/tanda baca berlebih di ujung nama
        text = re.sub(r"[\.,]+$", "", text)

        # 4. Rapikan spasi
        text = text.strip()
        text = " ".join(text.split())
        return text.title()

      # Buat salinan dataframe master
      df_processed = df_master.copy()

      # Bersihkan nama utama di master (tanpa gelar)
      df_processed["Nama_Bersih"] = df_processed[key_master].apply(
          remove_titles_and_clean
      )

      with st.spinner("Sedang memproses dan mencocokkan data..."):
        # Bersihkan list nama hadir untuk dicocokkan
        list_hadir_bersih = [
            remove_titles_and_clean(x).lower()
            for x in df_rekap[key_rekap].dropna().unique()
        ]

        df_processed["_key_clean"] = df_processed["Nama_Bersih"].str.lower()

        # Fungsi pencocokan fleksibel berbasis kata
        def check_partial_words(str1, str2):
          words1 = set(str1.split())
          words2 = set(str2.split())
          if not words1 or not words2:
            return False
          common = words1.intersection(words2)
          if len(words1) <= 2 or len(words2) <= 2:
            return len(common) >= min(len(words1), len(words2))
          else:
            return len(common) >= 2  # Minimal 2 kata yang sama

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

      # Filter hanya karyawan yang belum absen
      df_tidak_hadir = df_processed[
          df_processed["Sudah_Absen"] == False
      ].copy()

      # Ganti kolom nama asli dengan nama bersih tanpa gelar secara permanen di output
      df_tidak_hadir[key_master] = df_tidak_hadir["Nama_Bersih"]

      # Buang kolom bantu sementara
      cols_to_drop = [
          col
          for col in ["Sudah_Absen", "Nama_Bersih", "_key_clean"]
          if col in df_tidak_hadir.columns
      ]
      df_tidak_hadir = df_tidak_hadir.drop(columns=cols_to_drop)

      st.markdown("---")
      st.header("🔍 Filter Berdasarkan Unit / Departemen")

      # Deteksi kolom unit secara otomatis
      unit_cols = [
          col
          for col in df_master.columns
          if "unit" in col.lower()
          or "departemen" in col.lower()
          or "bagian" in col.lower()
      ]

      if unit_cols:
        selected_unit = st.selectbox(
            "Pilih Unit / Departemen:",
            ["Semua Unit"] + list(df_tidak_hadir[unit_cols[0]].dropna().unique()),
        )
        if selected_unit != "Semua Unit":
          df_tampil = df_tidak_hadir[
              df_tidak_hadir[unit_cols[0]] == selected_unit
          ]
        else:
          df_tampil = df_tidak_hadir
      else:
        df_tampil = df_tidak_hadir

      # Tampilkan Ringkasan Jumlah
      st.metric(
          "Total Karyawan Belum Absen",
          f"{len(df_tampil)} Orang (dari {len(df_master)} Total Karyawan)",
      )

      st.markdown("### Daftar Karyawan yang Belum Absen (Tanpa Gelar):")
      if len(df_tampil) > 0:
        st.dataframe(df_tampil, use_container_width=True)

        # Tombol Download Excel (.xlsx) dengan nama bersih tanpa gelar
        def convert_df_to_excel(df):
          output = io.BytesIO()
          with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Belum_Absen")
          return output.getvalue()

        excel_data = convert_df_to_excel(df_tampil)

        st.download_button(
            label="📥 Download Daftar Belum Absen (Format Excel .xlsx)",
            data=excel_data,
            file_name="karyawan_belum_absen_tanpa_gelar.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
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
elif menu == "🔄 Convert Data":
  st.title("🔄 Menu Konversi & Pilih Kolom Export")
  st.markdown(
      "Gunakan menu ini untuk merapikan file Excel, membersihkan gelar pada"
      " nama, serta memilih kolom apa saja yang ingin disimpan ke file baru."
  )
  st.markdown("---")

  convert_file = st.file_uploader(
      "Upload File Excel yang ingin dikonversi:", type=["xlsx", "xls"]
  )

  if convert_file:
    try:
      df_conv = pd.read_excel(convert_file)
      st.success("File berhasil dimuat!")
      st.subheader("Preview Data Asli")
      st.dataframe(df_conv.head(5), use_container_width=True)

      st.markdown("---")
      st.subheader("⚙️ Pengaturan Konversi")

      opsi_konversi = st.selectbox(
          "Pilih tindakan konversi (Opsional):",
          [
              "-- Tanpa Konversi (Hanya Pilih Kolom) --",
              "Hapus Gelar pada Kolom Nama",
          ],
      )

      kolom_target = None
      if opsi_konversi != "-- Tanpa Konversi (Hanya Pilih Kolom) --":
        kolom_target = st.selectbox(
            "Pilih kolom nama yang ingin dibersihkan gelarnya:",
            df_conv.columns.tolist(),
        )

      st.markdown("---")
      st.subheader("📋 Pilih Kolom yang Ingin Diexport")
      st.markdown(
          "Centang kolom-kolom di bawah ini yang ingin Anda masukkan ke dalam"
          " file Excel hasil download:"
      )

      kolom_terpilih = st.multiselect(
          "Daftar Kolom:",
          df_conv.columns.tolist(),
          default=df_conv.columns.tolist(),
      )

      if st.button("Proses & Download File"):
        if not kolom_terpilih:
          st.warning("⚠️ Harap pilih minimal 1 kolom untuk di-export!")
        else:
          df_hasil = df_conv.copy()

          if (
              opsi_konversi != "-- Tanpa Konversi (Hanya Pilih Kolom) --"
              and kolom_target
          ):
            if opsi_konversi == "Hapus Gelar pada Kolom Nama":

              def remove_titles_and_clean(val):
                if pd.isna(val):
                  return val
                text = str(val)
                text = re.sub(r"[,;\/].*", "", text)
                titles_pattern = r"\b(S\.Kom|M\.Kom|S\.E|S\.H|S\.T|M\.T|S\.Pd|M\.Pd|Dr|Ir|H\.|Hj\.)\.?"
                text = re.sub(titles_pattern, "", text, flags=re.IGNORECASE)
                text = re.sub(r"[\.,]+$", "", text)
                text = text.strip()
                text = " ".join(text.split())
                return text.title()

              df_hasil[kolom_target] = df_hasil[kolom_target].apply(
                  remove_titles_and_clean
              )

          df_final_export = df_hasil[kolom_terpilih]

          st.success("Konversi dan pemilihan kolom berhasil!")
          st.subheader("Preview Hasil Akhir:")
          st.dataframe(df_final_export.head(5), use_container_width=True)

          output = io.BytesIO()
          with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_final_export.to_excel(writer, index=False, sheet_name="Export_Data")
          excel_data = output.getvalue()

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