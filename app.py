import io
import re
import time
import pandas as pd
import requests
import schedule
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
  
  # PERBAIKAN DI SINI (menggunakan re.IGNORECASE secara langsung)
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
  output = io.BytesIO()
  with pd.ExcelWriter(output, engine="openpyxl") as writer:
    df.to_excel(writer, index=False, sheet_name=sheet_name)
  return output.getvalue()


def clean_nip_data(val):
  if pd.isna(val):
    return val
  text = str(val).replace("'", "").replace('"', "")
  return re.sub(r"\D", "", text)


# Fungsi untuk mengirim pesan WhatsApp via API Fonnte
def kirim_wa_fonnte(token_api, target_tujuan, pesan):
  url = "https://api.fonnte.com/send"
  payload = {"target": target_tujuan, "message": pesan, "country": "62"}
  headers = {"Authorization": token_api}
  try:
    response = requests.post(url, data=payload, headers=headers)
    return response.json()
  except Exception as e:
    return {"status": False, "reason": str(e)}


# ==========================================
# KONFIGURASI HALAMAN STREAMLIT
# ==========================================
st.set_page_config(
    page_title="Aplikasi Absensi Karyawan & Reminder",
    page_icon="📊",
    layout="wide",
)

# ==========================================
# NAVIGASI SIDEBAR
# ==========================================
with st.sidebar:
  menu = option_menu(
      menu_title="Main Menu",
      options=[
          "Beranda (Cek Absen)",
          "Convert Data",
          "WhatsApp Reminder",
          "Author",
      ],
      icons=["house-door", "arrow-repeat", "whatsapp", "person-badge"],
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
      " gelar otomatis."
  )
  st.markdown("---")

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
      df_master = pd.read_excel(master_file)
      df_rekap = pd.read_excel(rekap_file)
      st.success("File berhasil diunggah!")

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

      df_processed = df_master.copy()
      df_processed["Nama_Bersih"] = df_processed[key_master].apply(
          remove_titles_and_clean
      )

      with st.spinner("Sedang memproses dan mencocokkan data..."):
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

      st.markdown("---")
      st.header("🔍 Filter Berdasarkan Unit / Departemen")
      unit_cols = [
          col
          for col in df_master.columns
          if any(
              keyword in col.lower() for keyword in ["unit", "departemen", "bagian"]
          )
      ]

      if unit_cols:
        selected_unit = st.selectbox(
            "Pilih Unit / Departemen:",
            ["Semua Unit"] + list(df_tidak_hadir[unit_cols[0]].dropna().unique()),
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
          f"{len(df_tampil)} Orang (dari {len(df_master)} Total Karyawan)",
      )

      st.markdown("### Daftar Karyawan yang Belum Absen (Tanpa Gelar):")
      if len(df_tampil) > 0:
        st.dataframe(df_tampil, use_container_width=True)
        excel_data = convert_df_to_excel(df_tampil, sheet_name="Belum_Absen")
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
elif menu == "Convert Data":
  st.title("🔄 Menu Konversi & Pilih Kolom Export")
  st.markdown("Gunakan menu ini untuk merapikan file Excel dan data NIP.")
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

      if st.button("Proses & Download File"):
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

          df_final_export = df_hasil[kolom_terpilih]
          excel_data = convert_df_to_excel(df_final_export, sheet_name="Export_Data")

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
# MENU 3: WHATSAPP REMINDER (BARU)
# ==========================================
elif menu == "WhatsApp Reminder":
  st.title("⏰ Konfigurasi Pengingat Absensi WhatsApp")
  st.markdown(
      "Atur jadwal otomatis dan token API WhatsApp Gateway (contoh menggunakan"
      " **Fonnte**) untuk pengingat MagangHub."
  )
  st.markdown("---")

  with st.form("form_wa_setting"):
    st.subheader("🔑 Kredensial API & Tujuan")
    api_token = st.text_input(
        "Fonnte API Token:",
        type="password",
        placeholder="Masukkan token API Fonnte Anda di sini",
    )
    target_wa = st.text_input(
        "Nomor Tujuan / ID Grup WhatsApp:",
        placeholder="Contoh: 628123456789 atau ID Grup WhatsApp",
    )

    st.subheader("🕒 Pengaturan Waktu & Pesan")
    col_t1, col_t2 = st.columns(2)
    with col_t1:
      jam_kirim = st.time_input("Jam Pengiriman Harian:")
    with col_t2:
      status_aktif = st.checkbox("Aktifkan Pengingat Otomatis", value=False)

    pesan_default = (
        "Halo rekan-rekan MagangHub, diingatkan untuk segera melakukan absensi"
        " hari ini ya! Terima kasih. 🙏"
    )
    pesan_template = st.text_area("Template Pesan Pengingat:", value=pesan_default)

    submitted = st.form_submit_button("Simpan Pengaturan Reminder")

    if submitted:
      # Menyimpan konfigurasi ke Session State Streamlit
      st.session_state["wa_token"] = api_token
      st.session_state["wa_target"] = target_wa
      st.session_state["wa_jam"] = str(jam_kirim)
      st.session_state["wa_status"] = status_aktif
      st.success("✅ Pengaturan pengingat WhatsApp berhasil disimpan!")

  st.markdown("---")
  st.subheader("🧪 Uji Coba Kirim Pesan Manual")
  st.markdown(
      "Gunakan tombol di bawah ini untuk menguji apakah koneksi API WhatsApp"
      " Anda berfungsi dengan baik."
  )

  if st.button("🚀 Kirim Test Pesan Sekarang"):
    token_tersimpan = st.session_state.get("wa_token", "")
    target_tersimpan = st.session_state.get("wa_target", "")

    if not token_tersimpan or not target_tersimpan:
      st.warning(
          "⚠️ Harap isi dan simpan Token API serta Nomor Tujuan terlebih dahulu"
          " di form atas!"
      )
    else:
      with st.spinner("Mengirim pesan uji coba ke WhatsApp..."):
        hasil = kirim_wa_fonnte(
            token_tersimpan,
            target_tersimpan,
            "TEST MESSAGE: Bot Pengingat Absensi MagangHub aktif!",
        )
        if hasil.get("status") or hasil.get("true"):
          st.success("🎉 Pesan uji coba berhasil dikirim ke WhatsApp!")
        else:
          st.error(f"❌ Gagal mengirim pesan. Keterangan: {hasil}")

# ==========================================
# MENU 4: AUTHOR
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