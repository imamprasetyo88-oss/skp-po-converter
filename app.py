import io
import os
import re
import xml.etree.ElementTree as ET
import openpyxl
import pandas as pd
import pdfplumber
import pymysql
import streamlit as st
import streamlit.components.v1 as components

st.set_page_config(page_title="Sistem Upload PO", layout="centered")

# ==========================================
# KONEKSI DATABASE HOSTINGER
# ==========================================
db_config = {
    "host": "auth-db1322.hstrgr.io",
    "user": "u378956454_gudang",  # Sesuaikan username database Anda
    "password": "Belagu2388#",  # Sesuaikan password database Anda
    "database": "u378956454_gudang",
    "cursorclass": pymysql.cursors.DictCursor,
}


def get_db_connection():
  return pymysql.connect(**db_config)


# ==========================================
# CSS KUSTOM UNTUK MEMPERCANTIK UI
# ==========================================
st.markdown(
    """
    <style>
    [data-testid="stAppViewContainer"] { background-color: #F0F4F8 !important; }
    [data-testid="stHeader"] { background-color: #F0F4F8 !important; }
    .block-container { padding-top: 1.5rem !important; }
    html, body, [class*="css"] { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    hr { margin-top: 10px; margin-bottom: 20px; border: 0; border-top: 1px solid #d1d9e6; }
    [data-testid="stFileUploaderDropzone"] {
        background-color: #FFFFFF !important;
        border: 2px dashed #1E3A8A !important;
        border-radius: 12px !important;
        padding: 20px !important;
        box-shadow: 0px 4px 6px rgba(0, 0, 0, 0.05) !important;
    }
    [data-testid="stFileUploader"] button {
        background-color: #1E3A8A !important;
        color: white !important;
        border-radius: 6px !important;
        border: none !important;
        font-weight: bold !important;
        transition: all 0.2s ease-in-out !important;
    }
    [data-testid="stFileUploader"] button:hover {
        background-color: #152b65 !important;
        color: white !important;
        transform: scale(1.02);
    }
    </style>
""",
    unsafe_allow_html=True,
)

# ==========================================
# SIDEBAR: NAVIGASI MENU UTAMA & PIC
# ==========================================
st.sidebar.title("Navigasi Sistem")
menu = st.sidebar.radio("Pilih Menu:", ["Konverter PO", "Kelola Master Data (PIC)"])

# ==========================================
# MENU 2: KELOLA MASTER DATA (FITUR BARU UNTUK PIC)
# ==========================================
if menu == "Kelola Master Data (PIC)":
  st.title("🔧 Manajemen Master Data Customer")
  st.write(
      "Halaman khusus PIC untuk melihat, mencari, dan menambah item master"
      " baru secara *real-time* ke database MySQL."
  )
  st.markdown("---")

  try:
    conn = get_db_connection()
    query = "SELECT id, customer, item_customer, item_erp, isi_per_pak FROM master_customer_items"
    df_master_db = pd.read_sql(query, conn)
    conn.close()

    # Filter berdasarkan Customer
    list_customer = df_master_db["customer"].unique().tolist()
    selected_cust = st.selectbox(
        "Pilih Customer:", ["Semua Customer"] + list_customer
    )

    if selected_cust != "Semua Customer":
      df_filtered = df_master_db[df_master_db["customer"] == selected_cust]
    else:
      df_filtered = df_master_db

    # Kotak Pencarian Item
    search_query = st.text_input("🔍 Cari Kode Item Customer atau ERP:", "")
    if search_query:
      df_filtered = df_filtered[
          df_filtered["item_customer"]
          .astype(str)
          .str.contains(search_query, case=False, na=False)
          | df_filtered["item_erp"]
          .astype(str)
          .str.contains(search_query, case=False, na=False)
      ]

    # Tampilkan Tabel Data Master
    st.dataframe(df_filtered, use_container_width=True)

    # Form Tambah Item Baru
    st.markdown("---")
    st.subheader("➕ Tambah Item Master Baru")

    with st.form("form_tambah_item"):
      col1, col2, col3 = st.columns(3)
      with col1:
        new_customer = st.selectbox(
            "Nama Customer", ["Gramedia", "Promexx", "AEON", "GrandLucky"]
        )
      with col2:
        new_item_cust = st.text_input("Item Customer (Kode dari PO)")
      with col3:
        new_item_erp = st.text_input("Item ERP (Standar SKP)")

      new_isi_pak = st.number_input(
          "Isi per Pak", min_value=1, value=1, step=1
      )

      submit_button = st.form_submit_button("Simpan Item Baru")

      if submit_button:
        if new_item_cust and new_item_erp:
          try:
            conn = get_db_connection()
            with conn.cursor() as cursor:
              sql = "INSERT INTO master_customer_items (customer, item_customer, item_erp, isi_per_pak) VALUES (%s, %s, %s, %s)"
              cursor.execute(
                  sql,
                  (
                      new_customer,
                      new_item_cust,
                      new_item_erp,
                      new_isi_pak,
                  ),
              )
              conn.commit()
            conn.close()
            st.success(
                f"Berhasil! Item '{new_item_cust}' untuk customer"
                f" '{new_customer}' telah disimpan ke database."
            )
            st.rerun()
          except Exception as e:
            st.error(f"Gagal menyimpan ke database: {e}")
        else:
          st.warning(
              "Mohon lengkapi 'Item Customer' dan 'Item ERP' terlebih dahulu!"
          )

  except Exception as e:
    st.error(f"Gagal terhubung ke database MySQL: {e}")

# ==========================================
# MENU 1: KONVERTER PO (FITUR UTAMA)
# ==========================================
else:
  # ==========================================
  # JUDUL DAN HEADER (DENGAN LOGO)
  # ==========================================
  col_judul, col_logo = st.columns([4, 1])

  with col_judul:
    st.markdown(
        "<h1 style='font-size: 42px; font-weight: 800; margin-bottom: 0px;"
        " padding-bottom: 0px; color: #1E3A8A;'>PT. Sunda Kelapa Pustaka</h1>",
        unsafe_allow_html=True,
    )
    st.markdown(
        "<h3 style='font-size: 22px; font-weight: 400; margin-top: 0px;"
        " padding-top: 0px; color: #555555;'>Sistem Konversi PO ke ERP</h3>",
        unsafe_allow_html=True,
    )

  with col_logo:
    if os.path.exists("logo.png"):
      st.image("logo.png", use_container_width=True)
    elif os.path.exists("logo.jpg"):
      st.image("logo.jpg", use_container_width=True)

  st.markdown("---")

  # ==========================================
  # INISIALISASI MEMORI
  # ==========================================
  if "customer" not in st.session_state:
    st.session_state.customer = "Gramedia"
  if "daftar_manual" not in st.session_state:
    st.session_state.daftar_manual = []


  def ubah_customer(nama_customer):
    st.session_state.customer = nama_customer
    st.session_state.daftar_manual = []


  # ==========================================
  # TOMBOL PILIHAN CUSTOMER (6 KOLOM)
  # ==========================================
  st.write("**Pilih Customer:**")
  col1, col2, col3, col4, col5, col6 = st.columns(6)

  with col1:
    st.button(
        "Gramedia",
        use_container_width=True,
        type=(
            "primary"
            if st.session_state.customer == "Gramedia"
            else "secondary"
        ),
        on_click=ubah_customer,
        args=("Gramedia",),
    )
  with col2:
    st.button(
        "Promexx",
        use_container_width=True,
        type=(
            "primary" if st.session_state.customer == "Promexx" else "secondary"
        ),
        on_click=ubah_customer,
        args=("Promexx",),
    )
  with col3:
    st.button(
        "AEON",
        use_container_width=True,
        type=(
            "primary" if st.session_state.customer == "AEON" else "secondary"
        ),
        on_click=ubah_customer,
        args=("AEON",),
    )
  with col4:
    st.button(
        "Hypermart",
        use_container_width=True,
        type=(
            "primary"
            if st.session_state.customer == "Hypermart"
            else "secondary"
        ),
        on_click=ubah_customer,
        args=("Hypermart",),
    )
  with col5:
    st.button(
        "GrandLucky",
        use_container_width=True,
        type=(
            "primary"
            if st.session_state.customer == "GrandLucky"
            else "secondary"
        ),
        on_click=ubah_customer,
        args=("GrandLucky",),
    )
  with col6:
    st.button(
        "Jameson",
        use_container_width=True,
        type=(
            "primary" if st.session_state.customer == "Jameson" else "secondary"
        ),
        on_click=ubah_customer,
        args=("Jameson",),
    )

  customer = st.session_state.customer

  st.markdown("---")
  st.markdown(
      f"<div style='font-size: 18px; font-weight: 600; color: #4A4A4A;"
      " margin-bottom: 15px;'>🟢 Mode Aktif: <span style='color:"
      f" #1E3A8A;'>{customer}</span></div>",
      unsafe_allow_html=True,
  )

  # ==========================================
  # LOAD MASTER DATA DARI DATABASE MYSQL
  # ==========================================
  try:
    conn = get_db_connection()
    query = f"SELECT item_customer, item_erp, isi_per_pak FROM master_customer_items WHERE customer = '{customer}'"
    master_df = pd.read_sql(query, conn)
    conn.close()

    if not master_df.empty:
      master_df["Item Customer"] = (
          master_df["item_customer"].astype(str).str.strip()
      )
      master_df["Item ERP"] = master_df["item_erp"].astype(str).str.strip()
      master_df["Isi per Pak"] = pd.to_numeric(
          master_df["isi_per_pak"], errors="coerce"
      ).fillna(1)
      master_df = master_df.drop_duplicates(
          subset=["Item Customer"], keep="first"
      )
  except Exception as e:
    st.error(f"Gagal memuat master data dari database MySQL: {e}")
    master_df = pd.DataFrame()

  if master_df.empty and customer != "Hypermart":
    st.warning(
        f"⚠️ Belum ada data master untuk customer '{customer}' di database."
    )

  # ==========================================
  # 1. LOGIKA GRAMEDIA (XML)
  # ==========================================
  if customer == "Gramedia":
    if master_df.empty:
      st.stop()

    st.write(
        "Silakan upload file XML Purchase Order dari **Gramedia** di bawah ini:"
    )
    uploaded_xml = st.file_uploader("Upload File XML", type=["xml"])

    if uploaded_xml is not None:
      try:
        tree = ET.parse(uploaded_xml)
        root = tree.getroot()
        po_number = root.attrib.get("Textbox23", "Hasil_Konversi_PO")

        data_po = []
        for child in root.iter():
          if child.tag.endswith("Details"):
            qty = int(float(child.attrib.get("PurchQty", 0)))
            price = float(child.attrib.get("PurchPrice", 0))
            discount = child.attrib.get("LinePercent", 0)
            item_id = ""
            for subchild in child:
              if subchild.tag.endswith("Tablix10"):
                item_id = subchild.attrib.get("ItemId", "")

            if item_id:
              data_po.append(
                  {
                      "Item Customer": str(item_id).strip(),
                      "PurchQty": qty,
                      "Discount %": discount,
                      "Price List Rate": price,
                  }
              )

        po_df = pd.DataFrame(data_po)
        merged_df = pd.merge(
            po_df, master_df, on="Item Customer", how="left"
        )

        unmapped_items = merged_df[merged_df["Item ERP"].isna()][
            "Item Customer"
        ].unique()
        if len(unmapped_items) > 0:
          st.error(
              "🛑 **PROSES DIHENTIKAN! Terdapat Item yang Belum Terdaftar.**"
          )
          for item in unmapped_items:
            st.markdown(f"- Kode Customer: **{item}**")
          st.stop()

        merged_df["Quantity"] = (
            (merged_df["PurchQty"] // merged_df["Isi per Pak"])
            .fillna(0)
            .astype(int)
        )
        merged_df["Price List Rate"] = (
            merged_df["Price List Rate"] * merged_df["Isi per Pak"]
        )

        total_po_value = (
            merged_df["Quantity"] * merged_df["Price List Rate"]
        ).sum()
        total_po_format = (
            f"<span style='color:red; font-weight:bold;'>Rp"
            f" {total_po_value:,.2f}</span>"
        )

        output = io.BytesIO()
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Template-Quotation"

        ws.cell(row=1, column=1, value="Bulk Edit Items")
        headers_r2 = [
            "Item Code",
            "Quantity",
            "Discount %",
            "Warehouse",
            "Price List Rate",
            "UOM",
        ]
        for col, val in enumerate(headers_r2, start=1):
          ws.cell(row=2, column=col, value=val)
        headers_r3 = [
            "item_code",
            "qty",
            "discount_percentage",
            "warehouse",
            "price_list_rate",
            "uom",
        ]
        for col, val in enumerate(headers_r3, start=1):
          ws.cell(row=3, column=col, value=val)
        ws.cell(row=5, column=1, value="The format is case sensitive")
        ws.cell(
            row=6,
            column=1,
            value="Do not edit headers which are preset in the template",
        )
        ws.cell(row=7, column=1, value="------")

        start_row = 8
        for index, row in merged_df.iterrows():
          ws.cell(row=start_row + index, column=1, value=row["Item ERP"])
          ws.cell(row=start_row + index, column=2, value=row["Quantity"])
          ws.cell(row=start_row + index, column=3, value=row["Discount %"])
          ws.cell(row=start_row + index, column=5, value=row["Price List Rate"])

        wb.save(output)

        st.success("File PO Gramedia berhasil diproses!")
        st.markdown(
            f"**Nomor PO:** {po_number} | **Total PO (Sub Total):**"
            f" {total_po_format} *(Belum Termasuk PPN)*",
            unsafe_allow_html=True,
        )

        preview_df = pd.DataFrame(
            {
                "Item Code": merged_df["Item ERP"],
                "Quantity": merged_df["Quantity"],
                "Discount %": merged_df["Discount %"],
                "Price List Rate": merged_df["Price List Rate"],
            }
        )
        preview_df.index = preview_df.index + 1
        st.dataframe(preview_df, use_container_width=True)

        st.download_button(
            label=f"⬇️ Download File {po_number}.xlsx",
            data=output.getvalue(),
            file_name=f"{po_number}.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )
      except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses XML Gramedia: {e}")

  # ==========================================
  # 2. LOGIKA PROMEXX (PDF)
  # ==========================================
  elif customer == "Promexx":
    if master_df.empty:
      st.stop()

    st.write(
        "Silakan upload file PDF Purchase Order dari **Promexx** di bawah ini:"
    )
    uploaded_pdf = st.file_uploader("Upload File PDF", type=["pdf"])

    if uploaded_pdf is not None:
      try:
        data_po = []
        global_discount = 0
        po_number = "PO_Promexx"

        with pdfplumber.open(uploaded_pdf) as pdf:
          for page in pdf.pages:
            text = page.extract_text()
            if text:
              match_po = re.search(
                  r"No\.\s*PO:\s*([A-Za-z0-9_-]+)", text, re.IGNORECASE
              )
              if match_po and po_number == "PO_Promexx":
                po_number = match_po.group(1).strip()

              match_discount = re.search(
                  r"DF.*?([0-9\.]+)\s*%", text, re.IGNORECASE
              )
              if match_discount:
                global_discount = float(match_discount.group(1))

              lines = text.split("\n")
              for line in lines:
                line_clean = line.replace("|", " ").strip()
                parts = line_clean.split()
                if len(parts) >= 8:
                  try:
                    net_price_str = parts[-1].replace(",", "").replace(".", "")
                    qty_str = parts[-5].replace(",", "").replace(".", "")
                    barcode = (
                        parts[1]
                        if (
                            parts[0].replace(".", "").isdigit()
                            and len(parts[0]) <= 4
                        )
                        else parts[0]
                    )
                    if (
                        qty_str.isdigit()
                        and net_price_str.isdigit()
                        and barcode.isdigit()
                        and len(barcode) >= 4
                    ):
                      data_po.append(
                          {
                              "Item Customer": str(barcode).strip(),
                              "PurchQty": int(float(qty_str)),
                              "Discount %": global_discount,
                              "Price List Rate": float(net_price_str),
                          }
                      )
                  except Exception:
                    pass

        if not data_po:
          st.error("Gagal mengekstrak data pesanan dari PDF.")
          st.stop()

        po_df = pd.DataFrame(data_po)
        merged_df = pd.merge(
            po_df, master_df, on="Item Customer", how="left"
        )

        unmapped_items = merged_df[merged_df["Item ERP"].isna()][
            "Item Customer"
        ].unique()
        if len(unmapped_items) > 0:
          st.error(
              "🛑 **PROSES DIHENTIKAN! Terdapat Item yang Belum Terdaftar.**"
          )
          for item in unmapped_items:
            st.markdown(f"- Kode Customer: **{item}**")
          st.stop()

        merged_df["Quantity"] = (
            (merged_df["PurchQty"] // merged_df["Isi per Pak"])
            .fillna(0)
            .astype(int)
        )
        merged_df["Price List Rate"] = (
            merged_df["Price List Rate"] * merged_df["Isi per Pak"]
        )

        total_po_value = (
            merged_df["Quantity"]
            * merged_df["Price List Rate"]
            * (1 - (merged_df["Discount %"] / 100))
        ).sum()
        total_po_format = (
            f"<span style='color:red; font-weight:bold;'>Rp"
            f" {total_po_value:,.2f}</span>"
        )

        output = io.BytesIO()
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Template-Quotation"

        ws.cell(row=1, column=1, value="Bulk Edit Items")
        headers_r2 = [
            "Item Code",
            "Quantity",
            "Discount %",
            "Warehouse",
            "Price List Rate",
            "UOM",
        ]
        for col, val in enumerate(headers_r2, start=1):
          ws.cell(row=2, column=col, value=val)
        headers_r3 = [
            "item_code",
            "qty",
            "discount_percentage",
            "warehouse",
            "price_list_rate",
            "uom",
        ]
        for col, val in enumerate(headers_r3, start=1):
          ws.cell(row=3, column=col, value=val)
        ws.cell(row=5, column=1, value="The format is case sensitive")
        ws.cell(
            row=6,
            column=1,
            value="Do not edit headers which are preset in the template",
        )
        ws.cell(row=7, column=1, value="------")

        start_row = 8
        for index, row in merged_df.iterrows():
          ws.cell(row=start_row + index, column=1, value=row["Item ERP"])
          ws.cell(row=start_row + index, column=2, value=row["Quantity"])
          ws.cell(row=start_row + index, column=3, value=row["Discount %"])
          ws.cell(row=start_row + index, column=5, value=row["Price List Rate"])

        wb.save(output)

        st.success("File PO Promexx berhasil diproses!")
        st.markdown(
            f"**Nomor PO:** {po_number} | **Total PO (Sub Total):**"
            f" {total_po_format} *(Belum Termasuk PPN)*",
            unsafe_allow_html=True,
        )

        preview_df = pd.DataFrame(
            {
                "Item Code": merged_df["Item ERP"],
                "Quantity": merged_df["Quantity"],
                "Discount %": merged_df["Discount %"],
                "Price List Rate": merged_df["Price List Rate"],
            }
        )
        preview_df.index = preview_df.index + 1
        st.dataframe(preview_df, use_container_width=True)

        st.download_button(
            label=f"⬇️ Download File Promex_{po_number}.xlsx",
            data=output.getvalue(),
            file_name=f"Promex_{po_number}.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )
      except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses PDF Promexx: {e}")

  # ==========================================
  # 3. LOGIKA AEON (PDF)
  # ==========================================
  elif customer == "AEON":
    if master_df.empty:
      st.stop()

    st.write(
        "Silakan upload file PDF Purchase Order dari **AEON** di bawah ini:"
    )
    uploaded_pdf = st.file_uploader("Upload File PDF", type=["pdf"])

    if uploaded_pdf is not None:
      try:
        data_po = []
        po_number = "PO_AEON"

        with pdfplumber.open(uploaded_pdf) as pdf:
          full_text = ""
          for page in pdf.pages:
            extracted = page.extract_text()
            if extracted:
              full_text += extracted + "\n"

          match_po = re.search(r"(\d{14})", full_text[:500])
          if match_po:
            po_number = match_po.group(1).strip()

          words = full_text.split()
          for i, word in enumerate(words):
            if word.isdigit() and 6 <= len(word) <= 10:
              if (
                  i + 1 < len(words)
                  and words[i + 1].isdigit()
                  and 12 <= len(words[i + 1]) <= 14
              ):
                internal_code = word
                nums = []
                for w in words[i + 2 : min(i + 8, len(words))]:
                  w_clean = w.replace(",", "").replace(".", "")
                  if w_clean.isdigit():
                    nums.append(w)

                if len(nums) >= 3:
                  try:
                    qty_val = int(float(nums[0].replace(",", "")))
                    discount_val = float(nums[1].replace(",", ""))
                    p_clean = nums[2].replace(",", "")
                    if p_clean.endswith(".00"):
                      p_clean = p_clean[:-3]
                    price_val = float(p_clean.replace(".", ""))

                    if qty_val > 0:
                      data_po.append(
                          {
                              "Item Customer": str(internal_code).strip(),
                              "PurchQty": qty_val,
                              "Discount %": discount_val,
                              "Price List Rate": price_val,
                          }
                      )
                  except Exception:
                    pass

        if not data_po:
          st.error(
              "Gagal mengekstrak data pesanan. Pastikan PDF yang di-upload"
              " adalah milik AEON."
          )
          st.stop()

        po_df = pd.DataFrame(data_po)
        merged_df = pd.merge(
            po_df, master_df, on="Item Customer", how="left"
        )

        unmapped_items = merged_df[merged_df["Item ERP"].isna()][
            "Item Customer"
        ].unique()
        if len(unmapped_items) > 0:
          st.error(
              "🛑 **PROSES DIHENTIKAN! Terdapat Item yang Belum Terdaftar.**"
          )
          for item in unmapped_items:
            st.markdown(f"- Kode Internal: **{item}**")
          st.stop()

        merged_df["Quantity"] = (
            (merged_df["PurchQty"] // merged_df["Isi per Pak"])
            .fillna(0)
            .astype(int)
        )
        merged_df["Price List Rate"] = (
            merged_df["Price List Rate"] * merged_df["Isi per Pak"]
        )

        total_po_value = (
            merged_df["Quantity"]
            * merged_df["Price List Rate"]
            * (1 - (merged_df["Discount %"] / 100))
        ).sum()
        total_po_format = (
            f"<span style='color:red; font-weight:bold;'>Rp"
            f" {total_po_value:,.2f}</span>"
        )

        output = io.BytesIO()
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Template-Quotation"

        ws.cell(row=1, column=1, value="Bulk Edit Items")
        headers_r2 = [
            "Item Code",
            "Quantity",
            "Discount %",
            "Warehouse",
            "Price List Rate",
            "UOM",
        ]
        for col, val in enumerate(headers_r2, start=1):
          ws.cell(row=2, column=col, value=val)
        headers_r3 = [
            "item_code",
            "qty",
            "discount_percentage",
            "warehouse",
            "price_list_rate",
            "uom",
        ]
        for col, val in enumerate(headers_r3, start=1):
          ws.cell(row=3, column=col, value=val)
        ws.cell(row=5, column=1, value="The format is case sensitive")
        ws.cell(
            row=6,
            column=1,
            value="Do not edit headers which are preset in the template",
        )
        ws.cell(row=7, column=1, value="------")

        start_row = 8
        for index, row in merged_df.iterrows():
          ws.cell(row=start_row + index, column=1, value=row["Item ERP"])
          ws.cell(row=start_row + index, column=2, value=row["Quantity"])
          ws.cell(row=start_row + index, column=3, value=row["Discount %"])
          ws.cell(row=start_row + index, column=5, value=row["Price List Rate"])

        wb.save(output)

        st.success("File PO AEON berhasil diproses!")
        st.markdown(
            f"**Nomor PO:** {po_number} | **Total PO (Sub Total):**"
            f" {total_po_format} *(Belum Termasuk PPN)*",
            unsafe_allow_html=True,
        )

        preview_df = pd.DataFrame(
            {
                "Item Code": merged_df["Item ERP"],
                "Quantity": merged_df["Quantity"],
                "Discount %": merged_df["Discount %"],
                "Price List Rate": merged_df["Price List Rate"],
            }
        )
        preview_df.index = preview_df.index + 1
        st.dataframe(preview_df, use_container_width=True)

        st.download_button(
            label=f"⬇️ Download File AEON_{po_number}.xlsx",
            data=output.getvalue(),
            file_name=f"AEON_{po_number}.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )
      except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses PDF AEON: {e}")

  # ==========================================
  # 4. LOGIKA HYPERMART
  # ==========================================
  elif customer == "Hypermart":
    st.info(
        f"Modul konversi untuk **{customer}** sedang disiapkan. Silakan tunggu"
        " pembaruan selanjutnya."
    )
    st.file_uploader(
        f"Upload File PO {customer}",
        type=["xls", "xlsx", "xml", "pdf"],
        disabled=True,
    )

  # ==========================================
  # 5. LOGIKA GRANDLUCKY (PDF)
  # ==========================================
  elif customer == "GrandLucky":
    if master_df.empty:
      st.stop()

    st.write(
        "Silakan upload file PDF Purchase Order dari **GrandLucky** di bawah"
        " ini:"
    )
    uploaded_pdf = st.file_uploader("Upload File PDF GrandLucky", type=["pdf"])

    if uploaded_pdf is not None:
      try:
        data_po = []
        po_number = "PO_GrandLucky"

        with pdfplumber.open(uploaded_pdf) as pdf:
          for page in pdf.pages:
            text = page.extract_text()
            if text:
              match_po = re.search(
                  r"PO\s*#?\s*:?\s*([0-9]{4}\.[A-Za-z0-9\.\-_]+)",
                  text,
                  re.IGNORECASE,
              )
              if not match_po:
                match_po = re.search(
                    r"PO\s*#\s*:\s*([A-Za-z0-9\.\-_]+)", text, re.IGNORECASE
                )
              if match_po and po_number == "PO_GrandLucky":
                po_number = match_po.group(1).strip()

              lines = text.split("\n")
              for line in lines:
                parts = line.strip().split()
                if len(parts) >= 8:
                  no_str = parts[0].replace(".", "")
                  sku_str = parts[1]
                  qty_str = parts[-3].replace(",", "").replace(".", "")
                  price_str = parts[-2].replace(",", "").replace(".", "")
                  total_str = parts[-1].replace(",", "").replace(".", "")

                  if (
                      no_str.isdigit()
                      and sku_str.isdigit()
                      and len(sku_str) >= 6
                      and qty_str.isdigit()
                      and price_str.isdigit()
                      and total_str.isdigit()
                  ):
                    qty_val = int(qty_str)
                    price_val = float(price_str)

                    if qty_val > 0:
                      data_po.append(
                          {
                              "Item Customer": str(sku_str).strip(),
                              "PurchQty": qty_val,
                              "Discount %": 0.0,
                              "Price List Rate": price_val,
                          }
                      )

        if not data_po:
          st.error(
              "Gagal mengekstrak data pesanan. Pastikan PDF yang di-upload"
              " adalah format asli dari GrandLucky."
          )
          st.stop()

        po_df = pd.DataFrame(data_po)
        merged_df = pd.merge(
            po_df, master_df, on="Item Customer", how="left"
        )

        unmapped_items = merged_df[merged_df["Item ERP"].isna()][
            "Item Customer"
        ].unique()
        if len(unmapped_items) > 0:
          st.error(
              "🛑 **PROSES DIHENTIKAN! Terdapat Item yang Belum Terdaftar.**"
          )
          for item in unmapped_items:
            st.markdown(f"- SKU# Customer: **{item}**")
          st.stop()

        def hitung_konversi(row):
          qty_po = row["PurchQty"]
          harga_po = row["Price List Rate"]
          rasio = row["Isi per Pak"]
          item_erp = str(row["Item ERP"])

          if "PC" in item_erp and rasio > 0:
            q_erp = int(qty_po * rasio)
            h_erp = harga_po / rasio
          elif "PC" not in item_erp and rasio > 1:
            q_erp = int(qty_po // rasio)
            h_erp = harga_po * rasio
          else:
            q_erp = int(qty_po)
            h_erp = harga_po
          return pd.Series([q_erp, h_erp])

        merged_df[["Quantity", "Price List Rate"]] = merged_df.apply(
            hitung_konversi, axis=1
        )

        total_po_value = (
            merged_df["Quantity"] * merged_df["Price List Rate"]
        ).sum()
        total_po_format = (
            f"<span style='color:red; font-weight:bold;'>Rp"
            f" {total_po_value:,.2f}</span>"
        )

        output = io.BytesIO()
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Template-Quotation"

        ws.cell(row=1, column=1, value="Bulk Edit Items")
        headers_r2 = [
            "Item Code",
            "Quantity",
            "Discount %",
            "Warehouse",
            "Price List Rate",
            "UOM",
        ]
        for col, val in enumerate(headers_r2, start=1):
          ws.cell(row=2, column=col, value=val)
        headers_r3 = [
            "item_code",
            "qty",
            "discount_percentage",
            "warehouse",
            "price_list_rate",
            "uom",
        ]
        for col, val in enumerate(headers_r3, start=1):
          ws.cell(row=3, column=col, value=val)
        ws.cell(row=5, column=1, value="The format is case sensitive")
        ws.cell(
            row=6,
            column=1,
            value="Do not edit headers which are preset in the template",
        )
        ws.cell(row=7, column=1, value="------")

        start_row = 8
        for index, row in merged_df.iterrows():
          ws.cell(row=start_row + index, column=1, value=row["Item ERP"])
          ws.cell(row=start_row + index, column=2, value=row["Quantity"])
          ws.cell(row=start_row + index, column=3, value=row["Discount %"])
          ws.cell(row=start_row + index, column=5, value=row["Price List Rate"])

        wb.save(output)

        st.success("File PO GrandLucky berhasil diproses!")
        st.markdown(
            f"**Nomor PO:** {po_number} | **Total PO (Sub Total):**"
            f" {total_po_format} *(Belum Termasuk PPN)*",
            unsafe_allow_html=True,
        )

        preview_df = pd.DataFrame(
            {
                "Item Code": merged_df["Item ERP"],
                "Quantity": merged_df["Quantity"],
                "Discount %": merged_df["Discount %"],
                "Price List Rate": merged_df["Price List Rate"],
            }
        )
        preview_df.index = preview_df.index + 1
        st.dataframe(preview_df, use_container_width=True)

        st.download_button(
            label=f"⬇️ Download File GrandLucky-{po_number}.xlsx",
            data=output.getvalue(),
            file_name=f"GrandLucky-{po_number}.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
            use_container_width=True,
        )
      except Exception as e:
        st.error(f"Terjadi kesalahan saat memproses PDF GrandLucky: {e}")

  # ==========================================
  # 6. LOGIKA JAMESON (INPUT MANUAL)
  # ==========================================
  elif customer == "Jameson":
    if master_df.empty:
      st.stop()

    st.write(
        "Silakan input data pesanan untuk **Jameson** secara manual di bawah"
        " ini:"
    )

    with st.form("manual_form", clear_on_submit=True):
      col_f1, col_f2, col_f3, col_f4 = st.columns(4)
      m_kode = col_f1.text_input("Kode Item Customer")
      m_qty = col_f2.number_input("Qty (Pcs)", min_value=1)
      m_hrg = col_f3.number_input("Harga Netto", min_value=0)
      m_dsc = col_f4.number_input("Diskon (%)", min_value=0.0)

      if st.form_submit_button("Tambah ke Daftar"):
        if m_kode:
          st.session_state.daftar_manual.append(
              {
                  "Item Customer": str(m_kode).strip(),
                  "PurchQty": m_qty,
                  "Discount %": m_dsc,
                  "Price List Rate": float(m_hrg),
              }
          )
          st.rerun()
        else:
          st.warning("Kode Item tidak boleh kosong!")

    components.html(
        """
        <script>
        const inputField = window.parent.document.querySelector('input[aria-label="Kode Item Customer"]');
        if (inputField) {
            setTimeout(function() {
                inputField.focus();
            }, 100);
        }
        </script>
        """,
        height=0,
        width=0,
    )

    if st.session_state.daftar_manual:
      st.write("### Daftar Pesanan (Keranjang)")
      df_manual = pd.DataFrame(st.session_state.daftar_manual)

      preview_keranjang = df_manual.copy()
      preview_keranjang.index = preview_keranjang.index + 1
      st.dataframe(preview_keranjang, use_container_width=True)

      col_btn1, col_btn2 = st.columns([1, 4])
      if col_btn1.button("Hapus Semua"):
        st.session_state.daftar_manual = []
        st.rerun()

      if col_btn2.button(
          "🚀 Generate PO Jameson", type="primary", use_container_width=True
      ):
        merged_df = pd.merge(
            df_manual, master_df, on="Item Customer", how="left"
        )

        unmapped_items = merged_df[merged_df["Item ERP"].isna()][
            "Item Customer"
        ].unique()
        if len(unmapped_items) > 0:
          st.error(
              "🛑 **PROSES DIHENTIKAN! Terdapat Item yang Belum Terdaftar.**"
          )
          for item in unmapped_items:
            st.markdown(f"- Kode Customer: **{item}**")
        else:
          merged_df["Quantity"] = (
              (merged_df["PurchQty"] // merged_df["Isi per Pak"])
              .fillna(0)
              .astype(int)
          )
          merged_df["Price List Rate"] = (
              merged_df["Price List Rate"] * merged_df["Isi per Pak"]
          )

          total_po_value = (
              merged_df["Quantity"]
              * merged_df["Price List Rate"]
              * (1 - (merged_df["Discount %"] / 100))
          ).sum()
          format_rupiah = (
              f"Rp {total_po_value:,.2f}".replace(",", "X")
              .replace(".", ",")
              .replace("X", ".")
          )

          output = io.BytesIO()
          wb = openpyxl.Workbook()
          ws = wb.active
          ws.title = "Template-Quotation"

          ws.cell(row=1, column=1, value="Bulk Edit Items")
          headers_r2 = [
              "Item Code",
              "Quantity",
              "Discount %",
              "Warehouse",
              "Price List Rate",
              "UOM",
          ]
          for col, val in enumerate(headers_r2, start=1):
            ws.cell(row=2, column=col, value=val)
          headers_r3 = [
              "item_code",
              "qty",
              "discount_percentage",
              "warehouse",
              "price_list_rate",
              "uom",
          ]
          for col, val in enumerate(headers_r3, start=1):
            ws.cell(row=3, column=col, value=val)
          ws.cell(row=5, column=1, value="The format is case sensitive")
          ws.cell(
              row=6,
              column=1,
              value="Do not edit headers which are preset in the template",
          )
          ws.cell(row=7, column=1, value="------")

          start_row = 8
          for index, row in merged_df.iterrows():
            ws.cell(row=start_row + index, column=1, value=row["Item ERP"])
            ws.cell(row=start_row + index, column=2, value=row["Quantity"])
            ws.cell(row=start_row + index, column=3, value=row["Discount %"])
            ws.cell(row=start_row + index, column=5, value=row["Price List Rate"])

          wb.save(output)

          st.success("File PO Jameson berhasil di-generate!")
          st.markdown(
              f"**Total PO (Setelah Diskon):** <span style='color:red;"
              f" font-weight:bold;'>{format_rupiah}</span> *(Belum Termasuk"
              " PPN)*",
              unsafe_allow_html=True,
          )

          st.download_button(
              label="⬇️ Download File Jameson_PO.xlsx",
              data=output.getvalue(),
              file_name="Jameson_PO.xlsx",
              mime=(
                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              ),
              use_container_width=True,
          )