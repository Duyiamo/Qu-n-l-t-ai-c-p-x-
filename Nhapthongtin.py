import sqlite3
from datetime import datetime
import io
import os
import folium
from folium.plugins import Draw, LocateControl
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
import pandas as pd
import streamlit as st
from streamlit_folium import st_folium

DB_FILE = "quan_ly_dat_dai.db"


def init_db():
  conn = sqlite3.connect(DB_FILE)
  cursor = conn.cursor()
  cursor.execute("""
        CREATE TABLE IF NOT EXISTS thia_dat (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            ho_ten TEXT,
            sdt TEXT,
            dia_chi_thuong_tru TEXT,
            so_to TEXT,
            so_thua TEXT,
            dia_chi_thua_dat TEXT,
            dien_tich_khai_bao REAL,
            nguon_goc TEXT,
            hien_trang TEXT,
            hien_trang_chi_tiet TEXT,
            tinh_trang_so TEXT,
            ghi_chu TEXT,
            lat REAL,
            lon REAL,
            ngay_tao TEXT
        )
    """)
  conn.commit()
  conn.close()


init_db()

st.set_page_config(
    page_title="Quản lý Hiện trạng Đất đai Cấp Xã", layout="wide"
)

st.title("🌾 Hệ thống Thu thập & Quản lý Hiện trạng Đất đai Cấp Xã")
st.markdown(
    "Ứng dụng hỗ trợ ghi nhận vị trí canh tác của hộ dân và kết xuất báo cáo"
    " chuyên nghiệp."
)

tab1, tab2 = st.tabs(
    ["📝 1. Khai báo / Cập nhật thửa đất", "🔒 2. Khu vực Quản trị (Admin)"]
)

with tab1:
  st.header("Nhập thông tin và xác định vị trí thửa đất")

  with st.form("form_khai_bao"):
    col1, col2 = st.columns(2)
    with col1:
      ho_ten = st.text_input("Họ và tên chủ sử dụng *")
      sdt = st.text_input("Số điện thoại")
      dia_chi_thuong_tru = st.text_input(
          "Địa chỉ thường trú (Thôn/Xóm, Xã...)"
      )
      so_to = st.text_input("Số tờ bản đồ")
      so_thua = st.text_input("Số thửa đất")
      dia_chi_thua_dat = st.text_input(
          "Địa chỉ / Khu vực tọa lạc thửa đất (Ví dụ: Thôn 2, Khu Đồng Lớn...)"
      )

    with col2:
      dien_tich = st.number_input(
          "Diện tích tự khai báo (m²) *", min_value=0.0, value=0.0, step=10.0
      )
      nguon_goc = st.text_input(
          "Nguồn gốc sử dụng đất tự kê khai (Ví dụ: Khai hoang, Nhận chuyển"
          " nhượng...)"
      )
      hien_trang = st.selectbox(
          "Nhóm hiện trạng sử dụng đất *", [
              "Đất trồng lúa",
              "Đất trồng cây hàng năm khác",
              "Đất trồng cây lâu năm",
              "Đất nuôi trồng thủy sản",
              "Đất lâm nghiệp",
              "Đất phi nông nghiệp / Khác",
          ],
      )
      hien_trang_chi_tiet = st.text_input(
          "Cụ thể tên cây trồng / mục đích (Ví dụ: Cây mì, Cây điều, Cây keo,"
          " ...)"
      )
      tinh_trang_so = st.radio(
          "Tình trạng Giấy chứng nhận QSDĐ:",
          ["Đã có Giấy chứng nhận", "Chưa có / Đang sử dụng ổn định"],
          horizontal=True,
      )
      ghi_chu = st.text_area("Ghi chú thêm (nếu có)")

    st.markdown("---")
    st.markdown(
        "**Xác định vị trí trên bản đồ:** Bạn có thể bấm vào nút định vị GPS"
        " góc trên bản đồ để lấy vị trí hiện tại, hoặc tự di chuyển bản đồ đến"
        " thửa đất. (Có thể chuyển sang chế độ **Vệ tinh** ở góc trái bản đồ"
        " để dễ quan sát)."
    )

    m = folium.Map(location=[14.3305, 108.6472], zoom_start=15)

    folium.TileLayer(
        tiles="https://{s}.google.com/vt/lyrs=s&x={x}&y={y}&z={z}",
        attr="Google Satellite",
        name="Bản đồ Vệ tinh",
        subdomains=["mt0", "mt1", "mt2", "mt3"],
        overlay=True,
        control=True,
    ).add_to(m)

    folium.LayerControl().add_to(m)

    LocateControl(
        auto_start=False,
        position="topleft",
        strings={
            "title": "Tìm vị trí hiện tại của tôi",
            "popup": "Bạn đang ở đây",
        },
        locate_options={"maxZoom": 18},
    ).add_to(m)

    draw = Draw(
        export=False,
        draw_options={
            "polyline": False,
            "polygon": False,
            "rectangle": False,
            "circle": False,
            "marker": True,
            "circlemarker": False,
        },
    )
    draw.add_to(m)

    output = st_folium(m, width="100%", height=450, key="map_input")

    submit_button = st.form_submit_button(
        "Gửi thông tin thửa đất", type="primary"
    )

    if submit_button:
      if not ho_ten:
        st.error("Vui lòng nhập họ và tên chủ sử dụng!")
      else:
        lat, lon = None, None
        if output:
          if output.get("last_clicked"):
            lat = output["last_clicked"]["lat"]
            lon = output["last_clicked"]["lng"]
          elif output.get("all_drawings") and len(output["all_drawings"]) > 0:
            geometry = output["all_drawings"][-1]["geometry"]
            if geometry["type"] == "Point":
              lon, lat = geometry["coordinates"][0], geometry["coordinates"][1]

        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()

        is_duplicate = False
        if lat is not None and lon is not None:
          cursor.execute(
              """
                    SELECT COUNT(*) FROM thia_dat 
                    WHERE ABS(lat - ?) < 0.00001 AND ABS(lon - ?) < 0.00001
                """,
              (lat, lon),
          )
          count = cursor.fetchone()[0]
          if count > 0:
            is_duplicate = True

        if is_duplicate:
          st.error(
              "⚠️ Vị trí thửa đất này đã được kê khai vào hệ thống! Xin vui"
              " lòng nhập thửa đất khác và cập nhật vị trí khác trên bản đồ."
          )
        else:
          # Chỉ lưu ngày tháng năm (YYYY-MM-DD)
          ngay_hien_tai = datetime.now().strftime("%Y-%m-%d")
          cursor.execute(
              """
                    INSERT INTO thia_dat (ho_ten, sdt, dia_chi_thuong_tru, so_to, so_thua, dia_chi_thua_dat, dien_tich_khai_bao, nguon_goc, hien_trang, hien_trang_chi_tiet, tinh_trang_so, ghi_chu, lat, lon, ngay_tao)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
              (
                  ho_ten,
                  sdt,
                  dia_chi_thuong_tru,
                  so_to,
                  so_thua,
                  dia_chi_thua_dat,
                  dien_tich,
                  nguon_goc,
                  hien_trang,
                  hien_trang_chi_tiet,
                  tinh_trang_so,
                  ghi_chu,
                  lat,
                  lon,
                  ngay_hien_tai,
              ),
          )
          conn.commit()

          if lat and lon:
            st.success(
                f"Cảm ơn ông/bà **{ho_ten}**! Dữ liệu đã gửi về hệ thống quản"
                f" lý thành công (Tọa độ: {lat:.5f}, {lon:.5f})."
            )
          else:
            st.warning(
                f"Cảm ơn ông/bà **{ho_ten}**! Đã lưu thông tin (Chưa thấy bạn"
                " chấm điểm vị trí trên bản đồ)."
            )

        conn.close()

with tab2:
  st.header("Khu vực Quản trị dành cho Cán bộ địa chính")

  password = st.text_input(
      "Nhập mật khẩu quản lý để tiếp tục:", type="password"
  )
  ADMIN_PASSWORD = "admin123"

  if password == ADMIN_PASSWORD:
    st.success("Xác thực thành công! Chào mừng cán bộ quản lý.")

    conn = sqlite3.connect(DB_FILE)
    df = pd.read_sql_query("SELECT * FROM thia_dat", conn)
    conn.close()

    if df.empty:
      st.info("Chưa có dữ liệu khai báo nào từ người dân.")
    else:
      st.metric(
          label="Tổng số thửa đã cập nhật vào hệ thống",
          value=f"{len(df)} thửa đất",
      )

      # --- BỔ SUNG BỘ LỌC THEO NGÀY ---
      st.markdown("### Bộ lọc dữ liệu quản lý")
      danh_sach_ngay = ["Tất cả các ngày"] + sorted(
          df["ngay_tao"].dropna().unique().tolist()
      )
      chon_ngay = st.selectbox("Lọc danh sách theo ngày kê khai:", danh_sach_ngay)

      if chon_ngay != "Tất cả các ngày":
        df_hien_thi = df[df["ngay_tao"] == chon_ngay]
      else:
        df_hien_thi = df

      st.markdown(f"Đang hiển thị **{len(df_hien_thi)}** bản ghi.")

      st.dataframe(
          df_hien_thi[[
              "id",
              "ho_ten",
              "sdt",
              "dia_chi_thuong_tru",
              "so_to",
              "so_thua",
              "dia_chi_thua_dat",
              "dien_tich_khai_bao",
              "nguon_goc",
              "hien_trang",
              "hien_trang_chi_tiet",
              "tinh_trang_so",
              "lat",
              "lon",
              "ngay_tao",
          ]],
          use_container_width=True,
      )

      st.markdown("### Xuất dữ liệu báo cáo định dạng chuyên nghiệp")

      if st.button("📥 Tạo và Tải xuống File Excel Báo Cáo"):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Danh sách hiện trạng đất"

        ws.sheet_view.showGridLines = True

        ws.merge_cells("A1:P1")
        ws["A1"] = (
            "DANH SÁCH TỔNG HỢP HIỆN TRẠNG CANH TÁC ĐẤT ĐAI CẤP XÃ"
        ).upper()
        ws["A1"].font = Font(name="Times New Roman", size=14, bold=True)
        ws["A1"].alignment = Alignment(horizontal="center", vertical="center")

        headers = [
            "STT",
            "Họ và tên chủ sử dụng",
            "Số điện thoại",
            "Địa chỉ thường trú",
            "Số tờ",
            "Số thửa",
            "Địa chỉ thửa đất",
            "Diện tích (m²)",
            "Nguồn gốc tự kê khai",
            "Nhóm hiện trạng",
            "Tên cây trồng cụ thể",
            "Tình trạng Giấy chứng nhận",
            "Vĩ độ (Lat)",
            "Kinh độ (Lon)",
            "Link Google Maps",
            "Ngày kê khai",
        ]
        ws.append([])
        ws.append(headers)

        header_font = Font(
            name="Times New Roman", size=11, bold=True, color="FFFFFF"
        )
        header_fill = PatternFill(
            start_color="1F4E78", end_color="1F4E78", fill_type="solid"
        )
        header_align = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )

        for col_num in range(1, len(headers) + 1):
          cell = ws.cell(row=3, column=col_num)
          cell.font = header_font
          cell.fill = header_fill
          cell.alignment = header_align

        thin_border = Border(
            left=Side(style="thin", color="D9D9D9"),
            right=Side(style="thin", color="D9D9D9"),
            top=Side(style="thin", color="D9D9D9"),
            bottom=Side(style="thin", color="D9D9D9"),
        )

        data_font = Font(name="Times New Roman", size=11)

        # Xuất dữ liệu theo bảng đang lọc hoặc toàn bộ
        for idx, row in df_hien_thi.reset_index(drop=True).iterrows():
          lat_val = row["lat"]
          lon_val = row["lon"]
          map_link = (
              f"https://www.google.com/maps?q={lat_val},{lon_val}"
              if pd.notnull(lat_val) and pd.notnull(lon_val)
              else "Chưa có vị trí"
          )

          row_data = [
              idx + 1,
              row["ho_ten"],
              str(row["sdt"]) if pd.notnull(row["sdt"]) else "",
              str(row["dia_chi_thuong_tru"])
              if pd.notnull(row["dia_chi_thuong_tru"])
              else "",
              str(row["so_to"]) if pd.notnull(row["so_to"]) else "",
              str(row["so_thua"]) if pd.notnull(row["so_thua"]) else "",
              str(row["dia_chi_thua_dat"])
              if pd.notnull(row["dia_chi_thua_dat"])
              else "",
              row["dien_tich_khai_bao"],
              str(row["nguon_goc"]) if pd.notnull(row["nguon_goc"]) else "",
              row["hien_trang"],
              row["hien_trang_chi_tiet"]
              if pd.notnull(row["hien_trang_chi_tiet"])
              else "",
              row["tinh_trang_so"],
              lat_val if pd.notnull(lat_val) else "",
              lon_val if pd.notnull(lon_val) else "",
              map_link,
              str(row["ngay_tao"]),
          ]
          ws.append(row_data)

        for row_idx in range(4, 4 + len(df_hien_thi)):
          for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.font = data_font
            cell.border = thin_border

            if col_idx in [1, 5, 6, 13, 14, 16]:
              cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx == 8:
              cell.alignment = Alignment(horizontal="right", vertical="center")
              cell.number_format = "#,##0"
            elif col_idx == 15:
              cell.alignment = Alignment(horizontal="center", vertical="center")
              cell.font = Font(
                  name="Times New Roman",
                  size=10,
                  color="0563C1",
                  underline="single",
              )
            else:
              cell.alignment = Alignment(horizontal="left", vertical="center")

        for col in ws.columns:
          max_len = 0
          col_letter = openpyxl.utils.get_column_letter(col[0].column)
          for cell in col:
            if cell.row > 1:
              val_str = str(cell.value or "")
              if len(val_str) > max_len:
                max_len = len(val_str)
          ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

        local_file_name = "Bao_cao_hien_trang_dat_dai.xlsx"
        wb.save(local_file_name)

        with open(local_file_name, "rb") as f:
          st.download_button(
              label="📥 Nhấn vào đây để tải file Excel báo cáo đẹp",
              data=f,
              file_name=local_file_name,
              mime=(
                  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
              ),
          )

        st.success(f"Đã cập nhật file `{local_file_name}` mới nhất!")

  elif password != "":
    st.error("Sai mật khẩu quản lý! Vui lòng thử lại.")
  else:
    st.info("Vui lòng nhập mật khẩu quản lý để xem danh sách và xuất báo cáo.")
