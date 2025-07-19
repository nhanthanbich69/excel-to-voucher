import streamlit as st
import pandas as pd
import zipfile
from io import BytesIO
import traceback
from datetime import datetime

st.set_page_config(page_title="Tạo File Hạch Toán", layout="wide")
st.title("📋 Tạo File Hạch Toán - Chỉ KCB (KB NGOẠI TRÚ)")

tab1, tab2 = st.tabs(["🧾 Tạo File Hạch Toán", "🔍 So sánh khách bị thiếu"])

with tab1:
    uploaded_file = st.file_uploader("📂 Chọn file Excel (.xlsx)", type=["xlsx"])
    chu_hau_to = st.text_input("✍️ Hậu tố chứng từ (VD: A, B1, NV123)").strip().upper()

    def classify_department(value):
        if isinstance(value, str) and value.strip().upper() == "KB NGOẠI TRÚ":
            return "KCB"
        return None

    category_info = {
        "KCB": {"ma": "KHACHLE01", "ten": "Khách hàng lẻ - Khám chữa bệnh"}
    }

    if st.button("🚀 Tạo File Zip") and uploaded_file and chu_hau_to:
        try:
            xls = pd.ExcelFile(uploaded_file)
            st.success(f"📥 Đọc thành công file `{uploaded_file.name}` với {len(xls.sheet_names)} sheet.")

            data_by_category = {"KCB": {}}
            logs = []
            prefix = "T00_0000"
            all_original_data = pd.DataFrame()

            for sheet_name in xls.sheet_names:
                if not sheet_name.replace(".", "", 1).isdigit() and not sheet_name.replace(",", "", 1).isdigit():
                    logs.append(f"⏩ Bỏ qua sheet không hợp lệ: {sheet_name}")
                    continue

                df = xls.parse(sheet_name)
                df.columns = [str(col).strip().upper() for col in df.columns]

                if "KHOA/BỘ PHẬN" not in df.columns or "TIỀN MẶT" not in df.columns:
                    logs.append(f"⚠️ Sheet `{sheet_name}` thiếu cột cần thiết.")
                    continue

                df["TIỀN MẶT"] = pd.to_numeric(df["TIỀN MẶT"], errors="coerce")
                df = df[df["TIỀN MẶT"].notna() & (df["TIỀN MẶT"] != 0)]

                df["CATEGORY"] = df["KHOA/BỘ PHẬN"].apply(classify_department)
                df = df[df["CATEGORY"] == "KCB"]

                if df.empty:
                    logs.append(f"⏩ Sheet `{sheet_name}` không có dữ liệu KCB từ 'KB NGOẠI TRÚ'.")
                    continue

                category = "KCB"
                all_original_data = pd.concat([all_original_data, df], ignore_index=True)

                for mode in ["PT", "PC"]:
                    is_pt = mode == "PT"
                    df_mode = df[df["TIỀN MẶT"] > 0] if is_pt else df[df["TIỀN MẶT"] < 0]
                    if df_mode.empty:
                        continue

                    out_df = pd.DataFrame()
                    ngay_quy = pd.to_datetime(df_mode["NGÀY QUỸ"], errors="coerce")
                    ngay_kham = pd.to_datetime(df_mode["NGÀY KHÁM"], errors="coerce")

                    for date_series in [ngay_quy, ngay_kham]:
                        sample_date = date_series.dropna()
                        if not sample_date.empty:
                            prefix = f"T{sample_date.iloc[0].month:02}_{sample_date.iloc[0].year}"
                            break

                    out_df["Ngày hạch toán (*)"] = ngay_quy.dt.strftime("%d/%m/%Y")
                    out_df["Ngày chứng từ (*)"] = ngay_kham.dt.strftime("%d/%m/%Y")

                    def gen_so_chung_tu(date_str):
                        try:
                            d, m, y = date_str.split("/")
                            return f"{mode}_{'THUOC'}_{d}{m}{y}_{chu_hau_to}"
                        except:
                            return f"{mode}_INVALID_{chu_hau_to}"

                    out_df["Số chứng từ (*)"] = out_df["Ngày chứng từ (*)"].apply(gen_so_chung_tu)
                    out_df["Mã đối tượng"] = category_info[category]["ma"]
                    out_df["Tên đối tượng"] = category_info[category]["ten"]
                    out_df["Địa chỉ"] = ""
                    out_df["Lý do nộp"] = "Thu khác" if is_pt else "Chi khác"
                    noun = category_info[category]["ten"].split("-")[-1].strip().lower()
                    out_df["Diễn giải lý do nộp"] = ("Thu tiền" if is_pt else "Chi tiền") + f" {noun} ngày " + out_df["Ngày chứng từ (*)"]
                    out_df["Diễn giải (hạch toán)"] = out_df["Diễn giải lý do nộp"] + " - " + df_mode["HỌ VÀ TÊN"]
                    out_df["Loại tiền"] = ""
                    out_df["Tỷ giá"] = ""
                    out_df["TK Nợ (*)"] = "1111" if is_pt else "131"
                    out_df["TK Có (*)"] = "131" if is_pt else "1111"
                    out_df["Số tiền"] = df_mode["TIỀN MẶT"].abs()
                    out_df["Quy đổi"] = ""
                    out_df["Mã đối tượng (hạch toán)"] = ""
                    out_df["Số TK ngân hàng"] = ""
                    out_df["Tên ngân hàng"] = ""

                    data_by_category[category].setdefault(sheet_name, {})[mode] = out_df
                    logs.append(f"✅ {sheet_name} ({category}) [{mode}]: {len(out_df)} dòng")

            st.session_state["original_df"] = all_original_data[["HỌ VÀ TÊN", "KHOA/BỘ PHẬN", "NGÀY KHÁM"]].drop_duplicates()

            if all(not sheets for sheets in data_by_category.values()):
                st.warning("⚠️ Không có dữ liệu hợp lệ sau khi lọc.")
            else:
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                    all_exported_names = []

                    for category, sheets in data_by_category.items():
                        for day, data in sheets.items():
                            output = BytesIO()
                            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                                for mode in ["PT", "PC"]:
                                    if mode in data and not data[mode].empty:
                                        full_df = data[mode]
                                        all_exported_names += list(full_df["Diễn giải (hạch toán)"].str.extract(r"- (.*)")[0])
                                        chunks = [full_df[i:i+500] for i in range(0, len(full_df), 500)]
                                        for idx, chunk in enumerate(chunks):
                                            sheet_tab = mode if idx == 0 else f"{mode} {idx + 1}"
                                            chunk.to_excel(writer, sheet_name=sheet_tab, index=False)
                            output.seek(0)
                            zip_path = f"{prefix}_{category}/{day.replace(',', '.').strip()}.xlsx"
                            zip_file.writestr(zip_path, output.read())

                    st.session_state["exported_names"] = list(set(all_exported_names))

                st.success("🎉 Đã xử lý xong!")
                st.download_button("📦 Tải File Zip", data=zip_buffer.getvalue(), file_name=f"{prefix}.zip")

            st.markdown("### 📄 Nhật ký xử lý")
            st.markdown("\n".join([f"- {line}" for line in logs]))

        except Exception as e:
            st.error("❌ Đã xảy ra lỗi:")
            st.code(traceback.format_exc(), language="python")


with tab2:
    st.markdown("### 🧮 So sánh khách giữa file gốc và các file đầu ra")

    original_file = st.file_uploader("📂 Chọn file Excel GỐC", type=["xlsx"], key="goc")
    output_files = st.file_uploader("📂 Chọn các file Excel đầu ra để so sánh", type=["xlsx"], accept_multiple_files=True, key="daura")

    if original_file and output_files:
        try:
            # Đọc file gốc
            df_orig = pd.read_excel(original_file)
            df_orig.columns = [c.strip().upper() for c in df_orig.columns]
            df_orig = df_orig[df_orig["HỌ VÀ TÊN"].notna()]
            df_orig["NGÀY KHÁM"] = pd.to_datetime(df_orig["NGÀY KHÁM"], errors="coerce")

            if "KHOA/BỘ PHẬN" not in df_orig.columns:
                df_orig["KHOA/BỘ PHẬN"] = "Không rõ"

            original_guests = df_orig[["HỌ VÀ TÊN", "KHOA/BỘ PHẬN", "NGÀY KHÁM"]].drop_duplicates()

            # Đọc toàn bộ file đầu ra
            out_all = pd.DataFrame()
            for f in output_files:
                xls = pd.ExcelFile(f)
                for sheet in xls.sheet_names:
                    df_tmp = xls.parse(sheet)
                    df_tmp.columns = [c.strip().upper() for c in df_tmp.columns]
                    if "DIỄN GIẢI (HẠCH TOÁN)" in df_tmp.columns:
                        out_all = pd.concat([out_all, df_tmp], ignore_index=True)

            # Trích tên khách từ diễn giải
            out_all["HỌ VÀ TÊN"] = out_all["DIỄN GIẢI (HẠCH TOÁN)"].str.extract(r"- (.*)")
            output_guests = out_all["HỌ VÀ TÊN"].dropna().unique()

            original_names = set(original_guests["HỌ VÀ TÊN"])
            output_names = set(output_guests)

            missing_names = original_names - output_names
            extra_names = output_names - original_names

            def display_guest_list(title, name_list, color, full_df):
                if name_list:
                    st.markdown(f"### {title} ({len(name_list)} khách)")
                    df_display = full_df[full_df["HỌ VÀ TÊN"].isin(name_list)].copy()
                    df_display.sort_values("NGÀY KHÁM", inplace=True)
                    for date, group in df_display.groupby(df_display["NGÀY KHÁM"].dt.strftime("%d/%m/%Y")):
                        st.markdown(f"#### 📅 Ngày khám: `{date}`")
                        st.dataframe(group[["HỌ VÀ TÊN", "KHOA/BỘ PHẬN", "NGÀY KHÁM"]], use_container_width=True)
                else:
                    st.success(f"✅ Không có khách nào thuộc nhóm {title.lower()}.")

            display_guest_list("❌ Thiếu khách (có trong file gốc nhưng không có trong đầu ra)", missing_names, "red", original_guests)
            display_guest_list("⚠️ Dư khách (có trong đầu ra nhưng không có trong file gốc)", extra_names, "orange", pd.DataFrame({"HỌ VÀ TÊN": list(extra_names)}))

        except Exception as e:
            st.error("❌ Đã xảy ra lỗi:")
            st.code(traceback.format_exc(), language="python")
    else:
        st.info("⬆️ Vui lòng chọn đầy đủ file GỐC và file đầu ra để tiến hành so sánh.")
