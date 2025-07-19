import streamlit as st
import pandas as pd
import zipfile
from io import BytesIO
import traceback
from datetime import datetime

st.set_page_config(page_title="Tạo File Hạch Toán", layout="wide")
st.title("📋 Tạo File Hạch Toán - Full Dịch Vụ")

tab1, tab2 = st.tabs(["🧾 Tạo File Hạch Toán", "🔍 So sánh khách bị thiếu"])

# ====================== TAB 1 ======================
with tab1:
    uploaded_file = st.file_uploader("📂 Chọn file Excel (.xlsx)", type=["xlsx"])
    chu_hau_to = st.text_input("✍️ Hậu tố chứng từ (VD: A, B1, NV123)").strip().upper()

    def classify_category(value):
        value = str(value).strip().upper()
        if value == "KB NGOẠI TRÚ":
            return "KCB"
        elif "THUỐC" in value:
            return "THUOC"
        elif "TIÊM" in value or "VAC" in value:
            return "VACCINE"
        return None

    category_info = {
        "KCB": {"ma": "KHACHLE01", "ten": "Khách hàng lẻ - Khám chữa bệnh"},
        "THUOC": {"ma": "KHACHLE02", "ten": "Khách hàng lẻ - Thuốc"},
        "VACCINE": {"ma": "KHACHLE03", "ten": "Khách hàng lẻ - Vaccine"}
    }

    if st.button("🚀 Tạo File Zip") and uploaded_file and chu_hau_to:
        try:
            xls = pd.ExcelFile(uploaded_file)
            data_by_category = {key: {} for key in category_info}
            logs = []
            prefix = "T00_0000"
            all_original_data = pd.DataFrame()

            for sheet_name in xls.sheet_names:
                df = xls.parse(sheet_name)
                df.columns = [str(col).strip().upper() for col in df.columns]

                if not {"KHOA/BỘ PHẬN", "TIỀN MẶT", "NGÀY KHÁM", "NGÀY QUỸ", "HỌ VÀ TÊN"}.issubset(df.columns):
                    logs.append(f"⚠️ Sheet `{sheet_name}` thiếu cột cần thiết.")
                    continue

                df["TIỀN MẶT"] = pd.to_numeric(df["TIỀN MẶT"], errors="coerce")
                df = df[df["TIỀN MẶT"].notna() & (df["TIỀN MẶT"] != 0)]

                df["CATEGORY"] = df["KHOA/BỘ PHẬN"].apply(classify_category)
                df = df[df["CATEGORY"].isin(category_info.keys())]

                if df.empty:
                    logs.append(f"⏩ Sheet `{sheet_name}` không có dữ liệu hợp lệ.")
                    continue

                all_original_data = pd.concat([all_original_data, df], ignore_index=True)

                for category in df["CATEGORY"].unique():
                    df_cat = df[df["CATEGORY"] == category]

                    for mode in ["PT", "PC"]:
                        df_mode = df_cat[df_cat["TIỀN MẶT"] > 0] if mode == "PT" else df_cat[df_cat["TIỀN MẶT"] < 0]
                        if df_mode.empty:
                            continue

                        out_df = pd.DataFrame()
                        ngay_quy = pd.to_datetime(df_mode["NGÀY QUỸ"], errors="coerce")
                        ngay_kham = pd.to_datetime(df_mode["NGÀY KHÁM"], errors="coerce")

                        for d in [ngay_quy, ngay_kham]:
                            sample = d.dropna()
                            if not sample.empty:
                                prefix = f"T{sample.iloc[0].month:02}_{sample.iloc[0].year}"
                                break

                        out_df["Ngày hạch toán (*)"] = ngay_quy.dt.strftime("%d/%m/%Y")
                        out_df["Ngày chứng từ (*)"] = ngay_kham.dt.strftime("%d/%m/%Y")

                        def gen_so_chung_tu(date_str):
                            try:
                                d, m, y = date_str.split("/")
                                return f"{mode}_{category}_{d}{m}{y}_{chu_hau_to}"
                            except:
                                return f"{mode}_INVALID_{chu_hau_to}"

                        out_df["Số chứng từ (*)"] = out_df["Ngày chứng từ (*)"].apply(gen_so_chung_tu)
                        out_df["Mã đối tượng"] = category_info[category]["ma"]
                        out_df["Tên đối tượng"] = category_info[category]["ten"]
                        out_df["Địa chỉ"] = ""
                        out_df["Lý do nộp"] = "Thu khác" if mode == "PT" else "Chi khác"
                        noun = category_info[category]["ten"].split("-")[-1].strip().lower()
                        out_df["Diễn giải lý do nộp"] = ("Thu tiền" if mode == "PT" else "Chi tiền") + f" {noun} ngày " + out_df["Ngày chứng từ (*)"]
                        out_df["Diễn giải (hạch toán)"] = out_df["Diễn giải lý do nộp"] + " - " + df_mode["HỌ VÀ TÊN"]
                        out_df["Loại tiền"] = ""
                        out_df["Tỷ giá"] = ""
                        out_df["TK Nợ (*)"] = "1111" if mode == "PT" else "131"
                        out_df["TK Có (*)"] = "131" if mode == "PT" else "1111"
                        out_df["Số tiền"] = df_mode["TIỀN MẶT"].abs()
                        out_df["Quy đổi"] = ""
                        out_df["Mã đối tượng (hạch toán)"] = ""
                        out_df["Số TK ngân hàng"] = ""
                        out_df["Tên ngân hàng"] = ""

                        data_by_category[category].setdefault(sheet_name, {})[mode] = out_df
                        logs.append(f"✅ {sheet_name} ({category}) [{mode}]: {len(out_df)} dòng")

            st.session_state["original_df"] = all_original_data[["HỌ VÀ TÊN", "KHOA/BỘ PHẬN", "NGÀY KHÁM"]].drop_duplicates()

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

        except Exception:
            st.error("❌ Đã xảy ra lỗi:")
            st.code(traceback.format_exc(), language="python")

with tab2:
    st.markdown("### 🧮 So sánh khách giữa file gốc và các file đầu ra")
    original_file = st.file_uploader("📂 Chọn file Excel GỐC", type=["xlsx"], key="origin_file_compare")
    uploaded_files = st.file_uploader("📂 Chọn các file Excel đầu ra để so sánh", type=["xlsx"], accept_multiple_files=True, key="output_files_compare")

    if original_file and uploaded_files:
        try:
            # Đọc toàn bộ file gốc
            df_goc = pd.read_excel(original_file, sheet_name=None)
            df_all = pd.concat(df_goc.values(), ignore_index=True)
            df_all.columns = df_all.columns.str.upper()
            df_all = df_all[["HỌ VÀ TÊN", "KHOA/BỘ PHẬN", "NGÀY KHÁM", "DIỄN GIẢI"]].dropna(subset=["NGÀY KHÁM"])
            df_all["HỌ VÀ TÊN"] = df_all["HỌ VÀ TÊN"].astype(str).str.strip()
            df_all["NGÀY KHÁM"] = pd.to_datetime(df_all["NGÀY KHÁM"], errors="coerce")

            # Gắn loại dịch vụ
            def get_dich_vu(row):
                dien_giai = str(row.get("DIỄN GIẢI", "")).upper()
                if "VẮC XIN" in dien_giai:
                    return "VACCINE"
                elif "THUỐC" in dien_giai:
                    return "THUOC"
                else:
                    return "KCB"

            df_all["DỊCH VỤ"] = df_all.apply(get_dich_vu, axis=1)

            # Đọc các file đầu ra
            all_missing = []

            for file in uploaded_files:
                # Nhận diện loại dịch vụ từ tên file
                filename = file.name.upper()
                if "KCB" in filename:
                    dv = "KCB"
                elif "THUOC" in filename:
                    dv = "THUOC"
                elif "VACCINE" in filename:
                    dv = "VACCINE"
                else:
                    st.warning(f"⚠️ Không xác định được loại dịch vụ từ file: {file.name}")
                    continue

                df_out = pd.read_excel(file)
                if "Diễn giải (hạch toán)" not in df_out.columns:
                    st.warning(f"⚠️ Thiếu cột 'Diễn giải (hạch toán)' trong file: {file.name}")
                    continue

                df_out["HỌ VÀ TÊN"] = df_out["Diễn giải (hạch toán)"].str.extract(r"- (.*)")[0].str.strip()
                df_out = df_out[df_out["HỌ VÀ TÊN"].notna()]
                output_names = set(df_out["HỌ VÀ TÊN"])

                df_goc_dv = df_all[df_all["DỊCH VỤ"] == dv]
                grouped = df_goc_dv.groupby(df_goc_dv["NGÀY KHÁM"].dt.strftime("%d/%m/%Y"))

                for date_str, group in grouped:
                    names_in_goc = set(group["HỌ VÀ TÊN"])
                    missing_names = names_in_goc - output_names
                    if missing_names:
                        df_missing = group[group["HỌ VÀ TÊN"].isin(missing_names)].copy()
                        df_missing["NGÀY"] = date_str
                        df_missing["DỊCH VỤ"] = dv
                        all_missing.append(df_missing)

            # Hiển thị kết quả
            if all_missing:
                df_all_missing = pd.concat(all_missing, ignore_index=True)
                total_missing = len(df_all_missing)
                st.markdown(f"### ❌ Thiếu khách (có trong file gốc nhưng không có trong đầu ra) ({total_missing} khách)")

                for (dv, date_str), group in df_all_missing.groupby(["DỊCH VỤ", "NGÀY"]):
                    st.markdown(f"#### 📅 Ngày khám: `{date_str}` - 🧾 Dịch vụ: `{dv}`")
                    st.dataframe(group[["HỌ VÀ TÊN", "KHOA/BỘ PHẬN", "NGÀY KHÁM"]], use_container_width=True)
            else:
                st.success("✅ Không thiếu khách nào theo từng dịch vụ và ngày khám.")

        except Exception as e:
            st.error("❌ Lỗi khi xử lý so sánh:")
            st.code(traceback.format_exc())
    else:
        st.info("📥 Vui lòng chọn đầy đủ file gốc và các file đầu ra để tiến hành so sánh.")
