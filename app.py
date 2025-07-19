import streamlit as st
import pandas as pd
import zipfile
import tempfile
import os
import re
import traceback
from io import BytesIO

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
    zip_file = st.file_uploader("📂 Upload file ZIP chứa các file đầu ra (KCB, THUOC, VACCINE)", type=["zip"], key="zip_output_compare")

    if original_file and zip_file:
        try:
            # Đọc file gốc
            df_goc = pd.read_excel(original_file, sheet_name=None)
            df_goc_all = pd.concat(df_goc.values(), ignore_index=True)
            df_goc_all.columns = df_goc_all.columns.str.upper()
            df_goc_all["HỌ VÀ TÊN"] = df_goc_all["HỌ VÀ TÊN"].astype(str).str.strip().str.upper()
            df_goc_all["NGÀY KHÁM"] = pd.to_datetime(df_goc_all["NGÀY KHÁM"], errors="coerce")
            df_goc_all = df_goc_all.dropna(subset=["NGÀY KHÁM"])

            # Giải nén file zip
            with tempfile.TemporaryDirectory() as tmpdir:
                with zipfile.ZipFile(zip_file, "r") as zip_ref:
                    zip_ref.extractall(tmpdir)

                all_missing = {}

                for filename in os.listdir(tmpdir):
                    if not filename.lower().endswith(".xlsx"):
                        continue
                    file_path = os.path.join(tmpdir, filename)
                    match = re.search(r'(\d{2}-\d{2}-\d{4})', filename)
                    if not match:
                        continue
                    date_str = match.group(1)
                    date_obj = pd.to_datetime(date_str, dayfirst=True, errors="coerce")
                    if pd.isna(date_obj):
                        continue

                    df_out = pd.read_excel(file_path, sheet_name=None)
                    all_names = set()

                    for sheet in df_out:
                        df = df_out[sheet]
                        if "Diễn giải (hạch toán)" in df.columns:
                            extracted = df["Diễn giải (hạch toán)"].astype(str).str.extract(r"-\s*(.*)")
                            names = extracted[0].dropna().str.strip().str.upper()
                            all_names.update(names)

                    # Lọc dữ liệu gốc chỉ trong ngày và bộ phận tương ứng
                    df_day = df_goc_all[df_goc_all["NGÀY KHÁM"] == date_obj]
                    khoa = None
                    if "_KCB_" in filename.upper():
                        khoa = "KCB"
                    elif "_THUOC_" in filename.upper():
                        khoa = "THUỐC"
                    elif "_VACCINE_" in filename.upper():
                        khoa = "VACCINE"
                    if khoa:
                        df_day = df_day[df_day["KHOA/BỘ PHẬN"].str.upper().str.contains(khoa)]

                    guest_set = set(df_day["HỌ VÀ TÊN"])
                    missing_guests = guest_set - all_names

                    if missing_guests:
                        df_missing = df_day[df_day["HỌ VÀ TÊN"].isin(missing_guests)]
                        all_missing[date_str + f" ({khoa})"] = df_missing

            if all_missing:
                st.markdown(f"### ❌ Thiếu khách ({sum(len(df) for df in all_missing.values())} khách)")
                for date, df in all_missing.items():
                    st.markdown(f"#### 📅 Ngày khám: `{date}`")
                    st.dataframe(df[["HỌ VÀ TÊN", "KHOA/BỘ PHẬN", "NGÀY KHÁM"]], use_container_width=True)
            else:
                st.success("✅ Không thiếu khách nào trong các file bạn đã chọn!")

        except Exception as e:
            st.error("❌ Lỗi khi so sánh:")
            st.code(traceback.format_exc())
    else:
        st.info("📥 Vui lòng chọn file gốc và file zip đầu ra để so sánh.")
