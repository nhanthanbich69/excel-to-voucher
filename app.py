import streamlit as st
import pandas as pd
import zipfile
from io import BytesIO
import traceback
from datetime import datetime

st.set_page_config(page_title="Tạo File Hạch Toán", layout="wide")
st.title("📋 Tạo File Hạch Toán - Chỉ KCB (KB NGOẠI TRÚ)")

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
        all_original_names = set()
        all_final_names = set()
        prefix = "T00_0000"

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

            all_original_names.update(df["HỌ VÀ TÊN"].dropna().unique())

            if df.empty:
                logs.append(f"⏩ Sheet `{sheet_name}` không có dữ liệu KCB từ 'KB NGOẠI TRÚ'.")
                continue

            category = "KCB"

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
                        return f"{mode}_THUOC_{d}{m}{y}_{chu_hau_to}"
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

                all_final_names.update(df_mode["HỌ VÀ TÊN"].dropna().unique())

                data_by_category[category].setdefault(sheet_name, {})[mode] = out_df
                logs.append(f"✅ {sheet_name} ({category}) [{mode}]: {len(out_df)} dòng")

        missing_names = sorted(all_original_names - all_final_names)
        extra_names = sorted(all_final_names - all_original_names)

        if all(not sheets for sheets in data_by_category.values()):
            st.warning("⚠️ Không có dữ liệu hợp lệ sau khi lọc.")
        else:
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                for category, sheets in data_by_category.items():
                    for day, data in sheets.items():
                        output = BytesIO()
                        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                            for mode in ["PT", "PC"]:
                                if mode in data and not data[mode].empty:
                                    full_df = data[mode]
                                    chunks = [full_df[i:i+500] for i in range(0, len(full_df), 500)]
                                    for idx, chunk in enumerate(chunks):
                                        sheet_tab = mode if idx == 0 else f"{mode} {idx + 1}"
                                        chunk.to_excel(writer, sheet_name=sheet_tab, index=False)
                            
                            # ➕ Thêm tab thống kê khách bị thiếu/thừa
                            summary_df = pd.DataFrame({
                                "Khách bị lọc mất (gốc có - kết quả không)": pd.Series(missing_names),
                                "Khách xuất hiện thêm (kết quả có - gốc không)": pd.Series(extra_names)
                            })
                            summary_df.to_excel(writer, sheet_name="🔍 So sánh KH", index=False)

                        output.seek(0)
                        zip_path = f"{prefix}_{category}/{day.replace(',', '.').strip()}.xlsx"
                        zip_file.writestr(zip_path, output.read())

            st.success("🎉 Đã xử lý xong!")
            st.download_button("📦 Tải File Zip", data=zip_buffer.getvalue(), file_name=f"{prefix}.zip")

        st.markdown("### 📄 Nhật ký xử lý")
        st.markdown("\n".join([f"- {line}" for line in logs]))

    except Exception as e:
        st.error("❌ Đã xảy ra lỗi:")
        st.code(traceback.format_exc(), language="python")
