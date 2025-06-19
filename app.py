
import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
from docx import Document
from docxcompose.composer import Composer
import tempfile
import os

st.set_page_config(page_title="Tạo Biên Bản Trả Hàng", layout="centered")
st.title("📄 Tạo Biên Bản Trả Hàng Từ Excel & Word Mẫu")

uploaded_excel = st.file_uploader("📊 Tải lên file Excel (sheet 'ds')", type=["xlsx"])
uploaded_template = st.file_uploader("📄 Tải lên file Word mẫu (.docx)", type=["docx"])

if uploaded_excel and uploaded_template:
    with tempfile.TemporaryDirectory() as tmpdir:
        # Đọc dữ liệu trực tiếp từ file upload
        df = pd.read_excel(uploaded_excel, sheet_name="ds")
        grouped = df.groupby("Số hóa đơn")

        # Ghi template ra file tạm (docxtpl cần đường dẫn)
        template_path = os.path.join(tmpdir, "template.docx")
        with open(template_path, "wb") as f:
            f.write(uploaded_template.read())

        # Tạo tài liệu tổng hợp
        master_doc = Document()
        composer = Composer(master_doc)

        for so_hd, group in grouped:
            info = group.iloc[0]
            context = {
                "Tại_văn_phòng_": info["Tại văn phòng: "],
                "BÊN_A_Bên_mua": info["BÊN A (Bên mua)"],
                "Địa_chỉ": info["Địa chỉ"],
                "Mã_số_thuế": info["Mã số thuế"],
                "xuất_hóa_đơn_tài_chính_số": info["xuất hóa đơn tài chính số"],
                "Tên_hàng": group.iloc[0]["Tên hàng"],
                "ĐVT": group.iloc[0]["ĐVT"],
                "Số_lượng": group.iloc[0]["Số lượng"],
                "Đơn_giá": f"{group.iloc[0]['Đơn giá']:,.0f}",
                "Thành_tiền": f"{group['Thành tiền'].sum():,.0f}",
                "Thuế_suất": str(group.iloc[0]["Thuế suất"]),
                "Tiền_thuế_GTGT": f"{group['Tiền thuế GTGT'].sum():,.0f}",
                "Tổng_tiền_thanh_toán": f"{group['Tổng tiền thanh toán'].sum():,.0f}",
                "Số_tiền_bằng_chữ": info["Số tiền bằng chữ"],
            }

            tpl = DocxTemplate(template_path)
            tpl.render(context)
            rendered_path = os.path.join(tmpdir, f"bb_{so_hd}.docx")
            tpl.save(rendered_path)

            composer.append(Document(rendered_path))

        # Lưu biên bản tổng hợp
        final_path = os.path.join(tmpdir, "Bien_ban_tra_hang_tong_hop.docx")
        composer.save(final_path)

        # Cho phép tải về
        with open(final_path, "rb") as f:
            st.download_button("📥 Tải file biên bản tổng hợp", data=f, file_name="Bien_ban_tra_hang.docx")
