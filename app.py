from docxtpl import DocxTemplate
import pandas as pd
from docx import Document
from docxcompose.composer import Composer

# Đọc dữ liệu từ Excel
df = pd.read_excel("Mẫu Data excel.xlsx", sheet_name="ds")

# Nhóm theo Số hóa đơn
grouped = df.groupby("Số hóa đơn")

# Tạo tài liệu chính để ghép biên bản
master_doc = Document()
composer = Composer(master_doc)

for so_hd, group in grouped:
    info = group.iloc[0]

    # Tạo dữ liệu context thay vào mẫu
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

    # Tạo một bản từ template
    tpl = DocxTemplate("Mau bien ban tra lai hang.docx")
    tpl.render(context)
    temp_doc_path = f"temp_{so_hd}.docx"
    tpl.save(temp_doc_path)

    # Nạp và nối vào tài liệu chính
    sub_doc = Document(temp_doc_path)
    composer.append(sub_doc)

# Lưu toàn bộ file Word gộp lại
composer.save("Bien_ban_tra_hang_tong_hop.docx")
