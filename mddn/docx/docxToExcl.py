import docx
from openpyxl import Workbook



fn = r'C:\Users\CYY\Desktop\docx\表格.docx'
doc = docx.Document(fn)

wb = Workbook()
sheet = wb.active

header_list = []
# 按段落读取全部数据
for paragraph in doc.paragraphs:
    print(paragraph.text)
    header_list.append(paragraph.text)

# 按表格读取全部数据
for table in doc.tables:
    row_list = []
    for row in table.rows:
        cell_list = []
        for cell in row.cells:
            cell_list.append(cell.text)
        row_list.append(cell_list)
        # 写入行
        sheet.append(cell_list)
        # 保存为excl
        wb.save('2020.xlsx')




