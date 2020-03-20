import tabula
import PyPDF2
import xlwt
import threading

file = open("pdf/simple_table.pdf", 'rb')
file_pdf = PyPDF2.PdfFileReader(file)

len_doc = file_pdf.getNumPages()

# menyimpan data ke dalam file excel
worksheet = xlwt.Workbook()
sheet = worksheet.add_sheet('page 1')
baris_excel = 0

for i in range(len_doc):
    print("ekstrak page : " + str(i))

    # membaca table dari pdf
    data = tabula.read_pdf("pdf/simple_table.pdf", pages=i + 1, output_format='JSON')
    save_data = []
    for sub_data in data[0]['data']:
        sub_save = []
        for sub_json in sub_data:
            sub_save.append(sub_json['text'])
        save_data.append(sub_save)

    # memasukkan data kedalam file excel
    for sub in save_data:
        for i_sub in range(len(sub)):
            sheet.write(baris_excel, i_sub, sub[i_sub])
        baris_excel += 1

worksheet.save("excel/simple_table.xls")