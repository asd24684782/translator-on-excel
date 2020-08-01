from openpyxl import load_workbook
import googletrans



excel_name='sample.xlsx'        #excel's name

wb = load_workbook(excel_name)
sheet = wb.active 

translator = googletrans.Translator()


for row in sheet.rows:
    for cell in row:
        results = translator.translate(cell.value,dest='zh-tw').text
        cell.value=results

wb.save(excel_name)
