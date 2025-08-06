import xlwings as xw

wb = xw.Book(r'C:\Users\EfrénAlexisHernández\OneDrive - FANGIO COM\Escritorio\nuevo baseado\site_survey\prueba_checkbox.xlsx')
ws = wb.sheets[0]
ws.range('A1').value = True
wb.save(r'C:\Users\EfrénAlexisHernández\OneDrive - FANGIO COM\Escritorio\nuevo baseado\site_survey\prueba_checkbox_result.xlsx')
wb.close()
print("¡Listo! Revisa el archivo generado.")