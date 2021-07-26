#from openpyxl.workbook import Workbook
#from for_indexes_reading import get_index_vrp
#from for_vrp_reading import get_value_vrp
#from path_of_excel_files import AREA

#HEADERS = ['Название показателя', 'Область', 'Год', 'Значение']
#DATA = [get_index_vrp(areas=AREA), get_value_vrp(areas=AREA)]


#def writing_results():
#    wb = Workbook()
#    dest_filename = 'Results/output.xlsx'
#    ws1 = wb.active
#    ws1.title = 'Результат работы парсера'
#    ws1.append(HEADERS)
#    for row in DATA:
#        for name_pok, areas in row.items():
#            for area, year_value in areas.items():
#                for year, value in year_value.items():
#                    row = name_pok, area, year, value
#                    ws1.append(row)

#    wb.save(filename=dest_filename)


#writing_results()
