import openpyxl
from path_of_excel_files import VRP, BEGIN_DATE, AREA

VRP = openpyxl.load_workbook(filename=VRP)
VRP_SHEET = VRP['Лист1']


def get_name_of_indicator():
    indicator_name = ''
    for cell in VRP_SHEET['A']:
        if cell.value == 'Валовой региональный продукт по субъектам Российской Федерации (валовая добавленная ' \
                         'стоимость в основных ценах)':
            indicator_name += cell.value
    return indicator_name


def get_years():
    """
    Возвращает заголовки Excel-файла(годы)
    :return: Заголовки
    """
    headers = []
    for cell in VRP_SHEET[5]:
        headers.append(str(cell.value))
    headers = [char.replace("г.", "") for char in headers]
    headers = list(map(int, headers[1::]))
    begin_years = [year for year in headers if year > BEGIN_DATE]
    return begin_years


def get_areas():
    """
    Получить все значения областей из списка
    :return:
    """
    areas = []
    for cell in VRP_SHEET['A']:
        if cell.value is None:
            continue
        if 'федеральный' in cell.value:
            continue
        if 'Республика' in cell.value:
            areas.append(cell.value)
        if 'область' in cell.value:
            areas.append(cell.value)
        if 'округ' in cell.value:
            areas.append(cell.value)
    return areas


def get_value_vrp(areas):
    """
    Срез значений по годам для одной области
    :param area: область AREA
    :return:  {ГОД :{Область:Значение}}
    """
    area_name = {}
    years = get_years()
    areas = get_areas()
    indicator_name = get_name_of_indicator()

    for row in VRP_SHEET.iter_rows(min_row=1, min_col=1, values_only=True):
        for j in areas:
            for i in row[1:]:
                if i == '…' or isinstance(i, int) or isinstance(i, float):
                    print(i,j)

    return area_name


if __name__ == '__main__':
    get_years()
    print(get_value_vrp(areas=AREA))
