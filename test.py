from for_indexes_reading import get_index_vrp
from for_vrp_reading import get_value_vrp
from path_of_excel_files import AREA

parsing_result = [get_index_vrp(areas=AREA), get_value_vrp(areas=AREA)]

for i in parsing_result:
    print(i)