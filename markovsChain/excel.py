import openpyxl

def find_intersection_value(header_value, index_value):
    file_path = 'transition_matrix.xlsx'
    # Открываем Excel файл
    workbook = openpyxl.load_workbook(file_path)
    # Выбираем активный лист
    worksheet = workbook.active

    # Ищем индекс столбца по значению в первой строке
    col_idx = None
    for cell in worksheet[1]:
        if cell.value == header_value:
            col_idx = cell.column
            print(col_idx)
            break

    # Ищем индекс строки по значению в первом столбце
    row_idx = None
    for cell in worksheet['A']:
        if cell.value == index_value:
            row_idx = cell.row
            print(row_idx)
            break

    # Если найдены оба индекса, возвращаем значение в ячейке на их пересечении
    if col_idx is not None and row_idx is not None:
        intersection_cell = worksheet.cell(row=row_idx, column=col_idx)
        return intersection_cell.value

    # Если значение не найдено, возвращаем None
    return None

# Пример использования функции

header_value = 'й'  # Значение в первой строке
index_value = 'л'    # Значение в первом столбце
result = round(find_intersection_value(header_value, index_value), 1000000)
print(result)