def find_cell_indexes(sheet, target_value, case_sensitive=True):

    found_cells = []

    # Приведение к строке если необходимо
    if not case_sensitive and isinstance(target_value, str):
        target_value = target_value.lower()

    for row in sheet.iter_rows():
        for cell in row:
            cell_value = cell.value

            # Обработка регистра
            if not case_sensitive and isinstance(cell_value, str):
                cell_value = cell_value.lower()

            if cell_value == target_value:
                found_cells.append(cell.coordinate)
    return found_cells