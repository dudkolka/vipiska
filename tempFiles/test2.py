import openpyxl
from docxtpl import DocxTemplate
from num2words import num2words

# def find_and_remove(context):
#     test = 0
#     for i in context:
#         for j in i:
#             if(j == 'Производственная практика:'):
#                 test=context.index(i)
#                 context.pop(context.index(i))
#                 break
#             elif(j == None):
#                 context[context.index(i)][i.index(j)]='зачтено'
#             if (is_number(context[context.index(i)][2])):
#                 context[context.index(i)][2] = str(context[context.index(i)][2]) + '(' + num2words(int(context[context.index(i)][2]), lang='ru') + ')'
#     # context[test1][2]=str(context[test1][2]) + '(' + num2words(int(context[test-1][2]), lang='ru') + ')'
#     return context


def main():
    alphabet = generate_alphabet()
    sheet,doc = load_office_files()
    generate_files(alphabet, sheet, doc)
    # doc.save(context['name'] + ' выписка.docx')

def generate_files(alphabet,sheet,doc):
    for num in range(9, len(list(sheet.rows)) - 7):
        context = return_of_dicts(sheet,alphabet,num)
        context = dicts_redact(context)
        doc.render(context)
        doc.save('./generated_files/'+context['name'+str(num)] + ' выписка.docx')
    return 0

def generate_alphabet():
    """Создает массив букв алфавита и их комбинаций (AA, AB и т.д.)."""
    alphabet = [chr(i) for i in range(65, 91)]  # A-Z
    extended_alphabet = alphabet + [f"A{char}" for char in alphabet] + [f"B{char}" for char in alphabet]
    return extended_alphabet

def load_office_files():
    """Загружает Excel и Word файлы."""
    wb = openpyxl.load_workbook(filename='tdp.xlsx')
    sheet = wb['Лист1']
    doc = DocxTemplate('G.docx')
    return sheet, doc

def return_of_dicts(sheet, alphabet,num):
    context = {}
    # for num in range(9, len(list(sheet.rows)) - 7):
    name = sheet[alphabet[1] + str(num)].value
    context['name'+ str(num)] = name

    # Генерация оценок и часов для основных данных
    context.update(return_of_line(sheet, alphabet, num, range(2, 52), "grade", "hours",'item'))
    # Генерация оценок и часов для дополнительной информации
    context.update(return_of_line(sheet, alphabet, num, range(53, 57), "pgrade", "phours", 'item'))
    context.update(return_of_line(sheet, alphabet, num, range(58, 60), "pgrade", "phours", 'item'))
    print (context)
    return context

def dicts_redact(context):
    for key in context.keys():
        if(isinstance(context[key],int) and context[key]<=10 and ("grade" in key)):
            context[key]=str(context[key])+"("+num2words(context[key],lang='ru')+")"
        elif (context[key] == None):
            context[key] = 'зачтено'
    return context

def return_of_line(sheet, alphabet, row, column_range, grade_key, hours_key, item_key):
    """Извлекает оценки и часы из Excel по указанным диапазонам."""
    data = {}
    for col in column_range:
        grade = sheet[f"{alphabet[col]}{row}"].value
        hours = sheet[f"{alphabet[col]}31"].value
        item = sheet[f"{alphabet[col]}8"].value
        data[f"{grade_key}{col}"] = grade
        data[f"{hours_key}{col}"] = hours
        data[f"{item_key}{col}"] = item
    return data

def create_table(document, headers, rows, style='Table Grid'):
    cols_number = len(headers)

    table = document.add_table(rows=1, cols=cols_number)
    table.style = style

    hdr_cells = table.rows[0].cells
    for i in range(cols_number):
        hdr_cells[i].text = headers[i]

    for row in rows:
        row_cells = table.add_row().cells
        for i in range(cols_number):
            row_cells[i].text = str(row[i])

    return table

if __name__ == "__main__":
    main()