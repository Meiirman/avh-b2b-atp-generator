# import openpyxl
# from openpyxl import load_workbook

# # Load the Excel template
# wb = load_workbook('template.xlsx')
# sheet = wb.active

# # Define the variables to fill in the template
# context = {
#     "test1" : "22.12.2023",
#     "test2" : "01:11:59",
#     "test3" : "22.10.2022",
#     "test4" : "15",
#     "test5" : "15.00",
#     "test6" : "15,00",
#     "test7" : "15000,00",
#     "test8" : "15 000,00",
#     "test9" : "15000.00",
#     "test10" : "15 000.00",
#     "test11" : "15000",
#     "test12" : "15 000",
#     "test13" : 15.00,
#     "test14" : 15000.00,
#     "test15" : 15000,
#     "test16" : "qwer qwer qwer qwer qwer wsg",
#     "test17" : "wh156ewr5h1w864h14w365h4erjh6rh",
#     "test18" : "aerherb\nerhewrh\nerhn\n",
# }

# # Replace named ranges with actual data
# for cell_name, value in context.items():
#     sheet[cell_name].value = value

# # Save the resulting Excel file
# wb.save('output.xlsx')


import xlsxtpl

# Создаем шаблон Excel
template_path = "template.xlsx"

# Читаем шаблон
template = xlsxtpl.read_template(template_path)

# Создаем словарь с данными
data = {
    "title": "Отчет",
    "data": [
        [1, 2, 3],
        [4, 5, 6],
        [7, 8, 9],
    ],
}

# Генерируем файл Excel
xlsxtpl.render_template(template, data)

