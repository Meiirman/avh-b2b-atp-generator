import os
import traceback
from openpyxl import load_workbook
from shutil import copyfile

import pandas as pd

def generate(render_data, template_path, output_folder_path) -> dict:
    if render_data is None:
        return {"status" : "error", "message" : "Вызов метода не может быть пустым"}
    return render_and_save_excel(render_data, template_path, output_folder_path)

def render_and_save_excel(render_data, template_path, output_folder_path):
    print(render_data)

    render_data["data"]["clent"] = render_data["data"]["BS_COMPANY"]
    render_data["data"]["city"] = render_data["data"]["ORDER_REGION"]
    render_data["data"]["address"] = 	render_data["data"]["BS_ADDRESS"]
    render_data["data"]["number"] = render_data["data"]["BS_NUMBER"]
    
    try: render_data["data"]["date"] = render_data["data"]["ORDER_DOGOVOR_DATE"]
    except:
        traceback.print_exc()
        input()
        render_data["data"]["date"] = 0 

    
    try: render_data["data"]["budget"] =  render_data["data"]["TOTAL_NDS"]
    except:
        traceback.print_exc()
        input()
        render_data["data"]["budget"] = 0 

    try: render_data["data"]["budget_with_nds"] =  render_data["data"]["TOTAL_SUMM_NDS"]
    except:
        traceback.print_exc()
        input()
        render_data["data"]["budget_with_nds"] = 0 

    try: render_data["data"]["nds_budget"] = float(render_data["data"]["budget_with_nds"]) - float(render_data["data"]["budget"])
    except:
        pass

    list_table = render_data["data"]["TABLE"]
    
    if list_table:
        try:
            for i, row in enumerate(list_table):
                if i == 0 :
                    continue

                render_data["data"]["index_" + str(i)] = row["index"]
                render_data["data"]["table_number_" + str(i)] = row["number"]
                render_data["data"]["table_work_name_" + str(i)] = row["work_name"]
                render_data["data"]["table_measure_" + str(i)] = row["measure"]
                render_data["data"]["table_count_" + str(i)] = row["count"]
                render_data["data"]["table_price_" + str(i)] = row["price"]
                render_data["data"]["table_price_with_nds_" + str(i)] = row["price_with_nds"] 

            # Копируем шаблон в выходную директорию
            output_file_path = os.path.join(output_folder_path, "output.xlsx")
            copyfile(template_path, output_file_path)

            # Переименовываем копию в "output.xlsx"
            os.rename(output_file_path, os.path.join(output_folder_path, "output.xlsx"))

            # Загружаем данные из созданной копии
            wb = load_workbook(os.path.join(output_folder_path, "output.xlsx"))
            ws = wb.active

            # Проходимся по каждой ячейке в созданной копии
            for i, row in enumerate(ws.iter_rows(min_row=1, max_row=150, min_col=1, max_col=30)):
                for cell in row:
                    for key, value in render_data["data"].items():
                        if "table_count_" in key:
                            try: cell.value = cell.value.replace("{{" + key + "}}", str(value).replace(".", ","))
                            except: pass

                        elif "table_price_" in key:
                            try: cell.value = cell.value.replace("{{" + key + "}}", str(value).replace(".", ","))
                            except: pass

                        elif "table_price_with_nds_" in key:
                            try: cell.value = cell.value.replace("{{" + key + "}}", str(value).replace(".", ","))
                            except: pass

                        else:                        
                            try: cell.value = cell.value.replace("{{" + key + "}}", str(value))
                            except: pass


            del_rows = []
            del_rows_nums = []
            for idx, row in enumerate(ws.iter_rows(min_row=1, max_row=150, min_col=1, max_col=30)):
                row_text = " ".join([str(cell.value) for cell in row])
                if "{{" in row_text and "}}" in row_text and "index" in row_text:
                    del_rows_nums.append(idx)
                    del_rows.append(row)

            for row in del_rows:
                for cell in row:
                    cell.value = None

            ws.delete_rows(15+len(list_table), amount=64-len(list_table))    

            # ws[f'A{16+len(list_table)}:Q{15+len(list_table)+30}']
            start_row = 15+len(list_table)

            ws[f'L{start_row}'] = f'=SUM(L{15}:L{start_row-1})'
            ws[f'L{start_row+1}'] = f'=L{start_row}*1.12'

            ws[f'B11'] = str(ws[f'B11'].value).replace('L79', f'L{start_row}')

            ws.merge_cells(f'E{start_row+0}:K{start_row+0}')
            ws.merge_cells(f'E{start_row+1}:K{start_row+1}')
            
            ws.merge_cells(f'E{start_row+3}:L{start_row+3}')
            ws.merge_cells(f'E{start_row+4}:L{start_row+4}')
            ws.merge_cells(f'E{start_row+5}:L{start_row+5}')
            



            # # Применяем стили к новым строкам
            # for coord, style in styles_to_apply:
            #     ws[coord].style = style

            # for row_num in del_rows:
            #     ws.row_dimensions[row_num].height = 0


            # Сохраняем созданную копию с измененными значениями
            wb.save(os.path.join(output_folder_path, "output.xlsx"))






            return {"message": "Файл успешно создан и изменен: " + os.path.join(output_folder_path, "output.xlsx")}

        except Exception as e:
            return {"message": f"Произошла ошибка: {str(traceback.format_exc())}"}
    return {"message": f"Нет таблицы"}



# , {{clent}}, {{city}}, {{address}}	
# {{number}}
# {{date}}
# {{budget}} 
# {{budget_with_nds}} 
# {{nds_budget}}


# {{table_start}}
# {{table_end}}



# {{table_number}}
# {{table_work_name}}
# {{table_measure}} 
# {{table_count}}
# {{table_price}}
# {{table_price_with_nds}}


