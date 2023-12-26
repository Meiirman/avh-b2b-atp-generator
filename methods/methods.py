import atexit
import json
import os
import pathlib
import tkinter as tk
import traceback
from tkinter import filedialog, messagebox

from models.model import AutoClosingWindow

from . import atp, excel_generator, html_generator


def get_value(parameter):
    try:
        with open("settings/config.json", "r", encoding="utf-8") as f:
            data = json.load(f)
            # print(data)
            # print(data[parameter])
            return data[parameter]
    except:
        traceback.print_exc()
        return False


def send_message(message, message_type, out_of_queue=False):
    print(message)
    if out_of_queue:
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("АТП Генератор", message)
    elif get_value(message_type):
        root = tk.Tk()
        root.withdraw()
        messagebox.showinfo("АТП Генератор", message)
    else:
        pass


def send_closing_notification(message, message_type="show_info", out_of_queue=False):
    print(message)
    timeout_seconds = 2
    if out_of_queue:
        root = tk.Tk()
        app = AutoClosingWindow(root, timeout_seconds, message)
        root.mainloop()
    elif get_value(message_type):
        root = tk.Tk()
        app = AutoClosingWindow(root, timeout_seconds, message)
        root.mainloop()

    else:
        root = tk.Tk()
        timeout_seconds = 5
        app = AutoClosingWindow(root, timeout_seconds, message)
        root.mainloop()


def browse_folder(entry_var: tk.StringVar) -> None:
    folder_selected = filedialog.askdirectory()
    entry_var.set(folder_selected)
    set_work_folder(folder_selected)


def set_work_folder(folder_path: str):
    config_path = "settings/config.json"

    try:
        with open(config_path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        data = {}

    data["folder_path"] = folder_path

    with open(config_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    send_message(
        'Новое местоположение рабочей папки: "' + folder_path + '"', "show_info"
    )


def generate_b2b_excel():
    source_path = get_value("prices_list_path")
    work_folder = get_value("folder_path")

    if get_value("default_prices_list_path_in_folder_path"):
        try:
            files = os.listdir(work_folder)
            excel_files = [file for file in files if file.lower().endswith(".xlsx")]
            if excel_files:
                proposal_path = None
                for file in excel_files:
                    if file.lower().endswith(".xlsx") and "тцп" in file.lower():
                        source_path = work_folder + "/" + file
            else:
                send_message("В директории нет файлов Excel (.xlsx)", "show_info")
                return 0
            
        except PermissionError:
            send_message("Ошибка доступа к директории", "show_info")
            return 0
    else:
        source_path = get_value("prices_list_path")
    

    try:
        files = os.listdir(work_folder)
        excel_files = [file for file in files if file.lower().endswith(".xlsx")]
        if excel_files:
            proposal_path = None
            for file in excel_files:
                if file.lower().endswith(".xlsx") and "crq" in file.lower() and "заяв" in file.lower():
                    proposal_path = work_folder + "/" + file
        else:
            send_message("В директории нет файлов Excel (.xlsx)", "show_info")
            return 0
    except PermissionError:
        send_message("Ошибка доступа к директории", "show_info")
        return 0


    if proposal_path == None:
        send_message("В директории нет Заявки (.xlsx)", "show_info")
        return 0
    
    data = excel_generator.get_data(
        source_path=source_path, work_folder=work_folder, proposal_path=proposal_path
    )

    send_message(message=data["message"], message_type="show_info")

    if data["data"]:
        generate_message: dict = atp.generate(
            data,
            template_path="templates/b2b_template.xlsx",
            output_folder_path=work_folder,
        )
        send_message(generate_message["message"], message_type="show_info", out_of_queue=True)

    pass


def generate_b2b_html():
    # РАБОЧАЯ ПАПКА
    work_folder = get_value("folder_path")


    files = os.listdir(work_folder)

    html_file_path = None
    try:
        if files:
            for file in files:
                if file.endswith((".html")):
                    html_file_path = work_folder + "/" + file
            if html_file_path == None:
                send_message("В рабочей папке нет файлов HTML", "show_info")
                return 0
        else:
            send_message("В рабочей папке нет файлов", "show_info")
            return 0
    except PermissionError:
        send_message("Ошибка доступа к директории", "show_info")
        return 0

    proposal_path = None
    try:
        excel_files = [file for file in files if file.lower().endswith(".xlsx")]
        if excel_files:
            for file in excel_files:
                if file.lower().endswith(".xlsx") and "crq" in file.lower() and "заяв" in file.lower():
                    proposal_path = work_folder + "/" + file
        else:
            send_message("В директории нет файлов Excel (.xlsx)", "show_info")
            return 0
    except PermissionError:
        send_message("Ошибка доступа к директории", "show_info")
        return 0
    
    if proposal_path == None:
        send_message("В директории нет Заявки (.xlsx)", "show_info")
        return 0

    data = html_generator.get_data(source_path=html_file_path, work_folder=work_folder, proposal_path=proposal_path)
    print(data)


def change_excel_path(entry_var: tk.StringVar) -> None:
    excel_file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx;*.xls")]
    )

    # Получить относительный путь
    relative_path = (
        pathlib.Path(excel_file_path).relative_to(pathlib.Path.cwd()).as_posix()
    )
    # relative_path = os.path.relpath(excel_file_path, start=os.getcwd())

    entry_var.set(relative_path)
    excel_path = relative_path

    config_path = "settings/config.json"

    try:
        with open(config_path, "r") as f:
            data = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        data = {}

    data["prices_list_path"] = excel_path

    with open(config_path, "w") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    send_message(
        'Новое местоположение Excel файла: "' + excel_path + '"',
        "show_info",
        message_type="show_info",
    )
