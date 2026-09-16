import os
import re
from datetime import datetime, timedelta
from tkinter import Tk, filedialog, messagebox
from openpyxl import load_workbook, Workbook
from openpyxl.utils import column_index_from_string

# =============================================================================
# НАСТРОЙКИ
# =============================================================================
OUTPUT_SHEET = "Sheet1"     # Имя листа в выходном файле
START_ROW = 998             # С какой строки начать запись

# =============================================================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# =============================================================================
def get_cell_text(ws, row, col):
    cell = ws.cell(row=row, column=col)
    val = cell.value
    if val is None:
        return ""
    return str(val)


def val_func(s):
    s = str(s).strip().replace('\xa0', ' ')
    match = re.search(r'[-+]?\d*\.?\d+', s)
    if match:
        return float(match.group())
    return 0.0


def choose_folder(title="Выберите папку с протоколами"):
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    folder = filedialog.askdirectory(title=title)
    root.destroy()
    return folder


def choose_output_file(title="Выберите файл журнала"):
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    path = filedialog.askopenfilename(
        title=title,
        filetypes=[("Excel файлы", "*.xlsx"), ("Все файлы", "*.*")]
    )
    root.destroy()
    return path


# =============================================================================
# ОСНОВНАЯ ПРОЦЕДУРА
# =============================================================================
def main():
    # --- Выбор папки с протоколами ---
    folder_path = choose_folder()
    if not folder_path:
        print("Папка не выбрана. Выход.")
        return
    if not os.path.isdir(folder_path):
        print(f"Папка не существует: {folder_path}")
        return

    # --- Выбор файла журнала ---
    output_file = choose_output_file()
    if not output_file:
        print("Файл журнала не выбран. Выход.")
        return
    if not os.path.exists(output_file):
        print(f"Файл не найден: {output_file}")
        return

    print(f"Папка с протоколами: {folder_path}")
    print(f"Журнал: {output_file}")

    # --- Загружаем журнал ---
    wb_out = load_workbook(output_file)
    if OUTPUT_SHEET in wb_out.sheetnames:
        ws_out = wb_out[OUTPUT_SHEET]
    else:
        ws_out = wb_out.active
        ws_out.title = OUTPUT_SHEET

    current_row = START_ROW
    processed_count = 0

    all_files = os.listdir(folder_path)
    files = [f for f in all_files
             if f.lower().endswith(('.xlsx', '.xlsm')) and not f.startswith('~$')]

    print(f"К обработке: {len(files)} файлов")

    for file_name in files:
        file_path = os.path.join(folder_path, file_name)

        if os.path.abspath(file_path) == os.path.abspath(output_file):
            continue

        try:
            wb_src = load_workbook(file_path, read_only=True, data_only=True)
        except Exception as e:
            print(f"Ошибка открытия {file_name}: {e}")
            continue

        ws_src = wb_src.worksheets[0]

        try:
            # ---- СТРОКА 8: Номер протокола -> AW ----
            txt = get_cell_text(ws_src, 8, 1)
            pos = txt.find("№")
            if pos != -1:
                ws_out.cell(row=current_row,
                            column=column_index_from_string("AW"),
                            value=txt[pos + 1:].strip())

            # ---- СТРОКА 9: Модификация -> E ----
            txt = get_cell_text(ws_src, 9, 1)
            pos = txt.find("Мод.:")
            if pos != -1:
                temp = txt[pos + 5:]
                if "," in temp:
                    temp = temp.split(",")[0]
                ws_out.cell(row=current_row,
                            column=column_index_from_string("E"),
                            value=temp.strip())

            # ---- СТРОКА 10: № ФИФ ОЕИ -> B, Заводской № -> F, Год -> H ----
            txt = get_cell_text(ws_src, 10, 1)

            pos = txt.find("№ ФИФ ОЕИ")
            if pos != -1:
                temp = txt[pos + 9:].replace('\xa0', ' ').strip()
                pos2 = temp.find(" ")
                if pos2 != -1:
                    temp = temp[:pos2]
                ws_out.cell(row=current_row,
                            column=column_index_from_string("B"),
                            value=temp.strip())

            pos = txt.find("заводской №")
            if pos != -1:
                temp = txt[pos + 11:].replace('\xa0', ' ').strip()
                num_str = ""
                for ch in temp:
                    if ch.isdigit():
                        num_str += ch
                    elif num_str:
                        break
                ws_out.cell(row=current_row,
                            column=column_index_from_string("F"),
                            value=num_str)

            pos = txt.find("год изготовления:")
            if pos != -1:
                temp = txt[pos + 17:].replace(".", "").strip()
                try:
                    year_val = int(float(temp))
                except Exception:
                    year_val = 0
                ws_out.cell(row=current_row,
                            column=column_index_from_string("H"),
                            value=year_val)

            # ---- СТРОКА 21: Первый заводской номер -> Y ----
            txt = get_cell_text(ws_src, 21, 1)
            pos = txt.find("зав. №")
            if pos != -1:
                temp = txt[pos + 6:].replace('\xa0', ' ').strip()
                if "," in temp:
                    temp = temp.split(",")[0]
                ws_out.cell(row=current_row,
                            column=column_index_from_string("Y"),
                            value=temp.strip())

            # ---- СТРОКИ 27, 28, 29: (C + D) / 2 -> AE, AF, AG ----
            c27 = val_func(ws_src.cell(row=27, column=3).value)
            d27 = val_func(ws_src.cell(row=27, column=4).value)
            cell_ae = ws_out.cell(row=current_row,
                                  column=column_index_from_string("AE"),
                                  value=(c27 + d27) / 2)
            cell_ae.number_format = '0.0"°C"'

            c29 = val_func(ws_src.cell(row=29, column=3).value)
            d29 = val_func(ws_src.cell(row=29, column=4).value)
            cell_af = ws_out.cell(row=current_row,
                                  column=column_index_from_string("AF"),
                                  value=(c29 + d29) / 2)
            cell_af.number_format = '0.0"кПа"'

            c28 = val_func(ws_src.cell(row=28, column=3).value)
            d28 = val_func(ws_src.cell(row=28, column=4).value)
            cell_ag = ws_out.cell(row=current_row,
                                  column=column_index_from_string("AG"),
                                  value=((c28 + d28) / 2) / 100)
            cell_ag.number_format = '0.0%'

            # ---- A46: Результаты поверки -> O ----
            txt = get_cell_text(ws_src, 46, 1)
            ws_out.cell(row=current_row,
                        column=column_index_from_string("O"),
                        value="Да" if "пригоден" in txt.lower() else "Нет")

            # ---- СТРОКА 47: Дата поверки -> J, +6 лет -1 день -> K ----
            txt = get_cell_text(ws_src, 47, 1)
            match = re.search(r'\d{2}\.\d{2}\.\d{4}', txt)
            if match:
                date_str = match.group()
                try:
                    pDate = datetime.strptime(date_str, "%d.%m.%Y").date()
                    cell_j = ws_out.cell(row=current_row,
                                         column=column_index_from_string("J"),
                                         value=pDate)
                    cell_j.number_format = 'dd.mm.yyyy'

                    try:
                        next_date = pDate.replace(year=pDate.year + 6)
                    except ValueError:
                        next_date = pDate.replace(year=pDate.year + 6, day=28)
                    next_date = next_date - timedelta(days=1)

                    cell_k = ws_out.cell(row=current_row,
                                         column=column_index_from_string("K"),
                                         value=next_date)
                    cell_k.number_format = 'dd.mm.yyyy'
                except Exception as e:
                    print(f"Ошибка разбора даты '{date_str}' в {file_name}: {e}")

        except Exception as e:
            print(f"Ошибка обработки {file_name}: {e}")
        finally:
            wb_src.close()

        current_row += 1
        processed_count += 1
        print(f"  [{processed_count}] {file_name} -> строка {current_row - 1}")

    wb_out.save(output_file)
    print(f"Готово! Обработано файлов: {processed_count}")
    print(f"Результат сохранён в: {output_file}")

    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    messagebox.showinfo("Готово",
                        f"Обработано файлов: {processed_count}\n"
                        f"Результат сохранён в:\n{output_file}")
    root.destroy()


if __name__ == "__main__":
    main()