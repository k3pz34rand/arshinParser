import os
import re
import json
from datetime import datetime, timedelta, date
from tkinter import Tk, filedialog, messagebox
from openpyxl import load_workbook, Workbook
from openpyxl.utils import column_index_from_string

# ---- API ФГИС Аршин (файл request.py должен лежать рядом) ----
try:
    from request import search_arshin, ArshinError
    ARSHIN_AVAILABLE = True
except ImportError as _e:
    ARSHIN_AVAILABLE = False
    _ARSHIN_IMPORT_ERROR = str(_e)

# =============================================================================
# НАСТРОЙКИ
# =============================================================================
OUTPUT_SHEET = "Sheet1"     # Имя листа в выходном файле
START_ROW = 998             # С какой строки начать запись
DO_CHECK_ARSHIN = True      # Проверять каждую строку через API Аршина
ARSHIN_COL = "AX"           # Колонка для сообщений о проблемах

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


def _to_date(v):
    """Приводит значение ячейки к date (или None)."""
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip()
    for fmt in ("%d.%m.%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s[:10], fmt).date()
        except ValueError:
            continue
    return None

def _parse_iso_date(s):
    """'2032-01-07T00:00:00Z' -> date(2032,1,7). Также понимает 'dd.mm.yyyy'."""
    if not s:
        return None
    s = str(s).strip().replace("Z", "").split("T")[0].split(" ")[0]
    for fmt in ("%Y-%m-%d", "%d.%m.%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None

# =============================================================================
# ПРОВЕРКА ЧЕРЕЗ API АРШИНА
# =============================================================================
def check_arshin_for_row(ws_out, row, file_name, log):
    """
    Дёргает API для одной строки журнала.
    Параметры запроса: year из J, mitnumber из B, mi_number из F.

    Возвращает:
        ("ok", doc)      — ровно 1 запись
        ("none", None)   — 0 записей
        ("multi", n)     — >1 записей
        ("error", msg)   — ошибка API / не хватает данных
    """
    def val(col):
        return ws_out.cell(row=row,
                           column=column_index_from_string(col)).value

    ax = ws_out.cell(row=row, column=column_index_from_string(ARSHIN_COL))
    ax.number_format = '@'

    fif    = val("B")
    sn     = val("F")
    date_j = val("J")

    d = _to_date(date_j)

    # --- проверка исходных данных ---
    if not fif or not sn or not d:
        ax.value = "недостаточно данных (B/F/J) — ручная проверка"
        log.append({
            "file": file_name, "row": row, "reason": "missing_params",
            "B": fif, "F": sn, "J": date_j,
        })
        return ("error", "missing_params")

    year = d.year

    # --- запрос ---
    try:
        res = search_arshin(
            verification_year=year,
            mitnumber=str(fif).strip(),
            mi_number=str(sn).strip(),
            rows=10,
        )
    except ArshinError as e:
        ax.value = f"ошибка API: {e}"
        log.append({
            "file": file_name, "row": row, "reason": "api_error",
            "error": str(e),
            "params": {"B": fif, "F": sn, "year": year},
        })
        return ("error", str(e))
    except Exception as e:
        ax.value = f"ошибка API: {e}"
        log.append({
            "file": file_name, "row": row, "reason": "unexpected_error",
            "error": str(e),
            "params": {"B": fif, "F": sn, "year": year},
        })
        return ("error", str(e))

    n = res.num_found

    # --- разбор количества ---
    if n == 0:
        ax.value = "запись не найдена — ручная проверка"
        log.append({
            "file": file_name, "row": row, "reason": "not_found",
            "params": {"B": fif, "F": sn, "year": year},
        })
        return ("none", None)

    if n > 1:
        ax.value = f"несколько записей ({n}) — требует ручной проверки"
        log.append({
            "file": file_name, "row": row, "reason": "multiple",
            "count": n,
            "params": {"B": fif, "F": sn, "year": year},
        })
        return ("multi", n)

        # --- ровно одна ---
    ax.value = ""  # очищаем, если что-то было раньше
    doc = res.docs[0]

    # ---- K: дата из поля "Действительна до" (valid_date) ----
    valid_raw = doc.get("valid_date")
    vd = _parse_iso_date(valid_raw)
    if vd:
        cell_k = ws_out.cell(row=row,
                             column=column_index_from_string("K"),
                             value=vd)
        cell_k.number_format = 'dd.mm.yyyy'
    else:
        # в API пусто/непонятный формат — оставляем пусто и логируем
        ws_out.cell(row=row,
                    column=column_index_from_string("K"),
                    value=None)
        log.append({
            "file": file_name, "row": row,
            "reason": "no_valid_date",
            "valid_date_raw": valid_raw,
            "params": {"B": fif, "F": sn, "year": year},
        })

    return ("ok", doc)


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

    if DO_CHECK_ARSHIN and not ARSHIN_AVAILABLE:
        print(f"⚠ Не удалось импортировать arshin.py: {_ARSHIN_IMPORT_ERROR}")
        print("  Проверка через API будет пропущена.")
        do_check = False
    else:
        do_check = DO_CHECK_ARSHIN

    # --- Загружаем журнал ---
    wb_out = load_workbook(output_file)
    if OUTPUT_SHEET in wb_out.sheetnames:
        ws_out = wb_out[OUTPUT_SHEET]
    else:
        ws_out = wb_out.active
        ws_out.title = OUTPUT_SHEET

    current_row = START_ROW
    processed_count = 0
    arshin_log = []

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

            # ---- СТРОКА 47: Дата поверки -> J ----
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
                except Exception as e:
                    print(f"Ошибка разбора даты '{date_str}' в {file_name}: {e}")

        except Exception as e:
            print(f"Ошибка обработки {file_name}: {e}")
        finally:
            wb_src.close()

        # ================= ПРОВЕРКА ЧЕРЕЗ API =================
        if do_check:
            try:
                status, payload = check_arshin_for_row(
                    ws_out, current_row, file_name, arshin_log
                )
                if status == "ok":
                    print(f"    ✓ Аршин: одна запись найдена")
                elif status == "none":
                    print(f"    ✗ Аршин: запись не найдена")
                elif status == "multi":
                    print(f"    ⚠ Аршин: несколько записей ({payload})")
                elif status == "error":
                    print(f"    ! Аршин: ошибка ({payload})")
                # здесь позже можно добавить сверку полей AW/J/K/O/E
            except Exception as e:
                print(f"    ! Сверка упала для {file_name}: {e}")

        current_row += 1
        processed_count += 1
        print(f"  [{processed_count}] {file_name} -> строка {current_row - 1}")

    wb_out.save(output_file)
    print(f"Готово! Обработано файлов: {processed_count}")
    print(f"Результат сохранён в: {output_file}")

    # ================= ОТЧЁТ =================
    if arshin_log:
        report = os.path.join(os.path.dirname(output_file),
                              "arshin_problems.json")
        try:
            with open(report, "w", encoding="utf-8") as f:
                json.dump(arshin_log, f, ensure_ascii=False,
                          indent=2, default=str)
            print(f"⚠ Проблемных строк: {len(arshin_log)} — см. {report}")
        except Exception as e:
            print(f"Не удалось сохранить отчёт: {e}")
    else:
        print("✓ Проблемных строк по Аршину нет")

    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    messagebox.showinfo("Готово",
                        f"Обработано файлов: {processed_count}\n"
                        f"Проблемных строк по Аршину: {len(arshin_log)}\n"
                        f"Результат сохранён в:\n{output_file}")
    root.destroy()


if __name__ == "__main__":
    main()