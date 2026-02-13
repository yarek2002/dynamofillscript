# -*- coding: utf-8 -*-
import clr, csv, codecs
import re

clr.AddReference("RevitServices")
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory

doc = DocumentManager.Instance.CurrentDBDocument
path = IN[0]

KEY_NAME = "ADSK_Комплект чертежей"
TARGET_NAME = "Орг.ЗамечаниеКЛисту"

def get_p(el, n):
    if not el:
        return None
    search = n.strip().lower()
    # Проверка параметров экземпляра
    for p in el.Parameters:
        if p.Definition and p.Definition.Name.strip().lower() == search:
            return p
    # Проверка параметров типа
    t = doc.GetElement(el.GetTypeId())
    if t:
        for p in t.Parameters:
            if p.Definition and p.Definition.Name.strip().lower() == search:
                return p
    return None

def extract_sheet_number(col_a):
    """Извлекает номер листа из столбца A формата КУТ03-Р-ПИР-2-26-РД-МГ-100-ОВ01.01.00-ТЛ-0000_И_Р"""
    if not col_a:
        return None
    # Ищем все паттерны: дефис, затем 4 цифры, затем подчеркивание
    # Берем последнее совпадение (номер листа обычно в конце)
    matches = list(re.finditer(r'-(\d{4})_', col_a))
    if matches:
        # Берем последнее совпадение
        return matches[-1].group(1)
    return None

def extract_drawing_set(col_b):
    """Извлекает комплект чертежей из столбца B формата КУТ03-Р-ПИР-2-26-100-ОВ01.01.00"""
    if not col_b:
        return None
    # Ищем паттерн: ОВ + цифры + точка + цифры + точка + цифры в конце строки или перед концом
    match = re.search(r'(ОВ\d+\.\d+\.\d+)', col_b)
    if match:
        return match.group(1)
    return None

try:
    with codecs.open(path, 'r', encoding='utf-8-sig') as f:
        # Читаем CSV полностью
        data_rows = list(csv.reader(f, delimiter=',', quotechar='"'))
except Exception as e:
    OUT = "Ошибка CSV: " + str(e)
else:
    sheets = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sheets).WhereElementIsNotElementType().ToElements()
    tblocks = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsNotElementType().ToElements()
    
    # Собираем рамки в словарь
    tb_map = {}
    for tb in tblocks:
        p_num = tb.LookupParameter("Номер листа") or tb.LookupParameter("Sheet Number")
        if p_num:
            tb_map[p_num.AsString()] = tb
    
    updated, report = 0, []
    # НОВОЕ: словарь для подсчета листов по томам
    volume_counts = {}
    
    TransactionManager.Instance.EnsureInTransaction(doc)
    
    for s in sheets:
        sn = s.SheetNumber
        tb = tb_map.get(sn)
        
        # Ищем параметры (в листе или в рамке)
        pk = get_p(s, KEY_NAME) or get_p(tb, KEY_NAME)
        pt = get_p(s, TARGET_NAME) or get_p(tb, TARGET_NAME)
        p_volume = get_p(s, "ADSK_Штамп Раздел проекта") or get_p(tb, "ADSK_Штамп Раздел проекта")
        
        if not pk or not pt:
            continue
        
        # Ключ из Revit (например, ОВ02.02.02)
        vk = (pk.AsString() or "").strip().lower()
        vk = vk.replace('⠀', '')
        if not vk:
            continue
        
        res = None
        csv_sheet_number = None  # Номер листа из CSV для отладки
        debug_info = ""  # Отладочная информация
        # Убираем все пробелы из номера листа Revit (до и после)
        sn_trimmed = str(sn).strip() if sn else ""
        found_matches = []  # Для отладки: все найденные совпадения
        first_match = None  # Первое совпадение по комплекту чертежей
        
        # Функция нормализации номера листа для сравнения
        def normalize_sheet_number(sheet_num):
            """Нормализует номер листа для сравнения - убирает все пробелы и невидимые символы"""
            if not sheet_num:
                return None
            # Убираем все пробелы (в начале, в конце и внутри)
            normalized = str(sheet_num).replace(' ', '').replace('\t', '').strip()
            # Убираем невидимые символы (zero-width spaces и т.д.) - используем проверку через ord()
            # Оставляем только цифры, точки, дефисы и обычные буквы
            cleaned = ''
            for c in normalized:
                # Проверяем, что символ - это цифра, точка, дефис или обычная буква
                if c.isdigit() or c == '.' or c == '-' or (c.isalpha() and ord(c) < 128):
                    cleaned += c
            normalized = cleaned
            # Если это чисто цифры - убираем ведущие нули
            if normalized.isdigit():
                return str(int(normalized))
            # Если есть точка (например, "21.3"), извлекаем числовую часть
            # Для сравнения: "21.3" может соответствовать "0213" или "2130" в CSV
            # Пока оставляем как есть для точного сравнения
            return normalized
        
        # Нормализуем номер листа из Revit для сравнения
        revit_num_normalized = normalize_sheet_number(sn_trimmed)
        
        # Также создаем версию для поиска в CSV (умножаем на 10, если это просто число)
        # Например, "4" -> ищем "40" или "0040" в CSV
        revit_num_for_csv_search = None
        # Проверяем нормализованное значение, а не исходное (чтобы учесть удаление пробелов)
        if revit_num_normalized and revit_num_normalized.isdigit():
            try:
                revit_int = int(revit_num_normalized)
                revit_num_for_csv_search = str(revit_int * 10).zfill(4)  # "4" -> "0040"
            except (ValueError, TypeError):
                revit_num_for_csv_search = None
        
        for row in data_rows:
            if len(row) < 2:  # Минимум нужны столбцы A и B
                continue
            
            try:
                col_a = str(row[0]).strip()  # Столбец A: Орг.ЗамечаниеКЛисту и номер листа
                col_b = str(row[1]).strip()  # Столбец B: ADSK_Комплект чертежей
                
                if not col_a or not col_b:
                    continue
                
                # Извлекаем номер листа из столбца A
                sheet_num_from_csv = extract_sheet_number(col_a)
                
                # Извлекаем комплект чертежей из столбца B
                drawing_set = extract_drawing_set(col_b)
                if not drawing_set:
                    continue
                
                # Сравниваем с ключом из Revit (без учета регистра)
                if vk == drawing_set.lower():
                    # Сохраняем все найденные совпадения для отладки
                    match_info = {
                        'col_a': col_a[:80],
                        'sheet_num': sheet_num_from_csv
                    }
                    found_matches.append(match_info)
                    
                    # Сохраняем первое совпадение (на случай, если точного совпадения по номеру листа не найдется)
                    if first_match is None:
                        first_match = {
                            'col_a': col_a,
                            'sheet_num': sheet_num_from_csv
                        }
                    
                    # Проверяем совпадение номеров листов
                    if sheet_num_from_csv:
                        csv_num_normalized = normalize_sheet_number(sheet_num_from_csv)
                        # Сравниваем нормализованные значения
                        match_found = False
                        
                        # Вариант 1: если номер Revit - просто число, проверяем умноженный на 10 вариант
                        # Например, "4" в Revit должно соответствовать "0040" в CSV
                        if revit_num_for_csv_search and sheet_num_from_csv:
                            # Сравниваем напрямую с форматом CSV (4 цифры с ведущими нулями)
                            if revit_num_for_csv_search == sheet_num_from_csv:
                                match_found = True
                        
                        # Вариант 2: сравниваем числовые значения (если оба - числа)
                        # Это более надежный способ, так как сравнивает числовые значения
                        if not match_found and revit_num_normalized and csv_num_normalized:
                            if revit_num_normalized.isdigit() and csv_num_normalized.isdigit():
                                try:
                                    revit_int = int(revit_num_normalized)
                                    csv_int = int(csv_num_normalized)
                                    # Проверяем, соответствует ли номер Revit номеру CSV (умноженному на 10)
                                    if csv_int == revit_int * 10:
                                        match_found = True
                                except (ValueError, TypeError):
                                    pass
                        
                        # Вариант 3: прямое сравнение нормализованных значений (для случаев типа "21.3")
                        if not match_found and revit_num_normalized == csv_num_normalized:
                            match_found = True
                        
                        # Если номера листов совпадают - это точное совпадение, берем его
                        if match_found:
                            res = col_a
                            csv_sheet_number = sheet_num_from_csv
                            debug_info = " | Столбец A: '{}' | Точное совпадение по номеру листа! Revit:'{}' (норм:'{}', поиск:'{}') CSV:'{}' (норм:'{}')".format(
                                col_a[:100], sn_trimmed, revit_num_normalized or "N/A", revit_num_for_csv_search or "N/A", sheet_num_from_csv, csv_num_normalized or "N/A")
                            break
            except Exception as e:
                continue
        
        # Если точного совпадения не найдено, берем первое совпадение по комплекту чертежей
        if res is None and first_match is not None:
            res = first_match['col_a']
            csv_sheet_number = first_match['sheet_num']
            debug_info = " | Столбец A: '{}'".format(res[:100])
            if csv_sheet_number:
                csv_num_normalized = str(int(csv_sheet_number)) if csv_sheet_number.isdigit() else csv_sheet_number
                debug_info += " | Извлечено из CSV: '{}' (норм: '{}')".format(csv_sheet_number, csv_num_normalized)
                debug_info += " | Revit: '{}' (норм: '{}', для поиска: '{}') - совпадение только по комплекту. Проверено {} строк".format(
                    sn_trimmed, revit_num_normalized, revit_num_for_csv_search or "N/A", len(found_matches))
            else:
                debug_info += " | Извлечено: НИЧЕГО"
        
        # Добавляем информацию о всех найденных совпадениях
        if found_matches:
            debug_info += " | Всего совпадений по комплекту: {}".format(len(found_matches))
            if len(found_matches) > 1:
                debug_info += " | Номера листов в CSV: {}".format([m['sheet_num'] for m in found_matches])
        
        if res:
            try:
                pt.Set(res)
                updated += 1
                # Формируем информацию о сравнении номеров листов
                sheet_info = "Номер листа Revit: '{}'".format(sn_trimmed)
                if csv_sheet_number:
                    sheet_info += ", CSV: '{}'".format(csv_sheet_number)
                    # Нормализуем для сравнения (убираем ведущие нули)
                    revit_num_normalized = str(int(sn_trimmed)) if sn_trimmed.isdigit() else sn_trimmed
                    csv_num_normalized = str(int(csv_sheet_number)) if csv_sheet_number.isdigit() else csv_sheet_number
                    match_status = "✓ Совпадают" if revit_num_normalized == csv_num_normalized else "✗ НЕ совпадают"
                    sheet_info += " → {}".format(match_status)
                else:
                    sheet_info += ", CSV: не извлечен"
                
                # Добавляем отладочную информацию
                if debug_info:
                    sheet_info += debug_info
                
                report.append("✅ {}: Найдено '{}' | {}".format(sn, res, sheet_info))
                # ✅ СЧИТАЕМ ТОЛЬКО ОБНОВЛЕННЫЕ ЛИСТЫ
                if p_volume:
                    volume_name = p_volume.AsString() or "Без тома"
                    volume_counts[volume_name] = volume_counts.get(volume_name, 0) + 1
            except:
                report.append("⚠️ {}: Ошибка записи в параметр".format(sn))
        else:
            report.append("❓ {}: Ключ '{}' не найден в CSV".format(sn, vk))
    
    TransactionManager.Instance.TransactionTaskDone()
    
    # НОВОЕ: формируем второй вывод
    volume_report = []
    total_sheets = 0
    for volume, count in sorted(volume_counts.items()):
        volume_report.append("В томе {}: {} листов".format(volume, count))
        total_sheets += count
    
    # Добавляем итоговые строки
    volume_report.append("Обновлено листов: {}".format(updated))
    
    # ВОЗВРАЩАЕМ ДВА ЗНАЧЕНИЯ В РАЗНЫЕ ПОРТЫ
    log_text = "\n".join(report)
    OUT = [{"Обновлено": updated, "Лог": log_text}, "\n".join(volume_report)]
