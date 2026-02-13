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
        sn_trimmed = sn.strip()  # Обрезаем пробелы из номера листа Revit
        found_matches = []  # Для отладки: все найденные совпадения
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
                    found_matches.append({
                        'col_a': col_a[:80],
                        'sheet_num': sheet_num_from_csv
                    })
                    
                    # Отладочная информация: показываем что в столбце A и что извлекли
                    debug_info = " | Столбец A: '{}'".format(col_a[:100])  # Первые 100 символов
                    if sheet_num_from_csv:
                        debug_info += " | Извлечено из CSV: '{}'".format(sheet_num_from_csv)
                        # Нормализуем для сравнения
                        revit_num_normalized = str(int(sn_trimmed)) if sn_trimmed.isdigit() else sn_trimmed
                        csv_num_normalized = str(int(sheet_num_from_csv)) if sheet_num_from_csv.isdigit() else sheet_num_from_csv
                        debug_info += " | Revit: '{}' (норм: '{}') vs CSV: '{}' (норм: '{}')".format(
                            sn_trimmed, revit_num_normalized, sheet_num_from_csv, csv_num_normalized)
                    else:
                        debug_info += " | Извлечено: НИЧЕГО"
                    
                    res = col_a  # Орг.ЗамечаниеКЛисту берем из столбца A
                    csv_sheet_number = sheet_num_from_csv  # Сохраняем для отладки
                    break
        
        # Добавляем информацию о всех найденных совпадениях
        if found_matches:
            debug_info += " | Всего совпадений: {}".format(len(found_matches))
            if len(found_matches) > 1:
                debug_info += " | Номера листов: {}".format([m['sheet_num'] for m in found_matches])
            except Exception as e:
                continue
        
        if res:
            try:
                pt.Set(res)
                updated += 1
                # Формируем информацию о сравнении номеров листов
                sn_trimmed = sn.strip()  # Обрезаем пробелы
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
