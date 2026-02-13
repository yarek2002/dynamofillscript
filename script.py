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
    # Ищем паттерн: дефис, затем 4 цифры, затем подчеркивание
    match = re.search(r'-(\d{4})_', col_a)
    if match:
        return match.group(1)
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

def normalize_sheet_number(sheet_num):
    """Нормализует номер листа для сравнения (убирает ведущие нули)"""
    if not sheet_num:
        return None
    try:
        # Преобразуем в число и обратно в строку, чтобы убрать ведущие нули
        return str(int(str(sheet_num).strip()))
    except (ValueError, AttributeError):
        # Если не число, возвращаем как есть
        return str(sheet_num).strip()

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
    
    # Создаем словарь для быстрого поиска листов по комплекту чертежей и номеру листа
    sheets_dict = {}
    for s in sheets:
        sn = s.SheetNumber
        tb = tb_map.get(sn)
        
        # Ищем параметр комплекта чертежей (в листе или в рамке)
        pk = get_p(s, KEY_NAME) or get_p(tb, KEY_NAME)
        if pk:
            vk = (pk.AsString() or "").strip().lower()
            vk = vk.replace('⠀', '')
            if vk:
                # Нормализуем номер листа для ключа
                normalized_sn = normalize_sheet_number(sn)
                key = (vk, normalized_sn)
                if key not in sheets_dict:
                    sheets_dict[key] = []
                sheets_dict[key].append((s, tb))
    
    TransactionManager.Instance.EnsureInTransaction(doc)
    
    # Идем по CSV и ищем соответствующие листы
    for row in data_rows:
        if len(row) < 2:  # Минимум нужны столбцы A и B
            continue
        
        try:
            col_a = str(row[0]).strip()  # Столбец A: Орг.ЗамечаниеКЛисту и номер листа
            col_b = str(row[1]).strip()  # Столбец B: ADSK_Комплект чертежей
            
            if not col_a or not col_b:
                continue
            
            # Извлекаем номер листа из столбца A (например, "0120")
            csv_sheet_num = extract_sheet_number(col_a)
            csv_sheet_num_normalized = normalize_sheet_number(csv_sheet_num) if csv_sheet_num else None
            
            # Извлекаем комплект чертежей из столбца B (например, "ОВ01.01.00")
            drawing_set = extract_drawing_set(col_b)
            if not drawing_set or not csv_sheet_num_normalized:
                continue
            
            # Ищем листы с соответствующими параметрами
            search_key = (drawing_set.lower(), csv_sheet_num_normalized)
            matching_sheets = sheets_dict.get(search_key, [])
            
            for s, tb in matching_sheets:
                sn = s.SheetNumber
                pt = get_p(s, TARGET_NAME) or get_p(tb, TARGET_NAME)
                p_volume = get_p(s, "ADSK_Штамп Раздел проекта") or get_p(tb, "ADSK_Штамп Раздел проекта")
                
                if pt:
                    try:
                        pt.Set(col_a)  # Записываем значение из столбца A
                        updated += 1
                        report.append("✅ {}: Найдено '{}'".format(sn, col_a))
                        # ✅ СЧИТАЕМ ТОЛЬКО ОБНОВЛЕННЫЕ ЛИСТЫ
                        if p_volume:
                            volume_name = p_volume.AsString() or "Без тома"
                            volume_counts[volume_name] = volume_counts.get(volume_name, 0) + 1
                    except:
                        report.append("⚠️ {}: Ошибка записи в параметр".format(sn))
        except Exception as e:
            continue
    
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
