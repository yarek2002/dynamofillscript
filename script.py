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
        
        # Нормализуем номер листа из Revit для сравнения
        revit_sheet_num = normalize_sheet_number(sn)
        
        res = None
        comparison_info = []  # Для сбора информации о сравнениях
        
        for row in data_rows:
            if len(row) < 2:  # Минимум нужны столбцы A и B
                continue
            
            try:
                col_a = str(row[0]).strip()  # Столбец A: Орг.ЗамечаниеКЛисту и номер листа
                col_b = str(row[1]).strip()  # Столбец B: ADSK_Комплект чертежей
                
                if not col_a or not col_b:
                    continue
                
                # Извлекаем номер листа из столбца A
                csv_sheet_num = extract_sheet_number(col_a)
                csv_sheet_num_normalized = normalize_sheet_number(csv_sheet_num) if csv_sheet_num else None
                
                # Извлекаем комплект чертежей из столбца B
                drawing_set = extract_drawing_set(col_b)
                if not drawing_set:
                    continue
                
                # Сравниваем и комплект чертежей, и номер листа
                drawing_set_match = (vk == drawing_set.lower())
                sheet_num_match = (revit_sheet_num == csv_sheet_num_normalized) if csv_sheet_num_normalized else False
                
                # Сохраняем информацию о сравнении для отчета
                comp_msg = "  → CSV: Комплект='{}', Номер листа CSV='{}' (норм.='{}') vs Revit Номер='{}' (норм.='{}')".format(
                    drawing_set, csv_sheet_num or "не найден", csv_sheet_num_normalized or "N/A",
                    sn, revit_sheet_num or "N/A"
                )
                comp_msg += " | Комплект: {}, Номер: {}".format("✓" if drawing_set_match else "✗", "✓" if sheet_num_match else "✗")
                comparison_info.append(comp_msg)
                
                if drawing_set_match and sheet_num_match:
                    res = col_a  # Орг.ЗамечаниеКЛисту берем из столбца A
                    break
            except Exception as e:
                continue
        
        if res:
            try:
                pt.Set(res)
                updated += 1
                report.append("✅ {}: Найдено '{}'".format(sn, res))
                # Добавляем информацию о сравнении для успешных случаев
                if comparison_info:
                    report.append("   Сравнение: {}".format(comparison_info[-1] if comparison_info else ""))
                # ✅ СЧИТАЕМ ТОЛЬКО ОБНОВЛЕННЫЕ ЛИСТЫ
                if p_volume:
                    volume_name = p_volume.AsString() or "Без тома"
                    volume_counts[volume_name] = volume_counts.get(volume_name, 0) + 1
            except:
                report.append("⚠️ {}: Ошибка записи в параметр".format(sn))
        else:
            report.append("❓ {}: Ключ '{}' не найден в CSV".format(sn, vk))
            # Добавляем информацию о сравнениях для неудачных случаев (первые 3 попытки)
            if comparison_info:
                report.append("   Попытки сравнения:")
                for comp_msg in comparison_info[:3]:  # Показываем первые 3 попытки
                    report.append(comp_msg)
                if len(comparison_info) > 3:
                    report.append("   ... и еще {} попыток".format(len(comparison_info) - 3))
    
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
