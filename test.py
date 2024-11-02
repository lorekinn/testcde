import pandas as pd
from datetime import datetime

file_path = "D:\\downloads\\2024-10-25_Материалы в курсах_свод.xlsx"
df_main = pd.read_excel(file_path, sheet_name='не скрыты')

def validate_materials(data, institute=None, department=None, discipline=None, semester=None):
    """
    Проверка наличия и корректного размещения материалов по учебной дисциплине.
    - data (DataFrame): Данные по курсам.
    - institute (str): Институт для фильтрации.
    - department (str): Кафедра для фильтрации.
    - discipline (str): Дисциплина для фильтрации.
    - semester (str): Семестр для фильтрации.

    Возвращает:
    - report (dict): Словарь с результатами проверки (отсутствующие и некорректные элементы).
    """
    filtered_data = data
    if institute:
        filtered_data = filtered_data[filtered_data['Категория'].str.contains(institute, case=False, na=False)]
    if department:
        filtered_data = filtered_data[filtered_data['Категория'].str.contains(department, case=False, na=False)]
    if discipline:
        filtered_data = filtered_data[filtered_data['Курс'].str.contains(discipline, case=False, na=False)]
    if semester:
        semester_pattern = rf'\[{semester}\.'
        filtered_data = filtered_data[filtered_data['Курс'].str.contains(semester_pattern, case=False, na=False)]

    material_requirements = {
        "Лекция": ["file", "folder", "page"],
        "Практические": ["file", "folder", "page"],
        "Лабораторная": ["file", "folder", "page"],
        "Текущий контроль": ["assign"],
        "Самостоятельная": ["assign"],
        "Источники": ["url", "external_resource", "file", "folder"],
        "Тест": ["quiz"],
        "Курсовой проект": ["choice"],
        "Практика": ["file", "folder", "page"]
    }
    
    report = {"missing": [], "incorrect": [], "filled_correctly": []}

    search_keywords = {
        "Лекция": ["лекция", "лекционные материалы", "лекция №"],
        "Практические": ["практические", "практическое занятие"],
        "Лабораторная": ["лабораторная"],
        "Текущий контроль": ["текущий контроль"],
        "Самостоятельная": ["самостоятельная"],
        "Источники": ["источники", "литература"],
        "Тест": ["тест"],
        "Курсовой проект": ["курсовой проект"],
        "Практика": ["практика"]
    }

    for req_type, allowed_elements in material_requirements.items():
        keywords = "|".join(search_keywords[req_type])
        req_data = filtered_data[filtered_data['Название'].str.contains(keywords, case=False, na=False)]

        if req_data.empty:
            report["missing"].append({"Дисциплина": discipline, "Материал": req_type})
        else:
            for index, row in req_data.iterrows():
                element_type = row['Элемент'].lower()
                element_name = row['Название']

                if element_type not in allowed_elements:
                    report["incorrect"].append({"Дисциплина": discipline, "Материал": req_type, "Элемент": element_name, "Тип": element_type})
                else:
                    if element_type == "folder":
                        files_in_folder = row.get("Файл", "-")
                        report["filled_correctly"].append({
                            "Дисциплина": discipline, 
                            "Материал": req_type, 
                            "Элемент": element_name, 
                            "Тип": element_type, 
                            "Файлы": files_in_folder if files_in_folder != "-" else "Нет файлов"
                        })
                    else:
                        report["filled_correctly"].append({
                            "Дисциплина": discipline, 
                            "Материал": req_type, 
                            "Элемент": element_name, 
                            "Тип": element_type,
                            "Файлы": "N/A" 
                        })

    return report

def export_report_to_excel(report):
    current_datetime = datetime.now().strftime("%Y-%m-%d_%H-%M")
    output_path = f"D:\\Downloads\\Отчет_проверки_материалов_{current_datetime}.xlsx"
    
    missing_df = pd.DataFrame(report["missing"])
    incorrect_df = pd.DataFrame(report["incorrect"])
    filled_correctly_df = pd.DataFrame(report["filled_correctly"])

    with pd.ExcelWriter(output_path) as writer:
        missing_df.to_excel(writer, sheet_name="Отсутствующие материалы", index=False)
        incorrect_df.to_excel(writer, sheet_name="Некорректные материалы", index=False)
        filled_correctly_df.to_excel(writer, sheet_name="Заполнено корректно", index=False)
    
    print(f"Отчет успешно сохранен в {output_path}")

test_report = validate_materials(df_main, discipline="Алгоритмы машинного обучения для решения прикладных задач", semester="I")
export_report_to_excel(test_report)
