## Функции

- **Проверка типов учебных материалов** (лекции, практические занятия, лабораторные работы и т.д.).
- **Проверка соответствия типов размещения** (должны быть `file`, `folder` или `page`).
- **Проверка содержимого папок** для убедительности, что они содержат нужные файлы.
- **Формирование отчета в формате Excel** с тремя вкладками:
  - **Отсутствующие материалы** — материалы, которые должны присутствовать, но отсутствуют.
  - **Некорректные материалы** — материалы, которые присутствуют, но не соответствуют требованиям по типу.
  - **Заполнено корректно** — материалы, которые присутствуют и соответствуют требованиям. Для папок также указаны названия файлов внутри.

## Требования

- **Python** версии 3.8 и выше.
- **Библиотеки**:
  - `pandas` — для работы с данными и экспорта в Excel.
  - `openpyxl` — для записи отчетов в Excel.

### Установка библиотек

```bash
pip install pandas openpyxl
```

#### Использование
```python
file_path = "D:\\downloads\\2024-10-25_Материалы в курсах_свод.xlsx" необходимо изменить на свой путь к файлу с выгруженными данными.
```

Функция validate_materials позволяет задавать фильтры по следующим параметрам:

- institute — название института.
- department — название кафедры.
- discipline — название дисциплины.
- semester — семестр (например, I для первого семестра).

Также существует словарь для поиска схожих слов search_keywords.

Запуск скрипта происходит путем открытия .py файла

##### Пример использования
```python
test_report = validate_materials(df_main, discipline="Базы данных для индустриальных задач", semester="I")
```

```python
test_report = validate_materials(df_main, discipline="Алгоритмы машинного обучения для решения прикладных задач", semester="I")
```


