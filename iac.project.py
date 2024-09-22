import os
import pandas as pd
from datetime import datetime

# Получаем текущую дату
today = datetime.today()
current_day = today.strftime("%d.%B").lower()
current_month = today.strftime("%m %B").lower()
current_time = datetime.now().time()

# Путь до папки с планом производства 
folder_path_base = "/Users/piratejet/Documents/IAC.SERVER/Logistics/02 - PLANY VYROBY/Plany vyroby IM/2024"

# Строим полный путь до папки с текущим месяцем
month_folder_path = os.path.join(folder_path_base, current_month)

# Проверяем, существует ли папка с текущим месяцем
file_name = None
if os.path.exists(month_folder_path):
    # Ищем файл с текущим днём
    for file in os.listdir(month_folder_path):
        if current_day in file.lower() and file.endswith('.xlsx'):
            file_name = file
            break
else:
    print(f"Subor {current_month} neexistuje.")
    exit()

# Проверяем существует ли файл с планом на день
if file_name:
    # Путь к найденному файлу
    file_path = os.path.join(month_folder_path, file_name)
    
    # Читаем данные из файла
    data = pd.read_excel(file_path, sheet_name='INJECTION MOULDING')
else:
    print(f"Plan na {current_day} ešte neexistuje")

# Выбор смены (утренняя, вечерняя, ночная)
shift = input("Aka smena? (ran, pob, noc): ").strip().lower()

# Функция для извлечения деталей для выбранной смены
def get_parts_for_shift(data, shift):
    # Определяем колонку в зависимости от смены
    if shift == 'ran':
        shift_col = 12  # Колонка M
    elif shift == 'pob':
        shift_col = 15  # Колонка P
    elif shift == 'noc':
        shift_col = 18  # Колонка S
    else:
        print("Nespravna smena")
        return []
    
    # Фильтруем детали
    parts_indices = []
    for i in range(7, 244):  # Строки с 7 до 243
        if not pd.isna(data.iloc[i, shift_col]):  # Проверяем, что колонка смены не пуста
            part = data.iloc[i, 3]  # Колонка с деталями (D)
            if not pd.isna(part):  # Если деталь не пуста
                parts_indices.append(i)
    
    return parts_indices

# Функция для получения проектов и деталей
def get_project_with_parts(data, parts_for_shift_indices):
    project_with_parts = {}
    current_project = None
    
    for i in range(7, 244):
        project = data.iloc[i, 1]  # Колонка с проектами (B)
        if not pd.isna(project):  # Если название проекта есть, обновляем его
            current_project = project

        # Проверяем, что детали находятся в текущем проекте и он не пустой
        if i in parts_for_shift_indices and current_project:
            part = data.iloc[i, 3]  # Колонка с деталями (D)
            if current_project not in project_with_parts:
                project_with_parts[current_project] = []
            project_with_parts[current_project].append(part)

    return project_with_parts

# Извлекаем список индексов строк для выбранной смены
parts_for_shift_indices = get_parts_for_shift(data, shift)

# Получаем проекты с их деталями
project_parts = get_project_with_parts(data, parts_for_shift_indices)

# Выводим список проектов и деталей
for project, parts in project_parts.items():
    print(f"{project}:")
    for part in parts:
        print(f"  {part}")
    print('')
