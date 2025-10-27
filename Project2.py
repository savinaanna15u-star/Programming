# Импортируем нужные библиотеки
import pandas as pd  # для работы с таблицами
import matplotlib.pyplot as plt  # для создания графиков

print("=" * 50)
print("АНАЛИЗ РЕЗУЛЬТАТОВ ТЕСТИРОВАНИЯ")
print("=" * 50)

# ==================== ЭТАП 1: ЗАГРУЗКА И ОЧИСТКА ДАННЫХ ====================

print("\n ЭТАП 1: Загрузка и подготовка данных")

# Загружаем данные из Excel файла
file_path = r'C:\Users\ANNA\Desktop\ПРОЕКТ\итог25-11-22.xlsx'
df = pd.read_excel(file_path, sheet_name='Цифровые навыки - 7')

print(f"Загружено строк: {len(df)}")

# Оставляем ТОЛЬКО нужные столбцы: школа (ОУ) и задания M-P
columns_to_keep = [
    'ОУ',  # Столбец E - школы
    'В. 5 /1,00', 'В. 6 /2,00', 'В. 7 /1,00', 'В. 8 /1,00'  # Столбцы M-P
]

# Фильтруем таблицу - оставляем только выбранные столбцы
df = df[columns_to_keep]

# Переименовываем столбцы для удобства
df = df.rename(columns={
    'ОУ': 'Школа',
    'В. 5 /1,00': 'Задание_5',
    'В. 6 /2,00': 'Задание_6',
    'В. 7 /1,00': 'Задание_7',
    'В. 8 /1,00': 'Задание_8'
})

print("Сохранены столбцы: Школа и 4 задания (M-P)")

# Обрабатываем оценки - заменяем ошибки на 0
task_columns = ['Задание_5', 'Задание_6', 'Задание_7', 'Задание_8']
for column in task_columns:
    df[column] = pd.to_numeric(df[column], errors='coerce').fillna(0)

# Считаем общий балл за все задания (M-P)
df['Общий_балл'] = df[task_columns].sum(axis=1)

print(f"После обработки: {len(df)} учащихся")
print(f"Общий балл колеблется от {df['Общий_балл'].min()} до {df['Общий_балл'].max()}")

# Сохраняем очищенные данные в новый файл
output_file = r'C:\Users\ANNA\Desktop\ПРОЕКТ\анализ_результатов.xlsx'
df.to_excel(output_file, sheet_name='Очищенные_данные', index=False)
print(" Данные сохранены в файл: анализ_результатов.xlsx")

# ==================== ЭТАП 2: АНАЛИЗ ПО ШКОЛАМ ====================

print("\n ЭТАП 2: Анализ результатов по школам")

# Группируем данные по школам и считаем статистику
school_stats = df.groupby('Школа').agg({
    'Общий_балл': ['mean', 'count']  # средний балл и количество учащихся
}).round(2)

# Упрощаем названия столбцов
school_stats.columns = ['Средний_балл', 'Количество_учащихся']

# Сортируем школы по среднему баллу (от лучших к худшим)
school_stats = school_stats.sort_values('Средний_балл', ascending=False)

print(f"Проанализировано школ: {len(school_stats)}")
print("\nТоп-5 школ по среднему баллу:")
print(school_stats.head())

# Сохраняем рейтинг школ
with pd.ExcelWriter(output_file, mode='a') as writer:
    school_stats.to_excel(writer, sheet_name='Рейтинг_школ')

print(" Рейтинг школ сохранен")

# ==================== ЭТАП 3: ВИЗУАЛИЗАЦИЯ ====================

print("\n ЭТАП 3: Создание графиков")

# График 1: Рейтинг топ-10 школ
plt.figure(figsize=(12, 6))
top_schools = school_stats.head(10)

# Создаем сокращенные названия для графика
short_names = []
for school in top_schools.index:
    if 'Гимназия' in school:
        num = ''.join(filter(str.isdigit, school))
        short_names.append(f'Гим.{num}' if num else 'Гимназия')
    elif 'СОШ' in school:
        num = ''.join(filter(str.isdigit, school))
        short_names.append(f'Шк.{num}' if num else 'СОШ')
    elif 'Лицей' in school:
        num = ''.join(filter(str.isdigit, school))
        short_names.append(f'Лиц.{num}' if num else 'Лицей')
    else:
        short_names.append(school[:10] + '..' if len(school) > 12 else school)

# Строим столбчатую диаграмму
plt.bar(short_names, top_schools['Средний_балл'], color='skyblue', edgecolor='black')
plt.title('Топ-10 школ по среднему баллу за задания M-P', fontsize=14, fontweight='bold')
plt.xlabel('Школы')
plt.ylabel('Средний балл')
plt.xticks(rotation=45)
plt.grid(axis='y', alpha=0.3)

# Добавляем значения на столбцы
for i, value in enumerate(top_schools['Средний_балл']):
    plt.text(i, value + 0.05, f'{value:.2f}', ha='center', va='bottom')

plt.tight_layout()
plt.show()

# График 2: Распределение баллов всех учащихся
plt.figure(figsize=(10, 6))

# Строим гистограмму
plt.hist(df['Общий_балл'], bins=8, color='lightgreen', edgecolor='black', alpha=0.7)
plt.title('Распределение баллов учащихся за задания M-P', fontsize=14, fontweight='bold')
plt.xlabel('Общий балл (максимум 5 баллов)')
plt.ylabel('Количество учащихся')
plt.grid(axis='y', alpha=0.3)

plt.tight_layout()
plt.show()

# ==================== ИТОГИ ====================

print("\n" + "=" * 50)
print("ИТОГОВАЯ СТАТИСТИКА")
print("=" * 50)

total_students = len(df)
total_schools = df['Школа'].nunique()  # ← ИСПРАВЛЕНО: русская "а"
average_score = df['Общий_балл'].mean()
max_score = df['Общий_балл'].max()
min_score = df['Общий_балл'].min()

print(f"• Всего учащихся: {total_students}")
print(f"• Количество школ: {total_schools}")
print(f"• Средний балл по всем: {average_score:.2f}")
print(f"• Максимальный балл: {max_score}")
print(f"• Минимальный балл: {min_score}")
print(f"• Лучшая школа: {school_stats.index[0]}")
print(f"  - Средний балл: {school_stats.iloc[0]['Средний_балл']:.2f}")
print(f"  - Количество учащихся: {school_stats.iloc[0]['Количество_учащихся']}")

print("\n" + "=" * 50)
print("АНАЛИЗ ЗАВЕРШЕН! ")
print("=" * 50)