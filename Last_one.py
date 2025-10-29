# Импорт необходимых библиотек
import pandas as pd # pandas - для работы с табличными данными (как Excel)
import matplotlib.pyplot as plt # matplotlib - для построения графиков и диаграмм

file_path = r'C:\Users\ANNA\Desktop\ПРОЕКТ\итог25-11-22.xlsx'
sheet_name = 'Цифровые навыки - 7'

df = pd.read_excel(file_path, sheet_name=sheet_name) # Загружаем данные из Excel файла в переменную df (DataFrame - таблица данных)

columns_to_keep = ['ОУ', 'В. 5 /1,00', 'В. 6 /2,00', 'В. 7 /1,00', 'В. 8 /1,00']
df = df[columns_to_keep]  # Оставляем только выбранные столбцы

# Переименовываем столбцы
df = df.rename(columns={
    'ОУ': 'Школа',
    'В. 5 /1,00': 'Задание_5',
    'В. 6 /2,00': 'Задание_6',
    'В. 7 /1,00': 'Задание_7',
    'В. 8 /1,00': 'Задание_8'
})
# Обрабатываем столбцы с баллами: преобразуем в числа и заменяем ошибки на 0
for column in ['Задание_5', 'Задание_6', 'Задание_7', 'Задание_8']:
    # pd.to_numeric - преобразует в числа
    # errors='coerce' - текст и ошибки превращает в NaN (не число)
    df[column] = pd.to_numeric(df[column], errors='coerce').fillna(0)  # fillna(0) - заменяет все NaN на 0

# Сохраняем очищенные данные обратно в Excel файл на новый лист
with pd.ExcelWriter(file_path, mode='a', if_sheet_exists='replace') as writer:
    df.to_excel(writer, sheet_name='Лист2', index=False)  # index=False - не сохранять номера строк

# Создаем новый столбец с суммой баллов за все задания для каждого ученика
df['Сумма_баллов'] = df['Задание_5'] + df['Задание_6'] + df['Задание_7'] + df['Задание_8']

# Подготовка данных для построения гистограммы
values = df['Сумма_баллов']  # Берем все суммарные баллы

# Создаем гистограмму распределения баллов
plt.figure(figsize=(15, 6))
counts, bins, patches = plt.hist(values, bins=15, edgecolor='black')

# Добавляем подписи количества наблюдений над каждым столбцом гистограммы
for count, x in zip(counts, bins):
    plt.text(x + 0.05, count, str(int(count)), ha='center', va='bottom')

# Настраиваем внешний вид гистограммы
plt.title('Гистограмма распределения баллов (max балл=5)')  # Заголовок
plt.xlabel('баллы')  # Подпись оси X
plt.ylabel('кол-во школьников')  # Подпись оси Y
plt.grid(axis='x')  # Включаем сетку по оси X
plt.grid(axis='y', alpha=0)  # Отключаем видимую сетку по оси Y

# Показываем гистограмму
plt.show()

# Фильтруем данные: убираем учеников с нулевой суммой баллов
df = df[df['Сумма_баллов'] != 0]

# Группируем данные по школам для дальнейшего анализа
grouped = df.groupby('Школа')

# Создаем таблицу с результатами по школам
result = pd.DataFrame({
    'Учреждения': grouped.groups.keys(),  # Названия школ
    'среднее': grouped['Сумма_баллов'].mean(),  # Средний балл по школе
    'кол': grouped['Сумма_баллов'].count()  # Количество учеников в школе
})

# Сортируем школы по среднему баллу в порядке убывания (от лучшей к худшей)
result = result.sort_values(by='среднее', ascending=False)

# Сохраняем рейтинг школ в Excel файл
with pd.ExcelWriter(file_path, mode='a', if_sheet_exists='replace') as writer:
    result.to_excel(writer, sheet_name='Лист3', index=False)

# Загружаем сохраненные данные рейтинга для построения графика
data = pd.read_excel(file_path, sheet_name='Лист3')

# Подготавливаем данные для столбчатой диаграммы
categories = data['Учреждения']  # Названия школ для оси X
values = data['среднее']  # Средние баллы для оси Y

# создаем сокращенные названия для графика
short_names = []
for school in categories:
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

# Создаем столбчатую диаграмму рейтинга школ
plt.figure(figsize=(15, 6))  # Задаем размер графика
plt.bar(short_names, values, color='skyblue')

# Настраиваем внешний вид диаграммы
plt.title('Рейтинг школ города по результатам тестирования')
plt.xlabel('Школы')
plt.ylabel('Баллы')
plt.xticks(rotation=60)
plt.grid(axis='y')

plt.tight_layout()
plt.show()