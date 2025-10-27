import pandas as pd
import matplotlib.pyplot as plt

file_path = r'C:\Users\ANNA\Desktop\ПРОЕКТ\итог25-11-22.xlsx'
sheet_name = 'Цифровые навыки - 7'
df = pd.read_excel(file_path, sheet_name=sheet_name)

columns_to_keep = ['ОУ', 'В. 5 /1,00', 'В. 6 /2,00', 'В. 7 /1,00', 'В. 8 /1,00']
df = df[columns_to_keep]

df = df.rename(columns={
    'ОУ': 'Школа',
    'В. 5 /1,00': 'Задание_5',
    'В. 6 /2,00': 'Задание_6',
    'В. 7 /1,00': 'Задание_7',
    'В. 8 /1,00': 'Задание_8'
})

for column in ['Задание_5', 'Задание_6', 'Задание_7', 'Задание_8']:
    df[column] = pd.to_numeric(df[column], errors='coerce').fillna(0)

with pd.ExcelWriter(file_path, mode='a', if_sheet_exists='replace') as writer:
    df.to_excel(writer, sheet_name='Лист2', index=False)

df['Сумма_баллов'] = df['Задание_5'] + df['Задание_6'] + df['Задание_7'] + df['Задание_8']

values = df['Сумма_баллов']
plt.figure(figsize=(15, 6))
counts, bins, patches = plt.hist(values, bins=15, edgecolor='black', color='skyblue', alpha=0.7)

for count, x in zip(counts, bins):
    if count > 0:
        plt.text(x + 0.1, count + 5, str(int(count)), ha='center', va='bottom')

plt.title('Гистограмма распределения суммы баллов по заданиям M-P (max балл=5)')
plt.xlabel('Сумма баллов')
plt.ylabel('Количество школьников')
plt.grid(axis='x', alpha=0.3)
plt.grid(axis='y', alpha=0.3)
plt.tight_layout()
plt.show()

df = df[df['Сумма_баллов'] != 0]

grouped = df.groupby('Школа')

result = pd.DataFrame({
    'Учреждения': grouped.groups.keys(),
    'среднее': grouped['Сумма_баллов'].mean(),
    'кол': grouped['Сумма_баллов'].count()
})

result = result.sort_values(by='среднее', ascending=False)

with pd.ExcelWriter(file_path, mode='a', if_sheet_exists='replace') as writer:
    result.to_excel(writer, sheet_name='Лист3', index=False)

plt.figure(figsize=(16, 8))
plt.bar(result['Учреждения'], result['среднее'], color='skyblue')
plt.title('Рейтинг школ по среднему баллу')
plt.xlabel('Школы')
plt.ylabel('Средний балл')
plt.xticks(rotation=75, ha='right', fontsize=8)
plt.grid(axis='y', alpha=0.3)

for i, v in enumerate(result['среднее']):
    plt.text(i, v + 0.02, f'{v:.2f}', ha='center', va='bottom', fontsize=7)

plt.tight_layout()
plt.show()