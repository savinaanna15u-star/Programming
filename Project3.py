import pandas as pd
import matplotlib.pyplot as plt

print("АНАЛИЗ РЕЗУЛЬТАТОВ ТЕСТИРОВАНИЯ")

# ==================== ЭТАП 1: ЗАГРУЗКА ДАННЫХ ====================

# Загружаем и фильтруем данные
df = pd.read_excel(r'C:\Users\ANNA\Desktop\ПРОЕКТ\итог25-11-22.xlsx',
                   sheet_name='Цифровые навыки - 7')

# Оставляем только школу и задания M-P
df = df[['ОУ', 'В. 5 /1,00', 'В. 6 /2,00', 'В. 7 /1,00', 'В. 8 /1,00']]
df = df.rename(columns={'ОУ': 'Школа'})

# Обрабатываем оценки и считаем общий балл
for col in df.columns[1:]:
    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

df['Общий_балл'] = df.iloc[:, 1:].sum(axis=1)

print(f"Учащихся: {len(df)}, Школ: {df['Школа'].nunique()}")

# ==================== ЭТАП 2: АНАЛИЗ ШКОЛ ====================

# Группируем по школам
school_stats = df.groupby('Школа').agg(
    Средний_балл=('Общий_балл', 'mean'),
    Учащихся=('Общий_балл', 'count')
).round(2).sort_values('Средний_балл', ascending=False)

print("\nТоп-5 школ:")
print(school_stats.head())

# Сохраняем результаты
with pd.ExcelWriter(r'C:\Users\ANNA\Desktop\ПРОЕКТ\анализ_результатов.xlsx') as writer:
    df.to_excel(writer, sheet_name='Данные', index=False)
    school_stats.to_excel(writer, sheet_name='Рейтинг_школ')

# ==================== ЭТАП 3: ГРАФИКИ ====================

# График 1: Топ-10 школ
plt.figure(figsize=(12, 6))
top_schools = school_stats.head(10)

# Сокращаем названия
short_names = [school[:8] + '..' if len(school) > 10 else school
               for school in top_schools.index]

plt.bar(short_names, top_schools['Средний_балл'], color='skyblue')
plt.title('Топ-10 школ по среднему баллу')
plt.xticks(rotation=45)
plt.ylabel('Средний балл')

for i, v in enumerate(top_schools['Средний_балл']):
    plt.text(i, v + 0.05, f'{v:.2f}', ha='center')

plt.tight_layout()
plt.show()

# График 2: Распределение баллов
plt.figure(figsize=(10, 6))
plt.hist(df['Общий_балл'], bins=8, color='lightgreen', alpha=0.7)
plt.title('Распределение баллов учащихся')
plt.xlabel('Баллы')
plt.ylabel('Количество учащихся')
plt.show()

print("\nАнализ завершен!")