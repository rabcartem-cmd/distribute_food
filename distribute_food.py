import pandas as pd
import numpy as np
import math
import sys
import os

# ================= ПРОВЕРКА АРГУМЕНТОВ =================
if len(sys.argv) < 2:
    print("Ошибка: Вы не указали файл для обработки.")
    print("Использование: python distribute_food.py <путь_к_файлу.xlsx>")
    sys.exit(1)

FILE_PATH = sys.argv[1]

if not os.path.exists(FILE_PATH):
    print(f"Ошибка: Файл '{FILE_PATH}' не найден.")
    sys.exit(1)

base_name, ext = os.path.splitext(FILE_PATH)
OUTPUT_PATH = f"{base_name}_out{ext}"

# Исключения: { 'часть_названия_продукта': 'название_семьи' }
EXCEPTIONS = {
    'морква': 'семья 4'
}

# ================= ФУНКЦИИ =================

def classify_product(name):
    """Классификация продуктов для балансировки категорий."""
    name = str(name).lower()
    
    if any(x in name for x in ['балон', 'газ', 'мыла', 'мыцця', 'сродкі', 'пакет', 'смецця', 'губк', 'папер', 'посуд', 'запалк']):
        return 'Непродовольственные товары'
    if any(x in name for x in ['бульба', 'морква', 'цыбуля', 'агурок', 'памідор', 'капуста', 'перац', 'часнок', 'кабачок', 'айсберг', 'зеляніна', 'зелень']): 
        return 'Овощи и зелень'
    if any(x in name for x in ['лімон', 'яблык', 'арэхі', 'разынкі', 'курага', 'чарнасліў', 'фрукт', 'цукаты']): 
        return 'Фрукты и орехи'
    if any(x in name for x in ['курыца', 'каўбаса', 'тушонка', 'сала', 'кілбаскі', 'фарш', 'мяса', 'кіўбаскі']): 
        return 'Мясо'
    if any(x in name for x in ['рыб', 'шпроты', 'тунец', 'скумбрыя', 'кілька', 'сардзіна', 'крабавыя', 'аліўкі']): 
        return 'Рыба и консервы'
    if any(x in name for x in ['сыр', 'малако', 'згушчонка', 'смятана', 'масла', 'тварог', 'яйкі', 'яйка']): 
        return 'Молочные продукты и яйца'
    if any(x in name for x in ['хлеб', 'батон', 'лаваш', 'сушкі', 'сухары', 'печыва', 'пернікі', 'перапечкі']): 
        return 'Хлеб и сладости'
    
    return 'Бакалея и прочее'

def get_target_budget(age):
    """Расчет целевого взноса по возрасту."""
    try:
        age = float(age)
        return 70 if age >= 10 else 35
    except:
        return 0

# ================= ОСНОВНОЙ КОД =================

print(f"Читаем файл: {FILE_PATH}...")

# 1. Загрузка данных о людях (Колонки X, Y)
try:
    df_people = pd.read_excel(FILE_PATH, usecols="X,Y")
    df_people.columns = ['Age', 'Family']
    df_people = df_people.dropna(subset=['Family'])
    df_people['Family'] = df_people['Family'].astype(str).str.strip().str.lower()
except Exception as e:
    print(f"Ошибка при чтении колонок X, Y: {e}")
    sys.exit(1)

# Считаем бюджет и размер каждой семьи
df_people['Contribution'] = df_people['Age'].apply(get_target_budget)
family_budgets = df_people.groupby('Family')['Contribution'].sum().to_dict()
family_sizes = df_people.groupby('Family').size().to_dict()
total_people = sum(family_sizes.values())

# 2. Загрузка данных о продуктах
try:
    df_products = pd.read_excel(FILE_PATH, usecols="H,J,K,N")
    df_products.columns = ['Product', 'Qty', 'Unit', 'Price_per_Unit']
    df_products = df_products.dropna(subset=['Product', 'Qty'])
except Exception as e:
    print(f"Ошибка при чтении колонок H, J, K, N: {e}")
    sys.exit(1)

df_products['Product'] = df_products['Product'].astype(str).str.strip()
df_products['Unit'] = df_products['Unit'].astype(str).str.strip()
df_products['Qty'] = pd.to_numeric(df_products['Qty'], errors='coerce').fillna(0)
df_products['Price_per_Unit'] = pd.to_numeric(df_products['Price_per_Unit'], errors='coerce').fillna(0)
df_products['Category'] = df_products['Product'].apply(classify_product)

# 3. Дробление позиций
print("Классифицируем и дробим продукты...")
split_items = []
for _, row in df_products.iterrows():
    qty = row['Qty']
    if qty > 1:
        whole_parts = math.floor(qty)
        fractional_part = qty - whole_parts
        
        for _ in range(whole_parts - 1):
            split_items.append({
                'Product': row['Product'], 'Category': row['Category'], 
                'Qty': 1, 'Unit': row['Unit'], 
                'Cost': row['Price_per_Unit'] * 1
            })
        
        last_qty = 1 + fractional_part
        split_items.append({
            'Product': row['Product'], 'Category': row['Category'], 
            'Qty': last_qty, 'Unit': row['Unit'], 
            'Cost': row['Price_per_Unit'] * last_qty
        })
    elif qty > 0:
        split_items.append({
            'Product': row['Product'], 'Category': row['Category'], 
            'Qty': qty, 'Unit': row['Unit'], 
            'Cost': row['Price_per_Unit'] * qty
        })

df_items = pd.DataFrame(split_items)
df_items = df_items.sort_values(by='Cost', ascending=False).reset_index(drop=True)

# Считаем ожидаемое количество позиций (пропорционально людям)
total_items_count = len(df_items)
expected_items = {f: max(1, round((family_sizes.get(f, 1) / total_people) * total_items_count, 1)) for f in family_budgets}

# 4. Умное распределение (ПРЕДИКТИВНЫЙ СКОРИНГ)
print("Распределяем позиции по семьям...")
current_spent = {fam: 0.0 for fam in family_budgets.keys()}
family_items = {fam: [] for fam in family_budgets.keys()}
product_counts = {fam: {} for fam in family_budgets.keys()}
category_counts = {fam: {} for fam in family_budgets.keys()}

valid_families = [f for f, tgt in family_budgets.items() if tgt > 0]

def _add_item_to_family(f, item):
    """Служебная функция для добавления продукта в корзину."""
    family_items[f].append(item)
    current_spent[f] += item['Cost']
    
    prod_key = str(item['Product']).lower()
    cat_key = item['Category']
    product_counts[f][prod_key] = product_counts[f].get(prod_key, 0) + 1
    category_counts[f][cat_key] = category_counts[f].get(cat_key, 0) + 1

def candidate_score(f, item):
    """Предиктивная оценка: 'Что будет, если я отдам эту позицию именно этой семье?'"""
    new_budget = current_spent[f] + item['Cost']
    new_items_count = len(family_items[f]) + 1
    
    budget_fill = new_budget / family_budgets[f]
    items_fill = new_items_count / expected_items[f]
    
    imbalance_penalty = abs(budget_fill - items_fill) * 4.0 
    fill_penalty = (budget_fill ** 2) + (items_fill ** 2)
    
    cat_count = category_counts[f].get(item['Category'], 0) + 1
    cat_fill = cat_count / max(1.0, (expected_items[f] * 0.15)) 
    
    return fill_penalty + imbalance_penalty + (cat_fill * 0.5)

def assign_item(item, is_reassignment=False):
    """Основная рекурсивная функция распределения."""
    prod_name = str(item['Product']).lower()
    item_cost = item['Cost']
    
    for exc_keyword, exc_family in EXCEPTIONS.items():
        if exc_keyword in prod_name and exc_family in valid_families:
            _add_item_to_family(exc_family, item)
            return
            
    if not valid_families:
        return

    affordable_families = [f for f in valid_families if current_spent[f] + item_cost <= family_budgets[f]]
    
    if affordable_families:
        min_prod_count = min(product_counts[f].get(prod_name, 0) for f in affordable_families)
        candidates = [f for f in affordable_families if product_counts[f].get(prod_name, 0) == min_prod_count]
        
        assigned_family = min(candidates, key=lambda f: candidate_score(f, item))
        _add_item_to_family(assigned_family, item)
        return

    target_f = min(valid_families, key=lambda f: candidate_score(f, item))
    
    if not is_reassignment:
        rem_budget = family_budgets[target_f] - current_spent[target_f]
        deficit = item_cost - rem_budget
        
        fam_items = sorted(family_items[target_f], key=lambda x: x['Cost'])
        
        removed_items = []
        freed_amount = 0
        items_to_keep = []
        
        for existing_item in fam_items:
            is_exception = any(exc in str(existing_item['Product']).lower() for exc in EXCEPTIONS.keys())
            
            if freed_amount < deficit and not is_exception:
                removed_items.append(existing_item)
                freed_amount += existing_item['Cost']
            else:
                items_to_keep.append(existing_item)
        
        if removed_items:
            family_items[target_f] = items_to_keep
            current_spent[target_f] -= freed_amount
            for r_item in removed_items:
                product_counts[target_f][str(r_item['Product']).lower()] -= 1
                category_counts[target_f][r_item['Category']] -= 1
            
            _add_item_to_family(target_f, item)
            
            removed_items.sort(key=lambda x: x['Cost'], reverse=True)
            for r_item in removed_items:
                assign_item(r_item, is_reassignment=True)
            return

    _add_item_to_family(target_f, item)

for _, item in df_items.iterrows():
    assign_item(item.to_dict())

# 5. Сохранение результатов в Excel
print(f"Группируем позиции и сохраняем результат в {OUTPUT_PATH}...")
with pd.ExcelWriter(OUTPUT_PATH, engine='openpyxl') as writer:
    summary_data = []
    
    for fam in family_budgets.keys():
        if not family_items[fam]:
            continue
            
        df_fam = pd.DataFrame(family_items[fam])
        df_fam_grouped = df_fam.groupby(['Product', 'Category', 'Unit'], as_index=False).agg({
            'Qty': 'sum', 
            'Cost': 'sum'
        })
        
        sheet_name = fam[:31]
        df_fam_grouped.to_excel(writer, sheet_name=sheet_name, index=False)
        
        actual_items_count = len(df_fam)
        
        summary_data.append({
            'Семья': fam,
            'Кол-во человек': family_sizes.get(fam, 0),
            'Целевой бюджет (руб)': family_budgets[fam],
            'Реально закуплено (руб)': current_spent[fam],
            'Разница (Сдача/Долг)': family_budgets[fam] - current_spent[fam],
            '% выполнения бюджета': round((current_spent[fam] / family_budgets[fam]) * 100, 2) if family_budgets[fam] > 0 else 0,
            'Ожидаемое кол-во позиций': expected_items.get(fam, 0),
            'Реальное кол-во позиций': actual_items_count,
            'Уникальных строк в списке': len(df_fam_grouped)
        })
        
    df_summary = pd.DataFrame(summary_data)
    df_summary.to_excel(writer, sheet_name='Общая_Сводка', index=False)

print("Успешно завершено!")