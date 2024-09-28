import os
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
from datetime import datetime

# Ваши статусы
statuses = {
    "1": "В обработке", "2": "Подтвержден, ожидает отправки", "3": "Отклонен",
    "4": "Отправлен, ожидается доставка", "5": "Возвращен", "6": "Ошибка",
    "7": "Завершен", "8": "Возвращен полностью", "9": "Возвращен частично",
    "10": "Подтвержден, ожидает оплаты", "11": "Завершен, ожидает зачисления",
    "12": "Ожидает оплаты", "13": "Завершен, оплачен", "14": "Завершен, отменен",
    "15": "В процессе выполнения", "16": "Ждет подтверждения покупателем", "17": "Приглашен администратор"
}


# Вспомогательная функция для расчета и записи метрик
def calculate_metrics(vac):
    total_revenue = sum(item[8] for item in vac)
    total_profit = sum(item[9] for item in vac)
    total_sales = sum(item[5] for item in vac)
    total_transactions = len([item for item in vac if item[3] == 'Завершен'])
    avg_order_value = total_revenue / total_transactions if total_transactions else 0

    return {
        'Общая выручка': total_revenue,
        'Общая прибыль': total_profit,
        'Количество проданных товаров': total_sales,
        'Количество успешных транзакций': total_transactions,
        'Средняя стоимость заказа': avg_order_value
    }



# Анализ самых популярных товаров
def top_products_extended(vac):
    products_sales = {}
    products_buyers = {}
    products_returns = {}
    products_revenue = {}

    for item in vac:
        product_name = item[6]
        user_id = item[1]
        amount = item[5]
        summa = item[8]
        status = item[3]

        products_sales[product_name] = products_sales.get(product_name, 0) + amount
        products_revenue[product_name] = products_revenue.get(product_name, 0) + summa

        if product_name not in products_buyers:
            products_buyers[product_name] = set()
        products_buyers[product_name].add(user_id)

        if status in ["Возвращен полностью", "Возвращен частично"]:
            products_returns[product_name] = products_returns.get(product_name, 0) + 1

    avg_check = {product: products_revenue[product] / products_sales[product] for product in products_sales}

    products_df = pd.DataFrame({
        'Продукт': list(products_sales.keys()),
        'Количество продаж': list(products_sales.values()),
        'Уникальные покупатели': [len(products_buyers[p]) for p in products_sales],
        'Возвраты': [products_returns.get(p, 0) for p in products_sales],
        'Средний чек': [avg_check.get(p, 0) for p in products_sales]
    })

    top_by_unique_buyers = products_df[['Продукт', 'Уникальные покупатели']].sort_values(by='Уникальные покупатели',
                                                                                         ascending=False).head(10)
    top_by_sales = products_df[['Продукт', 'Количество продаж']].sort_values(by='Количество продаж',
                                                                             ascending=False).head(10)
    top_by_avg_check = products_df[['Продукт', 'Средний чек']].sort_values(by='Средний чек', ascending=False).head(10)
    top_by_returns = products_df[['Продукт', 'Возвраты']].sort_values(by='Возвраты', ascending=False).head(10)

    return {
        "Топ по уникальным покупателям": top_by_unique_buyers,
        "Топ по количеству продаж": top_by_sales,
        "Топ по среднему чеку": top_by_avg_check,
        "Топ по возвратам": top_by_returns
    }


def top_customers(vac):
    customers = {}
    for item in vac:
        user_id = item[1]
        customers[user_id] = customers.get(user_id, 0) + 1

    top_customers_df = pd.DataFrame(list(customers.items()), columns=['User ID', 'Количество покупок']).sort_values(
        by='Количество покупок', ascending=False)
    return top_customers_df


def top_categories(vac):
    categories_sales = {}
    categories_revenue = {}

    for item in vac:
        category = item[11]
        amount = item[5]
        summa = item[8]

        categories_sales[category] = categories_sales.get(category, 0) + amount
        categories_revenue[category] = categories_revenue.get(category, 0) + summa

    categories_df = pd.DataFrame({
        'Категория': list(categories_sales.keys()),
        'Количество продаж': list(categories_sales.values()),
        'Выручка': list(categories_revenue.values())
    })

    top_categories_df = categories_df.sort_values(by='Количество продаж', ascending=False).head(10)

    return top_categories_df


def generate_report(telegram_id, tax, transactions, period_suffix):
    vac = []
    for item in transactions:
        date_payed = item['date_payed']
        order_id = item['id']
        user_id = item['user_id']
        status_num = item['status']
        status_text = statuses.get(str(status_num), "Неизвестный статус")
        product_id = item['items'][0]['product_id']
        amount = item['items'][0]['count']
        product_name = item['items'][0]['product_name']
        price = item['items'][0]['price']
        summa = float(price) * int(amount)
        profit = -summa * (1-tax/100) if status_num == 8 else summa * (1-tax/100)
        discount = item['items'][0]['discount']
        category = item['items'][0]['category']['slug']

        vac.append([order_id, user_id, date_payed, status_text, product_id, amount, product_name, price, summa, profit,
                    discount, category])

    metrics = calculate_metrics(vac)

    sales_df = pd.DataFrame(vac, columns=['Order ID', 'User ID', 'Date Payed', 'Status', 'Product ID', 'Amount',
                                          'Product Name', 'Price', 'Summa', 'Profit', 'Discount', 'Category'])
    top_products = top_products_extended(vac)
    top_customers_df = top_customers(vac)
    top_categories_df = top_categories(vac)

    current_time = datetime.now().strftime('%Y%m%d_%H%M%S')
    file_name = f'report_{telegram_id}_{period_suffix}_{current_time}.xlsx'
    file_path = f'./reports/{file_name}'

    if not os.path.exists('./reports'):
        os.makedirs('./reports')

    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
        # Лист с продажами
        sales_df.to_excel(writer, sheet_name='Продажи', index=False)

        # Лист с метриками
        metrics_df = pd.DataFrame([metrics])
        metrics_df.to_excel(writer, sheet_name='Метрики', index=False)

        # Лист с топами товаров
        workbook = writer.book
        worksheet = workbook.add_worksheet('Топы товаров')

        row = 0
        worksheet.write(row, 0, 'Топы товаров', workbook.add_format({'bold': True, 'font_size': 14}))
        row += 2

        for title, df in top_products.items():
            worksheet.write(row, 0, title, workbook.add_format({'bold': True}))
            df.to_excel(writer, sheet_name='Топы товаров', startrow=row + 1, index=False)
            row += len(df) + 3

        # Топ покупателей
        worksheet.write(row, 0, 'Топ покупателей', workbook.add_format({'bold': True}))
        top_customers_df.to_excel(writer, sheet_name='Топы товаров', startrow=row + 1, index=False)
        row += len(top_customers_df) + 3

        # Топ категории
        worksheet.write(row, 0, 'Топ категории по количеству продаж', workbook.add_format({'bold': True}))
        top_categories_df.to_excel(writer, sheet_name='Топы товаров', startrow=row + 1, index=False)

    return file_path, metrics

def generate_comparison_report(telegram_id, metrics_period_1, metrics_period_2):
    comparison_data = {
        'Показатель': ['Общая выручка', 'Общая прибыль', 'Количество проданных товаров',
                       'Количество успешных транзакций', 'Средняя стоимость заказа'],
        'Период 1': [metrics_period_1['Общая выручка'], metrics_period_1['Общая прибыль'],
                     metrics_period_1['Количество проданных товаров'], metrics_period_1['Количество успешных транзакций'],
                     metrics_period_1['Средняя стоимость заказа']],
        'Период 2': [metrics_period_2['Общая выручка'], metrics_period_2['Общая прибыль'],
                     metrics_period_2['Количество проданных товаров'], metrics_period_2['Количество успешных транзакций'],
                     metrics_period_2['Средняя стоимость заказа']],
        'Изменение': [(metrics_period_2['Общая выручка'] - metrics_period_1['Общая выручка']),
                      (metrics_period_2['Общая прибыль'] - metrics_period_1['Общая прибыль']),
                      (metrics_period_2['Количество проданных товаров'] - metrics_period_1['Количество проданных товаров']),
                      (metrics_period_2['Количество успешных транзакций'] - metrics_period_1['Количество успешных транзакций']),
                      (metrics_period_2['Средняя стоимость заказа'] - metrics_period_1['Средняя стоимость заказа'])]
    }

    comparison_df = pd.DataFrame(comparison_data)

    current_time = datetime.now().strftime('%Y%m%d_%H%M%S')
    file_name = f'comparison_report_{telegram_id}_{current_time}.xlsx'
    file_path = f'./reports/{file_name}'

    with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
        comparison_df.to_excel(writer, sheet_name='Сравнительный отчет', index=False)

    return file_path


async def analitic(state):
    data = await state.get_data()
    telegram_id = data['telegram_id']
    tax = data['tax']
    valid_transactions = data.get('valid_transactions', [])
    valid_transactions_period_1 = data.get('valid_transactions_period_1', [])
    valid_transactions_period_2 = data.get('valid_transactions_period_2', [])

    if valid_transactions:
        return generate_report(telegram_id, tax, valid_transactions, "single_report")[0]
    elif valid_transactions_period_1 and valid_transactions_period_2:
        # Генерация отчета для первого периода
        file_path_1, metrics_period_1 = generate_report(telegram_id, tax, valid_transactions_period_1, "period_1")
        # Генерация отчета для второго периода
        file_path_2, metrics_period_2 = generate_report(telegram_id, tax, valid_transactions_period_2, "period_2")
        # Генерация сравнительного отчета
        comparison_file_path = generate_comparison_report(telegram_id, metrics_period_1, metrics_period_2)

        return [file_path_1, file_path_2, comparison_file_path]


