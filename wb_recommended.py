import pandas as pd

# Устанавливаем, сколько популярных товаров нам нужно
PRODUCTS_LIMIT = 10
IMPORT_SHEET_NAME = "Общий отчет"
EXPORT_SHEET_NAME = "Загрузка рекомендаций для WB"
HEADER = ["Артикул WB", "Связанный артикул WB"]
IN_STOCK_GOODS = "instock.xlsx"
OUT_STOCK_GOODS = "outstock.xlsx"
WBSTAT = "wbstat.xlsx"
STOPLIST = "stoplist.xlsx"
IN_STOCK_RECOMMEND = "recommend_instock"
OUT_STOCK_RECOMMEND = "recommend_outstock"
FILELIMIT = 10000


def get_active_wb_categories(data_values):
    """Возвращаем список активных категорий на ВБ"""
    categories_list = []
    for line in data_values:
        categories_list.append(line[1])
    unique_categories = set(categories_list)
    return sorted(list(unique_categories))


def get_products_by_categories(data_values, categories):
    """Готовим словарь словарей со всеми продуктами,
    разбитыми по категориям"""

    products_turnover = {}
    for category in categories:
        products_turnover[category] = {}

    for line in data_values:
        if line[9] == "Товар на сайте менее 30 дн.":
            turnover = 50
        else:
            turnover = int(line[9])
        # Иногда ВБ теряет оборачиваемость старых товаров, уберем ее
        if turnover == 0:
            turnover = 500
        products_turnover[line[1]][line[3]] = turnover
    return products_turnover


def create_recommended_dict(products_dict):
    """
    Формируем словарь с ключем-категорией и значением в
    виде листа с перечислением топовых товаров
    PRODUCTS_LIMIT укажет количество топа
    """
    top_products = {}
    for category_name in products_dict:
        sorted_products = sorted(
            products_dict[category_name],
            key=products_dict[category_name].get,
            reverse=True,
        )

        top_products[category_name] = sorted_products[:PRODUCTS_LIMIT]
    return top_products


def create_stock_recommendations(data_values, top10_in_categories):
    stock_recommendations = []
    # Убираем артикулы, которые ВБ не может найти
    stoplist = open_xlsx_file(STOPLIST, "Лист1")

    for line in data_values:
        # Используем wb_id теперь, а не артикул поставщика
        product = line[4]
        # проверяем на созданные ВБ артикулы (их может не оказаться)
        if [product] not in stoplist:
            top_products = top10_in_categories.get(line[1])
            if top_products is not None:
                for top_product in top_products:
                    # нельзя рекомендовать самого себя
                    if product != top_product:
                        stock_recommendations.append([product, top_product])
    return stock_recommendations


def gererate_xlsx_file(stock_recommendations, filename):
    """
    Создаем файл рекомендаций по заданному шаблону
    Отдельно для товаров в наличии и не в наличии
    ВБ не любит большие файлы, будем делить по 20к строк
    """
    i = 0
    finish = 0
    while finish < len(stock_recommendations):
        start = i * FILELIMIT
        finish = (i + 1) * FILELIMIT
        stock_recommendations_slice = stock_recommendations[start:finish]
        print(
            f"Saved {len(stock_recommendations_slice)}"
            f" rows to {filename}_{i + 1}.xlsx"
        )
        excel_dataframe = pd.DataFrame(stock_recommendations_slice)
        excel_dataframe.to_excel(
            f"{filename}_{i + 1}.xlsx",  # имя файла
            header=HEADER,  # заголовки столбцов
            index=False,  # не нумеруем строки
            sheet_name=EXPORT_SHEET_NAME,  # имя листа
        )
        i += 1


def open_xlsx_file(filename, sheetname):
    # Файл отчет о ценах на товары с остатком
    # Загружаем эксель в объект
    full_excel_file = pd.ExcelFile(filename)

    # Берем сразу нужный нам лист
    data_values = full_excel_file.parse(sheetname).values
    print(f"Opened {filename}: {len(data_values)} rows.")
    return data_values


def create_sku_category_dic(data_values):
    """Создаем словарь для категорий."""
    sku_category_dic = {}
    for item in data_values:
        # Теперь требуется не артикул поставщика, а номенклатура 1С от WB
        wb_id = item[4]
        category_name = item[1]
        sku_category_dic[wb_id] = category_name
    return sku_category_dic


def create_sku_wbstatrating(wbstat_sheet):
    """Создаем словарь wb_id -> рейтинг wbstat."""
    sku_wbstatrating_dic = {}
    for item_idx in range(1, len(wbstat_sheet)):
        item = wbstat_sheet[item_idx]
        in_stock = item[15]
        if int(in_stock) > 0:
            # Теперь требуется не артикул поставщика, а номенклатура 1С от WB
            wb_id = item[2]
            wbstat_rating = item[5]
            sku_wbstatrating_dic[wb_id] = wbstat_rating

    return sku_wbstatrating_dic


def create_category_sku_wbstat(sku_category_dic, sku_wbstatrating_dic, categories):
    category_sku_wbstat = {}
    for category in categories:
        category_sku_wbstat[category] = {}

    for item in sku_category_dic:
        category = sku_category_dic[item]
        wb_stat_rating = sku_wbstatrating_dic.get(str(item), 0)
        category_sku_wbstat[category][item] = wb_stat_rating

    return category_sku_wbstat


def main():
    """
    Популярные товары получаем из товаров в наличии
    Записываем их и туда, и ко всем товарам не в наличии
    """
    data_values = open_xlsx_file(IN_STOCK_GOODS, IMPORT_SHEET_NAME)
    data_values_out_stock = open_xlsx_file(OUT_STOCK_GOODS, IMPORT_SHEET_NAME)
    wbstat_sheet = open_xlsx_file(WBSTAT, "Аналитика Wildberries")

    # Получаем словарь Артикул - Категория для ВБСТАТ
    sku_category_dic = create_sku_category_dic(data_values)

    # Получаем словарь Аритул - Рейтинг ВБСТАТ
    sku_wbstatrating_dic = create_sku_wbstatrating(wbstat_sheet)

    # Получаем категории, где работает поставщик
    categories = get_active_wb_categories(data_values)

    # Получаем словарь Категория — Артикул — Рейтинг ВБСТАТ
    category_sku_wbstat = create_category_sku_wbstat(
        sku_category_dic, sku_wbstatrating_dic, categories
    )

    # Получаем топовые артикулы в каждой категории
    top_products = create_recommended_dict(category_sku_wbstat)

    # Получаем лист, каждый продукт получает все икс рекомендаций
    stock_recommendations = create_stock_recommendations(data_values, top_products)

    """
    # Заполняем словарь товарами по категориям
    products_dict = get_products_by_categories(data_values, categories)

    # Сокращаем словарь до топ продуктов (настраиваемое число)
    top_in_categories = create_recommended_dict(products_dict)

    # Получаем лист, каждый продукт получает все икс рекомендаций
    stock_recommendations = create_stock_recommendations(
        data_values,
        top_in_categories
    )
    """
    # Получаем лист, каждый продукт получает все икс рекомендаций
    out_stock_recommendations = create_stock_recommendations(
        data_values_out_stock, top_products
    )

    # Генерируем объект для сохранения его в виде XLSX (для наличия)
    gererate_xlsx_file(stock_recommendations, IN_STOCK_RECOMMEND)

    # Генерируем объект для сохранения его в виде XLSX (для неналичия)
    gererate_xlsx_file(out_stock_recommendations, OUT_STOCK_RECOMMEND)


if __name__ == "__main__":
    main()
