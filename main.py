from ozone_analyzer import OzoneAnalyzer

if __name__ == "__main__":
    input_file = input("Укажите путь к файлу с товарами: ")
    analyzer = OzoneAnalyzer(input_file)
    analyzer.filter_products()
    analyzer.generate_statistics()

    category_to_filter = input("Введите название категории для фильтрации (или нажмите Enter для пропуска): ")
    if category_to_filter:
        analyzer.filter_by_category(category_to_filter)