import os
import re
import shutil
import pdfplumber

"При запуске программы убедитесь что processed_contracts не содержит файлов, так как при запуске будут создоваться копии"
def extract_text_from_pdf(pdf_path):
    """Извлекает текст из PDF-файла с помощью pdfplumber."""
    text = ""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
    except Exception as e:
        print(f"Ошибка при чтении {pdf_path}: {str(e)}")
    return text


def parse_contract_data(text):
    """Извлекает ключевые данные из текста договора с улучшенными шаблонами."""
    data = {'type': '', 'counterparty': '', 'date': '', 'number': ''}

    # Разделяем текст на строки для анализа
    lines = [line.strip() for line in text.split('\n') if line.strip()]

    # 1. Поиск типа договора (первые 5 строк, содержащие слово "ДОГОВОР")
    contract_line_index = -1
    for i in range(min(5, len(lines))):
        if 'ДОГОВОР' in lines[i]:
            contract_line_index = i
            break

    if contract_line_index >= 0:
        # Нашли строку с типом договора
        contract_type_line = lines[contract_line_index]

        # Удаляем номер договора если он есть
        contract_type_line = re.sub(r'№\s*[^\s]+', '', contract_type_line)

        # Удаляем слово "договора" в родительном падеже
        contract_type_line = re.sub(r'\bдоговора\b', '', contract_type_line, flags=re.IGNORECASE)

        # Удаляем лишние символы
        contract_type_line = re.sub(r'[^a-zA-Zа-яА-Я0-9\s]', '', contract_type_line)

        # Заменяем пробелы на подчеркивания и удаляем лишние подчеркивания
        contract_type = re.sub(r'\s+', '_', contract_type_line.strip())
        contract_type = re.sub(r'_+', '_', contract_type)

        # Если в строке только слово "ДОГОВОР", проверяем следующую строку
        if contract_type == 'ДОГОВОР' and contract_line_index + 1 < len(lines):
            next_line = lines[contract_line_index + 1]

            # Проверяем, что следующая строка не содержит номер или дату
            if not re.search(r'№|\d{4}|г\.|год', next_line, re.IGNORECASE):
                # Добавляем следующую строку к типу договора
                next_line_clean = re.sub(r'[^a-zA-Zа-яА-Я0-9\s]', '', next_line)
                next_line_clean = re.sub(r'\s+', '_', next_line_clean.strip())
                contract_type += '_' + next_line_clean

        data['type'] = contract_type

    # 2. Поиск номера договора (первые 5 строк)
    number_patterns = [
        r'№\s*([^\s,]+)',
        r'номер[:\s]*([^\s,]+)',
    ]

    for i in range(min(5, len(lines))):
        for pattern in number_patterns:
            match = re.search(pattern, lines[i], re.IGNORECASE)
            if match:
                number = match.group(1).strip()
                number = re.sub(r'[^a-zA-Zа-яА-Я0-9\/\-]', '', number)
                if number and number != '':
                    data['number'] = number
                    break
        if data['number']:
            break

    # 3. Поиск даты (первые 5 строк)
    date_patterns = [
        r'«(\d{1,2})»\s*([а-я]+)\s*(\d{4})',
        r'"(\d{1,2})"\s*([а-я]+)\s*(\d{4})',
        r'(\d{1,2})\s*([а-я]+)\s*(\d{4})',
        r'(\d{1,2})\.(\d{1,2})\.(\d{4})',
        r'(\d{1,2})/(\d{1,2})/(\d{4})',
        r'(\d{4})[-–](\d{1,2})[-–](\d{1,2})',
        r'(\d{1,2})[-–](\d{1,2})[-–](\d{4})',
    ]

    months = {
        'января': '01', 'февраля': '02', 'марта': '03', 'апреля': '04',
        'мая': '05', 'июня': '06', 'июля': '07', 'августа': '08',
        'сентября': '09', 'октября': '10', 'ноября': '11', 'декабря': '12'
    }

    for i in range(min(5, len(lines))):
        for pattern in date_patterns:
            match = re.search(pattern, lines[i], re.IGNORECASE)
            if match:
                if len(match.groups()) == 3 and not match.group(2).isdigit():
                    # Формат с названием месяца
                    day, month_ru, year = match.groups()
                    month = months.get(month_ru.lower(), '01')
                    data['date'] = f"{year}-{month}-{day.zfill(2)}"
                    break
                elif len(match.groups()) == 3 and match.group(2).isdigit():
                    # Формат DD.MM.YYYY или DD/MM/YYYY
                    day, month, year = match.groups()
                    data['date'] = f"{year}-{month.zfill(2)}-{day.zfill(2)}"
                    break
        if data['date']:
            break

    # Если не нашли полную дату, ищем только год в первых 10 строках
    if not data['date']:
        for i in range(min(10, len(lines))):
            year_match = re.search(r'\b(20\d{2})\b', lines[i])
            if year_match:
                data['date'] = year_match.group(1)  # Только год
                break

    # 4. Поиск контрагента (первые 20 строк)
    counterparty_patterns = [
        r'Общество с ограниченной ответственностью\s*[«"]([^»"]+)[»"]',
        r'ООО\s*[«"]([^»"]+)[»"]',
        r'АО\s*[«"]([^»"]+)[»"]',
        r'ЗАО\s*[«"]([^»"]+)[»"]',
        r'индивидуальный предприниматель\s+([^,]+)',
        r'ИП\s+([^,]+)',
        r'Гражданин Российской Федерации\s+([^,]+)',
        r'Гражданка Российской Федерации\s+([^,]+)',
        r'Компания\s*[«"]([^»"]+)[»"]',
        r'Покупатель[:\s]*([^,]+)',
        r'Продавец[:\s]*([^,]+)',
    ]

    for i in range(min(20, len(lines))):
        for pattern in counterparty_patterns:
            match = re.search(pattern, lines[i], re.IGNORECASE)
            if match:
                counterparty = match.group(1)
                # Очищаем от лишних символов
                counterparty = re.sub(r'[«»"\\/:*?<>|\s]', '_', counterparty)
                data['counterparty'] = re.sub(r'_+', '_', counterparty)
                break
        if data['counterparty']:
            break

    return data


def categorize_contract(text, contract_type):
    """Определяет категорию договора на основе текста и типа с улучшенной логикой."""
    analysis_text = (contract_type + ' ' + text).lower()

    # Сначала проверяем особые случаи, которые должны быть в "Прочие"
    if any(word in analysis_text for word in ['инвестицион', 'товарищест', 'паев', 'актив', 'фонд']):
        return 'Прочие'

    # Проверяем, является ли это договором купли-продажи недвижимости
    if ('купл' in analysis_text or 'продаж' in analysis_text) and any(
            word in analysis_text for word in ['недвижим', 'складск', 'помещен', 'квартир', 'дом', 'здан']):
        return 'Прочие'

    # Улучшенные категории с ключевыми словами (корнями слов)
    categories = {
        'Трудовые': ['трудовой', 'труда', 'работник', 'занят', 'кадр', 'работодатель'],
        'Аренда': ['аренд', 'найм', 'жил', 'нежил'],
        'Поставка': ['поставк', 'товар', 'розничн'],
        'Услуги': ['услуг', 'подряд', 'строительн', 'перевоз', 'организац']
    }

    # Проверяем категории в порядке приоритета
    for category, keywords in categories.items():
        if any(keyword in analysis_text for keyword in keywords):
            return category

    return 'Прочие'


def process_contracts(input_folder):
    """Основная функция обработки договоров."""
    output_folder = "processed_contracts"
    os.makedirs(output_folder, exist_ok=True)

    # Создаем папки для категорий
    categories = ['Аренда', 'Поставка', 'Услуги', 'Трудовые', 'Прочие']
    for category in categories:
        os.makedirs(os.path.join(output_folder, category), exist_ok=True)

    processed_count = 0
    error_count = 0

    for filename in os.listdir(input_folder):
        if filename.lower().endswith('.pdf'):
            pdf_path = os.path.join(input_folder, filename)
            print(f"Обработка: {filename}")

            try:
                text = extract_text_from_pdf(pdf_path)
                if not text.strip():
                    print(f"  Предупреждение: не удалось извлечь текст из {filename}")
                    shutil.copy2(pdf_path, os.path.join(output_folder, 'Прочие', filename))
                    error_count += 1
                    continue

                # Извлечение данных
                data = parse_contract_data(text)
                print(f"  Извлеченные данные: {data}")

                # Формирование имени файла (только те поля, которые найдены)
                name_parts = []
                if data['type']:
                    name_parts.append(data['type'])
                if data['counterparty']:
                    name_parts.append(data['counterparty'])
                if data['date']:
                    name_parts.append(data['date'])
                if data['number']:
                    name_parts.append(f"№{data['number']}")

                new_name = "_".join(name_parts) + ".pdf"
                new_name = re.sub(r'_+', '_', new_name)
                new_name = re.sub(r'[\\/*?:"<>|]', '', new_name)

                # Определение категории
                category = categorize_contract(text, data['type'])
                print(f"  Определенная категория: {category}")

                # Проверка на дубликаты
                new_path = os.path.join(output_folder, category, new_name)
                counter = 1
                while os.path.exists(new_path):
                    name_parts = new_name.split('.')
                    name_parts[0] += f"_{counter}"
                    new_name = '.'.join(name_parts)
                    new_path = os.path.join(output_folder, category, new_name)
                    counter += 1

                # Копирование файла
                shutil.copy2(pdf_path, new_path)
                print(f"  Успешно: {filename} -> {new_name} в категории {category}")
                processed_count += 1

            except Exception as e:
                print(f"  Ошибка при обработке {filename}: {str(e)}")
                error_path = os.path.join(output_folder, 'Прочие', filename)
                shutil.copy2(pdf_path, error_path)
                error_count += 1

    print(f"\nОбработка завершена. Успешно: {processed_count}, с ошибками: {error_count}")

"При запуске программы убедитесь что processed_contracts не содержит файлов, так как при запуске будут создоваться копии"
if __name__ == "__main__":
    input_folder = input("Введите путь к папке с PDF-файлами: ")
    if os.path.exists(input_folder):
        process_contracts(input_folder)
    else:
        print("Указанная папка не существует!")