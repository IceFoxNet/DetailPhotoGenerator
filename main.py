import os

while True:
    try:
        from PIL import Image, ImageFont, ImageDraw
        import gspread, yadisk
        from rembg import remove
        import time, aiohttp, pathlib
        from configparser import ConfigParser
    except ImportError as e:
        package = e.msg.split()[-1][1:-1]
        os.system(f'python -m pip install {package}')
    else:
        break

workspace = pathlib.Path(__file__).parent.resolve()
downloaded_image_buffer = os.path.join(workspace, 'download_buffer.png')
config = ConfigParser()
config.read(os.path.join(workspace, 'config.ini'))
sheet_url = config.get('parser', 'url')

# Загрузка шрифтов
try:
    font_path_regular = os.path.join(workspace, "Inter.ttf")
    font_path_medium = os.path.join(workspace, "Inter-Medium.ttf")
    font_path_bold = os.path.join(workspace, "Inter-Bold.ttf")

    main_font_24 = ImageFont.truetype(font_path_regular, 24)
    main_font_12 = ImageFont.truetype(font_path_regular, 12)
    main_font_42_medium = ImageFont.truetype(font_path_medium, 42)
    main_font_82_bold = ImageFont.truetype(font_path_bold, 82)
    main_font_49_bold = ImageFont.truetype(font_path_bold, 49)

except IOError:
    print("Шрифты 'Inter.ttf', 'Inter-Medium.ttf' или 'Inter-Bold.ttf' не найдены. Проверьте пути.")

try:
    with Image.open(os.path.join(workspace, "sample.jpg")) as img1_template:
        img1_template.load()
except FileNotFoundError:
    print("Файл 'sample.jpg' не найден.")

# Функция для переноса текста по ширине
def wrap_text(text, font, max_width):
    words = text.split()
    lines = []
    current_line = ""
    for word in words:
        test_line = f"{current_line} {word}".strip()
        if font.getbbox(test_line)[2] <= max_width:
            current_line = test_line
        else:
            lines.append(current_line)
            current_line = word
    lines.append(current_line)
    return lines

# Асинхронная функция загрузки изображений
async def download_image(session, url, art):
    try:
        response = await session.get(url)
        if response.status == 200:
            file_path = os.path.join(workspace, f"buffer_{art}.jpg")
            with open(file_path, "wb") as f:
                f.write(await response.read())
            return file_path
        else:
            print(f"Ошибка загрузки изображения для артикула {art}. Статус: {response.status}")
            return None
    except Exception as e:
        print(f"Ошибка при загрузке изображения для артикула {art}: {e}")
        return None

# Асинхронная основная обработка изображений
async def main(start: int, end: int, setup: dict):

    if start < 3: start = 3

    creds = setup.get('GoogleCredentials')
    google_client = gspread.authorize(creds)
    spreadsheet = google_client.open_by_url(sheet_url)
    sheet = spreadsheet.worksheet('Детали')
    yandex: yadisk.YaDisk = setup.get('YandexDisk')
    history = set()    

    # Получаем все значения из нужных столбцов за один запрос
    arts = sheet.range(f'C{start}:C{end}')  # Все данные из столбца 3 (артикул)
    prices = sheet.range(f'I{start}:I{end}') # Все данные из столбца 9 (цена)
    names = sheet.range(f'A{start}:A{end}') # Все данные из столбца 1 (название)
    urls = sheet.range(f'O{start}:O{end}') # Все данные из столбца 15 (фотография)
    colors = sheet.range(f'B{start}:B{end}') # Все данные из столбца 2 (цвет)

    async with aiohttp.ClientSession(proxy='http://user258866:pe9qf7@166.0.211.142:7576') as session:
        for i in range(start, end+1):
            if i < len(arts):
                art = arts[i].value
            else:
                break  # Завершаем цикл, если индекс выходит за пределы
            if not art:
                print(f"Пропущена строка {i}: значение отсутствует.")
                continue

            typ = 'Part'
            price = (prices[i].value or "Не указана") + '₽'
            name = names[i].value or "Без названия"
            url = urls[i].value or None
            color = colors[i].value or "Без цвета"
            
            identity = art + '_' + color.replace(' ', '_')
            if identity in history:
                continue
            history.add(identity)

            # Проверка на допустимый URL
            if not url or url == '-':
                print(f"Пропущено изображение для артикула {art}: недопустимый URL.")
                continue

            # Загрузка изображения
            file_path = await download_image(session, url, art)
            if not file_path:
                continue

            def is_white_background(image_path, threshold=240):
                """
                Проверяет, является ли фон изображения белым.
                threshold — пороговое значение яркости для определения белого цвета.
                """
                try:
                    with Image.open(image_path) as img:
                        img = img.convert("RGBA")
                        # Берем верхнюю часть изображения для анализа фона
                        width, height = img.size
                        top_pixels = [img.getpixel((x, 0)) for x in range(width)]
                        # Проверяем, насколько пиксели близки к белому цвету
                        white_pixels = [pixel for pixel in top_pixels if all(val >= threshold for val in pixel[:3])]
                        # Если большинство пикселей белые, считаем фон белым
                        return len(white_pixels) / len(top_pixels) > 0.8
                except Exception as e:
                    print(f"Ошибка при проверке фона изображения: {e}")
                    return False

            # Основной блок обработки
            try:
                with open(file_path, 'rb') as f:
                    if is_white_background(file_path):
                        print(f"Фон изображения {art} белый. Пропускаем удаление фона.")
                        output_path = file_path  # Если фон белый, просто используем исходное изображение
                    else:
                        output_data = remove(f.read())  # Если фон не белый, удаляем фон
                        output_path = file_path.replace('.jpg', '_no_bg.png')
                        os.remove(file_path)
                        with open(output_path, 'wb') as out_file:
                            out_file.write(output_data)
            except Exception as e:
                print(f"Ошибка при удалении фона для изображения {art}: {e}")
                continue
            
            final_output_path = output_path.replace('buffer_', '').replace('_no_bg', '')
            if color != 'Без цвета':
                final_output_path = final_output_path.replace('.jpg', f'_{color.replace(' ', '_')}.jpg').replace('.png', f'_{color.replace(' ', '_')}.png')

            # Работа с фоном и наложением изображений
            img1 = img1_template.copy()

            try:
                with Image.open(output_path).convert("RGBA") as img2:
                    # Первичное изменение размера до высоты 500 пикселей
                    img2 = img2.resize((int(500 / img2.height * img2.width), 500))

                    # Ограничение ширины до 1020 пикселей
                    if img2.width > 1020:
                        scaling_factor = 1020 / img2.width
                        new_width = 1020
                        new_height = int(img2.height * scaling_factor)
                        img2 = img2.resize((new_width, new_height))

                    # Добавление белого фона
                    white_bg = Image.new("RGBA", img2.size, (255, 255, 255, 255))
                    img2 = Image.alpha_composite(white_bg, img2)

                    # Вычисление центра по оси Y, чтобы выровнять изображение между 727 и 227
                    center_y = (727 + 227) / 2
                    position_y = center_y - img2.height / 2

                    # Расчет позиции по оси X для центрирования
                    center_x = 1080 / 2
                    position_x = center_x - img2.width / 2

                    # Наложение изображения на шаблон
                    img1.paste(img2, (int(position_x), int(position_y)))
            except FileNotFoundError:
                print(f"Изображение {output_path} не найдено.")
                continue

            # Добавление рамок Frame_Gray.png и Frame_Green.png
            try:
                frame_gray_path = os.path.join(workspace, "Frame_Gray.png")
                frame_green_path = os.path.join(workspace, "Frame_Green.png")

                if os.path.exists(frame_gray_path):
                    with Image.open(frame_gray_path).convert("RGBA") as frame_gray:
                        gray_position = (60, 727)
                        img1.paste(frame_gray, gray_position, mask=frame_gray)

                if os.path.exists(frame_green_path):
                    with Image.open(frame_green_path).convert("RGBA") as frame_green:
                        green_position = (560, 894)
                        img1.paste(frame_green, green_position, mask=frame_green)
            except FileNotFoundError as e:
                print(f"Ошибка при добавлении рамок: {e}")
                continue

            # Наложение логотипа серии
            colors_for_generator_file = os.path.join(workspace, f'colors_for_generator/{color.replace(" ", "")}.png')
            if os.path.exists(colors_for_generator_file):
                try:
                    with Image.open(colors_for_generator_file) as img3:
                        img1.paste(img3.resize((960, 167)), (60, 60))
                except FileNotFoundError:
                    print(f"Файл {colors_for_generator_file} не найден.")
                    continue

            # Добавляем текст на изображение
            drawer = ImageDraw.Draw(img1)

            # Добавляем текст с подчеркиванием
            art_position = (102, 912)
            drawer.text(art_position, art, font=main_font_49_bold, fill='black')

            text_bbox = drawer.textbbox((0, 0), art, font=main_font_49_bold)
            text_width = text_bbox[2] - text_bbox[0]
            text_height = text_bbox[3] - text_bbox[1]

            underline_y = art_position[1] + text_height + 16
            line_width = 5

            drawer.line(
                [(art_position[0], underline_y), (art_position[0] + text_width, underline_y)],
                fill='black', width=line_width
            )

            # Многострочный текст для длинных названий
            lines = wrap_text(name, main_font_24, 400)
            y_offset = 0
            for line in lines:
                drawer.text((102, 756 + y_offset), line, font=main_font_24, fill='black')
                y_offset += 40

            # Текст "typ"
            drawer.text((602, 931.5), typ, font=main_font_42_medium, fill='black')

            # Текст с ценой
            drawer.text((560, 756), price, font=main_font_82_bold, fill='black')

            # Функция для разбивки текста на строки
            def wrap_text_to_box(text, font, max_width):
                words = text.split(' ')
                lines = []
                current_line = ""

                for word in words:
                    test_line = current_line + " " + word if current_line else word
                    line_width = drawer.textbbox((0, 0), test_line, font=font)[2] - \
                                 drawer.textbbox((0, 0), test_line, font=font)[0]

                    if line_width <= max_width:
                        current_line = test_line
                    else:
                        lines.append(current_line)
                        current_line = word

                if current_line:
                    lines.append(current_line)

                return lines

            # Параметры
            start_x = 60
            start_y = 1020
            box_width = 960  # Максимальная ширина блока
            box_height = 60  # Высота блока для текста
            line_height = 16  # Высота строки (интервал между строками)

            # Текст с товарным знаком
            trademark_text = (
                "LEGO, the LEGO logo, the Minifigure, DUPLO, the DUPLO logo, NINJAGO, "
                "the NINJAGO logo, the FRIENDS logo, the HIDDEN SIDE logo, the MINIFIGURES logo, "
                "MINDSTORMS and the MINDSTORMS logo are trademarks of the LEGO Group. ©2023 The LEGO Group. "
                "All rights reserved."
            )

            # Разбиваем текст на строки, учитывая максимальную ширину
            lines = wrap_text_to_box(trademark_text, main_font_12, box_width)

            # Общая высота текста
            total_text_height = len(lines) * line_height

            # Вычисляем смещение по вертикали для выравнивания по центру
            y_offset = start_y + (box_height - total_text_height) // 2

            # Рисуем каждую строку текста с выравниванием по ширине и высоте
            for line in lines:
                # Вычисляем ширину текущей строки
                line_width = drawer.textbbox((0, 0), line, font=main_font_12)[2] - \
                             drawer.textbbox((0, 0), line, font=main_font_12)[0]

                # Вычисляем смещение по горизонтали для выравнивания по центру
                x_offset = start_x + (box_width - line_width) // 2

                # Рисуем строку текста
                drawer.text((x_offset, y_offset), line, font=main_font_12, fill="#808080")

                # Увеличиваем Y для следующей строки
                y_offset += line_height

            # Сохраняем финальное изображение
            img1.save(final_output_path)
            try:
                yandex.makedirs(f'Авито/{pathlib.Path(final_output_path).stem}')
            except:
                pass
            yandex.upload(final_output_path, f'Авито/{pathlib.Path(final_output_path).stem}/{art}.{final_output_path.split('.')[-1]}', overwrite=True)
            os.remove(final_output_path)
            os.remove(output_path)
            print(art, typ, name, color, price)

            # Добавляем задержку между запросами
            time.sleep(5)  # Задержка на 1 секунду между запросами