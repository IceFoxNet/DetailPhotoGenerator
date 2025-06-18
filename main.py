from PIL import Image, ImageFont, ImageDraw
import time, aiohttp, pathlib
import gspread.spreadsheet
from rembg import remove
import gspread, yadisk
import os

from sqlalchemy import Column, String, UUID
from sqlalchemy.orm import declarative_base, sessionmaker, Session
from sqlalchemy import create_engine
from uuid import uuid4 as uuid
from sqlalchemy import and_, delete

Base = declarative_base()
class Media(Base):
    __tablename__ = 'media'
    id = Column(UUID, nullable=False, primary_key=True, comment='ID Созданного медиа')
    author_id = Column(UUID, nullable=False, comment='ID Приложения, создавшего медиа')
    name = Column(String, nullable=False, comment='Название файла')
    url = Column(String, comment='Прямая ссылка на медиа')
    author_ver = Column(String, comment='Версия приложения, создавшего медиа')
    resource_id = Column(String, nullable=False, comment='Идентификатор, по которому можно найти медиа')
    product_id = Column(String, nullable=False, comment='Идентификатор товара')
    description = Column(String, comment='Дополнительная информация о медиа')

class DBConnect:
    def __init__(self, appInfo: dict):
        self.login = appInfo.get('DBLogin')
        self.password = appInfo.get('DBPassword')
        self.appVer = appInfo.get('AppVer')
        self.id = appInfo.get('DBID')
        try:
            engine = create_engine(url=f"postgresql+psycopg2://{self.login}:{self.password}@scope-db-lego-invest-group.db-msk0.amvera.tech:5432/LEGOSystems")
            Session = sessionmaker(bind=engine)
            self.session = Session()
        except Exception as e:
            raise SystemError(f'Ошибка авторизации в базе данных: {e}')

    def create_media(self, url: str, filename: str, resource_id: str, product_id: str, description: str | None):
        new_media = Media(
            id = uuid(),
            author_id = self.id,
            author_ver = self.appVer,
            resource_id = resource_id,
            product_id = product_id,
            url = url,
            name = filename,
            description = description
        )
        self.session.add(new_media)
        self.session.commit()

    def is_actual_media_generated(self, resource_id: str):
        results = self.session.query(Media.author_ver).where(and_(Media.author_id == self.id, Media.resource_id == resource_id)).all()
        if len(results) == 0: return False
        return all(res[0] == self.appVer for res in results)

    def delete_media(self, resource_id: str, filename: str):
        media = self.session.query(Media).where(and_(Media.author_id == self.id, Media.resource_id == resource_id, Media.name == filename)).one_or_none()
        if media is not None:
            self.session.execute(delete(Media).where(Media.id == media.id))
    
    def close(self):
        self.session.close()

workspace = pathlib.Path(__file__).parent.resolve()
downloaded_image_buffer = os.path.join(workspace, 'download_buffer.png')

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

    spreadsheet: gspread.spreadsheet.Spreadsheet = setup.get('DetailsSheet')
    sheet = spreadsheet.worksheet('Детали')
    yandex: yadisk.YaDisk = setup.get('YandexDisk')
    dbconn = DBConnect(setup.get('AppInfo'))

    # Получаем все значения из нужных столбцов за один запрос
    arts = sheet.range(f'C{start}:C{end}')  # Все данные из столбца 3 (артикул)
    names = sheet.range(f'A{start}:A{end}') # Все данные из столбца 1 (название)
    urls = sheet.range(f'O{start}:O{end}') # Все данные из столбца 15 (фотография)
    colors = sheet.range(f'B{start}:B{end}') # Все данные из столбца 2 (цвет)

    async with aiohttp.ClientSession(proxy='http://user258866:pe9qf7@166.0.211.142:7576') as session:
        for i in range(len(arts)):
            art = arts[i].value
            if not art:
                print(f"Пропущена строка {i}: значение отсутствует.")
                continue

            typ = 'Part'
            name = names[i].value or "Без названия"
            url = urls[i].value or None
            color = colors[i].value or "Без цвета"
            
            identity = art + '_' + color.replace(' ', '_')
            if dbconn.is_actual_media_generated(identity):
                print(f'Пропущен артикул {identity}: актуальные карточки деталей уже сгенерированы')
                continue

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
                if is_white_background(file_path):
                    print(f"Фон изображения {art} белый. Пропускаем удаление фона.")
                    output_path = file_path  # Если фон белый, просто используем исходное изображение
                else:
                    output_data = remove(Image.open(file_path))  # Если фон не белый, удаляем фон
                    output_path = file_path.replace('.jpg', '_no_bg.png')
                    os.remove(file_path)
                    output_data.save(output_path)
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
                yandex.remove(f'Авито/{pathlib.Path(final_output_path).stem}')
            except:
                pass
            try:
                yandex.mkdir(f'Авито/{pathlib.Path(final_output_path).stem}')
            except:
                pass
            disk_path = f'Авито/{pathlib.Path(final_output_path).stem}/{art}.{final_output_path.split('.')[-1]}'
            yandex.upload(final_output_path, disk_path, overwrite=True)
            yandex.publish(disk_path)
            media_url = yandex.get_meta(disk_path).public_url
            if media_url is not None:
                media_url = media_url.replace('yadi.sk', 'disk.yandex.ru')
            dbconn.delete_media(art, disk_path)
            dbconn.create_media(media_url, disk_path, identity, f'ID-P-{art}-0-0', f'Карточка с деталью, фотография BrickLink, {art}, {color} {name}')
            os.remove(final_output_path)
            os.remove(output_path)
    dbconn.close()

if __name__ == '__main__':
    from Setup.setup import setup
    import asyncio
    asyncio.run(main(3, 23, setup))