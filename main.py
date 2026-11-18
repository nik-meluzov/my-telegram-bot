import io
import asyncio
import os 
from aiogram import Bot, Dispatcher, types, F
from aiogram.types import BufferedInputFile
import pdfplumber
import openpyxl

# Вставь сюда свой токен от @id199142634 (@BotFather)
API_TOKEN = os.getenv("API_TOKEN")

bot = Bot(token=API_TOKEN)
dp = Dispatcher()

# Твоя логика извлечения таблиц (без изменений, только обернули в async)
async def convert_pdf_to_excel(file_content: bytes) -> io.BytesIO:
    output_excel = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Данные из PDF"
    
    with pdfplumber.open(io.BytesIO(file_content)) as pdf:
        first = True
        found_tables = False
        
        for i, page in enumerate(pdf.pages):
            tables = page.extract_tables()
            for table in tables:
                if not table:
                    continue
                found_tables = True
                
                # Если это первая таблица, заголовки уже есть, если нет - можно добавить
                for row in table:
                    ws.append(row)
                
                # Добавим пустую строку между таблицами для красоты
                ws.append() 

    if not found_tables:
        raise ValueError("Таблицы не найдены")
        
    wb.save(output_excel)
    output_excel.seek(0)
    return output_excel

# Хендлер на команду /start
from aiogram.filters import Command
async def cmd_start(message: types.Message):
    await message.answer("Привет! Отправь мне PDF-файл с таблицами, я конвертирую его в Excel.")

# Хендлер на получение файла
@dp.message(F.document)
async def handle_file(message: types.Message):
    document = message.document
    
    # Проверка расширения
    if not document.file_name.lower().endswith('.pdf'):
        await message.answer("Это не PDF! Выбери файл с расширением .pdf")
        return

    status_msg = await message.answer("⏳ Конвертирую... Пожалуйста, подожди.")
    
    try:
        # Скачиваем файл в память
        file = await bot.download(document)
        file_content = file.read()
        
        # Запускаем твою логику
        excel_file = await convert_pdf_to_excel(file_content)
        
        # Отправляем результат
        await message.answer_document(
            document=BufferedInputFile(excel_file.read(), filename="result.xlsx"),
            caption="✅ Готово! Вот твой Excel файл."
        )
        await status_msg.delete()
        
    except Exception as e:
        await status_msg.edit_text(f"❌ Произошла ошибка: {e}")

async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())