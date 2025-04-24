# API_TOKEN = '7764315553:AAF4FwNA1Zanv22FmGru56bB67NmhTT_DnE'

import logging
import pandas as pd
from aiogram import Bot, Dispatcher, types
from aiogram.enums import ParseMode
from aiogram.types import Message
from aiogram.filters import CommandStart
from aiogram import F
from aiogram.client.default import DefaultBotProperties
import asyncio
import re

BOT_TOKEN = "7764315553:AAF4FwNA1Zanv22FmGru56bB67NmhTT_DnE"

logging.basicConfig(level=logging.INFO)

bot = Bot(token=BOT_TOKEN, default=DefaultBotProperties(parse_mode=ParseMode.HTML))
dp = Dispatcher()

def load_data():
    df_sheets = pd.read_excel("List_Priority.xlsx", sheet_name=None)
    return df_sheets['List'].fillna(''), df_sheets['Priority'].fillna('')

list_data, priority_data = load_data()

@dp.message(CommandStart())
async def send_welcome(message: Message):
    await message.answer("Добро пожаловать! Введите код ОКВЭД 2 или ОКПД 2 (например, 24 или 24.10.22):")

@dp.message()
async def handle_code(message: Message):
    code = message.text.strip()

    # Определение типа кода: ОКВЭД или ОКПД
    is_okpd = bool(re.match(r'^\d+\.\d+', code))

    if is_okpd:
        list_filtered = list_data[list_data.iloc[:, 2].astype(str) == code]
        priority_filtered = priority_data[priority_data.iloc[:, 2].astype(str) == code]
    else:
        list_filtered = list_data[list_data.iloc[:, 1].astype(str) == code]
        priority_filtered = priority_data[priority_data.iloc[:, 1].astype(str) == code]

    if list_filtered.empty and priority_filtered.empty:
        await message.answer(f"Код '{code}' не найден в базе данных. Попробуйте снова.")
        return

    if not list_filtered.empty:
        list_responses = []
        for _, row in list_filtered.iterrows():
            row_text = (
                f"<b>Код ОКПД 2:</b> {row.iloc[2]}\n"
                f"<b>Описание:</b> {row.iloc[3]}\n"
                f"<b>Ссылка:</b> {row.iloc[4]}"
            )
            list_responses.append(row_text)
        await message.answer("<b>Результаты из листа List:</b>\n\n" + "\n\n".join(list_responses))

    if not priority_filtered.empty:
        priority_responses = []
        for _, row in priority_filtered.iterrows():
            row_text = (
                f"<b>Код ОКПД 2:</b> {row.iloc[2]}\n"
                f"<b>Описание:</b> {row.iloc[3]}\n"
                f"<b>Ссылка:</b> {row.iloc[4]}"
            )
            priority_responses.append(row_text)
        await message.answer("<b>Результаты из листа Priority:</b>\n\n" + "\n\n".join(priority_responses))

    await message.answer("Введите следующий код ОКВЭД 2 или ОКПД 2:")

if __name__ == "__main__":
    import asyncio
    asyncio.run(dp.start_polling(bot))