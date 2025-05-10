import pandas as pd
from aiogram import Bot, Dispatcher, F, types
from aiogram.types import Message
from aiogram.fsm.storage.memory import MemoryStorage
import asyncio
import os

BOT_TOKEN = os.getenv("BOT_TOKEN")

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher(storage=MemoryStorage())

@dp.message(F.document)
async def handle_excel(message: Message):
    document = message.document
    if not document.file_name.endswith(".xlsx"):
        await message.reply("❗️")
        return

    file_path = f"{document.file_id}.xlsx"
    await bot.download(document, destination=file_path)

    try:
        df = pd.read_excel(file_path)
        df.columns = df.columns.str.strip()

        if df.shape[1] < 2:
            await message.reply("❗️")
            return

        df = df.iloc[:, :2]
        df.columns = ['Nom', 'Artikul']
        df = df.dropna()

        counts = df.groupby(['Nom', 'Artikul']).size().reset_index(name='Soni')
        javob = ""
        umumiy_soni = len(df)
        yagona_nomlar = df['Nom'].nunique()
        name_count= df.groupby('Nom').size().reset_index(name='Nechta')

        for _, row in counts.iterrows():
            javob += f"{row['Nom']} {row['Artikul']} > {int(row['Soni'])}ta\n"

        javob += f"\n📊 Jami: {umumiy_soni} ta"
        javob += f"\n🔢 Nomlar: {yagona_nomlar} ta"

        javob += "\n📦 Xar biridan:\n"
        for _, row in name_count.iterrows():
            javob += f"- {row['Nom']} > {row['Nechta']} ta\n"

        await message.reply(javob.strip())

    except Exception as e:
        await message.reply(f"❌ : {e}")
    finally:
        if os.path.exists(file_path):
            os.remove(file_path)

@dp.message(F.text)
async def start_message(message: Message):
    await message.reply("✅ I'm here send .xlsx")

async def main():
    await dp.start_polling(bot)

if __name__ == "__main__":
    asyncio.run(main())
