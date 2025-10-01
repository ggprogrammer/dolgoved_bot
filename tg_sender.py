from telethon import TelegramClient, events
import asyncio
import os
import time

# Ваши данные из my.telegram.org
api_id = '25651117'  # число
api_hash = '57f4fec3c64805c000d5cee4a180be0c'  # строка
bot_username = '@vika_iit_bot'
client = TelegramClient('anon', api_id, api_hash)
BASE_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'Dolgoved')
READFILE_DIR = os.path.join(BASE_DIR, 'READ_FILE')

async def send_receive_xlsx():
    await client.send_message(bot_username, '/start')
    print("Start отправлено")
    time.sleep(3)
    await client.send_message(bot_username, 'Расписание пересдач')
    print("Расписание пересдач отправлено")

    result = await client.get_messages(bot_username, limit=4)
    for i in range(4):
        if "Инструкция по ликвидации задолженности" in result[i].text:
            data = await result[i].download_media()
            os.rename(os.path.join(BASE_DIR, data), os.path.join(READFILE_DIR, data))
            return data
def run_sender():
    with client:
        return client.loop.run_until_complete(send_receive_xlsx())