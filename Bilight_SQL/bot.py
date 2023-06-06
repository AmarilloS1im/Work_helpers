import os
from aiogram import Bot, Dispatcher, executor, types
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.dispatcher import FSMContext
from aiogram.dispatcher.filters.state import State, StatesGroup
from aiogram.types import ContentType
from dotenv import find_dotenv, load_dotenv
import emoji

load_dotenv(find_dotenv())

bot = Bot(token=os.getenv('API_TOKEN'))

storage = MemoryStorage()

dp = Dispatcher(bot, storage=storage)


class FSMAdmin(StatesGroup):
    upload_new_data = State()
    download_new_data = State()
    upload_certificates = State()


next_button = types.InlineKeyboardButton('ДАЛЕЕ', callback_data='next')

button_back = types.InlineKeyboardButton('НАЗАД', callback_data='back')
markup_back = types.InlineKeyboardMarkup(row_width=1)
markup_back.add(button_back)

upload_button = types.InlineKeyboardButton('Загрузить данные', callback_data='upload_data')
get_data_button = types.InlineKeyboardButton('Получить данные из базы', callback_data='get_data')
markup_start_screen = types.InlineKeyboardMarkup(row_width=2)
markup_start_screen.add(upload_button, get_data_button)

upload_certificates = types.InlineKeyboardButton('Загрузить сертификаты в базу', callback_data='certificates_upload')
upload_products = types.InlineKeyboardButton('Загрузить артикулы в базу', callback_data='products_upload')
markup_upload_certificates = types.InlineKeyboardMarkup(row_width=2)
markup_upload_certificates.add(upload_certificates, next_button, button_back)

button_back = types.InlineKeyboardButton('НАЗАД', callback_data='back')
markup_back = types.InlineKeyboardMarkup(row_width=1)
markup_back.add(button_back)


@dp.message_handler(commands=['start'])
async def send_welcome(message: types.Message):
    print('bot  on line')
    user_full_name = message.from_user.full_name
    await message.reply(f"Привет {user_full_name}\nЯ Bilight_Bot\n", reply_markup=markup_start_screen)


@dp.callback_query_handler(text='upload_data')
async def callback_upload(callback: types.CallbackQuery):
    await callback.message.answer(f"СНАЧАЛА ЗАГРУЗИТЕ СЕРТИФИКАТЫ!\n"
                                  f"ЕСЛИ СЕРТИФИКАТЫ УЖЕ ЗАГРУЖЕННЫ, НАЖМИТЕ 'ДАЛЕЕ' ",
                                  reply_markup=markup_upload_certificates)


@dp.callback_query_handler(text='certificates_upload')
async def callback_certificates_upload(callback: types.CallbackQuery):
    await FSMAdmin.upload_certificates.set()
    await callback.message.answer(f"ЧТОБЫ ЗАГРУЗИТЕ ДОКУМЕНТ\n"
                                  f"НАЖМИТЕ НА {emoji.emojize(':paperclip:')}", reply_markup=markup_back)

@dp.message_handler(content_types=types.ContentType.ANY, state=FSMAdmin.upload_certificates)
async def load_certificates_to_postgresql(message: types.Message, state: FSMContext):
    if message.content_type != 'document':
        await FSMAdmin.upload_certificates.set()
        await message.answer('Загружать файлы можно только в формате .xlsx',
                             reply_markup=markup_back)
    else:
        print('ok')
        await state.finish()



@dp.callback_query_handler(text='back', state=[FSMAdmin.upload_new_data,FSMAdmin.download_new_data,
                                               FSMAdmin.upload_certificates, None])
async def callback_back_button(callback: types.CallbackQuery, state: FSMContext):
    await state.finish()
    user_full_name = callback.message.from_user.full_name
    await callback.message.reply(f"Привет {user_full_name}\nЯ Bilight_Bot\n", reply_markup=markup_start_screen)


executor.start_polling(dp, skip_updates=True)
