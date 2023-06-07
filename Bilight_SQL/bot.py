import os
from aiogram import Bot, Dispatcher, executor, types
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.dispatcher import FSMContext
from aiogram.dispatcher.filters.state import State, StatesGroup
from aiogram.types import ContentType
from dotenv import find_dotenv, load_dotenv
import emoji
from main import *

load_dotenv(find_dotenv())

bot = Bot(token=os.getenv('API_TOKEN'))

storage = MemoryStorage()

dp = Dispatcher(bot, storage=storage)


class FSMAdmin(StatesGroup):
    upload_new_data = State()
    download_new_data = State()
    upload_certificates = State()
    upload_products = State()


button_back = types.InlineKeyboardButton('НАЗАД', callback_data='back')
markup_back = types.InlineKeyboardMarkup(row_width=1)
markup_back.add(button_back)

button_yes = types.InlineKeyboardButton('ДА', callback_data='yes')
button_no = types.InlineKeyboardButton('НЕТ', callback_data='no')
markup_question = types.InlineKeyboardMarkup(row_width=2)
markup_question.add(button_yes, button_no, button_back)

upload_button = types.InlineKeyboardButton('Загрузить данные', callback_data='upload_data')
get_data_button = types.InlineKeyboardButton('Получить данные из базы', callback_data='get_data')
markup_start_screen = types.InlineKeyboardMarkup(row_width=2)
markup_start_screen.add(upload_button, get_data_button)

upload_certificates = types.InlineKeyboardButton('Загрузить сертификаты в базу', callback_data='certificates_upload')
upload_products = types.InlineKeyboardButton('Загрузить артикулы в базу', callback_data='products_upload')
markup_upload_certificates = types.InlineKeyboardMarkup(row_width=1)
markup_upload_certificates.add(upload_certificates, upload_products, button_back)

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
    await callback.message.answer(f"ВНИМАНИЕ! СНАЧАЛА ЗАГРУЗИТЕ СЕРТИФИКАТЫ!\nЕСЛИ СЕРТИФИКАТЫ УЖЕ ЗАГРУЖЕННЫ,"
                                  f" МОЖЕТЕ ЗАГРУЖАТЬ АРТИКУЛЫ ", reply_markup=markup_upload_certificates)


@dp.callback_query_handler(text='certificates_upload')
async def callback_certificates_upload(callback: types.CallbackQuery):
    await FSMAdmin.upload_certificates.set()
    await callback.message.answer(f"ЧТОБЫ ЗАГРУЗИТЕ ДОКУМЕНТ\n"
                                  f"НАЖМИТЕ НА {emoji.emojize(':paperclip:')}", reply_markup=markup_back)


@dp.callback_query_handler(text='products_upload')
async def callback_products_upload(callback: types.CallbackQuery):
    await FSMAdmin.upload_products.set()
    await callback.message.answer(f"ЧТОБЫ ЗАГРУЗИТЕ ДОКУМЕНТ\n"
                                  f"НАЖМИТЕ НА {emoji.emojize(':paperclip:')}", reply_markup=markup_back)


@dp.message_handler(content_types=types.ContentType.ANY, state=FSMAdmin.upload_certificates)
async def load_certificates_to_postgresql(message: types.Message, state: FSMContext):
    if message.content_type != 'document':
        await FSMAdmin.upload_certificates.set()
        await message.answer('Загружать файлы можно только в формате .xlsx',
                             reply_markup=markup_back)
    else:
        file_extantion = '.' + message.document.file_name.split('.')[-1]
        if file_extantion != '.xlsx' and file_extantion != '.xls':
            await FSMAdmin.upload_certificates.set()
            await message.answer('Документ должен быть в формате .xls или .xlsx', reply_markup=markup_back)
        else:
            doc = message.document.file_name
            cert_data_from_user = get_cert_info_from_user(doc)
            duplicate_user_data = \
                find_unique_and_duplicate_data(universal_query('certificates', 'certificate_id'), cert_data_from_user)[
                    -1]
            unique_user_data = \
                find_unique_and_duplicate_data(universal_query('certificates', 'certificate_id'), cert_data_from_user)[
                    0]
            if duplicate_user_data:
                await message.reply(f"В базе данные обнаружены дубликаты данных по следующим ID"
                                     f" {duplicate_user_data}. Заменить данные? y/n?", reply_markup=markup_question)

                @dp.callback_query_handler(text='yes', state=[FSMAdmin.upload_certificates])
                async def callback_yes(callback: types.CallbackQuery):
                    await state.finish()
                    await add_duplictes(duplicate_user_data, unique_user_data, cert_data_from_user, 'y')

            else:
                add_duplictes(duplicate_user_data, unique_user_data, cert_data_from_user, 'no')
            await state.finish()


@dp.callback_query_handler(text='back', state=[FSMAdmin.upload_new_data, FSMAdmin.download_new_data,
                                               FSMAdmin.upload_certificates, FSMAdmin.upload_products, None])
async def callback_back_button(callback: types.CallbackQuery, state: FSMContext):
    await state.finish()
    user_full_name = callback.message.from_user.full_name
    await callback.message.reply(f"Привет {user_full_name}\nЯ Bilight_Bot\n", reply_markup=markup_start_screen)


executor.start_polling(dp, skip_updates=True)
