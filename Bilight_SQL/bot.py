from aiogram import Bot, Dispatcher, executor, types
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.dispatcher import FSMContext
from aiogram.dispatcher.filters.state import State, StatesGroup

import emoji
from Refactor import *

load_dotenv(find_dotenv())

bot = Bot(token=os.getenv('API_TOKEN'))

storage = MemoryStorage()

dp = Dispatcher(bot, storage=storage)


# region FSMAdmin
class FSMAdmin(StatesGroup):
    upload_new_data = State()
    upload_certificates = State()
    upload_products = State()
    check_dup = State()
    get_data = State()
# endregion


# region buttons

# region BACK
button_back = types.InlineKeyboardButton('НАЗАД', callback_data='back')
markup_back = types.InlineKeyboardMarkup(row_width=1)
markup_back.add(button_back)
# endregion

# region YES/NO DUPLICATES
duplicate_button_yes = types.InlineKeyboardButton('ДА', callback_data='duplicate_yes')
duplicate_button_no = types.InlineKeyboardButton('НЕТ', callback_data='duplicate_no')
markup_duplicate_question = types.InlineKeyboardMarkup(row_width=2)
markup_duplicate_question.add(duplicate_button_yes, duplicate_button_no, button_back)
# endregion

# region YES/NO
button_yes = types.InlineKeyboardButton('ДА', callback_data='yes')
button_no = types.InlineKeyboardButton('НЕТ', callback_data='no')
markup_question = types.InlineKeyboardMarkup(row_width=2)
markup_question.add(button_yes, button_no, button_back)
# endregion

# region UPLOAD/DOWNLOAD
upload_button = types.InlineKeyboardButton('Загрузить данные', callback_data='upload_data')
download_button = types.InlineKeyboardButton('Получить данные из базы', callback_data='get_data')
markup_start_screen = types.InlineKeyboardMarkup(row_width=2)
markup_start_screen.add(upload_button, download_button)
# endregion

# region UPLOAD/DOWNLOAD CERTS OR PROD
upload_certificates = types.InlineKeyboardButton('Загрузить сертификаты в базу', callback_data='certificates_upload')
upload_products = types.InlineKeyboardButton('Загрузить артикулы в базу', callback_data='products_upload')
markup_upload_certificates_products = types.InlineKeyboardMarkup(row_width=1)
markup_upload_certificates_products.add(upload_certificates, upload_products, button_back)

duplicate_cert_button_yes = types.InlineKeyboardButton('Да', callback_data='cert_duplicate_yes')
duplicate_cert_button_no = types.InlineKeyboardButton('Нет', callback_data='cert_duplicate_no')
markup_cert_duplicate_question = types.InlineKeyboardMarkup(row_width=2)
markup_cert_duplicate_question.add(duplicate_cert_button_yes, duplicate_cert_button_no, button_back)


get_data_by_supplier_button = types.InlineKeyboardButton('Получить информацию о поставщике', callback_data='supplier_info')
get_data_by_cert_button = types.InlineKeyboardButton('Получить информацию о сертификатах', callback_data='cert_info')
markup_get_data = types.InlineKeyboardMarkup(row_width=1)
markup_get_data.add(get_data_by_supplier_button, get_data_by_cert_button, button_back)


# endregion

# endregion


# region start screen
@dp.message_handler(commands=['start'])
async def send_welcome(message: types.Message):
    print('bot  on line')
    user_full_name = message.from_user.full_name
    await message.reply(f"Привет {user_full_name}\nЯ Bilight_Bot\n", reply_markup=markup_start_screen)


# endregion


# region callback handlers
@dp.callback_query_handler(text='upload_data')
async def callback_upload(callback: types.CallbackQuery):
    await callback.message.answer(f"ВНИМАНИЕ! СНАЧАЛА ЗАГРУЗИТЕ СЕРТИФИКАТЫ!\nЕСЛИ СЕРТИФИКАТЫ УЖЕ ЗАГРУЖЕННЫ,"
                                  f" МОЖЕТЕ ЗАГРУЖАТЬ АРТИКУЛЫ ", reply_markup=markup_upload_certificates_products)


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

@dp.callback_query_handler(text='get_data')
async def callback_products_upload(callback: types.CallbackQuery):
    await FSMAdmin.get_data.set()
    await callback.message.answer(f"ЧТОБЫ ЗАГРУЗИТЕ ДОКУМЕНТ\n"
                                  f"НАЖМИТЕ НА {emoji.emojize(':paperclip:')}", reply_markup=markup_get_data)


# endregion


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
            unique_data = find_cert_unique_data(universal_query('certificates', '*'), cert_data_from_user)
            duplicate_data = find_cert_duplicate_data(universal_query('certificates', '*'), doc)
            async with state.proxy() as data:
                data['doc'] = doc
                data['cert_data_from_user'] = cert_data_from_user
                data['unique_data'] = unique_data
                data['duplicate_data'] = duplicate_data
            if duplicate_data:
                await message.reply(f"В базе данные обнаружены дубликаты данных по следующим ID"
                                    f" {duplicate_data}. Заменить данные? y/n?",
                                    reply_markup=markup_cert_duplicate_question)

            else:
                add_certificates(unique_data)
                await message.reply(f"Cертификаты закгружены",
                                    reply_markup=markup_back)
                await state.finish()


@dp.callback_query_handler(text='cert_duplicate_yes', state=[FSMAdmin.upload_certificates])
async def callback_yes_cert(callback: types.CallbackQuery, state=FSMContext):
    async with state.proxy() as data:
        unique_data = data['unique_data']
        duplicate_data = data['duplicate_data']
    add_certificates(unique_data)
    add_cert_duplicate_user_data(duplicate_data)
    await callback.message.reply(f"Уникальные сертификаты закгружены, дубликаты сертификатов в базе обновлены",
                                 reply_markup=markup_back)
    await state.finish()


@dp.callback_query_handler(text='cert_duplicate_no', state=[FSMAdmin.upload_certificates])
async def callback_no_cert(callback: types.CallbackQuery, state=FSMContext):
    async with state.proxy() as data:
        unique_data = data['unique_data']
    add_certificates(unique_data)
    await callback.message.reply(f"Уникальные сертификаты успешно добавлены в базу данных", reply_markup=markup_back)
    await state.finish()


@dp.message_handler(content_types=types.ContentType.ANY, state=[FSMAdmin.upload_products, FSMAdmin.check_dup])
async def load_products_to_postgresql(message: types.Message, state: FSMContext):
    if message.content_type != 'document':
        await FSMAdmin.upload_products.set()
        await message.answer('Загружать файлы можно только в формате .xlsx',
                             reply_markup=markup_back)
    else:
        file_extantion = '.' + message.document.file_name.split('.')[-1]
        if file_extantion != '.xlsx' and file_extantion != '.xls':
            await FSMAdmin.upload_products.set()
            await message.answer('Документ должен быть в формате .xls или .xlsx', reply_markup=markup_back)
        else:
            doc = message.document.file_name
            data_from_user = get_product_info_from_user(doc)
            manufacturers_dict = make_dict(universal_query('manufacturers', '*'))
            approved_manufacterers_data = approved_manufactureres_data_list(data_from_user, manufacturers_dict)
            manufacturers_data_to_check = manufacturers_to_check_data_list(data_from_user)
            unique_data = find_unique_data(universal_query('bilight_products', '*'), data_from_user)
            duplicate_data = find_duplicate_data(universal_query('bilight_products', '*'), doc)
            possible_change = make_possible_change_dict(manufacturers_dict, data_from_user)
            yes = False
            yes_without_make_pos_change_file = False
            async with state.proxy() as data:
                data['doc'] = doc
                data['data_from_user'] = data_from_user
                data['manufacturers_dict'] = manufacturers_dict
                data['approved_manufacterers_data'] = approved_manufacterers_data
                data['manufacturers_data_to_check'] = manufacturers_data_to_check
                data['unique_data'] = unique_data
                data['duplicate_data'] = duplicate_data
                data['possible_change'] = possible_change
                data['yes'] = yes
                data['yes_without_make_pos_change_file'] = yes_without_make_pos_change_file
            if possible_change:
                await FSMAdmin.upload_products.set()
                await message.reply(f"В вашем файле присуствуют поставщики которых"
                                    f" нет в списке поставщиков в базе даных\n"
                                    f"Если хотите проверить данные, нажмите 'ДА' и программа"
                                    f" отправит вам файл с возможными заменами, артикулы с корректными "
                                    f"поставщиками будут згружены\n"
                                    f"если вы уверены в своем выборе нажмите 'НЕТ' и программа добавит в"
                                    f"таблицу новых поставщиков", reply_markup=markup_question)

            else:
                if duplicate_data:
                    await FSMAdmin.check_dup.set()
                    yes_without_make_pos_change_file = True
                    async with state.proxy() as data:
                        data['yes_without_make_pos_change_file'] = yes_without_make_pos_change_file

                    await message.reply(f"В базе данные обнаружены дубликаты данных"
                                        f" по следующим ID {duplicate_data}."
                                        f" Заменить данные?", reply_markup=markup_duplicate_question)
                else:
                    converted_manufacturers_data = convert_manufacturers_to_digit(data_from_user)
                    add_products(converted_manufacturers_data)
                    await message.reply(
                        f"Артикулы успешно загружены", reply_markup=markup_back)
                    await state.finish()


@dp.callback_query_handler(text='yes', state=[FSMAdmin.check_dup, FSMAdmin.upload_products])
async def callback_yes_prod(callback: types.CallbackQuery, state=FSMContext):
    async with state.proxy() as data:
        doc = data['doc']
        approved_manufacterers_data = data['approved_manufacterers_data']
        duplicate_data = data['duplicate_data']
        possible_change = data['possible_change']
    if duplicate_data:
        yes = True
        async with state.proxy() as data:
            data['yes'] = yes
        await FSMAdmin.check_dup.set()
        await callback.message.reply(f"В базе данные обнаружены дубликаты данных"
                                     f" по следующим ID {duplicate_data}."
                                     f" Заменить данные?", reply_markup=markup_duplicate_question)
    else:
        get_replace_file(doc, possible_change)
        converted_manufacturers_data = convert_manufacturers_to_digit(approved_manufacterers_data)
        add_products(converted_manufacturers_data)

        reply_possibel_changes = open(r"possible_changes.xlsx", 'rb')
        await callback.message.reply_document(reply_possibel_changes)
        await callback.message.reply(
            f"Артикулы с корректными поставщиками  загружены, "
            f"предполагаемые замены поставщиков в подготовленном файле", reply_markup=markup_back)
        await state.finish()


@dp.callback_query_handler(text='no', state=[FSMAdmin.upload_products, FSMAdmin.check_dup])
async def callback_no_prod(callback: types.CallbackQuery, state=FSMContext):
    add_permition = True
    async with state.proxy() as data:
        data['add_permition'] = add_permition
    async with state.proxy() as data:
        data_from_user = data['data_from_user']
        manufacturers_dict = data['manufacturers_dict']
        unique_data = data['unique_data']
        duplicate_data = data['duplicate_data']

    if duplicate_data:
        await FSMAdmin.check_dup.set()
        await callback.message.reply(f"В базе данные обнаружены дубликаты данных"
                                     f" по следующим ID {duplicate_data}."
                                     f" Заменить данные?", reply_markup=markup_duplicate_question)
    else:
        add_new_manufacturers(manufacturers_dict, data_from_user)
        converted_manufacturers_data = convert_manufacturers_to_digit(unique_data)
        add_products(converted_manufacturers_data)
        await callback.message.reply(f"Артикулы и новые поставщики  успешно добавлены в базу данных",
                                     reply_markup=markup_back)
        await state.finish()


@dp.callback_query_handler(text='duplicate_yes', state=[FSMAdmin.check_dup, FSMAdmin.upload_products])
async def callback_yes_duplicate(callback: types.CallbackQuery, state=FSMContext):
    await FSMAdmin.check_dup.set()
    async with state.proxy() as data:
        yes = data['yes']
        doc = data['doc']
        duplicate_data = data['duplicate_data']
        possible_change = data['possible_change']
        manufacturers_dict = data['manufacturers_dict']
        data_from_user = data['data_from_user']
        unique_data = data['unique_data']
        yes_without_make_pos_change_file = data['yes_without_make_pos_change_file']
    if yes:
        get_replace_file(doc, possible_change)
        reply_possibel_changes = open(r"possible_changes.xlsx", 'rb')
        converted_manufacturers_data = convert_manufacturers_to_digit(duplicate_data)
        add_duplicate_user_data(converted_manufacturers_data)
        await callback.message.reply_document(reply_possibel_changes)
        await callback.message.reply(
            f"Артикулы с корректными поставщиками  загружены, "
            f"предполагаемые замены поставщиков в подготовленном файле", reply_markup=markup_back)
        await state.finish()
    elif yes_without_make_pos_change_file:
        converted_manufacturers_data = convert_manufacturers_to_digit(duplicate_data)
        add_duplicate_user_data(converted_manufacturers_data)
        await callback.message.reply(f"Артикулы успешно заменены",
                                     reply_markup=markup_back)
        await state.finish()
    else:
        add_new_manufacturers(manufacturers_dict, data_from_user)
        converted_manufacturers_data = convert_manufacturers_to_digit(duplicate_data)
        add_duplicate_user_data(converted_manufacturers_data)
        converted_manufacturers_data = convert_manufacturers_to_digit(unique_data)
        add_unique_user_data(converted_manufacturers_data)

        await callback.message.reply(f"Артикулы и новые поставщики  успешно добавлены в базу данных",
                                     reply_markup=markup_back)
        await state.finish()


@dp.callback_query_handler(text='duplicate_no', state=[FSMAdmin.check_dup, FSMAdmin.upload_products])
async def callback_yes_duplicate(callback: types.CallbackQuery, state=FSMContext):
    async with state.proxy() as data:
        unique_data = data['unique_data']
        add_permition = data['add_permition']
        manufacturers_dict = data['manufacturers_dict']
        data_from_user = data['data_from_user']
    if add_permition:
        add_new_manufacturers(manufacturers_dict, data_from_user)
        converted_manufacturers_data = convert_manufacturers_to_digit(unique_data)
        add_unique_user_data(converted_manufacturers_data)
        await callback.message.reply(
            f"Работа завершена, дубликаты данныех пользователя в базе остались без изменений, уникальные артикулы "
            f"добавлены, новые поставщики добавлены",
            reply_markup=markup_back)
        await state.finish()
    else:
        converted_manufacturers_data = convert_manufacturers_to_digit(unique_data)
        add_unique_user_data(converted_manufacturers_data)
        await callback.message.reply(
            f"Работа завершена, дубликаты данныех пользователя в базе остались без изменений, уникальные артикулы "
            f"добавлены",
            reply_markup=markup_back)
        await state.finish()


# region back button
@dp.callback_query_handler(text='back', state=[FSMAdmin.upload_new_data, FSMAdmin.get_data,
                                               FSMAdmin.upload_certificates, FSMAdmin.upload_products,
                                               FSMAdmin.check_dup, None])
async def callback_back_button(callback: types.CallbackQuery, state: FSMContext):
    await state.finish()
    user_full_name = callback.message.from_user.full_name
    await callback.message.reply(f"Привет {user_full_name}\nЯ Bilight_Bot\n", reply_markup=markup_start_screen)


# endregion


executor.start_polling(dp, skip_updates=True)
