# region IMPORT
from aiogram import Bot, Dispatcher, executor, types
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.dispatcher import FSMContext
from aiogram.dispatcher.filters.state import State, StatesGroup
import emoji
from core_functions import *

load_dotenv(find_dotenv())
# endregion

# region INITIALIZE BOT
bot = Bot(token=os.getenv('API_TOKEN'))
storage = MemoryStorage()
dp = Dispatcher(bot, storage=storage)


# endregion


# region FSMAdmin
class FSMAdmin(StatesGroup):
    upload_new_data = State()
    upload_certificates = State()
    scan_upload = State()
    upload_products = State()
    check_dup = State()
    get_manufacturers_data = State()
    get_certificates_data = State()
    get_tnved_data = State()
    get_all_together = State()
    download_templates = State()
    get_info_template = State()
    get_scan_by_article = State()


# endregion


# region BUTTONS


# region BACK
button_back = types.InlineKeyboardButton('НАЗАД', callback_data='back')
markup_back = types.InlineKeyboardMarkup(row_width=1)
markup_back.add(button_back)
# endregion

# region YES/NO DUPLICATES
duplicate_button_yes = types.InlineKeyboardButton('Заменить', callback_data='duplicate_yes')
duplicate_button_no = types.InlineKeyboardButton('НЕТ', callback_data='duplicate_no')
markup_duplicate_question = types.InlineKeyboardMarkup(row_width=2)
markup_duplicate_question.add(duplicate_button_yes, duplicate_button_no, button_back)
# endregion

# region YES/NO
button_yes = types.InlineKeyboardButton('Проверить', callback_data='yes')
button_no = types.InlineKeyboardButton('Я уверен', callback_data='no')
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
templates_for_upload_button = types.InlineKeyboardButton('Скачать шаблоны для загрузки артикулов или сертификатов',
                                                         callback_data='templates')
upload_scan_button = types.InlineKeyboardButton('Загрузить сканы сертификатов', callback_data='scan_upload')
markup_upload_certificates_products = types.InlineKeyboardMarkup(row_width=1)
markup_upload_certificates_products.add(upload_certificates, upload_scan_button, upload_products,
                                        templates_for_upload_button, button_back)

button_scan_yes = types.InlineKeyboardButton('ЗАМЕНИТЬ', callback_data='replace_scan_yes')
button_scan_no = types.InlineKeyboardButton('НЕТ', callback_data='replace_scan_no')
markup_scan = types.InlineKeyboardMarkup(row_width=1)
markup_scan.add(button_scan_yes, button_scan_no, button_back)

download_template_articles = types.InlineKeyboardButton('Скачать шаблон для загрузки артикулов',
                                                        callback_data='download_template_art')
download_template_certificates = types.InlineKeyboardButton('Скачать шаблон для загрузки сертификатов',
                                                            callback_data='download_template_cert')
markup_download_templates = types.InlineKeyboardMarkup(row_width=1)
markup_download_templates.add(download_template_articles, download_template_certificates, button_back)

get_info_template = types.InlineKeyboardButton("Скачать шаблон для получения данных", callback_data='get_info_template')
markup_get_info_template = types.InlineKeyboardMarkup(row_width=1)
markup_get_info_template.add(get_info_template, button_back)

duplicate_cert_button_yes = types.InlineKeyboardButton('Заменить', callback_data='cert_duplicate_yes')
duplicate_cert_button_no = types.InlineKeyboardButton('Нет', callback_data='cert_duplicate_no')
markup_cert_duplicate_question = types.InlineKeyboardMarkup(row_width=2)
markup_cert_duplicate_question.add(duplicate_cert_button_yes, duplicate_cert_button_no, button_back)

get_tnvd_code_button = types.InlineKeyboardButton("Получить коды ТНВЭД", callback_data='tnvd_code_upload')

get_data_by_supplier_button = types.InlineKeyboardButton('Получить информацию о производителе',
                                                         callback_data='manufacturer_info')
get_data_by_cert_button = types.InlineKeyboardButton('Получить информацию о сертификатах', callback_data='cert_info')

get_all_data_button = types.InlineKeyboardButton('ВСЕ ВМЕСТЕ', callback_data='all_together')

get_scan_certificate_by_article_button = types.InlineKeyboardButton('ПОЛУЧИТЬ СКАН СЕРТИФИКАТА ПО АРТИКУЛУ',
                                                                    callback_data='get_scan_by_article')

markup_get_data = types.InlineKeyboardMarkup(row_width=1)
markup_get_data.add(get_data_by_supplier_button, get_data_by_cert_button, get_tnvd_code_button, get_all_data_button,
                    get_scan_certificate_by_article_button, button_back)


# endregion

# endregion

# region Checkig errors function
def is_error(uuid):
    if f'{uuid}_error.log' in os.listdir():
        with open(f'{uuid}_error.log', mode='r', encoding='utf-8') as error_file:
            error_info = error_file.readline()
        error_file.close()
        return error_info
    else:
        return False


# endregion

# region delete doc and uuid file function
def delete_doc_and_uuid_file(uuid_file_name, doc):
    if uuid_file_name in os.listdir():
        os.remove(uuid_file_name)
    if doc in os.listdir():
        os.remove(doc)


# endregion


# region START SCREEN
@dp.message_handler(commands=['start'])
async def send_welcome(message: types.Message):
    print('bot  on line')
    user_full_name = message.from_user.username
    await message.reply(f"Привет {user_full_name}\nЯ Bilight_Bot\n"
                        f"Я помогаю найти информацию о сертификатах, производителях, кодах ТНВЭД"
                        f" по артикулам из базы ТД Билайт.\n"
                        f"Если хочешь получить данные о произвоидтелях, сертификатах или кодах ТНВЭД"
                        f" нажми кнопку 'Получть данные из базы' или кнопку 'Загрузить данные',"
                        f" если хочешь добавить новую информацию в базу.", reply_markup=markup_start_screen)


# endregion


# region CALLBACK HANDLERS
@dp.callback_query_handler(text='upload_data')
async def callback_upload(callback: types.CallbackQuery):
    await callback.message.answer(f"ВНИМАНИЕ!\n\nСНАЧАЛА ЗАГРУЗИТЕ СЕРТИФИКАТЫ!\n\nЕСЛИ СЕРТИФИКАТЫ УЖЕ ЗАГРУЖЕНЫ,"
                                  f" МОЖЕТЕ ЗАГРУЖАТЬ АРТИКУЛЫ ", reply_markup=markup_upload_certificates_products)


@dp.callback_query_handler(text='certificates_upload')
async def callback_certificates_upload(callback: types.CallbackQuery):
    await FSMAdmin.upload_certificates.set()
    await callback.message.answer(f"ЧТОБЫ ЗАГРУЗИТЕ ДОКУМЕНТ\n"
                                  f"НАЖМИТЕ НА {emoji.emojize(':paperclip:')}", reply_markup=markup_back)


@dp.callback_query_handler(text='scan_upload')
async def callback_scan_upload(callback: types.CallbackQuery):
    await FSMAdmin.scan_upload.set()
    await callback.message.answer(f"ЧТОБЫ ЗАГРУЗИТЬ СКАН СЕРТИФИКАТА\n"
                                  f"НАЖМИТЕ НА {emoji.emojize(':paperclip:')}\n\n"
                                  f"ВНИМАНИЕ!\nНАЗВАНИЕ ФАЙЛА ОТСКАНИРОВАННОГО СЕРТИФИКАТА"
                                  f" ДОЛЖНО БЫТЬ ИДЕНТИЧНО УНИКАЛЬНОМУ НОМЕРУ "
                                  f"СЕРТИФИКАТА КОТОРЫЙ ВЫ ЗАГРУЗИЛИ\n(см. столбец A) в шаблоне"
                                  f" 'cert_upload_template.xlsx'"
                                  f" скачать шаблон можно в соотвествующем разделе бота.", reply_markup=markup_back)


@dp.callback_query_handler(text='products_upload')
async def callback_products_upload(callback: types.CallbackQuery):
    await FSMAdmin.upload_products.set()
    await callback.message.answer(f"ЧТОБЫ ЗАГРУЗИТЕ ДОКУМЕНТ\n"
                                  f"НАЖМИТЕ НА {emoji.emojize(':paperclip:')}", reply_markup=markup_back)


@dp.callback_query_handler(text='templates')
async def callback_templates_download(callback: types.CallbackQuery):
    await FSMAdmin.download_templates.set()
    await callback.message.answer(f"Нажмите на соотвествующие кнопки, чтобы скачать шаблон для"
                                  f" загрузки артикулов или сертификатов", reply_markup=markup_download_templates)


@dp.callback_query_handler(text='download_template_art', state=[FSMAdmin.download_templates])
async def callback_download_templates_art(callback: types.CallbackQuery, state=FSMContext):
    await FSMAdmin.download_templates.set()
    reply_template_art = open(r"products_upload_template.xlsx", 'rb')
    await callback.message.reply_document(reply_template_art)

    await callback.message.answer(f"Шаблон готов к скачиванию", reply_markup=markup_back)


@dp.callback_query_handler(text='download_template_cert', state=[FSMAdmin.download_templates])
async def callback_download_templates_cert(callback: types.CallbackQuery, state=FSMContext):
    await FSMAdmin.download_templates.set()
    reply_template_cert = open(r"cert_upload_template.xlsx", 'rb')
    await callback.message.reply_document(reply_template_cert)

    await callback.message.answer(f"Шаблон готов к скачиванию", reply_markup=markup_back)


@dp.callback_query_handler(text='get_info_template',
                           state=[FSMAdmin.get_manufacturers_data, FSMAdmin.get_certificates_data,
                                  FSMAdmin.get_tnved_data, FSMAdmin.get_all_together])
async def callback_get_info_template(callback: types.CallbackQuery, state=FSMContext):
    reply_template = open(r"get_info_template.xlsx", 'rb')
    await callback.message.reply_document(reply_template)
    await callback.message.answer(f"Шаблон готов к скачиванию", reply_markup=markup_back)


@dp.callback_query_handler(text='get_data')
async def callback_products_upload(callback: types.CallbackQuery):
    await callback.message.answer(f"Нажмите на соотвествующую кнопку, для получения информации о "
                                  f"сертифмкатах,производетелях или кодах ТНВЭД\n"
                                  f"Вы можете получить информацию поарткульно, если их не очень много,"
                                  f" или сразу загрузить файл с интересующими позициями", reply_markup=markup_get_data)


@dp.callback_query_handler(text='manufacturer_info')
async def callback_get_manufacturers_info(callback: types.CallbackQuery):
    await FSMAdmin.get_manufacturers_data.set()
    await callback.message.answer(
        f"Введите артикул, либо, если артикулов много, загрузите файл для выгрузки данных в эксель.\n"
        f"Расширение файла должно быть .xls или .xlsx\n"
        f"Воспользуйтесь печатной формой заказа из 1С или загрузите свой файл в правильном формате:\n\n"
        f"ПЕРВЫЙ АРТИКУЛ В СПИСКЕ ДОЛЖЕН НАХОДИТЬСЯ ВО 2 КОЛОНКЕ И В 6 СТРОКЕ ОСТАЛЬНЫЕ ЯЧЕЙКИ ДОЛЖЫ БЫТЬ ПУСТЫМИ\n\n"
        f"В качестве артикула можно использовать код для заказа или код 1С\n\n"
        f"НАЖМИТЕ НА {emoji.emojize(':paperclip:')}", reply_markup=markup_get_info_template)


@dp.callback_query_handler(text='cert_info')
async def callback_get_cert_info(callback: types.CallbackQuery):
    await FSMAdmin.get_certificates_data.set()
    await callback.message.answer(
        f"Введите артикул, либо, если артикулов много, загрузите файл для выгрузки данных в эксель.\n"
        f"Расширение файла должно быть .xls или .xlsx\n"
        f"Воспользуйтесь печатной формой заказа из 1С или загрузите свой файл в правильном формате:\n\n"
        f"ПЕРВЫЙ АРТИКУЛ В СПИСКЕ ДОЛЖЕН НАХОДИТЬСЯ ВО 2 КОЛОНКЕ И В 6 СТРОКЕ ОСТАЛЬНЫЕ ЯЧЕЙКИ ДОЛЖЫ БЫТЬ ПУСТЫМИ\n\n"
        f"В качестве артикула можно использовать код для заказа или код 1С\n\n"
        f"НАЖМИТЕ НА {emoji.emojize(':paperclip:')}", reply_markup=markup_get_info_template)


@dp.callback_query_handler(text='tnvd_code_upload')
async def callback_get_tnved_info(callback: types.CallbackQuery):
    await FSMAdmin.get_tnved_data.set()
    await callback.message.answer(
        f"Введите артикул, либо, если артикулов много, загрузите файл для выгрузки данных в эксель.\n"
        f"Расширение файла должно быть .xls или .xlsx\n"
        f"Воспользуйтесь печатной формой заказа из 1С или загрузите свой файл в правильном формате:\n\n"
        f"ПЕРВЫЙ АРТИКУЛ В СПИСКЕ ДОЛЖЕН НАХОДИТЬСЯ ВО 2 КОЛОНКЕ И В 6 СТРОКЕ ОСТАЛЬНЫЕ ЯЧЕЙКИ ДОЛЖЫ БЫТЬ ПУСТЫМИ\n\n"
        f"В качестве артикула можно использовать код для заказа или код 1С\n\n"
        f"НАЖМИТЕ НА {emoji.emojize(':paperclip:')}", reply_markup=markup_get_info_template)


@dp.callback_query_handler(text='all_together')
async def callback_get_tnved_info(callback: types.CallbackQuery):
    await FSMAdmin.get_all_together.set()
    await callback.message.answer(
        f"Введите артикул, либо, если артикулов много, загрузите файл для выгрузки данных в эксель.\n"
        f"Расширение файла должно быть .xls или .xlsx\n"
        f"Воспользуйтесь печатной формой заказа из 1С или загрузите свой файл в правильном формате:\n\n"
        f"ПЕРВЫЙ АРТИКУЛ В СПИСКЕ ДОЛЖЕН НАХОДИТЬСЯ ВО 2 КОЛОНКЕ И В 6 СТРОКЕ ОСТАЛЬНЫЕ ЯЧЕЙКИ ДОЛЖЫ БЫТЬ ПУСТЫМИ\n\n"
        f"В качестве артикула можно использовать код для заказа или код 1С\n\n"
        f"НАЖМИТЕ НА {emoji.emojize(':paperclip:')}", reply_markup=markup_get_info_template)


@dp.callback_query_handler(text='get_scan_by_article')
async def callback_get_scan(callback: types.CallbackQuery):
    await FSMAdmin.get_scan_by_article.set()
    await callback.message.answer(
        f"Введите артикул.\n"
        f"В качестве артикула можно использовать код для заказа или код 1С\n\n", reply_markup=markup_back)


# region back button
@dp.callback_query_handler(text='back', state=[FSMAdmin.upload_new_data, FSMAdmin.get_certificates_data,
                                               FSMAdmin.get_manufacturers_data, FSMAdmin.upload_certificates,
                                               FSMAdmin.upload_products, FSMAdmin.get_tnved_data, FSMAdmin.check_dup,
                                               FSMAdmin.get_manufacturers_data, FSMAdmin.get_certificates_data,
                                               FSMAdmin.get_tnved_data, FSMAdmin.get_all_together,
                                               FSMAdmin.download_templates, FSMAdmin.get_info_template,
                                               FSMAdmin.scan_upload, FSMAdmin.get_scan_by_article, None])
async def callback_back_button(callback: types.CallbackQuery, state: FSMContext):
    await state.finish()
    user_full_name = callback.from_user.username
    await callback.message.reply(f"Привет {user_full_name}\nЯ Bilight_Bot\n", reply_markup=markup_start_screen)


# endregion

# endregion


# region MAIN BOT ACTIONS

@dp.message_handler(content_types=types.ContentType.ANY, state=FSMAdmin.upload_certificates)
async def load_certificates_to_postgresql(message: types.Message, state: FSMContext):
    if message.content_type != 'document':
        await FSMAdmin.upload_certificates.set()
        await message.answer('Загружать файлы можно только в формате .xlsx или .xls',
                             reply_markup=markup_back)
    else:
        file_extension = '.' + message.document.file_name.split('.')[-1]
        if file_extension != '.xlsx' and file_extension != '.xls':
            await FSMAdmin.upload_certificates.set()
            await message.answer('Документ должен быть в формате .xls или .xlsx', reply_markup=markup_back)
        else:
            doc = message.document.file_name
            uuid_name = str(uuid.uuid4())
            uuid_dict = {}
            if uuid_name in uuid_dict.keys():
                uuid_name = str(uuid.uuid4())
            else:
                uuid_dict[uuid_name] = doc
            doc = rf"{uuid_name}_{doc}"
            destination = rf"{os.getcwd()}\{doc}"
            await message.document.download(destination_file=destination)
            cert_data_from_user = get_cert_info_from_user(doc)
            error_info = is_error(uuid_name)
            if error_info:
                os.remove(f'{uuid_name}_error.log')
                delete_doc_and_uuid_file(uuid_name, doc)
                await message.reply(f"{error_info}", reply_markup=markup_back)
            else:
                unique_data = find_cert_unique_data(universal_query('certificates', '*'), cert_data_from_user)
                duplicate_data = find_cert_duplicate_data(universal_query('certificates', '*'), doc)
                async with state.proxy() as data:
                    data['doc'] = doc
                    data['cert_data_from_user'] = cert_data_from_user
                    data['unique_data'] = unique_data
                    data['duplicate_data'] = duplicate_data
                    data['uuid_name'] = uuid_name
                if duplicate_data:
                    await message.reply(f"В базе данные обнаружены дубликаты данных по следующим ID"
                                        f" {duplicate_data}. Заменить данные?",
                                        reply_markup=markup_cert_duplicate_question)

                else:
                    add_certificates(unique_data, doc)
                    error_info = is_error(uuid_name)
                    if error_info:
                        os.remove(f'{uuid_name}_error.log')
                        delete_doc_and_uuid_file(uuid_name, doc)
                        await message.reply(f"{error_info}\nСертификаты не загружены или загружены не полностью\n"
                                            f"Проверьте корректность вносимых данных и повторите попытку",
                                            reply_markup=markup_back)
                    else:
                        await message.reply(f"Cертификаты загружены",
                                            reply_markup=markup_back)
                    delete_doc_and_uuid_file(uuid_name, doc)
                    await state.finish()


@dp.callback_query_handler(text='cert_duplicate_yes', state=[FSMAdmin.upload_certificates])
async def callback_yes_cert(callback: types.CallbackQuery, state=FSMContext):
    async with state.proxy() as data:
        unique_data = data['unique_data']
        duplicate_data = data['duplicate_data']
        doc = data['doc']
        uuid_name = data['uuid_name']
    add_certificates(unique_data, doc)
    add_cert_duplicate_user_data(duplicate_data, doc)
    error_info = is_error(uuid_name)
    if error_info:
        os.remove(f'{uuid_name}_error.log')
        delete_doc_and_uuid_file(uuid_name, doc)
        await callback.message.reply(f"{error_info}\nСертификаты не загружены или загружены не полностью\n"
                                     f"Проверьте корректность вносимых данных и повторите попытку",
                                     reply_markup=markup_back)
    else:
        await callback.message.reply(f"Уникальные сертификаты загружены, дубликаты сертификатов в базе обновлены",
                                     reply_markup=markup_back)
        if doc in os.listdir():
            os.remove(doc)
        await state.finish()


@dp.callback_query_handler(text='cert_duplicate_no', state=[FSMAdmin.upload_certificates])
async def callback_no_cert(callback: types.CallbackQuery, state=FSMContext):
    async with state.proxy() as data:
        unique_data = data['unique_data']
        doc = data['doc']
        uuid_name = data['uuid_name']
    add_certificates(unique_data, doc)
    error_info = is_error(uuid_name)
    if error_info:
        os.remove(f'{uuid_name}_error.log')
        delete_doc_and_uuid_file(uuid_name, doc)
        await callback.message.reply(f"{error_info}\nСертификаты не загружены или загружены не полностью\n"
                                     f"Проверьте корректность вносимых данных и повторите попытку",
                                     reply_markup=markup_back)
    else:
        await callback.message.reply(f"Уникальные сертификаты успешно добавлены в базу данных",
                                     reply_markup=markup_back)
        if doc in os.listdir():
            os.remove(doc)
        await state.finish()


@dp.message_handler(content_types=types.ContentType.ANY, state=FSMAdmin.scan_upload)
async def load_scan_to_postgresql(message: types.Message, state: FSMContext):
    if message.content_type != 'document':
        await FSMAdmin.scan_upload.set()
        await message.answer('Загружать файлы можно только в формате .pdf',
                             reply_markup=markup_back)
    else:
        file_extension = '.' + message.document.file_name.split('.')[-1]
        if file_extension != '.pdf':
            await FSMAdmin.scan_upload.set()
            await message.answer("Документ должен быть в формате .pdf", reply_markup=markup_back)
        else:
            doc = message.document.file_name
            uuid_name = str(uuid.uuid4())
            uuid_dict = {}
            if uuid_name in uuid_dict.keys():
                uuid_name = str(uuid.uuid4())
            else:
                uuid_dict[uuid_name] = doc
            destination = rf"certificates\{doc}"
            tmp_destination = rf"tmp\{uuid_name}{doc}"
            async with state.proxy() as data:
                data['doc'] = doc
                data['uuid_name'] = uuid_name
            if doc in os.listdir(f"certificates"):
                await FSMAdmin.scan_upload.set()
                await message.document.download(destination_file=tmp_destination)
                await message.reply(f"ФАЙЛ С ТАКИМ ИМЕНЕМ УЖЕ ЗАГРУЖЕН В БАЗУ,"
                                    f"ХОТИТЕ ЗАМЕНИТЬ НА НОВЫЙ?", reply_markup=markup_scan)
            else:
                await message.document.download(destination_file=destination)
                await message.reply(f"НОВЫЙ СКАН СЕРТИФИКАТА ЗАГРУЖЕН В БАЗУ", reply_markup=markup_back)
                await state.finish()


@dp.callback_query_handler(text='replace_scan_yes', state=[FSMAdmin.scan_upload])
async def callback_scan_yes(callback: types.CallbackQuery, state=FSMContext):
    async with state.proxy() as data:
        doc = data['doc']
        uuid_name = data['uuid_name']
    destination = rf"certificates\{doc}"
    tmp_destination = rf"tmp\{uuid_name}{doc}"
    os.replace(tmp_destination, destination)
    await FSMAdmin.scan_upload.set()
    await callback.message.reply(f"ЗАГРУЗКА ПРОШЛА УСПЕШНО,СКАН СЕРТИФИКАТА В БАЗЕ"
                                 f"ЗАМЕНЕН НА НОВЫЙ", reply_markup=markup_back)
    await state.finish()


@dp.callback_query_handler(text='replace_scan_no', state=[FSMAdmin.scan_upload])
async def callback_scan_no(callback: types.CallbackQuery, state=FSMContext):
    async with state.proxy() as data:
        doc = data['doc']
        uuid_name = data['uuid_name']
    tmp_destination = rf"tmp\{uuid_name}{doc}"
    os.remove(tmp_destination)
    await FSMAdmin.scan_upload.set()
    await callback.message.reply(f"ЗАГРУЗКА ОТМЕНЕНА,СКАН СЕРТИФИКАТА В БАЗЕ ОСТАЛСЯ БЕЗ ИЗМЕНЕНИЙ",
                                 reply_markup=markup_back)
    await state.finish()


@dp.message_handler(content_types=types.ContentType.ANY, state=[FSMAdmin.upload_products, FSMAdmin.check_dup])
async def load_products_to_postgresql(message: types.Message, state: FSMContext):
    if message.content_type != 'document':
        await FSMAdmin.upload_products.set()
        await message.answer('Загружать файлы можно только в формате .xlsx или .xls',
                             reply_markup=markup_back)
    else:
        file_extension = '.' + message.document.file_name.split('.')[-1]
        if file_extension != '.xlsx' and file_extension != '.xls':
            await FSMAdmin.upload_products.set()
            await message.answer('Документ должен быть в формате .xls или .xlsx', reply_markup=markup_back)
        else:
            doc = message.document.file_name
            uuid_name = str(uuid.uuid4())
            uuid_dict = {}
            if uuid_name in uuid_dict.keys():
                uuid_name = str(uuid.uuid4())
            else:
                uuid_dict[uuid_name] = doc
            doc = rf"{uuid_name}_{doc}"
            destination = rf"{os.getcwd()}\{doc}"
            await message.document.download(destination_file=destination)
            data_from_user = get_product_info_from_user(doc)
            error_info = is_error(uuid_name)
            if error_info:
                os.remove(f'{uuid_name}_error.log')
                delete_doc_and_uuid_file(uuid_name, doc)
                await message.reply(f"{error_info}", reply_markup=markup_back)
            else:
                manufacturers_dict = make_dict(universal_query('manufacturers', '*'))
                approved_manufacturers_data = approved_manufacturers_data_list(data_from_user, manufacturers_dict)
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
                    data['approved_manufacturers_data'] = approved_manufacturers_data
                    data['manufacturers_data_to_check'] = manufacturers_data_to_check
                    data['unique_data'] = unique_data
                    data['duplicate_data'] = duplicate_data
                    data['possible_change'] = possible_change
                    data['yes'] = yes
                    data['yes_without_make_pos_change_file'] = yes_without_make_pos_change_file
                    data['uuid_name'] = uuid_name
                if possible_change:
                    await FSMAdmin.upload_products.set()
                    await message.reply(f"В вашем файле присутствуют производители которых"
                                        f" нет в списке производителей в базе данных.\n\n"
                                        f"Если хотите проверить корректность указаных вами производителей,"
                                        f" нажмите 'Проверить' и программа"
                                        f" отправит вам файл с возможными заменами, артикулы с корректными "
                                        f"производителями будут загружены.\n\n"
                                        f"Если вы уверены, что производители в файле указанны корректно"
                                        f" нажмите 'Я уверен' и программа добавит в"
                                        f" базу данных новых производителей", reply_markup=markup_question)
                else:
                    if duplicate_data:
                        await FSMAdmin.check_dup.set()
                        yes_without_make_pos_change_file = True
                        add_permition = False
                        async with state.proxy() as data:
                            data['yes_without_make_pos_change_file'] = yes_without_make_pos_change_file
                            data['add_permition'] = add_permition
                        await message.reply(f"В базе данные обнаружены дубликаты данных"
                                            f" по следующим ID {duplicate_data}."
                                            f" Заменить данные?", reply_markup=markup_duplicate_question)
                    else:
                        converted_manufacturers_data = convert_manufacturers_to_digit(data_from_user)
                        add_products(converted_manufacturers_data, doc)
                        error_info = is_error(uuid_name)
                        if error_info:
                            os.remove(f'{uuid_name}_error.log')
                            delete_doc_and_uuid_file(uuid_name, doc)
                            await message.reply(f"{error_info}\nАртикулы не загружены или загружены не полностью\n"
                                                f"Проверьте корректность вносимых данных и повторите попытку",
                                                reply_markup=markup_back)
                        else:
                            await message.reply(
                                f"Артикулы успешно загружены", reply_markup=markup_back)
                        if doc in os.listdir():
                            os.remove(doc)
                        await state.finish()


@dp.callback_query_handler(text='yes', state=[FSMAdmin.check_dup, FSMAdmin.upload_products])
async def callback_yes_prod(callback: types.CallbackQuery, state=FSMContext):
    async with state.proxy() as data:
        doc = data['doc']
        approved_manufacturers_data = data['approved_manufacturers_data']
        duplicate_data = data['duplicate_data']
        possible_change = data['possible_change']
        uuid_name = data['uuid_name']
    if duplicate_data:
        yes = True
        add_permition = True
        async with state.proxy() as data:
            data['yes'] = yes
            data['add_permition'] = add_permition
        await FSMAdmin.check_dup.set()
        await callback.message.reply(f"В базе данные обнаружены дубликаты данных"
                                     f" по следующим ID {duplicate_data}."
                                     f" Заменить данные?", reply_markup=markup_duplicate_question)
    else:
        user_name = callback.from_user.username
        extension = doc.split('.')[-1]
        make_replace_file(doc, possible_change, user_name, uuid_name)
        converted_manufacturers_data = convert_manufacturers_to_digit(approved_manufacturers_data)
        add_products(converted_manufacturers_data, doc)
        reply_possible_changes = open(rf"possible_changes_{user_name}.{extension}", 'rb')
        await callback.message.reply_document(reply_possible_changes)
        error_info = is_error(uuid_name)
        if error_info:
            os.remove(f'{uuid_name}_error.log')
            delete_doc_and_uuid_file(uuid_name, doc)
            await callback.message.reply(f"{error_info}", reply_markup=markup_back)
        else:
            await callback.message.reply(
                f"Артикулы с корректными производителями  загружены, "
                f"предполагаемые замены производителей в подготовленном файле", reply_markup=markup_back)
        if rf"possible_changes_{user_name}.{extension}" in os.listdir():
            os.remove(rf"possible_changes_{user_name}.{extension}")
        if rf"possible_changes_{uuid_name}.{extension}" in os.listdir():
            os.remove(rf"possible_changes_{uuid_name}.{extension}")
        if doc in os.listdir():
            os.remove(doc)
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
        doc = data['doc']
        uuid_name = data['uuid_name']
    if duplicate_data:
        await FSMAdmin.check_dup.set()
        await callback.message.reply(f"В базе данные обнаружены дубликаты данных"
                                     f" по следующим ID {duplicate_data}."
                                     f" Заменить данные?", reply_markup=markup_duplicate_question)
    else:
        add_new_manufacturers(manufacturers_dict, data_from_user, doc)
        converted_manufacturers_data = convert_manufacturers_to_digit(unique_data)
        add_products(converted_manufacturers_data, doc)
        error_info = is_error(uuid_name)
        if error_info:
            os.remove(f'{uuid_name}_error.log')
            delete_doc_and_uuid_file(uuid_name, doc)
            await callback.message.reply(f"{error_info}", reply_markup=markup_back)
        else:
            await callback.message.reply(f"Артикулы и новые производители  успешно добавлены в базу данных",
                                         reply_markup=markup_back)
        if doc in os.listdir():
            os.remove(doc)
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
        uuid_name = data['uuid_name']
    if yes:
        user_name = callback.from_user.username
        extension = doc.split('.')[-1]
        make_replace_file(doc, possible_change, user_name, uuid_name)
        reply_possible_changes = open(rf"possible_changes_{user_name}.{extension}", 'rb')
        converted_manufacturers_data = convert_manufacturers_to_digit(duplicate_data)
        tmp_unique_list = []
        for x in unique_data:
            if x[3] not in possible_change.keys():
                tmp_unique_list.append(x)
        add_unique_user_data(tmp_unique_list, doc)
        add_duplicate_user_data(converted_manufacturers_data, doc)
        await callback.message.reply_document(reply_possible_changes)
        if rf"possible_changes_{user_name}.{extension}" in os.listdir():
            os.remove(rf"possible_changes_{user_name}.{extension}")
        if rf"possible_changes_{uuid_name}.{extension}" in os.listdir():
            os.remove(rf"possible_changes_{uuid_name}.{extension}")
        error_info = is_error(uuid_name)
        if error_info:
            os.remove(f'{uuid_name}_error.log')
            delete_doc_and_uuid_file(uuid_name, doc)
            await callback.message.reply(f"{error_info}\nАртикулы не загружены или загружены не полностью\n"
                                         f"Проверьте корректность вносимых данных и повторите попытку",
                                         reply_markup=markup_back)
        else:
            await callback.message.reply(
                f"Артикулы с корректными производителями  загружены, "
                f"предполагаемые замены производителей в подготовленном файле", reply_markup=markup_back)
            if rf"possible_changes_{user_name}.{extension}" in os.listdir():
                os.remove(rf"possible_changes_{user_name}.{extension}")
            if rf"possible_changes_{uuid_name}.{extension}" in os.listdir():
                os.remove(rf"possible_changes_{uuid_name}.{extension}")
            if doc in os.listdir():
                os.remove(doc)
            await state.finish()
    elif yes_without_make_pos_change_file:
        converted_manufacturers_data = convert_manufacturers_to_digit(duplicate_data)
        add_duplicate_user_data(converted_manufacturers_data, doc)
        converted_manufacturers_data = convert_manufacturers_to_digit(unique_data)
        add_unique_user_data(converted_manufacturers_data, doc)
        error_info = is_error(uuid_name)
        if error_info:
            os.remove(f'{uuid_name}_error.log')
            delete_doc_and_uuid_file(uuid_name, doc)
            await callback.message.reply(f"{error_info}\nАртикулы не загружены или загружены не полностью\n"
                                         f"Проверьте корректность вносимых данных и повторите попытку",
                                         reply_markup=markup_back)
        else:
            await callback.message.reply(f"Артикулы успешно заменены",
                                         reply_markup=markup_back)
        if doc in os.listdir():
            os.remove(doc)
        await state.finish()
    else:
        add_new_manufacturers(manufacturers_dict, data_from_user, doc)
        converted_manufacturers_data = convert_manufacturers_to_digit(duplicate_data)
        add_duplicate_user_data(converted_manufacturers_data, doc)
        converted_manufacturers_data = convert_manufacturers_to_digit(unique_data)
        add_unique_user_data(converted_manufacturers_data, doc)
        error_info = is_error(uuid_name)
        if error_info:
            os.remove(f'{uuid_name}_error.log')
            delete_doc_and_uuid_file(uuid_name, doc)
            await callback.message.reply(f"{error_info}\nАртикулы не загружены или загружены не полностью\n"
                                         f"Проверьте корректность вносимых данных и повторите попытку",
                                         reply_markup=markup_back)
        else:
            await callback.message.reply(f"Артикулы и новые производители  успешно добавлены в базу данных",
                                         reply_markup=markup_back)
        if doc in os.listdir():
            os.remove(doc)
        await state.finish()


@dp.callback_query_handler(text='duplicate_no', state=[FSMAdmin.check_dup, FSMAdmin.upload_products])
async def callback_no_duplicate(callback: types.CallbackQuery, state=FSMContext):
    async with state.proxy() as data:
        unique_data = data['unique_data']
        add_permition = data['add_permition']
        manufacturers_dict = data['manufacturers_dict']
        data_from_user = data['data_from_user']
        doc = data['doc']
        uuid_name = data['uuid_name']
    if add_permition:
        add_new_manufacturers(manufacturers_dict, data_from_user, doc)
        converted_manufacturers_data = convert_manufacturers_to_digit(unique_data)
        add_unique_user_data(converted_manufacturers_data, doc)
        error_info = is_error(uuid_name)
        if error_info:
            os.remove(f'{uuid_name}_error.log')
            delete_doc_and_uuid_file(uuid_name, doc)
            await callback.message.reply(f"{error_info}\nАртикулы не загружены или загружены не полностью\n"
                                         f"Проверьте корректность вносимых данных и повторите попытку",
                                         reply_markup=markup_back)
        else:
            await callback.message.reply(
                f"Работа завершена, дубликаты данных пользователя в базе остались без изменений, уникальные артикулы "
                f"добавлены, новые производители добавлены",
                reply_markup=markup_back)
        if doc in os.listdir():
            os.remove(doc)
        await state.finish()
    else:
        converted_manufacturers_data = convert_manufacturers_to_digit(unique_data)
        add_unique_user_data(converted_manufacturers_data, doc)
        await callback.message.reply(
            f"Работа завершена, дубликаты данных пользователя в базе остались без изменений, уникальные артикулы "
            f"добавлены",
            reply_markup=markup_back)
        if doc in os.listdir():
            os.remove(doc)
        await state.finish()


@dp.message_handler(content_types=types.ContentType.ANY, state=FSMAdmin.get_manufacturers_data)
async def get_manufacturer_by_article(message: types.Message, state: FSMContext):
    await FSMAdmin.get_manufacturers_data.set()
    article = message.text
    all_article_from_db = make_global_info_table()
    if message.content_type == 'document':
        file_extension = '.' + message.document.file_name.split('.')[-1]
        if file_extension != '.xlsx' and file_extension != '.xls':
            await FSMAdmin.get_manufacturers_data.set()
            await message.answer('Документ должен быть в формате .xls или .xlsx', reply_markup=markup_back)
        else:
            file_name = message.document.file_name
            user_name = message.from_user.username
            extension = file_name.split('.')[-1]
            destination = rf"{os.getcwd()}\{file_name}"
            await message.document.download(destination_file=destination)
            uuid_file_name = make_manufacturer_list_file(file_name, user_name)
            manufacturer_list_by_articles = open(rf"manufacturer_list_by_articles_{user_name}.{extension}", 'rb')
            await message.reply_document(manufacturer_list_by_articles)
            await message.reply('Документ с производителями готов к скачиванию', reply_markup=markup_back)
            if rf"manufacturer_list_by_articles_{user_name}.{extension}" in os.listdir():
                os.remove(rf"manufacturer_list_by_articles_{user_name}.{extension}")
            if uuid_file_name in os.listdir():
                os.remove(uuid_file_name)
            if file_name in os.listdir():
                os.remove(file_name)
    elif message.content_type == 'text':
        match = None
        all_data = None
        for i in range(len(all_article_from_db)):
            for j in range(len(all_article_from_db[i])):
                if str(all_article_from_db[i][j]).upper() == str(article).upper():
                    match = all_article_from_db[i][j]
                    all_data = all_article_from_db[i]
                else:
                    pass
        if match is not None:
            await message.reply(f"{all_data[2]}")
            await message.reply(
                f"Введите следующий артикул, либо, если артикулов много загрузите"
                f" файл для выгрузки данных в эксель\n"
                f"НАЖМИТЕ НА {emoji.emojize(':paperclip:')}", reply_markup=markup_back)
        else:
            await message.reply(f"Данный артикул в базе данных не найден, либо артикул введен не корректно\n"
                                f"Проверьте правильно ли написан артикул и повторите попытку\n"
                                f"либо, если артикулов много загрузите файл для выгрузки данных в эксель\n"
                                f"НАЖМИТЕ НА {emoji.emojize(':paperclip:')}", reply_markup=markup_back)


@dp.message_handler(content_types=types.ContentType.ANY, state=FSMAdmin.get_certificates_data)
async def get_certificates_by_article(message: types.Message, state: FSMContext):
    await FSMAdmin.get_certificates_data.set()
    article = message.text
    all_article_from_db = make_global_info_table()
    if message.content_type == 'document':
        file_extension = '.' + message.document.file_name.split('.')[-1]
        if file_extension != '.xlsx' and file_extension != '.xls':
            await FSMAdmin.get_certificates_data.set()
            await message.answer('Документ должен быть в формате .xls или .xlsx', reply_markup=markup_back)
        else:
            file_name = message.document.file_name
            user_name = message.from_user.username
            extension = file_name.split('.')[-1]
            destination = rf"{os.getcwd()}\{file_name}"
            await message.document.download(destination_file=destination)
            uuid_file_name = make_certificate_list_file(file_name, user_name)
            certificates_list_by_articles = open(rf"certificates_list_by_article_{user_name}.{extension}", 'rb')
            await message.reply_document(certificates_list_by_articles)
            await message.reply('Документ с сертификатами готов к скачиванию', reply_markup=markup_back)
            if rf"certificates_list_by_article_{user_name}.{extension}" in os.listdir():
                os.remove(rf"certificates_list_by_article_{user_name}.{extension}")
            if uuid_file_name in os.listdir():
                os.remove(uuid_file_name)
            if file_name in os.listdir():
                os.remove(file_name)
    elif message.content_type == 'text':
        match = None
        all_data = None
        for i in range(len(all_article_from_db)):
            for j in range(len(all_article_from_db[i])):
                if str(all_article_from_db[i][j]).upper() == str(article).upper():
                    match = all_article_from_db[i][j]
                    all_data = all_article_from_db[i]
                else:
                    pass
        if match is not None:
            await message.reply(f"Сертификат номер - {all_data[3]} ТИП:  {all_data[4]} ДАТА НАЧАЛА ДЕЙСТВИЯ:"
                                f" {all_data[5]} ДАТА ОКОНЧАНИЯ: {all_data[6]}")
            await message.reply(
                f"Введите следующий артикул, либо, если артикулов много загрузите"
                f" файл для выгрузки данных в эксель\n"
                f"НАЖМИТЕ НА {emoji.emojize(':paperclip:')}", reply_markup=markup_back)
        else:
            await message.reply(f"Данный артикул в базе данных не найден, либо артикул введен не корректно\n"
                                f"Проверьте правильно ли написан артикул и повторите попытку\n"
                                f"либо, если артикулов много загрузите файл для выгрузки данных в эксель\n"
                                f"НАЖМИТЕ НА {emoji.emojize(':paperclip:')}", reply_markup=markup_back)


@dp.message_handler(content_types=types.ContentType.ANY, state=FSMAdmin.get_tnved_data)
async def get_tnved_data_by_article(message: types.Message, state: FSMContext):
    await FSMAdmin.get_tnved_data.set()
    article = message.text
    all_article_from_db = make_global_info_table()
    if message.content_type == 'document':
        file_extension = '.' + message.document.file_name.split('.')[-1]
        if file_extension != '.xlsx' and file_extension != '.xls':
            await FSMAdmin.get_tnved_data.set()
            await message.answer('Документ должен быть в формате .xls или .xlsx', reply_markup=markup_back)
        else:
            file_name = message.document.file_name
            user_name = message.from_user.username
            extension = file_name.split('.')[-1]
            destination = rf"{os.getcwd()}\{file_name}"
            await message.document.download(destination_file=destination)
            uuid_file_name = make_tnved_list_file(file_name, user_name)
            tnved_list_by_articles = open(rf"tnved_list_by_article_{user_name}.{extension}", 'rb')
            await message.reply_document(tnved_list_by_articles)
            await message.reply('Документ с кодами ТНВЭД готов к скачиванию', reply_markup=markup_back)
            if rf"tnved_list_by_article_{user_name}.{extension}" in os.listdir():
                os.remove(rf"tnved_list_by_article_{user_name}.{extension}")
            if uuid_file_name in os.listdir():
                os.remove(uuid_file_name)
            if file_name in os.listdir():
                os.remove(file_name)
    elif message.content_type == 'text':
        match = None
        all_data = None
        for i in range(len(all_article_from_db)):
            for j in range(len(all_article_from_db[i])):
                if str(all_article_from_db[i][j]).upper() == str(article).upper():
                    match = all_article_from_db[i][j]
                    all_data = all_article_from_db[i]
                else:
                    pass
        if match is not None:
            await message.reply(f"Код ТНВЭД - {all_data[7]} ОПИСАНИЕ:  {all_data[8]}")
            await message.reply(
                f"Введите следующий артикул, либо, если артикулов много загрузите"
                f" файл для выгрузки данных в эксель\n"
                f"НАЖМИТЕ НА {emoji.emojize(':paperclip:')}", reply_markup=markup_back)
        else:
            await message.reply(f"Данный артикул в базе данных не найден, либо артикул введен не корректно\n"
                                f"Проверьте правильно ли написан артикул и повторите попытку\n"
                                f"либо, если артикулов много загрузите файл для выгрузки данных в эксель\n"
                                f"НАЖМИТЕ НА {emoji.emojize(':paperclip:')}", reply_markup=markup_back)


@dp.message_handler(content_types=types.ContentType.ANY, state=FSMAdmin.get_all_together)
async def get_all_together_data_by_article(message: types.Message, state: FSMContext):
    await FSMAdmin.get_all_together.set()
    article = message.text
    all_article_from_db = make_global_info_table()
    if message.content_type == 'document':
        file_extension = '.' + message.document.file_name.split('.')[-1]
        if file_extension != '.xlsx' and file_extension != '.xls':
            await FSMAdmin.get_tnved_data.set()
            await message.answer('Документ должен быть в формате .xls или .xlsx', reply_markup=markup_back)
        else:
            file_name = message.document.file_name
            user_name = message.from_user.username
            extension = file_name.split('.')[-1]
            destination = rf"{os.getcwd()}\{file_name}"
            await message.document.download(destination_file=destination)
            uuid_file_name = make_total_list_file(file_name, user_name)
            tnved_list_by_articles = open(rf"total_list_by_article_{user_name}.{extension}", 'rb')
            await message.reply_document(tnved_list_by_articles)
            await message.reply('Документ с данными готов к скачиванию', reply_markup=markup_back)
            if rf"total_list_by_article_{user_name}.{extension}" in os.listdir():
                os.remove(rf"total_list_by_article_{user_name}.{extension}")
            if uuid_file_name in os.listdir():
                os.remove(uuid_file_name)
            if file_name in os.listdir():
                os.remove(file_name)
    elif message.content_type == 'text':
        match = None
        all_data = None
        for i in range(len(all_article_from_db)):
            for j in range(len(all_article_from_db[i])):
                if str(all_article_from_db[i][j]).upper() == str(article).upper():
                    match = all_article_from_db[i][j]
                    all_data = all_article_from_db[i]
                else:
                    pass
        if match is not None:
            await message.reply(f"{all_data}")
            await message.reply(
                f"Введите следующий артикул, либо, если артикулов много загрузите"
                f" файл для выгрузки данных в эксель\n"
                f"НАЖМИТЕ НА {emoji.emojize(':paperclip:')}", reply_markup=markup_back)
        else:
            await message.reply(f"Данный артикул в базе данных не найден, либо артикул введен не корректно\n"
                                f"Проверьте правильно ли написан артикул и повторите попытку\n"
                                f"либо, если артикулов много загрузите файл для выгрузки данных в эксель\n"
                                f"НАЖМИТЕ НА {emoji.emojize(':paperclip:')}", reply_markup=markup_back)


@dp.message_handler(content_types=types.ContentType.ANY, state=FSMAdmin.get_scan_by_article)
async def get_scan_cert_by_article(message: types.Message, state: FSMContext):
    await FSMAdmin.get_scan_by_article.set()
    article = message.text
    if message.content_type != 'text':
        await FSMAdmin.get_scan_by_article.set()
        await message.answer('Сертификаты присылаются только по артикулу, нужно ввести текст', reply_markup=markup_back)
    else:
        if article.isdigit():
            cert_id = get_cert_id_for_scan_query(article, 'product_id')
        else:
            cert_id = get_cert_id_for_scan_query(article, 'order_code')
        if cert_id is not None:
            if f"{cert_id}.pdf" in os.listdir(f'certificates'):
                scan = open(rf"certificates/{cert_id}.pdf", 'rb')
                await message.reply_document(scan)
                await message.reply(f"Скан сертификата готов к скачиванию\nВведите следующий артикул.",
                                    reply_markup=markup_back)
            else:
                await message.reply(f"Скан сертификата отсутсвует в базе данных\nВведите следующий артикул.",
                                    reply_markup=markup_back)
        else:
            await message.reply(f"Данный артикул в базе данных не найден, либо артикул введен не корректно\n"
                                f"Проверьте правильно ли написан артикул и повторите попытку\n",
                                reply_markup=markup_back)


# endregion

if __name__ == '__main__':
    executor.start_polling(dp, skip_updates=True)
