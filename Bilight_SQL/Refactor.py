# region Import
import os
import psycopg2
import openpyxl
import shutil
from dotenv import find_dotenv, load_dotenv

load_dotenv(find_dotenv())
from openpyxl.styles import colors
from openpyxl.styles import Font, Color, Alignment
from openpyxl.utils import get_column_letter


# endregion

# region Levinstain algo
def levinstain_algo(i, j, s1, s2, matrix):
    if i == 0 and j == 0:
        return 0
    if i == 0 and j > 0:
        return j
    if j == 0 and i > 0:
        return i
    else:
        if s1[i - 1] == s2[j - 1]:
            m = 0
        else:
            m = 1
    return min(matrix[i][j - 1] + 1, matrix[i - 1][j] + 1, matrix[i - 1][j - 1] + m)


def calculate_levinstain_distance(s1, s2):
    str_1 = len(s1)
    str_2 = len(s2)
    matrix = [[0 for x in range(str_2 + 1)] for j in range(str_1 + 1)]
    for i in range(str_1 + 1):
        for j in range(str_2 + 1):
            matrix[i][j] = levinstain_algo(i, j, s1, s2, matrix)
    return matrix[str_1][str_2]


# endregion

# region Get Editorial number
def get_editorial_num(input_string):
    if 2 <= len(input_string) <= 6:
        editorial_number = 2
    else:
        editorial_number = 3
    return editorial_number


# endregion

# region Query Helpers
def universal_query(table_name, *args):
    args = ','.join(args)
    try:
        connect = psycopg2.connect(dbname=os.getenv('db_name'), user=os.getenv('user'),
                                   password=os.getenv('password'), host=os.getenv('host'))
        connect.autocommit = True
        cursor = connect.cursor()
        cursor.execute(f"SELECT COUNT(*) FROM {table_name}")
        table_len = cursor.fetchone()[0]
        cursor.execute(f"SELECT {args} FROM {table_name}")
        query_data = cursor.fetchmany(table_len)
        return query_data
    except Exception as _ex:
        print(f"[INFO] ERROR while working with data base {_ex}")
    finally:
        cursor.close()
        connect.close()


def get_column_name_list(table_name, where_filter, *args):
    args = ','.join(args)
    try:
        connect = psycopg2.connect(dbname=os.getenv('db_name'), user=os.getenv('user'),
                                   password=os.getenv('password'), host=os.getenv('host'))
        connect.autocommit = True
        cursor = connect.cursor()
        cursor.execute(f"SELECT COUNT(*) FROM {table_name}"
                       f" WHERE table_name = '{where_filter}'")
        table_len = cursor.fetchone()[0]
        cursor.execute(f"SELECT {args} FROM {table_name} "
                       f"WHERE table_name = '{where_filter}'"
                       f"ORDER BY ordinal_position ASC")
        column_name = cursor.fetchmany(table_len)
        column_name = [x[0] for x in column_name]
        return column_name
    except Exception as _ex:
        print(f"[INFO] ERROR while working with data base {_ex}")
    finally:
        cursor.close()
        connect.close()


# endregion

# region Helpers
def make_dict(query):
    output_dict = {}
    for x in query:
        if x[1] not in output_dict:
            output_dict[x[1]] = x[0]
        else:
            pass
    return output_dict


def make_possible_change_dict(dict_to_compression, user_data_list):
    possible_change_list = []
    for i in range(len(user_data_list)):
        tmp_list = []
        editorial_number = get_editorial_num(user_data_list[i][3])
        for key in dict_to_compression.keys():
            if editorial_number >= calculate_levinstain_distance(key, user_data_list[i][3].upper()):
                tmp_list.append(key.upper())
            else:
                continue
        tmp_list.append((user_data_list[i][3]))
        possible_change_list.append(tmp_list)
    possible_change_dict = {}
    for j in range(len(possible_change_list)):
        if possible_change_list[j][-1].upper() not in dict_to_compression.keys():
            possible_change_dict[possible_change_list[j][-1]] = possible_change_list[j][:-1]
        else:
            pass
    return possible_change_dict


def make_replacement(existing_pkey, user_data, table_name, where_filter):
    columns_name_list = get_column_name_list('information_schema.columns', table_name, 'column_name')
    try:
        connect = psycopg2.connect(dbname=os.getenv('db_name'), user=os.getenv('user'),
                                   password=os.getenv('password'), host=os.getenv('host'))
        connect.autocommit = True
        cursor = connect.cursor()
        for i in range(len(existing_pkey)):
            for j in range(1, len(columns_name_list)):
                cursor.execute(f" UPDATE {table_name}"
                               f" SET {columns_name_list[j]} = %s"
                               f" WHERE {where_filter} = %s ", (user_data[i][j], user_data[i][0]))
    except Exception as _ex:
        print(f"[INFO] ERROR while working with data base {_ex}")
    finally:
        cursor.close()
        connect.close()


def convert_manufacturers_to_digit(data):
    manufacturers_dict = make_dict(universal_query('manufacturers', '*'))
    for i in range(len(data)):
        if str(data[i][3]).upper() in manufacturers_dict.keys():
            data[i][3] = manufacturers_dict[data[i][3].upper()]
    return data


def add_new_manufacturers(dict_to_compression, data_from_user):
    new_manufacturers = make_possible_change_dict(dict_to_compression, data_from_user)
    try:
        connect = psycopg2.connect(dbname=os.getenv('db_name'), user=os.getenv('user'),
                                   password=os.getenv('password'), host=os.getenv('host'))
        connect.autocommit = True
        cursor = connect.cursor()
        for i in new_manufacturers.keys():
            if i.upper() not in dict_to_compression.keys():
                cursor.execute(f"INSERT INTO manufacturers (manufacturer_name) "
                               f"VALUES ('{i.upper()}')")
    except Exception as _ex:
        print(f"[INFO] ERROR while working with data base {_ex}")

    finally:
        cursor.close()
        connect.close()
        print("Postgre SQL Connection closed")


def get_replace_file(file_path, possible_change):
    shutil.copy(rf"{file_path}", rf"possible_changes.xlsx")
    book = openpyxl.open(rf"possible_changes.xlsx", read_only=False, data_only=True)
    sheet = book.active
    font_data_to_check = Font(
        name='Tahoma',
        size=9,
        bold=True,
        italic=False,
        vertAlign=None,
        underline='none',
        strike=False,
        color='FFAA0000'
    )
    alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    sheet.column_dimensions['D'].width = 50
    for row in range(2, sheet.max_row + 1):
        if sheet[row][3].value not in possible_change.keys():
            sheet.delete_rows(idx=row)
        else:
            pass
    for row in range(2, sheet.max_row + 1):
        if sheet[row][3].value in possible_change.keys():
            sheet[row][3].font = font_data_to_check
            sheet[row][3].alignment = alignment
            sheet[row][3].value = f" Список замен для производителя {sheet[row][3].value}" \
                                  f" следующий:\n" \
                                  f"{possible_change[sheet[row][3].value]}\n" \
                                  f"Вставьте в ячейку производителя из предложенных и повторите загрузку\n" \
                                  f"Либо вставьте своего производителя, если уверенны в корректности названия"

    book.save(rf"possible_changes.xlsx")
    book.close()


# endregion


# region Info from user
def get_product_info_from_user(file_path):
    products_data_list = []
    book = openpyxl.open(file_path, read_only=True, data_only=True)
    sheet = book.active
    for row in range(2, sheet.max_row + 1):
        tmp_list = []
        for column in range(0, 7):
            tmp_list.append(sheet[row][column].value)
        products_data_list.append(tmp_list)
    book.close()
    return products_data_list


# endregion


# region Check existing manufacturers


def approved_manufactureres_data_list(products_data_list, manufacturer_dict):
    approved_data_list = []
    for i in range(len(products_data_list)):
        if products_data_list[i][3].upper() in manufacturer_dict.keys():
            approved_data_list.append(products_data_list[i])
        else:
            pass
    return approved_data_list


def manufacturers_to_check_data_list(products_data_list):
    manufacturer_dict = make_dict(universal_query('manufacturers', '*'))
    data_list_to_check = []
    for i in range(len(products_data_list)):
        if str(products_data_list[i][3]).upper() not in manufacturer_dict:
            data_list_to_check.append(products_data_list[i])
        else:
            pass
    return data_list_to_check


# endregion

# region Unique or Not
def find_duplicate_data(universal_query, file_name):
    data_from_user = get_product_info_from_user(file_name)
    existing_pkey = universal_query
    existing_pkey = [x[0] for x in existing_pkey]
    duplicate_user_data = []

    for i in range(len(data_from_user)):
        if data_from_user[i][0] in existing_pkey:
            duplicate_user_data.append(data_from_user[i])
        else:
            pass
    return duplicate_user_data


def find_unique_data(universal_query, data_from_user):
    existing_pkey = universal_query
    existing_pkey = [x[0] for x in existing_pkey]
    unique_user_data = []
    for i in range(len(data_from_user)):
        if data_from_user[i][0] not in existing_pkey:
            unique_user_data.append(data_from_user[i])
        else:
            pass
    return unique_user_data


# endregion


def add_unique_user_data(unique_data):
    try:
        connect = psycopg2.connect(dbname=os.getenv('db_name'), user=os.getenv('user'),
                                   password=os.getenv('password'), host=os.getenv('host'))
        connect.autocommit = True
        cursor = connect.cursor()
        if unique_data != '':
            cursor.executemany(f""" INSERT INTO bilight_products (product_id,order_code,article,
                                                            manufacturer_id,product_name,tnved_id,certificate_id)
                                                            VALUES
                                                            (%s,%s,%s,%s,%s,%s,%s) """, unique_data)
        else:
            pass
    except Exception as _ex:
        print(f"[INFO] ERROR while working with data base {_ex}")
    finally:
        cursor.close()
        connect.close()
        print("Postgre SQL Connection closed")


def add_duplicate_user_data(duplicate_data):
    try:
        connect = psycopg2.connect(dbname=os.getenv('db_name'), user=os.getenv('user'),
                                   password=os.getenv('password'), host=os.getenv('host'))
        connect.autocommit = True
        cursor = connect.cursor()
        if duplicate_data:
            existing_pkey = duplicate_data
            make_replacement(existing_pkey, duplicate_data, 'bilight_products', 'product_id')
        else:
            pass
    except Exception as _ex:
        print(f"[INFO] ERROR while working with data base {_ex}")
    finally:
        cursor.close()
        connect.close()
        print("Postgre SQL Connection closed")


# endregion

def add_products(products_data_from_user):
    try:
        connect = psycopg2.connect(dbname=os.getenv('db_name'), user=os.getenv('user'),
                                   password=os.getenv('password'), host=os.getenv('host'))
        connect.autocommit = True
        cursor = connect.cursor()
        cursor.executemany(f""" INSERT INTO bilight_products (product_id,order_code,article,
                                                                manufacturer_id,product_name,tnved_id,certificate_id)
                                                                VALUES
                                                                (%s,%s,%s,%s,%s,%s,%s) """, products_data_from_user)
    except Exception as _ex:
        print(f"[INFO] ERROR while working with data base {_ex}")
    finally:
        cursor.close()
        connect.close()
        print("Postgre SQL Connection closed")


# region Certification
def get_cert_info_from_user(file_path):
    certification_data_list = []
    book = openpyxl.open(file_path, read_only=True, data_only=True)
    sheet = book.active
    for row in range(3, sheet.max_row + 1):
        tmp_list = []
        for column in range(0, 5):
            tmp_list.append(sheet[row][column].value)
        if None in tmp_list:
            print("Columns can't be empty")
            return None
        else:
            certification_data_list.append(tmp_list)
    return certification_data_list


def add_certificates(cert_data_from_user):
    try:
        connect = psycopg2.connect(dbname=os.getenv('db_name'), user=os.getenv('user'),
                                   password=os.getenv('password'), host=os.getenv('host'))
        connect.autocommit = True
        cursor = connect.cursor()
        cursor.executemany(f""" INSERT INTO certificates (certificate_id,certificate_number,certificate_type_id,
                                                start_date,end_date)
                                                VALUES
                                                (%s,%s,%s,%s,%s) """, cert_data_from_user)
    except Exception as _ex:
        print(f"[INFO] ERROR while working with data base {_ex}")
    finally:
        cursor.close()
        connect.close()
        print("Postgre SQL Connection closed")


def find_cert_duplicate_data(universal_query, file_name):
    data_from_user = get_cert_info_from_user(file_name)
    existing_pkey = universal_query
    existing_pkey = [x[0] for x in existing_pkey]
    duplicate_user_data = []

    for i in range(len(data_from_user)):
        if data_from_user[i][0] in existing_pkey:
            duplicate_user_data.append(data_from_user[i])
        else:
            pass
    return duplicate_user_data


def find_cert_unique_data(universal_query, data_from_user):
    existing_pkey = universal_query
    existing_pkey = [x[0] for x in existing_pkey]
    unique_user_data = []
    for i in range(len(data_from_user)):
        if data_from_user[i][0] not in existing_pkey:
            unique_user_data.append(data_from_user[i])
        else:
            pass
    return unique_user_data


def add_cert_duplicate_user_data(duplicate_data):
    try:
        connect = psycopg2.connect(dbname=os.getenv('db_name'), user=os.getenv('user'),
                                   password=os.getenv('password'), host=os.getenv('host'))
        connect.autocommit = True
        cursor = connect.cursor()
        if duplicate_data:
            existing_pkey = duplicate_data
            make_replacement(existing_pkey, duplicate_data, 'certificates', 'certificate_id')
        else:
            pass
    except Exception as _ex:
        print(f"[INFO] ERROR while working with data base {_ex}")
    finally:
        cursor.close()
        connect.close()
        print("Postgre SQL Connection closed")


def get_all_article_from_db():
    try:
        connect = psycopg2.connect(dbname=os.getenv('db_name'), user=os.getenv('user'),
                                   password=os.getenv('password'), host=os.getenv('host'))
        connect.autocommit = True
        cursor = connect.cursor()
        cursor.execute(f"SELECT COUNT(*) FROM bilight_products")
        table_len = cursor.fetchone()[0]
        cursor.execute(f"SELECT product_id,order_code FROM bilight_products ")
        all_articles = cursor.fetchmany(table_len)
        return all_articles

    except Exception as _ex:
        print(f"[INFO] ERROR while working with data base {_ex}")
    finally:
        cursor.close()
        connect.close()
        print("Postgre SQL Connection closed")


def make_sertificate_dict(query):
    output_dict = {}
    for x in query:
        if x[0] not in output_dict:
            output_dict[x[0]] = x[1:]
        else:
            pass

    return output_dict


def get_certificate_id_by_article_query(article):
    try:
        connect = psycopg2.connect(dbname=os.getenv('db_name'), user=os.getenv('user'),
                                   password=os.getenv('password'), host=os.getenv('host'))
        connect.autocommit = True
        cursor = connect.cursor()
        cursor.execute(
            f"SELECT certificate_id FROM bilight_products "
            f"WHERE product_id = '{article}' or order_code = '{article}'")
        certificate_id = cursor.fetchone()[0]
        return certificate_id
    except Exception as _ex:
        print(f"[INFO] ERROR while working with data base  {_ex}")
    finally:
        cursor.close()
        connect.close()


def make_certificate_list_file(file_path, all_article_from_db, certificates_dict):
    shutil.copy(rf"{file_path}", rf"certificates_list_by_article.xlsx")
    book = openpyxl.open(rf"certificates_list_by_article.xlsx", read_only=False, data_only=True)
    sheet = book.active
    font_header = Font(
        name='Tahoma',
        size=9,
        bold=True,
        italic=False,
        vertAlign=None,
        underline='none',
        strike=False,
        color='FF000000'
    )
    font_data_exsist = Font(
        name='Tahoma',
        size=9,
        bold=True,
        italic=False,
        vertAlign=None,
        underline='none',
        strike=False,
        color='FF005500'
    )
    font_data_not_exsist = Font(
        name='Tahoma',
        size=9,
        bold=True,
        italic=False,
        vertAlign=None,
        underline='none',
        strike=False,
        color='FFAA0000'
    )
    alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    max_col = sheet.max_column
    tmp_list = []
    for i in range(1, 5):
        for j in range(5, sheet.max_row):
            x = sheet.cell(row=j, column=max_col + i)
            x.font = font_header
            x.alignment = alignment

    sheet[5][7].value = 'Сертификат'
    sheet[5][7].font = font_header
    sheet[5][7].alignment = alignment
    sheet.column_dimensions['H'].width = 20

    sheet[5][8].value = 'Тип'
    sheet[5][8].font = font_header
    sheet[5][8].alignment = alignment

    sheet[5][9].value = 'Дата начала действия сертификата'
    sheet[5][9].font = font_header
    sheet[5][9].alignment = alignment

    sheet[5][10].value = 'Дата окончания действия сертификата'
    sheet[5][10].font = font_header
    sheet[5][10].alignment = alignment

    max_col = sheet.max_column
    for i in range(len(all_article_from_db)):
        for j in range(len(all_article_from_db[i])):
            tmp_list.append(all_article_from_db[i][j])
    all_article_from_db = tmp_list
    decode_dict = decode_certificate_types()
    for row in range(6, sheet.max_row):
        if sheet[row][1].value in all_article_from_db:
            certificate_id = get_certificate_id_by_article_query(sheet[row][1].value)
            sheet[row][max_col - 4].value = certificates_dict[certificate_id][0]
            sheet[row][max_col - 4].font = font_data_exsist

            sheet[row][max_col - 3].value = decode_dict[certificates_dict[certificate_id][1]]
            sheet[row][max_col - 3].font = font_data_exsist

            sheet[row][max_col - 2].value = certificates_dict[certificate_id][2]
            sheet[row][max_col - 2].font = font_data_exsist

            sheet[row][max_col - 1].value = certificates_dict[certificate_id][3]
            sheet[row][max_col - 1].font = font_data_exsist

        else:
            sheet[row][max_col - 4].value = 'Нет данных'
            sheet[row][max_col - 4].font = font_data_not_exsist

            sheet[row][max_col - 3].value = 'Нет данных'
            sheet[row][max_col - 3].font = font_data_not_exsist

            sheet[row][max_col - 2].value = 'Нет данных'
            sheet[row][max_col - 2].font = font_data_not_exsist

            sheet[row][max_col - 1].value = 'Нет данных'
            sheet[row][max_col - 1].font = font_data_not_exsist

    book.save(rf"certificates_list_by_article.xlsx")
    book.close()


def decode_certificate_types():
    decode_dict = make_dict(universal_query('certificates_types', '*'))
    reverse_decode_dict = dict((v, k) for k, v in decode_dict.items())
    return reverse_decode_dict


# endregion


# region Manufacturers

def get_manufacturer_id_by_article_query(article, ):
    try:
        connect = psycopg2.connect(dbname=os.getenv('db_name'), user=os.getenv('user'),
                                   password=os.getenv('password'), host=os.getenv('host'))
        connect.autocommit = True
        cursor = connect.cursor()
        cursor.execute(
            f"SELECT manufacturer_id FROM bilight_products WHERE product_id = '{article}' or order_code = '{article}'")
        manufacturer_name = cursor.fetchone()[0]
        return manufacturer_name
    except Exception as _ex:
        print(f"[INFO] ERROR while working with data base  {_ex}")
    finally:
        cursor.close()
        connect.close()

        def make_manufacturer_list_file(file_path, all_article_from_db, reverse_manufacturer_dict):
            shutil.copy(rf"{file_path}", rf"manufacturer_list_by_articles.xlsx")
            book = openpyxl.open(rf"manufacturer_list_by_articles.xlsx", read_only=False, data_only=True)
            sheet = book.active
            font_header = Font(
                name='Tahoma',
                size=9,
                bold=True,
                italic=False,
                vertAlign=None,
                underline='none',
                strike=False,
                color='FF000000'
            )
            font_data_exsist = Font(
                name='Tahoma',
                size=9,
                bold=True,
                italic=False,
                vertAlign=None,
                underline='none',
                strike=False,
                color='FF005500'
            )
            font_data_not_exsist = Font(
                name='Tahoma',
                size=9,
                bold=True,
                italic=False,
                vertAlign=None,
                underline='none',
                strike=False,
                color='FFAA0000'
            )
            alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            max_col = sheet.max_column
            tmp_list = []
            sheet.cell(row=5, column=max_col + 1)
            sheet[5][max_col].value = 'Производитель'
            sheet[5][max_col].font = font_header
            sheet[5][max_col].alignment = alignment
            sheet.column_dimensions['H'].width = 20

            for i in range(len(all_article_from_db)):
                for j in range(len(all_article_from_db[i])):
                    tmp_list.append(all_article_from_db[i][j])
            all_article_from_db = tmp_list
            for row in range(6, sheet.max_row):
                if sheet[row][1].value in all_article_from_db:
                    manufacturer_id = get_manufacturer_id_by_article_query(sheet[row][1].value)
                    sheet.cell(row=row, column=max_col + 1)
                    sheet[row][max_col].value = reverse_manufacturer_dict[manufacturer_id]
                    sheet[row][max_col].font = font_data_exsist
                    sheet[row][max_col].alignment = alignment

                else:
                    sheet.cell(row=row, column=max_col + 1)
                    sheet[row][max_col].value = 'Нет данных'
                    sheet[row][max_col].font = font_data_not_exsist
                    sheet[row][max_col].alignment = alignment

            book.save(rf"manufacturer_list_by_articles.xlsx")
            book.close()


# endregion


# region TNVED
def get_tnved_by_article(article):
    try:
        connect = psycopg2.connect(dbname=os.getenv('db_name'), user=os.getenv('user'),
                                   password=os.getenv('password'), host=os.getenv('host'))
        connect.autocommit = True
        cursor = connect.cursor()
        cursor.execute(
            f"SELECT tnved_id FROM bilight_products WHERE product_id = '{article}' or order_code = '{article}'")
        tnved_id = cursor.fetchone()[0]
        return tnved_id
    except Exception as _ex:
        print(f"[INFO] ERROR while working with data base  {_ex}")
    finally:
        cursor.close()
        connect.close()


def make_tnved_dict():
    tnved_dict = make_dict(universal_query('tnved', '*'))
    reverse_tnved_dict = dict((v, k) for k, v in tnved_dict.items())
    return reverse_tnved_dict


def make_tnved_list_file(file_path, all_article_from_db, tnved_dict):
    shutil.copy(rf"{file_path}", rf"tnved_list_by_article.xlsx")
    book = openpyxl.open(rf"tnved_list_by_article.xlsx", read_only=False, data_only=True)
    sheet = book.active
    font_header = Font(
        name='Tahoma',
        size=9,
        bold=True,
        italic=False,
        vertAlign=None,
        underline='none',
        strike=False,
        color='FF000000'
    )
    font_data_exsist = Font(
        name='Tahoma',
        size=9,
        bold=True,
        italic=False,
        vertAlign=None,
        underline='none',
        strike=False,
        color='FF005500'
    )
    font_data_not_exsist = Font(
        name='Tahoma',
        size=9,
        bold=True,
        italic=False,
        vertAlign=None,
        underline='none',
        strike=False,
        color='FFAA0000'
    )
    alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    max_col = sheet.max_column
    tmp_list = []
    for i in range(1, 5):
        for j in range(5, sheet.max_row):
            x = sheet.cell(row=j, column=max_col + i)
            x.font = font_header
            x.alignment = alignment

    sheet[5][7].value = 'Код ТНВЭД'
    sheet[5][7].font = font_header
    sheet[5][7].alignment = alignment
    sheet.column_dimensions['H'].width = 20

    sheet[5][8].value = 'ОПИСАНИЕ'
    sheet[5][8].font = font_header
    sheet[5][8].alignment = alignment
    sheet.column_dimensions['I'].width = 150

    max_col = sheet.max_column
    for i in range(len(all_article_from_db)):
        for j in range(len(all_article_from_db[i])):
            tmp_list.append(all_article_from_db[i][j])
    all_article_from_db = tmp_list
    for row in range(6, sheet.max_row):
        if sheet[row][1].value in all_article_from_db:
            tned_id = get_tnved_by_article(sheet[row][1].value)
            sheet[row][max_col - 4].value = tned_id
            sheet[row][max_col - 4].font = font_data_exsist

            sheet[row][max_col - 3].value = tnved_dict[tned_id]
            sheet[row][max_col - 3].font = font_data_exsist

        else:
            sheet[row][max_col - 4].value = 'Нет данных'
            sheet[row][max_col - 4].font = font_data_not_exsist

            sheet[row][max_col - 3].value = 'Нет данных'
            sheet[row][max_col - 3].font = font_data_not_exsist

    book.save(rf"tnved_list_by_article.xlsx")
    book.close()
# endregion
