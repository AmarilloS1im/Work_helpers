import os
import psycopg2
import openpyxl
from dotenv import find_dotenv, load_dotenv

load_dotenv(find_dotenv())


def main():
    add_certificates(get_cert_info_from_user("cert_upload_template.xlsx"))
    # add_products(get_product_info_from_user('products_upload_template.xlsx'))
    # column_name_query('information_schema.columns','certificates','column_name')
    # pass
    # table_data_query('information_schema.columns', 'certificates', 'column_name')


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


def table_data_query(table_name, where_filter, *args):
    args = ','.join(args)
    try:
        connect = psycopg2.connect(dbname=os.getenv('db_name'), user=os.getenv('user'),
                                   password=os.getenv('password'), host=os.getenv('host'))
        connect.autocommit = True
        cursor = connect.cursor()
        cursor.execute(f"SELECT COUNT({args}) FROM {table_name}"
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


def make_replacement(existing_pkey, user_data, flag):
    columns_name_list = table_data_query('information_schema.columns', 'certificates', 'column_name')
    if flag:
        try:
            connect = psycopg2.connect(dbname=os.getenv('db_name'), user=os.getenv('user'),
                                       password=os.getenv('password'), host=os.getenv('host'))
            connect.autocommit = True
            cursor = connect.cursor()
            for i in range(len(existing_pkey)):
                for j in range(1, len(columns_name_list)):
                    cursor.execute(f" UPDATE certificates"
                                   f" SET {columns_name_list[j]} = %s"
                                   f" WHERE certificate_id = %s ", (user_data[i][j], user_data[i][0]))
        except Exception as _ex:
            print(f"[INFO] ERROR while working with data base {_ex}")
        finally:
            cursor.close()
            connect.close()
    else:
        pass


def get_cert_info_from_user(file_path):
    if os.path.isfile(file_path):
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
    else:
        raise ValueError("The function argument must be a file")


def add_certificates(cert_data_from_user):
    existing_pkey = universal_query('certificates', 'certificate_id')
    existing_pkey = [x[0] for x in existing_pkey]
    unique_user_data = []
    duplicate_user_data = []

    for i in range(len(cert_data_from_user)):
        if cert_data_from_user[i][0] in existing_pkey:
            duplicate_user_data.append(cert_data_from_user[i])
        else:
            unique_user_data.append(cert_data_from_user[i])
    print(duplicate_user_data)
    print(unique_user_data)

    try:
        connect = psycopg2.connect(dbname=os.getenv('db_name'), user=os.getenv('user'),
                                   password=os.getenv('password'), host=os.getenv('host'))
        connect.autocommit = True
        cursor = connect.cursor()
        if len(duplicate_user_data) > 0:
            if unique_user_data != '':
                cursor.executemany(f""" INSERT INTO certificates (certificate_id,certificate_number,certificate_type_id,
                                                        start_date,end_date)
                                                        VALUES
                                                        (%s,%s,%s,%s,%s) """, unique_user_data)
            else:
                pass
            add_duplicates_question = str(input(f'В базе данные обнаружены дубликаты данных'
                                                f' по следующим ID {duplicate_user_data}. Заменить данные? y/n?'))
            if add_duplicates_question == 'y':
                existing_pkey = duplicate_user_data
                make_replacement(existing_pkey,duplicate_user_data,flag=True)
            else:
                pass
        else:
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


def get_product_info_from_user(file_path):
    def make_manufacrurer_dict(query):
        manufacturer_dict = {}
        for x in query:
            if x[1] not in manufacturer_dict:
                manufacturer_dict[x[1]] = x[0]
            else:
                pass
        return manufacturer_dict
    if os.path.isfile(file_path):
        manufacturer_dict = make_manufacrurer_dict(universal_query('manufacturers', '*'))
        products_data_list = []
        book = openpyxl.open(file_path, read_only=True, data_only=True)
        sheet = book.active
        for row in range(2, sheet.max_row + 1):
            tmp_list = []
            for column in range(0, 7):
                tmp_list.append(sheet[row][column].value)
            products_data_list.append(tmp_list)
        for i in range(len(products_data_list)):
            if products_data_list[i][3] in manufacturer_dict:
                products_data_list[i][3] = manufacturer_dict[products_data_list[i][3]]
            else:
                try:
                    connect = psycopg2.connect(dbname=os.getenv('db_name'), user=os.getenv('user'),
                                               password=os.getenv('password'), host=os.getenv('host'))
                    connect.autocommit = True
                    cursor = connect.cursor()
                    print("Postgre SQL successfully conected")
                    cursor.execute(f"INSERT INTO manufacturers (manufacturer_name) "
                                   f"VALUES ('{products_data_list[i][3]}')")
                    manufacturer_dict = make_manufacrurer_dict(universal_query('manufacturers', '*'))
                    products_data_list[i][3] = manufacturer_dict[products_data_list[i][3]]
                except Exception as _ex:
                    print(f"[INFO] ERROR while working with data base {_ex}")
                finally:
                    cursor.close()
                    connect.close()
                    print("Postgre SQL Connection closed")
        return products_data_list
    else:
        raise ValueError("The function argument must be a file")


def add_products(products_data_from_user):
    try:
        connect = psycopg2.connect(dbname=os.getenv('db_name'), user=os.getenv('user'),
                                   password=os.getenv('password'), host=os.getenv('host'))
        connect.autocommit = True
        cursor = connect.cursor()
        print("Postgre SQL successfully conected")

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


if __name__ == '__main__':
    main()
