import os
import psycopg2
import openpyxl
from dotenv import find_dotenv, load_dotenv

load_dotenv(find_dotenv())


def main():
    # add_certificates(get_cert_info_from_user("cert_upload_template.xlsx"))
    get_product_info_from_user('products_upload_template.xlsx')



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
        print(certification_data_list)
        return certification_data_list
    else:
        raise ValueError("The function argument must be a file")


def add_certificates(cert_data_from_user):
    try:
        connect = psycopg2.connect(dbname=os.getenv('db_name'), user=os.getenv('user'),
                                   password=os.getenv('password'), host=os.getenv('host'))
        connect.autocommit = True
        cursor = connect.cursor()
        print("Postgre SQL successfully conected")

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

def query(table_name,*args):
    args = ','.join(args)
    try:
        connect = psycopg2.connect(dbname=os.getenv('db_name'), user=os.getenv('user'),
                                   password=os.getenv('password'), host=os.getenv('host'))
        connect.autocommit = True
        cursor = connect.cursor()
        print("Postgre SQL successfully conected")
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
        print("Postgre SQL Connection closed")


def get_product_info_from_user(file_path):
    if os.path.isfile(file_path):
        manufacturer_dict = {}
        for x in query('manufacturers','*'):
            if x[1] not in manufacturer_dict:
                manufacturer_dict[x[1]] = x[0]
            else:
                pass
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
    print(products_data_list)





if __name__ == '__main__':
    main()
