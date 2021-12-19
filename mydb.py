import mysql.connector as mysql

def db_login(hn, un, pw):
    result = None

    try:
        db = mysql.connect(
            host=hn,
            username=un,
            password=pw
        )
        c = db.cursor()
        query = "SHOW DATABASES"
        c.execute(query)

        rm_list = ['information_schema', 'performance_schema', 'sys', 'mysql']
        result = []

        for item in c.fetchall():
            if item[0] in rm_list: # check if item is in rm_list
                continue;

            result.append(item)

        db.close()
    except mysql.Error as err:
        result = str(err)

    return result

def get_tables(hn, un, pw, db_name):

    try:
        db = mysql.connect(
            host=hn,
            username=un,
            password=pw,
            database=db_name
        )
        c = db.cursor()
        query = "SHOW TABLES"
        c.execute(query)
        records = c.fetchall()

        db.close()
    except mysql.Error as err:
        str(err)

    return records

def get_records(hn, un, pw, db_name, table_name):
    records = None

    try:
        db = mysql.connect(
            host=hn,
            username=un,
            password=pw,
            database=db_name
        )
        c = db.cursor()
        query = f"SELECT * FROM {table_name}"
        c.execute(query)
        records = c.fetchall()
        column_names = c.description

        result = records, column_names

        db.close()
    except mysql.Error as err:
        result = str(err)

    return result
