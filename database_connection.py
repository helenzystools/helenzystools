import mysql.connector


def database_connection():
    db_connection = mysql.connector.connect(
        host="localhost",
        user="root",
        password="Power1234",
        database="test",
        auth_plugin='mysql_native_password'
    )

    db_cursor = db_connection.cursor()
    return db_connection, db_cursor


if __name__ == '__main__':
    database_connection()
