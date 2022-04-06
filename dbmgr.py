from PyQt5 import QtSql,QtWidgets
import sqlite3
class DbManager:
    shop_dict_cache = {}
    __database = 0
    __conn = 0
    def __init__(self,database_file):
        self.__conn = sqlite3.connect(database_file)
        # database = QtSql.QSqlDatabase.addDatabase("QSQLITE")
        # database.setDatabaseName(database_file)
        # db = database.open()
        # if not db:
        #     QtWidgets.QMessageBox.critical(None,"错误","连接数据库错误")
        #     return
        # self.__database = db

    # def getsqlquery(self):
    #     if self.__database == 0:
    #         QtWidgets.QMessageBox.critical(None, "错误", "连接数据库错误，不能获取连接")
    #     return QtSql.QSqlQuery()

    def getconn(self):
        return self.__conn

    def close(cursor, connection):
        cursor.close()
        connection.close
    # def getshop_dict_cache(self):
    #     return self.__shop_dict_cache