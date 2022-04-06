class InitData:

    __target_wb_path = 0

    __src_wb_path = 0

    __date = 0

    def __init__(self,target_wb_path,src_wb_path,date):
        self.__target_wb_path = target_wb_path
        self.__src_wb_path = src_wb_path
        self.__date = date

    def settarget_wb_path(self, target_wb_path):
        self.__target_wb_path = target_wb_path

    def gettarget_wb_path(self):
        return self.__target_wb_path

    def setsrc_wb_path(self, src_wb_path):
        self.__src_wb_path = src_wb_path

    def getsrc_wb_path(self):
        return self.__src_wb_path

    def setdate(self,date):
        self.__date = date

    def getdate(self):
        return self.__date