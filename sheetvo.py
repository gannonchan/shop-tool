
class SheetVo:
    __datadict = 0
    __day = 0
    __shopName = 0

    def __init__(self,datadict,day,shopName):
        self.__datadict = datadict
        self.__day = day
        self.__shopName = shopName

    def setdict(self,dict):
        self.__datadict = dict

    def getdict(self):
        return self.__datadict

    def setday(self, day):
        self.__day = day

    def getday(self):
        return self.__day

    def setshopName(self,shop_name):
        self.__shopName = shop_name

    def getshopName(self):
        return self.__shopName
