class CellValueCoord():
    def __init__(self,row,column):
        self.__row=row
        self.__column=column

    @property
    def row(self):
        return self.__row
    @property
    def column(self):
        return self.__column
