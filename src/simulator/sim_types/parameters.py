import os
from .coords import CellValueCoord

class Parameters():
    def __init__(self):
        self._num_simulations = 100
        self._data_path = os.path.join(os.path.abspath(os.getcwd()), 'data', 'INPUT.xlsx')
        self.__sheet_table_name = 'BaseCase'
        self.__inversion_row=4
        self.__ingresos_row=6
        self.__costes_totales_row=7
        self.__tir_cell=CellValueCoord(row=18,column=2)
        self.__multp_cell={
            "inversion":CellValueCoord(row=33,column=2),
            "ingresos":CellValueCoord(row=34,column=2),
            "costes_totales":CellValueCoord(row=35,column=2)
        }

    @property
    def num_simulations(self):
        return self._num_simulations
    @property
    def data_path(self):
        return self._data_path
    @property
    def sheet_name(self):
        return self.__sheet_table_name
    @property
    def inversion_row(self):
        return self.__inversion_row
    @property
    def ingresos_row(self):
        return self.__ingresos_row
    @property
    def costes_totales_row(self):
        return self.__costes_totales_row
    @property
    def tir_cell(self):
        return self.__tir_cell
    @property
    def multp_cell(self):
        return self.__multp_cell