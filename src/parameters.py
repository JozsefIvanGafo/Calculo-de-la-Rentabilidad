import os

class Parameters():
    def __init__(self):
        self._num_simulations = 10000
        self._data_path = os.path.join(os.path.abspath(os.getcwd()), 'data', 'INPUT.xlsx')
        self._sheet_name = 'BaseCase'

    @property
    def num_simulations(self):
        return self._num_simulations
    @property
    def data_path(self):
        return self._data_path
    @property
    def sheet_name(self):
        return self._sheet_name