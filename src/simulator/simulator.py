from .functions import number_generator as ng
import win32com.client

class MonteCarloSimulator():

    def __init__(self,params):
        self.__params = params
        self.__extract_excel()
        pass

    
    def __run(self):
        print("Running Monte Carlo Simulation")
        self.__calculate_an_iteration()
        pass

    def __plot(self):
        print("Plotting Monte Carlo Simulation")
        pass

    def __calculate_an_iteration(self):
        inversion_multp,ingresos_multp,costes_totales_mutp=ng.generate_random_numbers(3)
        self.__excel.Sheets(1).


    
    def __extract_excel(self):
        try:
            excel=win32com.client.Dispatch("Excel.Application")
            self.__excel=excel.Workbooks.Open(self.__params.data_path)
        except FileNotFoundError:
            print(f'[ERROR ID=001]: File not found {self.__params.data_path}')
        except Exception as e:
            print(f'[ERROR ID=002] {self.__params.data_path}: {e}')
    

    @property
    def run(self):
        return self.__run
    @property
    def plot(self):
        return self.__plot