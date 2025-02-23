import win32com.client
import numpy as np
from tqdm import tqdm
from .functions import generate_random_numbers 
import seaborn as sns
from concurrent.futures import ThreadPoolExecutor

class MonteCarloSimulator():
    #TODO: Make it faster
    #TODO: Refactor the code
    #TODO: Paralelize the code


    def __init__(self,params):
        self.__params = params
        self.__debug = False
        self.__see_excel=False
        self.__tir_values = []
        self.__workbook = None
        self.__sheet = None
        self.__excel_file = None
        self.__extract_excel()


    
    def __run(self):
        print("Running Monte Carlo Simulation")
        num_iterations = self.__params.num_simulations
        random_num_list=[generate_random_numbers(3) for _ in range(num_iterations)]

        
        for random_list in tqdm(random_num_list, desc="Simulations progress"):
            self.__tir_values.append(self.__calculate_an_iteration(random_list))
        self.__close_excel()
        
        self.__statistics()


    def __plot(self):
        
        print("Plotting Monte Carlo Simulation")
        #using seaborn create a graphic with the tir_values
        import matplotlib.pyplot as plt

        sns.set(style="whitegrid")
        # Create a histogram with KDE of the TIR values
        sns.histplot(self.__tir_values, kde=True, bins=30, color='blue')
        plt.title("Distribution of TIR Values")
        plt.xlabel("TIR")
        plt.ylabel("Frequency")
        plt.show()
        pass

    def __calculate_an_iteration(self,random_nums):
        
        inversion_multp,ingresos_multp,costes_totales_multp=random_nums
        if self.__debug:
            print(f"Inversion multiplier: {inversion_multp}")
            print(f"Ingresos multiplier: {ingresos_multp}")
            print(f"Costes totales multiplier: {costes_totales_multp}")

        self.__multiply_row(self.__params.inversion_row,inversion_multp)
        self.__multiply_row(self.__params.ingresos_row,ingresos_multp)
        self.__multiply_row(self.__params.costes_totales_row,costes_totales_multp)

        #update the Workbook
        self.__sheet.Calculate()

        #get the TIR
        tir=self.__sheet.Cells(self.__params.tir_cell[0],self.__params.tir_cell[1]).Value
        if self.__debug:
            print(f'TIR: {tir}')

        #reset workbook and excel file
        self.__reset_sheet()

        return tir

    def __statistics(self):
        # Mean (Average)
        mean_np = np.mean(self.__tir_values)
        print(f'Mean (numpy): {mean_np}')

        # Median
        median_np = np.median(self.__tir_values)
        print(f'Median (numpy): {median_np}')

        # Standard Deviation
        std_dev_np = np.std(self.__tir_values, ddof=1)  # ddof=1 for sample standard deviation
        print(f'Standard Deviation (numpy): {std_dev_np}')

        # Variance
        variance_np = np.var(self.__tir_values, ddof=1)
        print(f'Variance (numpy): {variance_np}')

    def __multiply_row(self, row, multiplier):
        column = 2
        while True:
            cell_value = self.__sheet.Cells(row, column).Value
            if cell_value is None or cell_value == "":
                break
            self.__sheet.Cells(row, column).Value = cell_value * multiplier
            column += 1
        
    def __reset_sheet(self):
        """ Closes and reopens the workbook to reset all changes. """
        self.__workbook.Close(SaveChanges=False)
        self.__workbook = self.__excel_file.Workbooks.Open(self.__params.data_path)
        self.__sheet = self.__workbook.Worksheets(1)
        

    
    def __extract_excel(self):
        try:
            self.__excel_file = win32com.client.Dispatch("Excel.Application")
            self.__excel_file.Visible = self.__debug or self.__see_excel# Ensure Excel window is open (optional)
            self.__workbook = self.__excel_file.Workbooks.Open(self.__params.data_path)
            self.__sheet = self.__workbook.Worksheets(1)
        except FileNotFoundError:
            print(f'[ERROR ID=001]: File not found {self.__params.data_path}')
        except Exception as e:
            print(f'[ERROR ID=002] {self.__params.data_path}: {e}')
    
    def __close_excel(self):
        if hasattr(self, '__workbook'):
            self.__workbook.Close(SaveChanges=False)
        if hasattr(self, '__excel_file'):
            self.__excel_file.Quit()

    @property
    def run(self):
        return self.__run
    @property
    def plot(self):
        return self.__plot