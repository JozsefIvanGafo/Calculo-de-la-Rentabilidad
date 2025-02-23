import win32com.client
import multiprocessing
import numpy as np
import seaborn as sns
from tqdm import tqdm
from concurrent.futures import ThreadPoolExecutor
#Project imports
from .functions import generate_random_numbers 
from .excel_handler import ExcelHandler
from .sim_types import Parameters


class MonteCarloSimulator():
    #TODO: Parallelize the code


    def __init__(self,params:Parameters):
        self.__params = params
        #Debugging options
        self.__debug = False
        self.__see_excel=True
        #declare variables need it for the simulation
        self.__tir_values = []
        #Variables of the excel
        self.__excel_obj = ExcelHandler(self.__params.data_path)

        #We count the number of cores and substract 2
        cores=multiprocessing.cpu_count()
        if cores>2:
            self.__useful_cores = cores-2
        else:
            self.__useful_cores=1

        self.__e_file,self.__wkbk,self.__sheet=\
            self.__excel_obj.extract_excel(visible=self.__see_excel)

    
    def run(self):
        print("Running Monte Carlo Simulation")
        num_iterations = self.__params.num_simulations
        random_num_list=[generate_random_numbers(3) for _ in range(num_iterations)]

        #TODO: We use the ThreadPoolExecutor to parallelize the simulation

        
        for random_list in tqdm(random_num_list, desc="Simulations progress"):
            self.__tir_values.append(self.__calculate_an_iteration(random_list,self.__sheet))

        #We close the excel file
        self.__excel_obj.close_excel(self.__wkbk,self.__e_file)
        #print statistics of the simulation
        self.__statistics()
        #plot the simulation
        self.__plot()


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

    def __calculate_an_iteration(self,random_nums,sheet):
        
        inversion_multp,ingresos_multp,costes_totales_multp=random_nums
        if self.__debug:
            print(f"Inversion multiplier: {inversion_multp}")
            print(f"Ingresos multiplier: {ingresos_multp}")
            print(f"Costes totales multiplier: {costes_totales_multp}")

        self.__change_multipliers(random_nums,sheet)


        #update the Workbook
        sheet.Calculate()

        #get the TIR
        tir=self.__excel_obj.get_cell_value(
            sheet,self.__params.tir_cell)
        
        
        if self.__debug:
            print(f'TIR: {tir}')

        return tir
    
    def __change_multipliers(self,random_nums,sheet):
        #Change the multipliers in the excel file
        for i, (_, multp_coord) in enumerate(self.__params.multp_cell.items()):
            number_change = 1 + random_nums[i]
            self.__excel_obj.write_on_cell(sheet, multp_coord, number_change)

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

        