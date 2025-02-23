import win32com.client
#Project import
from simulator.exception import SimulationException

class ExcelHandler():
    def __init__(self,data_path):
        print("ExcelHandler Initialized ...")
        self.__data_path=data_path

    def __reset_sheet(self,wkbk,sheet,e_file):
        """ Closes and reopens the workbook to reset all changes. """
        wkbk.Close(SaveChanges=False)
        wkbk = e_file.Workbooks.Open(self.__data_path)
        sheet = wkbk.Worksheets(1)
        return wkbk,sheet
    
    def __extract_excel(self,visible=False):
        """Extracts the Excel file and returns the Excel file, workbook and sheet"""
        try:
            e_file = win32com.client.Dispatch("Excel.Application")
            e_file.Visible = visible# Ensure Excel window is open (optional)
            wkbk = e_file.Workbooks.Open(self.__data_path)
            sheet = wkbk.Worksheets(1)
        except FileNotFoundError:
             raise SimulationException(
                f'[EXCEL ERROR ID=001]: File not found {self.__data_path}')

        except Exception as e:
            raise SimulationException(
                f'[EXCEL ERROR ID=002] {self.__data_path}: {e}')

        return e_file,wkbk,sheet
    
    def __close_excel(self,wkbk,e_file):
        try:
            wkbk.Close(SaveChanges=False)
            e_file.Quit()
        except Exception as e:
            raise SimulationException(f'[EXCEL ERROR ID=003]: {e}')
        
    @property
    def close_excel(self):
        return self.__close_excel
    @property  
    def extract_excel(self):
        return self.__extract_excel
    @property
    def reset_sheet(self):
        return self.__reset_sheet
