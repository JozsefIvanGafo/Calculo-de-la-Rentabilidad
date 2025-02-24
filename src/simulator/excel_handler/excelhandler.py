import os
import win32com.client
#Project import
from simulator.exception import SimulationException
from ..sim_types import CellValueCoord

class ExcelHandler():
    def __init__(self,data_path):
        print("ExcelHandler Initialized ...")
        if not os.path.exists(data_path):
            raise SimulationException(f'[EXCEL ERROR ID=000]: Path does not exist {data_path}')
        self.__data_path=data_path
        

    def extract_excel(self,visible=False):
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
        

    def reset_sheet(self,wkbk,sheet,e_file):
        """ Closes and reopens the workbook to reset all changes. """
        wkbk.Close(SaveChanges=False)
        wkbk = e_file.Workbooks.Open(self.__data_path)
        sheet = wkbk.Worksheets(1)
        return wkbk,sheet
    
    def copy_sheet(self, wkbk, sheet, e_file):
        """
        Copies the specified sheet and returns the new sheet.
        """
        try:
            # Copy the sheet after itself
            sheet.Copy(After=sheet)
            
            # Get the newly created sheet (it will be immediately after the original sheet)
            new_sheet_index = sheet.Index + 1
            new_sheet = wkbk.Worksheets(new_sheet_index)
            
        except Exception as e:
            raise SimulationException(f'[EXCEL ERROR ID=005]: {e}')
        
        return new_sheet
        

    
    def close_excel(self,wkbk,e_file):
        try:
            wkbk.Close(SaveChanges=False)
            e_file.Quit()
        except Exception as e:
            raise SimulationException(f'[EXCEL ERROR ID=003]: {e}')
        
    def write_on_cell(self,sheet,coord:CellValueCoord,value):
        try:
            sheet.Cells(coord.row, coord.column).Value = value

        except Exception as e:
            raise SimulationException(
            f'[EXCEL ERROR ID=004] Error writing to cell ({coord.row}, {coord.column}): {e}')
        
    def get_cell_value(self,sheet,coord:CellValueCoord):
        try:
            return sheet.Cells(coord.row, coord.column).Value
        except Exception as e:
            raise SimulationException(
            f'[EXCEL ERROR ID=006] Error reading cell ({coord.row}, {coord.column}): {e}')
        
