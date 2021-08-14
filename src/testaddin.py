#opens the excel sheet with our macros in the background. Note this should be stored in the data
#part of the github repo, which we can then find again automatically. 


import xlwings as xw
import sys
from pathlib import Path
DATA_DIR = Path(sys.executable).resolve().parent / 'data'


"""
For UDF function to be added to add-in, which uses RunPython to call these

TO-DO: dynamic paths (currently static) by installing to the data section of github repo


ALSO: currently not working
"""


"""
@xw.func
def open_macros():
    xw.apps.active.books.open(r"C:\Users\ethan\Documents\excel stuff\testMacros.xlsm")

@xw.func
def close_macros():
    wb = xw.apps.active.books.open(r"C:\Users\ethan\Documents\excel stuff\testMacros.xlsm") 
    #if already open the above does nothing. I.e. just sets wb to reference the instance
    #if not open, it's a little inefficient in that it opens and closes it again, but it's fast 

    wb.close()


"""