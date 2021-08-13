import xlwings as xw
from index_helpers import block_to_list
from test_gui import file_address_gui
import datetime
import re

import webbrowser

"""
Link capabilities
"""

def add_link():
    file_address = file_address_gui()[0] #gets the address of the file clicked on
    xw.apps.active.selection.api.NoteText(file_address)

def open_link():
    """
    TO-DO:
        Error handling for if a valid link isn't provided
    """
    file_name = xw.apps.active.selection.note.text
    webbrowser.open(file_name, new=2) #new=2 opens it in a new tab



"""
Coloring expiry dates

"""
def char_lim_255(address):
    """
    breaks a string of addresses into ones which are sub<255 chars

    Used in, for example, color_cells, in case the range selected has addresses of length >=255 chars
    """

    if len(address) > 255:
        for i in reversed(range(256)):
            if address[i] == ",":
                return [address[:i]] + char_lim_255(address[i+1:])
            else:
                pass
    else:
        return [address]


def color_cells(color_dict, wb):
    """
    color_dict maps integers between 1 and 56 inclusive to lists of cells which are to be colored that color
    
    
    (1 to 56 are the numbers excel uses for its color chart)
    
    !wb should already be defined in the python file to be the current workbook!
    """
    for color in color_dict:
        if color_dict[color] != []:
            addresses = ",".join(color_dict[color])
            for addresses_sub_255 in char_lim_255(addresses):
                wb.sheets.active.range(addresses_sub_255).api.Interior.ColorIndex = color
        else:
            #if the array is empty, there is nothing to color, so we pass
            pass
    return

def highlight_dates_to_expiry():
    """
    Given a selection of cells which are dates
    TO-DO:
        ERROR HANDLING FOR IF THE VALUES ARE NOT DATETIMES (partially done 13.08.21)
        Change the color coding based on datetimes to ones requested. (Currently: 1 year green, 6mths-1yr yellow, less than 
        6 months is red. Should probably be more of a gradation, perhaps with red set aside for if it has expired?)
    
    """
    wb = xw.apps.active.books.active #high cost
    user_selection = wb.selection #high cost cost
    values = user_selection.value #low cost
    if type(values) != list:
        #e.g., if we select a single cell, this puts it in an array, so that the code runs in the same way 
        #for multiple values and for single values
        values = [values]
    else:
        pass
    
    address = user_selection.address #low cost
    address_as_list = block_to_list(address).split(",")

    green_list = []
    yellow_list = []
    red_list = []
    purple_list = [] #for cells whose values we couldn't read as a date

    for i, val in enumerate(values):
        if val == None:
            continue
        else:
            pass
        
        if type(val) != datetime.datetime:
            #tries to convert to a datetime if the input formula is a bit off e.g., 18/9/18 instead of 18/9/2018
            a = re.match(r"([0-2]{0,1}[0-9]||30||31)[/.]([0]{0,1}[1-9]||[1][0-2])[/.][0-9][0-9]", val)
            if a:
                #we matched the regex
                try:
                    #try to make a datetime
                    a = re.split("[.]|/",a.group())
                    val = datetime.datetime(day=a[0], month=a[1], year=int("20"+a[2]))
                except:
                    purple_list.append(address_as_list[i])
            else:
                purple_list.append(address_as_list[i])


        else:
            pass


        #below we check the dates and compare to now. 
        if type(val) == datetime.datetime:
            #for val's which were a datetime or we managed to make a datetime, we execute this
            timedif =val - datetime.datetime.today() #val is the expiry date of the DBS
            if timedif > datetime.timedelta(days=360):
                green_list.append(address_as_list[i])
            elif (timedif <= datetime.timedelta(days=360))and(timedif > datetime.timedelta(days=180)):
                yellow_list.append(address_as_list[i])
            else:
                red_list.append(address_as_list[i])
        else:
            pass

    
    color_cells(color_dict={4:green_list, 3: red_list, 13:purple_list,6:yellow_list}, wb=wb)
if __name__ == '__main__':
    # Expects the Excel file next to this source file, adjust accordingly.
    xw.Book(r"C:\Users\ethan\Documents\Mo Forgan demo\Mentor student template.xlsm").set_mock_caller()
    open_link()


