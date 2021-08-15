import xlwings as xw
from index_helpers import block_to_list
from test_gui import file_address_gui, simple_error
import datetime
import re

import webbrowser

"""
Link capabilities
"""

def add_link():
    file_address = file_address_gui()[0] #gets the address of the file clicked on
    if file_address == "":
        error_title = "A small hiccup"
        error_body = "You forgot to enter a link!"
        simple_error(error_title=error_title, error_body=error_body)
    else:
        create_comment = xw.books["MACROS_SHFX.xlsm"].macro(r'AddComment') #this assumes MACROS_SHFX is open
        create_comment(file_address)

def open_link():
    """
    TO-DO:
        slightly more sophisticated error handling
    """
    try:
        file_name = xw.apps.active.selection.note.text
        webbrowser.open(file_name, new=2) #new=2 opens it in a new tab
    except:
        error_title = "A small hiccup..."
        error_body = "We couldn't read your file link. Try adding again, or contact support!"
        simple_error(error_title=error_title, error_body=error_body)



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
    color_dict maps RGB tuples (0:255, 0:255, 0:255) to lists of cells which are to be colored that color
    
    empty lists are ignored and so can be passed in safely

    !wb should already be defined in the python file to be the current workbook!
    """
    for color in color_dict:
        if color_dict[color] != []:
            addresses = ",".join(color_dict[color])
            for addresses_sub_255 in char_lim_255(addresses):
                wb.sheets.active.range(addresses_sub_255).color = color
        else:
            #if the array is empty, there is nothing to color, so we pass
            pass
    return

def helper_date_to_expiry(values, address_as_list, purple_list, red_list, yellow_list, green_list):
    """
    Helper code for highlight_dates_to_expiry, for code which is needed in both the case where the user selects
    multiple blocks, and the case where the user selects only one block
    """    
    
    if type(values) != list:
    #e.g., if we select a single cell, this puts it in an array, so that the code runs in the same way 
    #for multiple values and for single values
        values = [values]
    else:
        pass
    
    if type(values[0]) == list:
        #if it is an array then we raise an error message, as this is meant to act on columns and single cells
        #the reason for this is twofold: (1) from the sheets shown, it seems if an array is selected it's not actually likely
        #to be helpful for this function and (2) it makes P(bug) higher in my experience, so if it is useful/requested
        #we should do it, but not pre-emptively
        simple_error(error_title="A minor hiccup...",error_body="Oops - please only select columns or individual cells")
        return
    else:
        pass
    
    for i, val in enumerate(values):
        if val == None:
            continue
        else:
            pass

        if type(val) != datetime.datetime:
            #tries to convert to a datetime if the input formula is a bit off e.g., 18/9/18 instead of 18/9/2018
            #str(val) is because regex requires a string to be passed in
            a = re.match(r"([0-2]{0,1}[0-9]||30||31)[/.]([0]{0,1}[1-9]||[1][0-2])[/.][0-9][0-9]", str(val))
            if a:
                #we matched the regex
                try:
                    #try to make a datetime, e.g. for 12.05.20 
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


def highlight_dates_to_expiry():
    """
    Given a selection of cells which are dates
    TO-DO:
        ERROR HANDLING FOR IF THE VALUES ARE NOT DATETIMES
        Change the color coding based on datetimes to ones requested. (Currently: 1 year green, 6mths-1yr yellow, less than 
        6 months is red. Should probably be more of a gradation, perhaps with red set aside for if it has expired?)
    
    """
    wb = xw.apps.active.books.active #high cost
    user_selection = wb.selection #high cost cost
    green_list = []
    yellow_list = []
    red_list = []
    purple_list = [] #for cells whose values we couldn't read as a date



    address = user_selection.address #low cost

    if "," in address:
        #means multiple columns have been passed in, and the handling is slower, and more complicated
        address_columns = address.split(",")
        for block in address_columns:
            address_as_list = block_to_list(block).split(",")
            values = wb.sheets.active.range(block).value
            helper_date_to_expiry(values=values, address_as_list=address_as_list, 
                          purple_list=purple_list, green_list=green_list, red_list=red_list, yellow_list=yellow_list)


    else:
        #a single block was passed in, and handling is simpler
        values = user_selection.value #low cost
        address_as_list = block_to_list(address).split(",")


        helper_date_to_expiry(values=values, address_as_list=address_as_list, 
                              purple_list=purple_list, green_list=green_list, red_list=red_list, yellow_list=yellow_list)




    color_cells(color_dict={(102,255,102):green_list, (255,102,102): red_list, (255,153,255):purple_list,(255,255,153):yellow_list}, wb=wb)

if __name__ == '__main__':
    # Expects the Excel file next to this source file, adjust accordingly.
    xw.Book(r"C:\Users\ethan\Documents\Mo Forgan demo\Mentor student template.xlsm").set_mock_caller()
    open_link()


