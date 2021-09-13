import os
import xlwings as xw
import index_helpers
from test_gui import file_address_gui, simple_error
import datetime
import re
import webbrowser
from variables import MACROS_SHFX_location, undo_dict, undo_control
import platform
from undo import extract_range_fill_data
from doubllist import doubllistNode, doubllist


"""
UNDO function
"""

def undo_wrapper():
    undo_control["UNDO"] = True
    undo_control["TIME"] = datetime.datetime.now()



"""
TEST FUNCTIONS - SHOULD REMOVE AT SOME POINT!
"""

def nuclear_dashboard():   
    webbrowser.open(r"https://www.blitzortung.org/en/live_lightning_maps.php", new=2)


"""
Link capabilities
"""
def add_link():
    response = file_address_gui()
    if response != "":
        macro_address = MACROS_SHFX_location[platform.system()]
        if platform.system() == "Windows":
            create_comment = xw.Book(macro_address+"\\" + "MACROS_SHFX.xlam").macro(r'AddComment')
        else:
            create_comment = xw.Book("/Users/fourthuser/Downloads/MACROS_SHFX-3.xlam").macro(r'AddComment')
        create_comment(response)
    else:
        #this means they cancelled the form, and so we add nothing
        pass

def open_link():
    """
    TO-DO:
        slightly more sophisticated error handling
    """
    try:
        if platform.system() == "Windows":
            file_name = xw.apps.active.selection.note.text

            webbrowser.open(file_name, new=2) #new=2 opens it in a new tab
        else:
            macro_address = MACROS_SHFX_location[platform.system()]
            read_comment = xw.Book(os.path.join(macro_address, "MACROS_SHFX.xlam")).macro(r'ReadComment')
            file_name = read_comment(xw.apps.active.books.active.selection.address)
            webbrowser.open(file_name, new=2) #new=2 opens it in a new tab
            #mac code should go here
    except:
        error_title = "A small hiccup..."
        error_body = "We couldn't read your file link. Try adding again, or contact support!"
        simple_error(error_title=error_title, error_body=error_body)



"""
Coloring expiry dates

"""



def color_cells(color_dict, wb):
    """
    color_dict maps RGB tuples (0:255, 0:255, 0:255) to lists of cells which are to be colored that color
    
    empty lists are ignored and so can be passed in safely

    !wb should already be defined in the python file to be the current workbook!
    """
    for color in color_dict:
        if color_dict[color] != []:
            addresses = ",".join(color_dict[color])
            for addresses_sub_255 in index_helpers.char_lim_255(addresses):
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
    user_selection = wb.selection #high cost
    active_sheet = wb.sheets.active

    ### UNDO/REDO STUFF

    #if the UNDO/REDO list doesn't exist, we create it and add our root node
    if active_sheet not in undo_dict["highlight_dates_to_expiry"]:
        starting_node = doubllistNode(parent=None, child=None, value={})
        undo_dict["highlight_dates_to_expiry"][active_sheet] = doubllist(starting_node)
        linked_list_of_actions = undo_dict["highlight_dates_to_expiry"][active_sheet]
        if platform.system() == "Windows":
            extract_range_fill_data(file_address=wb.api.Application.ActiveWorkbook.FullName, 
                    sheet_name=wb.api.ActiveSheet.Name, selection=user_selection.address, 
                    color_fill_dictionary=linked_list_of_actions.current_node.value)
        else:
            #for "Darwin", aka Mac
            macro_address = MACROS_SHFX_location[platform.system()] 
            workbook_address_macro = xw.Book("/Users/fourthuser/Downloads/MACROS_SHFX-3.xlam").macro(r'ActiveWorkbookAddress') 
            filename = workbook_address_macro()
            extract_range_fill_data(file_address=filename, 
                    sheet_name=wb.api.ActiveSheet.Name, selection=user_selection.address, 
                    color_fill_dictionary=linked_list_of_actions.current_node.value)
                
            
    else:
        pass


    ###

    green_list = []
    yellow_list = []
    red_list = []

    purple_list = [] #for cells whose values we couldn't read as a date



    address = user_selection.address #low cost

    if "," in address:
        #means multiple columns have been passed in, and the handling is slower, and more complicated
        address_columns = address.split(",")
        for block in address_columns:
            address_as_list = index_helpers.block_to_list(block).split(",")
            values = active_sheet.range(block).value
            helper_date_to_expiry(values=values, address_as_list=address_as_list, 
                          purple_list=purple_list, green_list=green_list, red_list=red_list, yellow_list=yellow_list)


    else:
        #a single block was passed in, and handling is simpler
        values = user_selection.value #low cost
        address_as_list = index_helpers.block_to_list(address).split(",")


        helper_date_to_expiry(values=values, address_as_list=address_as_list, 
                              purple_list=purple_list, green_list=green_list, red_list=red_list, yellow_list=yellow_list)




    color_cells(color_dict={(102,255,102):green_list, (255,102,102): red_list, (255,153,255):purple_list,(255,255,153):yellow_list}, wb=wb)


    ###UNDO/REDO STUFF






    linked_list_of_actions = undo_dict["highlight_dates_to_expiry"][active_sheet]
    linked_list_of_actions.insert_ahead({}) #creates a node ahead, with an empty 
                        #dictionary to be filled by extract_range_fill_data
    linked_list_of_actions.step_forward()
    linked_list_of_actions.current_node.del_descendants() #removes descendants
                    #as a new action has replaced those descendants
    extract_range_fill_data(file_address=wb.api.Application.ActiveWorkbook.FullName, 
            sheet_name=wb.api.ActiveSheet.Name, selection=user_selection.address, 
            color_fill_dictionary=linked_list_of_actions.current_node.value)


def save_workbook(wb):
    """
    saves the current workbook
    """
    #WILL NEED TO SORT OUT THE SAVING STUFF FOR RE-DO CAPABILITIES
    #MAYBE NEED TO CREATE CUSTOM EVENTS TO DEAL WITH THE 
    
    #create a snapshot of the new colors for undo/redo purposes:
    if platform.system() == "Windows":
        wb.api.Save #need to save the workbook so we can load up the new snapshot with openpyxl
        print("here")
    else:
        #this method seems to be c. 10x slower than the .api method 0.05s vs 0.005s
        macro_address = MACROS_SHFX_location[platform.system()] 
        save_wb_macro = xw.Book(macro_address+"\\" + "MACROS_SHFX.xlam").macro(r'saveActiveWorkbook') 
        save_wb_macro()

def blank_cell_background():
    """
    Gives all the given cells a blank background fill
    """
    xw.apps.active.selection.color = None


"""
Implement undo and redo functionality for color fill
"""
def undo_highlight_date_to_expiry():
    wb=xw.apps.active.books.active

    active_sheet = wb.sheets.active

    #get the current node in the linked list, step back to the step before, and load it
    if active_sheet in undo_dict["highlight_dates_to_expiry"]:
        linkedList= undo_dict["highlight_dates_to_expiry"][active_sheet]

        print(linkedList)
        linkedList.step_backward()
        color_cells(linkedList.current_node.value, wb=wb)

    else:
        pass

def redo_highlight_date_to_expiry():
    wb = xw.apps.active.books.active

    undo_dict["highlight_dates_to_expiry"][wb.sheets.active].step_forward()

    color_cells(undo_dict["highlight_dates_to_expiry"][wb.sheets.active].current_node.value, wb=wb)


def highlight_dates_to_expiry_wrapper():
    """this is what get called from within
    the excel sheet
    
    
    TO-DO. The undo linkedlist works, but I think the implementation could be improved. It relies on us
    saving the file with a macro (which might be annoying for the user?) - but is the only way to get
    the data in read only mode using openpyxl, which is fast
    """
    if undo_control["UNDO"] == True and (datetime.datetime.now() - undo_control["TIME"])<datetime.timedelta(seconds=1.5):
        undo_highlight_date_to_expiry()
    else:
        highlight_dates_to_expiry()
    undo_control["UNDO"] = False




"""
FOR SARAH DEMO

"""
from for_sarah import extract_emails, attach_data, find_record









"""
the below are used primarily for debugging
"""
def main():
    add_link()

    return


if __name__ == '__main__':
    # Expects the Excel file next to this source file, adjust accordingly.

    open_link()

    # macro_address = MACROS_SHFX_location[platform.system()] 
    # add_link_macro = xw.Book(macro_address+"\\" + "MACROS_SHFX.xlam").macro(r'add_link') 
    # add_link_macro()



