"""
this is for functions used to undo commands
"""

import openpyxl
import io
from variables import InteriorColorToRGBdictOpenpyxl
from doubllist import doubllist, doubllistNode
from color_utilities import theme_and_tint_to_rgb, RGBA_to_rgbTuple
from test_gui import simple_error



def handle_cell_for_extract_range_fill_data(cell, color_fill_dictionary, wb):
    #means that it is a single cell
    cell_address = cell.coordinate

    cell_color=cell.fill.start_color.index


    if type(cell_color) == int:
        #this deals with the case where the cell_color is a theme and a tint, and converts it to RGB
        cell_color=theme_and_tint_to_rgb(wb=wb, theme=cell.fill.start_color.theme, tint=cell.fill.start_color.tint)  
    else:
        cell_color = RGBA_to_rgbTuple(cell.fill.start_color.rgb)
    if cell_color not in color_fill_dictionary:
        color_fill_dictionary[cell_color] = [cell_address]
    else:
        color_fill_dictionary[cell_color].append(cell_address) #LINKED LIST IMPLEMENTATION WOULD BE MORE EFFICIENT   


def extract_range_fill_data(file_address, sheet_name, selection, color_fill_dictionary):
    """
    Given a selection for a sheet, this updates dictionary mapping fill colors (in RGB tuples) to cell addresses
    """
    xlsx_filename=file_address

    try:
        with open(xlsx_filename, "rb") as f:
            #this needs to be done as well as wb.close() to make sure the workbook doesnt get broken by saving and other actions

            in_mem_file = io.BytesIO(f.read())
    except FileNotFoundError:
        error_title = "A small problem..."
        error_body = "You need to save your file first!"
        simple_error(error_title=error_title, error_body=error_body)
        return


    wb = openpyxl.load_workbook(in_mem_file, read_only=True)
    ws = wb[sheet_name]
    

    block_addresses = selection.split(",")
    for block_address in block_addresses:
        cells = ws[block_address]

        if type(cells) != tuple:
            #means that it is a single cell
            cell = cells

            handle_cell_for_extract_range_fill_data(cell=cell, color_fill_dictionary=color_fill_dictionary, wb=wb)
        
        elif type(cells) == tuple and cells == ():
            #to catch the error of when cells is a tuple but the empty tuple
            return

        elif type(cells) == tuple and type(cells[0]) != tuple:
            #means it is a single row but assumes cells is not length zero
            for cell in cells:
            #means that it is a single cell
                handle_cell_for_extract_range_fill_data(cell=cell, color_fill_dictionary=color_fill_dictionary, wb=wb)
            return
        else:
            #means it is multiple rows  
            print("checkpoint 3")
            for row in cells:
                for cell in row:
                    handle_cell_for_extract_range_fill_data(cell=cell, color_fill_dictionary=color_fill_dictionary, wb=wb)
                
            return

    wb.close() #needed to avoid breaking the file

