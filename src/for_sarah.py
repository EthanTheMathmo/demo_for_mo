#for PDF reading
from pyxpdf import Document, Page, Config
from pyxpdf.xpdf import TextControl
from variables import MACROS_SHFX_location
import PySimpleGUI as sg
import re
import webbrowser
import index_helpers
import test_gui
import xlwings as xw
import os
import platform

#CHECK_FOR_MAC left if there is something I think specifically needs to be checked for mac 
#which might be hard to spot. 


def find_record():
    target = xw.apps.active.books.active.selection.value

    openable_file_types = [".pdf", ".txt", ".html", ".docx", ".img", ".jpeg", ".xlsm", ".csv"] #can add to this as needed

    if type(target) == list:
        error_title = "Selection error"
        error_body = "You need to select a single cell input"
        test_gui.simple_error(error_title=error_title, error_body=error_body)
        return
    else:
        try:
            target = str(target)
        except TypeError:
            error_title = "Selection error"
            error_body = "Your input was invalid. If the error is ours, contact support and we'll make it work"
            test_gui.simple_error(error_title=error_title, error_body=error_body)
            return
    

    files_opened = []

    target_directory = test_gui.folder_address_gui()

    if target_directory == "" or target_directory == None:
        #nothing entered. 
        # CHECK_FOR_MAC because 
        # titlebar is enabled on MAC, I don't know what is returned to target_directory
        #if that cross in the top right hand corner is used
        return

    for dirpath, dirnames, filenames in os.walk(target_directory):
        for filename in [f for f in filenames if f.endswith(".pdf")]:
            if filename in [target + end for end in openable_file_types]:
                webbrowser.open(os.path.join(dirpath, filename))
                files_opened.append(os.path.join(dirpath, filename))
            break

    if files_opened == []:
        error_title = "No files found!"

        error_body = "No files were found - would you like us to do a similarity search?"

        test_gui.simple_error(error_title=error_title, error_body=error_body)


        return

    else:

        with open("opened.txt", "w") as text_file:
                for index, file in enumerate(files_opened):
                    index+=1 #as zero indexed
                    text_file.write(f"File number {index} \n")
                    text_file.write("file name: "+os.path.splitext(os.path.basename(file))[0] + "\n") #.basename is a function
                            #which from /.../file.txt returns file.txt or .../../file.docx returns file.docx
                            #and .splitext takes something with a .txt of .docx or .pdf (etc) ending and returns the bit before that and the bit after in an array
                    text_file.write("file location: " +file + "\n\n")
        
        webbrowser.open("opened.txt")


def extract_emails():
    #given pdfs, exctracts the emails from it
    import os
    #GUI
    layout = [[sg.Text("Files"), sg.Input(visible=True), sg.FilesBrowse()],
            [sg.Button('Go'), sg.Button('Exit')]  ]
    if platform.system() == "Windows":
        event, values = sg.Window('Window Title', layout, no_titlebar=True, keep_on_top=True).read(close=True)
    else:
        event, values = sg.Window('Window Title', layout, keep_on_top=True).read(close=True)
    files = values[0].split(";")

    files_to_emails = {}

    for index, file in enumerate(files):
        with open(file, 'rb') as fp:
            doc = Document(fp)

        all_text = doc.text()
        emails = re.findall(r"""(?:[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*|"(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21\x23-\x5b\x5d-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])*")@(?:(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?|\[(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?|[a-z0-9-]*[a-z0-9]:(?:[\x01-\x08\x0b\x0c\x0e-\x1f\x21-\x5a\x53-\x7f]|\\[\x01-\x09\x0b\x0c\x0e-\x7f])+)\])""",
            all_text)
        files_to_emails[file] = "\t"+"\n\t".join(set(emails)) #so xlwings returns as a column
        sg.OneLineProgressMeter('Creating worksheet', index+1, len(files), 'single')
    with open("emails.txt", "w") as text_file:
        for file in files:
            text_file.write("file name: "+os.path.splitext(os.path.basename(file))[0] + "\n") #.basename is a function
                    #which from /.../file.txt returns file.txt or .../../file.docx returns file.docx
                    #and .splitext takes something with a .txt of .docx or .pdf (etc) ending and returns the bit before that and the bit after in an array
            text_file.write("file location: " +file + "\n")
            text_file.write(files_to_emails[file])
            text_file.write("\n\n")
            
        
    webbrowser.open("emails.txt")

def attach_data():
    macro_address = MACROS_SHFX_location[platform.system()]
    create_comment = xw.Book(os.path.join(macro_address, "MACROS_SHFX.xlam")).macro(r'AddCommentToSpecificCell')
    
    sht = xw.apps.active.books.active.sheets.active

    active_book = xw.apps.active.books.active
    
    range1, range2 = active_book.selection.address.split(",")
    
    active_sheet = active_book.sheets.active
    
    if "," in range1 or "," in range2:
        #TO-DO. ERROR MESSAGE AS WE ONLY WANT THE USER TO ENTER RANGE AS SINGLE BLOCK
        return
    else:
        pass
    
    #names are the names we'll look for in files
    #write_address is the cell address we'll first write to, and then move down
    if ":" in range1 and ":" in range2:
        #TO-DO ERROR MESSAGE. ONE should be a row, one should be a single cell
        return
    elif ":" in range1 and ":" not in range2:
        names = sht.range(range1).value
        write_address = range2
    else:
        names = sht.range(range2).value
        write_address = range1     
    
    openable_file_types = [".pdf", ".txt", ".html", ".docx", ".img", ".jpeg", ".xlsm", ".csv"] #can add to this as needed
   

    names_with_no_file_found = []
    
    names_with_multiple_files_found = {}

    target_directory = test_gui.folder_address_gui()

    if target_directory == "" or target_directory == None:
        #nothing entered. 
        # CHECK_FOR_MAC because 
        # titlebar is enabled on MAC, I don't know what is returned to target_directory
        #if that cross in the top right hand corner is used
        return
    
    names_to_files = {}

    for dirpath, dirnames, filenames in os.walk(target_directory):
        for filename in filenames:
            name = os.path.splitext(filename)[0]
            if name in names_to_files:
                names_to_files[name] += "\n" + os.path.join(dirpath,filename)
                names_with_multiple_files_found[name] = True
            else:
                names_to_files[name] = os.path.join(dirpath, filename)
                names_with_multiple_files_found[name] = False
    
    index=1
    
    for name in names:
        if name in names_to_files:
            create_comment(names_to_files[name], write_address)
            #sht.range(write_address).api.NoteText(names_to_files[name])
        else:
            names_with_no_file_found.append(name)
            
        write_address = index_helpers.next_down(write_address)
        sg.OneLineProgressMeter('Adding file information', index, len(names))
        index+=1
    

# def attach_data():
#     sht = xw.apps.active.books.active.sheets.active
#     range1, range2 = xw.apps.active.books.active.selection.address.split(",")

#     layout = [[sg.Text("Folder"), sg.Input(visible=True), sg.FolderBrowse()],
#             [sg.Button('Go'), sg.Button('Exit')]  ]

#     event, values = sg.Window('Window Title', layout, no_titlebar=True, keep_on_top=True).read(close=True)
    
#     file_folder = values[0]

#     range2_as_list = index_helpers.block_to_list(range2).split(",")

#     index=1
#     for row_name, target_address in zip(sht.range(range1).value, range2_as_list):
#         target_file = os.path.join(file_folder, row_name + ".pdf")
#         if os.path.isfile(target_file): #only adds the link if the file exists
#             sht[target_address].api.NoteText(target_file)
#         sg.OneLineProgressMeter('Creating worksheet', index, len(range2_as_list), 'single')
#         index+=1

#     return


if __name__ == '__main__':
    # Expects the Excel file next to this source file, adjust accordingly.

    attach_data()

    # macro_address = MACROS_SHFX_location[platform.system()] 
    # add_link_macro = xw.Book(macro_address+"\\" + "MACROS_SHFX.xlam").macro(r'add_link') 
    # add_link_macro()
