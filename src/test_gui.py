from tkinter.constants import TRUE
import PySimpleGUI as sg
import platform

def folder_address_gui():
        gui_text = "Folder name"
        if platform.system() == "Windows": 
                event, values = sg.Window("",[[sg.Text(gui_text)], [sg.Input(),sg.FolderBrowse()],
                        [sg.OK(), sg.Cancel()]], no_titlebar=True, keep_on_top=True, grab_anywhere=True).read(close=True)
                return values[0]
        else:
                event, values = sg.Window("",[[sg.Text('Filename')], [sg.Input(),sg.FolderBrowse()],
                        [sg.OK(), sg.Cancel()]], keep_on_top=True, grab_anywhere=True).read(close=True)
                return values[0]

def file_address_gui():
        
        if platform.system() == "Windows": 
                event, values = sg.Window("",[[sg.Text('Filename')], [sg.Input(),sg.FileBrowse()],
                        [sg.OK(), sg.Cancel()]], no_titlebar=True, keep_on_top=True, grab_anywhere=True).read(close=True)
                return values[0]
        else:
                event, values = sg.Window("",[[sg.Text('Filename')], [sg.Input(),sg.FileBrowse()],
                        [sg.OK(), sg.Cancel()]], keep_on_top=True, grab_anywhere=True).read(close=True)
                return values[0]


        

def simple_error(error_title, error_body):
    event = sg.Window(error_title,
                [[sg.Text(error_body), sg.OK()]]).read(close=TRUE)
