from tkinter.constants import TRUE
import PySimpleGUI as sg


def file_address_gui():
        event, values = sg.Window('Select File',
                [[sg.Text('Filename')], [sg.Input(),sg.FileBrowse()],
                [sg.OK(), sg.Cancel()]]).read(close=TRUE)
        return values

def simple_error(error_title, error_body):
    event = sg.Window(error_title,
                [[sg.Text(error_body), sg.OK()]]).read(close=TRUE)
