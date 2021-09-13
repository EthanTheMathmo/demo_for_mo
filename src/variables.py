"""
To make the compatibility for MAC and WINDOWS as smooth as possible,
and for automatic installtion

"""
import os
import datetime
import platform

undo_control = {"UNDO":False, "TIME":datetime.datetime.min} #execute undo action only if UNDO is true and TIME is a suitably small difference

#Note: Darwin is the name of the Mac operating system when we use the python module platform and run platform.module()
if platform.system() == "Windows":
    MACROS_SHFX_location = {"Windows": os.getenv("APPDATA") + r"\microsoft\excel\xlstart", "Darwin":""}
else:
    MACROS_SHFX_location = {"Windows": r"\microsoft\excel\xlstart", "Darwin":""}


"""
FOR tracing back and undoing actions
"""
undo_dict = {"highlight_dates_to_expiry":{}, "color_border":{}, "font_color":{}}


"""
Mapping Interior.Color to RGB values

TO-DO: double check all of these...
"""
InteriorColorToRGBdict = {0:None,
    1:(0,0,0), 2:(255,255,255), 3:(255,0,0), 4:(0,255,0), 5:(0,0,255), 6:(255,255,0),
    7:(255,0,255), 8:(0,255,255), 9:(128,0,0), 10:(0,128,0), 11:(0,0,128), 12:(128,128,0),
    13:(128,0,128), 14:(0,128,128), 15:(192,192,192), 16:(128,128,128), 17:(153,153,255),
    18:(153,51,102), 19:(255,255,204), 20:(204,255,255), 21:(102,0,102), 22:(255,128,128),
    23:(0,102,204), 24:(204,204,255), 25:(0,0,128), 26:(255,0,255), 27:(255,255,0),
    28:(0,255,255), 29:(128,0,128), 30:(128,0,0), 31:(0,128,128), 32:(0,0,255),
    33: (0,204,255), 34:(204,255,255), 35:(204,255,204), 36: (255,255,153),
    37:(153,204,255), 38:(255,153,204), 39:(204,153,255), 40:(255,204,153),
    41:(51, 102, 255), 42:(51,204,204), 43:(153,204,0), 44:(255,204,0), 45:(255,153,0),
    46:(255,102,0), 47:(102,102,153), 48:(150,150,150), 49:(0,51,102), 50:(51,153,102),
    51:(0,51,0), 52:(51,51,0), 53:(153,51,0), 54:(153,51,102), 55:(51,51,153), 56:(51,51,51)
}


#see https://openpyxl.readthedocs.io/en/stable/_modules/openpyxl/styles/colors.html #I dont think this is needed anymore
InteriorColorToRGBdictOpenpyxl={0: '00000000',
 1: '00FFFFFF',
 2: '00FF0000',
 3: '0000FF00',
 4: '000000FF',
 5: '00FFFF00',
 6: '00FF00FF',
 7: '0000FFFF',
 8: '00000000',
 9: '00FFFFFF',
 10: '00FF0000',
 11: '0000FF00',
 12: '000000FF',
 13: '00FFFF00',
 14: '00FF00FF',
 15: '0000FFFF',
 16: '00800000',
 17: '00008000',
 18: '00000080',
 19: '00808000',
 20: '00800080',
 21: '00008080',
 22: '00C0C0C0',
 23: '00808080',
 24: '009999FF',
 25: '00993366',
 26: '00FFFFCC',
 27: '00CCFFFF',
 28: '00660066',
 29: '00FF8080',
 30: '000066CC',
 31: '00CCCCFF',
 32: '00000080',
 33: '00FF00FF',
 34: '00FFFF00',
 35: '0000FFFF',
 36: '00800080',
 37: '00800000',
 38: '00008080',
 39: '000000FF',
 40: '0000CCFF',
 41: '00CCFFFF',
 42: '00CCFFCC',
 43: '00FFFF99',
 44: '0099CCFF',
 45: '00FF99CC',
 46: '00CC99FF',
 47: '00FFCC99',
 48: '003366FF',
 49: '0033CCCC',
 50: '0099CC00',
 51: '00FFCC00',
 52: '00FF9900',
 53: '00FF6600',
 54: '00666699',
 55: '00969696',
 56: '00003366',
 57: '00339966',
 58: '00003300',
 59: '00333300',
 60: '00993300',
 61: '00993366',
 62: '00333399',
 63: '00333333'}
