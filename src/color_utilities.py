"""The color utility below is to convert ARGB to RGB"""

def RGBA_to_rgbTuple(rgb, background=(255,255,255)):
    """
    takes an rgba value and converts to to rgb
    see:https://stackoverflow.com/questions/2049230/convert-rgba-color-to-rgb
    """
    source = (int(rgb[0:2],16),int(rgb[2:4],16), int(rgb[4:6], 16), int(rgb[6:8],16))
    SourceA = source[0]/255
    SourceR = source[1]/255
    SourceG = source[2]/255
    SourceB = source[3]/255
    TargetR = (1 - SourceA)*background[0]/255 + (SourceA * SourceR)
    TargetG = (1 - SourceA)*background[1]/255 + (SourceA * SourceG)
    TargetB = (1 - SourceA)*background[2]/255 + (SourceA * SourceB)
    x=(min(255, int(255*TargetR)), min(255, int(255*TargetG)), min(255, int(255*TargetB)))
    return (min(255, 255*int(TargetR)), min(255, 255*int(TargetG)), min(255, 255*int(TargetB)))




"""This contains COLOR utilities
which convert themes and shades into RGB
the source is below, but has been slightly changed to

(1) in get_theme_colors, the get_children() method was deprecated, and I replaced it with list,
also see the answer by andrew pate here https://stackoverflow.com/questions/55883322/why-is-if-statement-not-working-in-elementtree-parsing

(2) rgb to hex was changed so that we get rgb tuples which are not normalised, as this is the input
format for xlwings color


"""

from colorsys import rgb_to_hls, hls_to_rgb
#https://bitbucket.org/openpyxl/openpyxl/issues/987/add-utility-functions-for-colors-to-help
 
RGBMAX = 0xff  # Corresponds to 255
HLSMAX = 240  # MS excel's tint function expects that HLS is base 240. see:
# https://social.msdn.microsoft.com/Forums/en-US/e9d8c136-6d62-4098-9b1b-dac786149f43/excel-color-tint-algorithm-incorrect?forum=os_binaryfile#d3c2ac95-52e0-476b-86f1-e2a697f24969
 
def rgb_to_ms_hls(red, green=None, blue=None):
    """Converts rgb values in range (0,1) or a hex string of the form '[#aa]rrggbb' to HLSMAX based HLS, (alpha values are ignored)"""
    if green is None:
        if isinstance(red, str):
            if len(red) > 6:
                red = red[-6:]  # Ignore preceding '#' and alpha values
            blue = int(red[4:], 16) / RGBMAX
            green = int(red[2:4], 16) / RGBMAX
            red = int(red[0:2], 16) / RGBMAX
        else:
            red, green, blue = red
    h, l, s = rgb_to_hls(red, green, blue)
    return (int(round(h * HLSMAX)), int(round(l * HLSMAX)), int(round(s * HLSMAX)))
 
def ms_hls_to_rgb(hue, lightness=None, saturation=None):
    """Converts HLSMAX based HLS values to rgb values in the range (0,1)"""
    if lightness is None:
        hue, lightness, saturation = hue
    return hls_to_rgb(hue / HLSMAX, lightness / HLSMAX, saturation / HLSMAX)
 
def rgb_normalised_to_rgb_tuple(red, green=None, blue=None):
    """Converts (0,1) based RGB values to a tuple of non normalised rgbs'"""
    if green is None:
        red, green, blue = red
    return (int(round(red * RGBMAX)), int(round(green * RGBMAX)), int(round(blue * RGBMAX)))
 
def get_theme_colors(wb):
    """Gets theme colors from the workbook"""
    # see: https://groups.google.com/forum/#!topic/openpyxl-users/I0k3TfqNLrc
    from openpyxl.xml.functions import QName, fromstring
    xlmns = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    root = fromstring(wb.loaded_theme)
    themeEl = root.find(QName(xlmns, 'themeElements').text)
    colorSchemes = themeEl.findall(QName(xlmns, 'clrScheme').text)
    firstColorScheme = colorSchemes[0]
 
    colors = []
 
    for c in ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6']:
        accent = firstColorScheme.find(QName(xlmns, c).text)
 
        if 'window' in list(accent)[0].attrib['val']:
            colors.append(list(accent)[0].attrib['lastClr'])
        else:
            colors.append(list(accent)[0].attrib['val'])
 
    return colors
 
def tint_luminance(tint, lum):
    """Tints a HLSMAX based luminance"""
    # See: http://ciintelligence.blogspot.co.uk/2012/02/converting-excel-theme-color-and-tint.html
    if tint < 0:
        return int(round(lum * (1.0 + tint)))
    else:
        return int(round(lum * (1.0 - tint) + (HLSMAX - HLSMAX * (1.0 - tint))))
    
def theme_and_tint_to_rgb(wb, theme, tint):
    """Given a workbook, a theme number and a tint return an rgb tuple"""
    rgb = get_theme_colors(wb)[theme]
    h, l, s = rgb_to_ms_hls(rgb)
    return rgb_normalised_to_rgb_tuple(ms_hls_to_rgb(h, tint_luminance(tint, l), s))