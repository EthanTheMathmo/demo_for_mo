import re
import string

def convert_to_tuple(cell_address):
    """
    e.g. converts A1 to (1,1), converts AC1 to (1*26 + 3,1) DE1 to (4*26+5,1), CZE to (3*26^2 + 26*26^1 + 5*26^0)
    bit of a headache, but with some messing about it makes sense!
    """
    letters = re.search("[A-Z]{1,}",cell_address).group()
    numbers = re.search("[0-9]{1,}",cell_address).group()
    return (sum([(ord(x[1])-64)*26**x[0] for x in enumerate(reversed(letters))]), int(numbers))

digs = "0" + string.ascii_letters[-26:]
#from SO
def int2base(x, base):
    """
    function from SO. Benefit that is definitely works. Was a bit lazy to take it but oh well
    """
    if x < 0:
        sign = -1
    elif x == 0:
        return digs[0]
    else:
        sign = 1

    x *= sign
    digits = []

    while x:
        digits.append(digs[int(x % base)])
        x = int(x / base)

    if sign < 0:
        digits.append('-')

    digits.reverse()

    return ''.join(digits)

def h(chunk):
    """
    helper function for convert from tuple
    """
    chunk=chunk.group()
    if chunk == "":
        return ""
    elif chunk[0] == "A":
        return "Z"*len(chunk[:-1])
    else:
        char = chunk[0]
        return chr(ord(char)-1)+"Z"

def convert_from_tuple(tuple_address):
    """
    Takes a cell, say, (1,1) and returns it in excel notation A1
    """
    return re.sub("A{1,}0||[B-Z]0",h,int2base(tuple_address[0],26))+str(tuple_address[1])


def block_to_list(block_address):
    """
    E.g., converts $A$1:$B$2 to $A$1,$A$2,$B$1,$B$2
    """
    first_and_last = block_address.split(":")

    if len(first_and_last) == 1:
        #i.e. is just a single cell passed in.
        #we cover this case to deal with edge cases where a user's 
        #block selection is just a single cell
        return block_address

    first = first_and_last[0]
    last = first_and_last[1]
    first_tuple = convert_to_tuple(first)
    last_tuple = convert_to_tuple(last)
    return ",".join([",".join([convert_from_tuple((i,j)) for j in range(first_tuple[1], last_tuple[1]+1)]) for i in range(first_tuple[0], last_tuple[0]+1)])


