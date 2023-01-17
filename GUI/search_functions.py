# Author: yzjnxsantiago
# Start Date: Tuesday, January 11, 2023

alphabet = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG']

# Search through an excel sheet and return the cell locations of keywords
def search_sheet(sheet, keywords, range_x, range_y):

    len_keywords = len(keywords) # Length of keywords array
    keywords_cell = [] # cell of which the keywords are on the excel spreadsheet

    # Go through each row and column of the excel sheet within the range
    for i in range(0, range_x):
        for j in range(0 , range_y):
            for k in range(0, len_keywords): # Check for keywords 
                if sheet[alphabet[j] + str(i)] == keywords[k]:
                    keywords_cell.append(alphabet[j] + str(i))
                if len(keywords_cell) == len_keywords:
                    return keywords_cell
    
    return keywords_cell


def inc_alpha(char):
    x = ord(char)
    x += 1
    char = chr(x)
    return char

def dec_alpha(char):

    if char == 'A' or char == 'a':
        return char
            
    x = ord(char)
    x -= 1
    char = chr(x)
    return char


