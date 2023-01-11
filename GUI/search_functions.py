# Author: yzjnxsantiago
# Start Date: Tuesday, January 11, 2023

alphabet = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG']

# Search through an excel sheet and return the cell locations of keywords
def search_sheet(sheet, keywords, range_x, range_y):

    len_keywords = len(keywords)
    keywords_cell = []

    for i in range(0, range_x):
        for j in range(0 , range_y):
            for k in range(0, len_keywords):
                if sheet[alphabet[j] + str(i)] == keywords[k]:
                    keywords_cell.append(alphabet[j] + str(i))
    
    return keywords_cell

