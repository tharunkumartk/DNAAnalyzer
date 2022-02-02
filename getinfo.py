import requests
from bs4 import BeautifulSoup
import randomheaders
import xlrd
import xlsxwriter

# global variables that are used to index ANIMAL vs FUNGI vs PLANT
ANIMAL = 0
FUNGI = 1
PLANT = 2


# a private function used in the process of returning the correct species of the DNA
# parameter 'dna' is the string of dna that is being checked
# parameter 'datatype' is the type of dna (ANIMAL vs FUNGI vs PLANT)
def input_data(dna, datatype):
    data = {}
    if datatype == ANIMAL:
        data = {
            'tabtype': 'animalTabPane',
            'historicalDB': '',
            'searchdb': 'COX1_SPECIES',
            'sequence': dna
        }
    if datatype == PLANT:
        data = {
            'tabtype': 'plantTabPane',
            'historicalDB': '',
            'searchdb': 'COX1_SPECIES',
            'sequence': dna
        }
    if datatype == FUNGI:
        data = {
            'tabtype': 'fungiTabPane',
            'historicalDB': '',
            'searchdb': 'COX1_SPECIES',
            'sequence': dna
        }
    link = 'http://v4.boldsystems.org/index.php/IDS_IdentificationRequest'

    # connect to website and send ur data
    return requests.post(link, data, headers=randomheaders.LoadHeader(), allow_redirects=True)

# a function that gets the closest match to the DNA.
# parameter 'source' is the source website of the database (returned by the input_data() function)
def get_data(source):
    soup = BeautifulSoup(source.text, 'html.parser')
    closest_match = soup.find_all('div', class_='ibox-title')[1].text
    return closest_match

# a function that combines the private functions 'input_data' and 'get_data'. this is the method that other classes can call. returns the actual closest match.
# parameter 'dna' is the string of dna that is being checked
def get_predicted_animal(dna):
    return get_data(input_data(dna))

# a function that reads an excel sheet at location 'file_location'. returns a list variable with the table.
# parameter 'file_location' is the location of the file
def read_table(file_location):
    wb = xlrd.open_workbook(file_location)
    sheet = wb.sheet_by_index(0)
    temptable = []
    for row in range(1, sheet.nrows):
        temptable2 = []
        for col in range(0, sheet.ncols):
            temptable2.append(sheet.cell_value(row, col))
        temptable.append(temptable2)
    return temptable

# runner code
inp = input("Enter the file address that you would like to read >>> ")
out = input("Enter the file name of the output >>> ")
workbook = xlsxwriter.Workbook("D:/Downloads D/OriginalWorkStuff/" + out + ".xlsx")
worksheet = workbook.add_worksheet()
worksheet.write('A1', 'INDEX')
worksheet.write('B1', 'TYPE')
worksheet.write('C1', 'DNA STRAND')
table = read_table(inp)
testinput = "D:/Downloads D/OriginalWorkStuff/DNAFile.xlsx"
for row in range(len(table)):
    dna_type = table[row][1]
    if dna_type == 'ANIMAL':
        dna_type = ANIMAL
    if dna_type == 'PLANT':
        dna_type = PLANT
    if dna_type == 'FUNGI':
        dna_type = FUNGI
    source = input_data(dna=table[row][2], datatype=dna_type)
    worksheet.write('A' + str(row + 2), table[row][0])
    worksheet.write('B' + str(row + 2), table[row][1])
    worksheet.write('C' + str(row + 2), get_data(source=source)[37:])
workbook.close()
