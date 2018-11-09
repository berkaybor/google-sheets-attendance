import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

credentials = ServiceAccountCredentials.from_json_keyfile_name('Compec-creds.json', scope)

gc = gspread.authorize(credentials)

device = int(input('Iphone: 1\nAndroid: 2\n'))
haftanumarasi = str(input('Kaçıncı hafta?\n'))
altk = int(input('Dev-Team: 1\nEğitimler: 2\nBBO: 3\n'))

# Open a worksheet from spreadsheet:
if altk == 1:
    wks = gc.open('DevTeam Yoklama').sheet1
elif altk == 2:
    sps = gc.open_by_key('1dsQmZ7mOrJO8GaYOKgIlDuGbh3JGVH1pqNhcPXosBd8')
    wegitim = int(input('Python: 1\nWeb: 2\nC: 3\nJava: 4\n'))
    if wegitim == 1:
        wks = sps.get_worksheet(1)
    elif wegitim == 2:
        wks = sps.get_worksheet(2)
    elif wegitim == 3:
        wks = sps.get_worksheet(3)
    elif wegitim == 4:
        wks = sps.get_worksheet(4)
    else:
        raise ValueError('Wrong value entered.')
elif altk == 3:
    wks = gc.open_by_key('1IIu4KhAw86wKGc0jPp3NnTzsjBQc-YJg8EIy8tDsW54').sheet1

# Get the position of the column:
cell_list = wks.range('A1:Z1')
for cell in cell_list:
    if cell.value == (str(haftanumarasi) + '. Hafta'):
        hafta_col = cell.col

# Pop up: select input csv
from tkinter.filedialog import askopenfilename
filename = askopenfilename()

# Select Iphone or Android and parse file:
df = pd.read_csv(filename)  
if device == 1:
    QRgelen = df[' Text'].to_string(index=False).split('\n')  
elif device == 2:
    QRgelen = df['text'].to_string(index=False).split('\n')

# Create a list form csv file:
gelenler = []
for gelen in QRgelen:
    gelenler.append(gelen.split(',')[2])

# Update yoklama sheet and set apart missing names: 
no_match = []
for gelen in gelenler:
    try:
        wks.update_cell(wks.find(gelen).row, hafta_col, '1')
        print('Added', gelen)
    except:
        no_match.append(gelen)


# Add missing names from the first worksheet:
hicbiryerde_olmayan_numaralar = []

def next_available_row(worksheet):
    return(len(worksheet.col_values(1)) + 1)

for num in no_match:
    wks_hepsi = sps.get_worksheet(0)
    
    # while wks.cell(next_row, 1).value:
        # next_row += 1

    try:
        row_to_write = next_available_row(wks)
        for i, element in enumerate(wks_hepsi.row_values(wks_hepsi.find(num).row)):
            wks.update_cell((row_to_write), i+1, element)

        wks.update_cell(row_to_write, hafta_col, '1')
    except:
        hicbiryerde_olmayan_numaralar.append(num)

   


print(no_match)
print('Bu numaralar bulunamadı')
if(len(hicbiryerde_olmayan_numaralar)>0):
    print(hicbiryerde_olmayan_numaralar)
    print('Bu numaralar bulunamadı ve eklenemedi')
    