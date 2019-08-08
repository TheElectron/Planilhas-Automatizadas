# Importing Packages
import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials
# Configuration Of Spreadsheets
scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',
         "https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
# Access Credentials
creds = ServiceAccountCredentials.from_json_keyfile_name("Credenciais.json", scope)
client = gspread.authorize(creds)
# Open The Spreadhseet.
sheet = client.open("Secratario-PET").sheet1
# Add. Members Name With Intercalation of New Members.
List_Novatos = [['Bianca','Gabriel','Clórys'],
                    ['Gabriel','Clórys','Guilherme A.'],
                    ['Clórys','Guilherme A.','Guilherme M.'],
                    ['Guilherme A.','Guilherme M.','Henrique'],
                    ['Guilherme M.','Henrique','Gustavo'],
                    ['Henrique','Gustavo','Kelly'],
                    ['Gustavo','Kelly','Johnnatan'],
                    ['Kelly','Johnnatan','Ligia'],
                    ['Johnnatan','Ligia','Milene'],
                    ['Ligia','Milene','Matheus'],
                    ['Milene','Matheus','Nilson'],
                    ['Matheus','Nilson','Vinicius'],
                    ['Nilson','Vinicius','Bianca'],
                    ['Vinicius','Bianca','Gabriel']]
# Add. Members Name.
List =[['Bianca','Clórys','Gabriel'],
        ['Clórys','Gabriel','Guilherme A.'],
        ['Gabriel','Guilherme A.','Guilherme M.'],
        ['Guilherme A.','Guilherme M.','Gustavo'],
        ['Guilherme M.','Gustavo','Henrique'],
        ['Gustavo','Henrique','Johnnatan'],
        ['Henrique','Johnnatan','Kelly'],
        ['Johnnatan','Kelly','Ligia'],
        ['Kelly','Ligia','Matheus'],
        ['Ligia','Matheus','Milene'],
        ['Matheus','Milene','Nilson'],
        ['Milene','Nilson','Vinicius'],
        ['Nilson','Vinicius','Bianca'],
        ['Vinicius','Bianca','Clórys']]
# Add. Columns Titles.
Titulos = ['Nº','Secretário','Revisor', 'Revisor']
Line = 7
Column = 3
for elem in Titulos:
    sheet.update_cell(Line, Column, elem)
    Column += 1
print('Cabeçalho 1 - OK')
N_Ata = 167
Line += 1
Column = 3
for elem in List_Novatos:
    sheet.update_cell(Line, Column, N_Ata)
    Column += 1
    for name in elem:
        sheet.update_cell(Line, Column, name)
        Column += 1
    Line += 1
    Column = 3
    N_Ata += 1
for elem in List:
    sheet.update_cell(Line, Column, N_Ata)
    Column += 1
    for name in elem:
        sheet.update_cell(Line, Column, name)
        Column += 1
    Line += 1
    Column = 3
    N_Ata += 1
print('Listagem - OK')
time.sleep(2)
print(Line)
print(N_Ata)
print('Finalização - OK')