#Finnomena Assignment
#Phanuwat Ekjeen
#14 June 2021

from time import time
from trello import TrelloClient
from trello import label
from trello.label import Label
import datetime

import gspread
from oauth2client.service_account import ServiceAccountCredentials
from pprint import pprint
scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
cerds = ServiceAccountCredentials.from_json_keyfile_name("cerds.json", scope)
client = gspread.authorize(cerds)
sheet = client.open('Not Qualify Sheet').worksheet('sheet1')

client = TrelloClient(
    api_key='98af34584fd3f4a483eb33e4196f3b39',
    api_secret='5f815adb22efbe94ffc5da5db54b68e3235cd81c1ff3b41312e3143af15c36e5',
    token='0fcd3c8423746dcd2d0e4c7bae9682728d8c8ac40a58965487cde2701213bc24',
)

all_boards = client.list_boards()
last_board = all_boards[-1]
first_board = all_boards[0]

first_board.list_lists()
last_board.list_lists()

leads_list = first_board.get_list('60c30f2961b5f8826c63e34b')
sale_list = first_board.get_list('60c30f2961b5f8826c63e34c')
notqualified_list = first_board.get_list('60c30f2961b5f8826c63e34d')

qualified_list = last_board.get_list('60c30feefbf8914f938ffd59')

tmp = 1
col_name = 1
col_label = 2
col_due = 3
tmp_row = 2

print("Ready")

def sheet_update():
    global col_name, col_label, col_due, tmp_row, tmp_id 
    
    card_nqf_list = notqualified_list.list_cards()
    
    tmp_json = {card.dateLastActivity: card for card in card_nqf_list}
    sorted_json = sorted(tmp_json, reverse=True)
    new_card_idx = sorted_json.pop(0)
    tmp_name = tmp_json[new_card_idx].name
    tmp_id = tmp_json[new_card_idx].id
    tmp_due = tmp_json[new_card_idx].due_date
    tmp_due = tmp_due.astimezone().strftime("%d/%m/%Y %I:%M %p")
    tmp_label = tmp_json[new_card_idx].labels
    tmp_labels = ", ".join([label.name for label in tmp_label])

    sheet.update_cell(tmp_row,col_name,str(tmp_name))
    sheet.update_cell(tmp_row,col_label,str(tmp_labels))
    sheet.update_cell(tmp_row,col_due,str(tmp_due))

    tmp_row = tmp_row + 1

def qualify_update():
    global col_name, col_label, col_due, tmp_row, tmp_id

    card_sale_list = sale_list.list_cards()

    tmp_json = {card.dateLastActivity: card for card in card_sale_list}
    sorted_json = sorted(tmp_json, reverse=True)
    new_card_idx = sorted_json.pop(0)
    tmp_name = tmp_json[new_card_idx].name
    tmp_id = tmp_json[new_card_idx].id
    tmp_due = tmp_json[new_card_idx].due_date
    tmp_due = tmp_due.astimezone().strftime("%d/%m/%Y %I:%M %p")
    
    qualified_list.add_card(tmp_name,source=tmp_id,keep_from_source="all")

while(True):
    if tmp == 1:
        num_sale = sale_list.cardsCnt()
        num_notqf = notqualified_list.cardsCnt()
        tmp = 0
    
    if num_sale < sale_list.cardsCnt():
        qualify_update()
        tmp = 1
    
    elif num_sale > sale_list.cardsCnt():
        tmp = 1
    
    elif num_notqf < notqualified_list.cardsCnt():
        sheet_update()
        tmp = 1
    
    elif num_notqf > notqualified_list.cardsCnt():
        tmp = 1
    
    else:
        pass
