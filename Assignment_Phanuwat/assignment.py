#Finnomena Assignment
#Phanuwat Ekjeen
#14 June 2021

from trello import TrelloClient
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from pprint import pprint

scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
cerds = ServiceAccountCredentials.from_json_keyfile_name("cerds.json", scope)
client = gspread.authorize(cerds)

#เปิดหน้า Google Sheet ที่ต้องการ
sheet = client.open('Not Qualify Sheet').worksheet('sheet1')

#กำหนด key และ token ของ Trello
client = TrelloClient(
    api_key='98af34584fd3f4a483eb33e4196f3b39',
    api_secret='5f815adb22efbe94ffc5da5db54b68e3235cd81c1ff3b41312e3143af15c36e5',
    token='0fcd3c8423746dcd2d0e4c7bae9682728d8c8ac40a58965487cde2701213bc24',
)

#define board
all_boards = client.list_boards()
sale_team_board = all_boards[-1]
call_center_board = all_boards[0]

#define list
leads_list = call_center_board.get_list('60c30f2961b5f8826c63e34b')
sale_list = call_center_board.get_list('60c30f2961b5f8826c63e34c')
notqualified_list = call_center_board.get_list('60c30f2961b5f8826c63e34d')
qualified_list = sale_team_board.get_list('60c30feefbf8914f938ffd59')

tmp = 1
col_name = 1
col_label = 2
col_due = 3
tmp_row = 2

print("Ready")

def sheet_update():
    global col_name, col_label, col_due, tmp_row
    
    card_nqf_list = notqualified_list.list_cards()
    
    #เลือก card ล่าสุดที่เพิ่มเข้ามาใน list
    tmp_json = {card.dateLastActivity: card for card in card_nqf_list}
    sorted_json = sorted(tmp_json, reverse=True)
    new_card_idx = sorted_json.pop(0)
    
    #copy ข้อมูล card
    tmp_name = tmp_json[new_card_idx].name
    tmp_due = tmp_json[new_card_idx].due_date
    tmp_due = tmp_due.astimezone().strftime("%d/%m/%Y %I:%M %p")
    tmp_label = tmp_json[new_card_idx].labels
    tmp_labels = ", ".join([label.name for label in tmp_label])

    #update data in Google sheet
    sheet.update_cell(tmp_row,col_name,str(tmp_name))
    sheet.update_cell(tmp_row,col_label,str(tmp_labels))
    sheet.update_cell(tmp_row,col_due,str(tmp_due))

    tmp_row = tmp_row + 1

def qualify_update():
    global col_name, col_label, col_due, tmp_row

    card_sale_list = sale_list.list_cards()

    #เลือก card ล่าสุดที่เพิ่มเข้ามาใน list
    tmp_json = {card.dateLastActivity: card for card in card_sale_list}
    sorted_json = sorted(tmp_json, reverse=True)
    new_card_idx = sorted_json.pop(0)
    
    #copy ข้อมูล card
    tmp_name = tmp_json[new_card_idx].name
    tmp_id = tmp_json[new_card_idx].id
    
    #เพิ่ม card ลงใน Qualified list
    qualified_list.add_card(tmp_name,source=tmp_id,keep_from_source="all")

while(True):
    #update จำนวน card ใน sale list และ not qualified list 
    if tmp == 1:
        num_sale = sale_list.cardsCnt()
        num_notqf = notqualified_list.cardsCnt()
        tmp = 0
    
    #เมื่อ card ใน sale list มีจำนวนเพิ่มขึ้น ให้ update card ที่ถูกเพิ่มเข้ามาใน qualified list
    if num_sale < sale_list.cardsCnt():
        qualify_update()
        tmp = 1
    
    elif num_sale > sale_list.cardsCnt():
        tmp = 1
    
    #เมื่อ card ใน not qualified list มีจำนวนเพิ่มขึ้น ให้ update ข้อมูลใน Google sheet
    elif num_notqf < notqualified_list.cardsCnt():
        sheet_update()
        tmp = 1
    
    elif num_notqf > notqualified_list.cardsCnt():
        tmp = 1
    
    else:
        pass
