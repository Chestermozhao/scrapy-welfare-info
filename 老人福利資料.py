import requests
from bs4 import BeautifulSoup as bs
import re
import xlwt
import xlrd
from xlutils.copy import copy

head = {"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/62.0.3202.89 Safari/537.36"}
url = "http://www.dosw2.gov.taipei/gmap/callNearData.aspx?do=5&address=台灣南港區重陽路187巷&type=04&type2=all&lat=25.0575383&lng=121.59855600000003&far=2000"
html = requests.get(url, headers = head)
html = html.text
#抓sno
sno_list=[]
html2 = html.split(",")
#print(html2[8])
for i in range(len(html2)):
    if i == 0:
        sno_list.append(html2[i][15:])
    else:
        if i % 8 == 0:
            sno_list.append(html2[i])
for i in range(len(sno_list)):
    sno_list[i] = sno_list[i][7:].replace('"','')
print(sno_list)
#namelist機構名稱
url = "http://www.dosw2.gov.taipei/gmap/callNearData.aspx?do=5&address=台灣南港區重陽路187巷&type=04&type2=all&lat=25.0575383&lng=121.59855600000003&far=2000"
html = requests.get(url, headers = head)
html = html.text
name_list=[]
html2 = html.split(",")
for i in range(len(html2)):
    if i % 8 == 1:
        name_list.append(html2[i])
for i in range(len(name_list)):
    name_list[i] = name_list[i][8:].replace('"','')
##print(name_list)
#sno加進詳細資料API
#抓有電話的機構名稱跟電話
content_list = []
tele_list = []
address_list = []
for i in sno_list:
    url_detail = "http://www.dosw2.gov.taipei/gmap/callOneInfoData.aspx?sno=" + i + "&l1_code=04"
    content = requests.get(url_detail , headers = head)
    content_list.append(content.text)
    bscontent = bs(content.text,"html.parser")
    for j in range(len(content.text)):
        if content.text[j] == "電" and content.text[j+1] == "話":
            tele_list.append(bscontent.find("b"))
            tele_list.append(content.text[j+4:j+17])
    #抓機構地址
        if content.text[j] == "地" and content.text[j+1] == "址":
            address_list.append(bscontent.find("b"))
            address_list.append(content.text[j+4:j+28])
for i in range(len(address_list)):
    if i % 2 == 1:
        address_list[i] = address_list[i].replace("<","")
        address_list[i] = address_list[i].replace("b","")
        address_list[i] = address_list[i].replace("r","")
        address_list[i] = address_list[i].replace(" ","")
        address_list[i] = address_list[i].replace("/","")
        address_list[i] = address_list[i].replace(">","")
    else:
        address_list[i] = str(address_list[i])
        address_list[i] = address_list[i].replace("<b>機構名稱: ","")
        address_list[i] = address_list[i].replace("</b>","")
print(address_list)
#判斷整數或非整數
def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        pass

    try:
        import unicodedata
        unicodedata.numeric(s)
        return True
    except (TypeError, ValueError):
        pass

    return False
#清理電話數列的內容
for i in range(len(tele_list)):
    for j in range(len(tele_list[i])):
        if i % 2 == 1:
            if is_number(tele_list[i][j]) == False or tele_list[i][j] == "三":
                tele_list[i] = tele_list[i].replace(tele_list[i][j]," ")
for i in range(len(tele_list)):
    if i % 2 == 1:
        tele_list[i] = tele_list[i].replace(" ","")
    else:
        tele_list[i] = str(tele_list[i])
        tele_list[i] = tele_list[i].replace("<b>機構名稱: ","")
        tele_list[i] = tele_list[i].replace("</b>","")
print(tele_list)

###讀兩公里內所有機構名稱
##name = xlrd.open_workbook('機構名稱.xls')
##table = name.sheet_by_name('infant_service')
##nrows = table.nrows
##name_list = []
##for i in range(nrows):
##    name_list.append(table.row_values(i))
##print(name_list)

#機構名稱寫入excel
wb = xlwt.Workbook()
ws = wb.add_sheet('old_service', cell_overwrite_ok=True)
ws.write(0, 0, "機構名稱")
for i in range(len(name_list)):
    ws.write(i+1 , 0 , name_list[i])
ws.write(0 ,1, "連絡電話")
ws.write(0 ,2, "機構地址")
#比對機構名稱，若有電話者則寫入，無者鍵入N
for i in range(len(name_list)):
    for j in range(len(tele_list)):
        try:
            if name_list[i] == tele_list[j]:
                if tele_list[j+1] == "":
                    ws.write(i+1 , 1, "N")
                    ws.write(i+1 , 2, "N")
                else:
                    ws.write(i+1 , 1, tele_list[j+1])
                    ws.write(i+1 , 2, address_list[j+1])
    ##            print(address_list[j])
            elif name_list[i] not in tele_list:
                ws.write(i+1 , 1, "N")
                ws.write(i+1 , 2, "N")
        except:
            continue
            
wb.save("銀髮族服務機構名稱.xls")

