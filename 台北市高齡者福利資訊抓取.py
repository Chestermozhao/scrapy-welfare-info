import requests
from bs4 import BeautifulSoup as bs
import re
import xlwt
import xlrd
from xlutils.copy import copy

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

def welfareinfoTP(address,distance):
    head = {"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/62.0.3202.89 Safari/537.36"}
    url = "http://www.dosw2.taipei.gov.tw/gmap/callNearData.aspx?do=5&address="+address+"&type=0&type2=--%20ALL%20--&lat=25.020307&lng=121.52848670000003&far="+distance
    html = requests.get(url, headers = head)
    html = html.text
    html_split = html.split(",")
    sno_list = []
    name_list=[]
    #抓sno
    for i in range(len(html_split)):
        if len(re.findall(r"sno...(\d+)" , html_split[i])) != 0:
            sno_list.append(re.findall(r"sno...(\d+)" , html_split[i])[0])
    #namelist機構名稱
        Name = re.findall(r"pName...(\w+)" , html_split[i])
        if len(Name) != 0:
            name_list.append(Name[0])
    #sno取得詳細資料
    #抓有電話的機構名稱跟電話
    content_list = []
    con_name = []
    for i in range(120):
        url_detail = "http://www.dosw2.gov.taipei/gmap/callOneInfoData.aspx?sno=" + sno_list[i] + "&l1_code=04"
        content = requests.get(url_detail , headers = head)
        content_list.append(content.text)
        tele_bs = bs(content.text,"html.parser")
        if tele_bs.find_all("b") is not None:
            con_name.append(tele_bs.find_all("b")[0].get_text()[6:])
        else:
            con_name.append("None")
    tele_list = []
    address_list = []
    for i in range(len(content_list)):
        #抓電話
        tele_number = re.findall(r"電話..\b(\d+.+)傳真\b" , content_list[i])
        if len(tele_number) == 0:
            tele_number = re.findall(r"電話..\b(\d+.+)地址\b" , content_list[i])
        #比對資料
        tele_list.append(con_name[i])
        if len(tele_number) != 0:
            tele_list.append(tele_number[0][:12])
        else:
            tele_list.append("無電話資料")
        tele_list[i] = re.sub(r"<|b|r| |/","", tele_list[i]).strip()
        #抓地址
        address_data = re.findall(r"地址..\b(.+\w+.+)<br /><div>\b", content_list[i])
        address_list.append(address_data[0])

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
                    if tele_list[j+1] == "無電話資料":
                        ws.write(i+1 , 1, "無電話資料")
                        ws.write(i+1 , 2, address_list[i])
                    else:
                        ws.write(i+1 , 1, tele_list[j+1])
                        ws.write(i+1 , 2, address_list[i])
                elif name_list[i] not in tele_list:
                    ws.write(i+1 , 1, "目前沒有詳細資訊")
                    ws.write(i+1 , 2, "目前沒有詳細資訊")
            except:
                continue
                
    wb.save("銀髮族服務機構名稱.xls")

welfareinfoTP("台灣大安區羅斯福路四段","2000")

