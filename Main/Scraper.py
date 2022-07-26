import requests
from bs4 import BeautifulSoup
from datetime import datetime
from openpyxl import *   
from trial import *


pagedata=[]
flag=0
def Initialize():
    global pagedata
    global flag
    
    wb2 = load_workbook('Data.xlsx')
    sheet = wb2.active
    for i in range(5):
        name=sheet['A'+str(i+2)].value
        temp=trial(name)
        pagedata.append(temp)
    flag=1
    wb2.save("Data.xlsx")

def data():
    global pagedata
    global flag

    now = datetime.now()
    current_time = now.strftime("%H:%M:%S")
    print("Current Time =", current_time)

    wb2 = load_workbook('Data.xlsx')
    sheet = wb2.active
    
    for i in range(len(pagedata)):
        try:
            r=requests.get(pagedata[i][0],timeout=5)
            soup = BeautifulSoup(r.text, "html5lib")
            price=soup.find_all('div',class_="My(6px) Pos(r) smartphone_Mt(6px)")[0].find('span').text
            if ',' in price:
                price=float(price.replace(',',""))
            if type(price) == str:
                price=float(price)
            print(sheet['A'+str(i+2)].value,price)
            
            if (sheet['B'+str(i+2)].value)!=0:
                previous=sheet['B'+str(i+2)].value
            else:
                previous=price        
            sheet['C'+str(i+2)]=previous
            sheet['B'+str(i+2)]=price
            sheet['D'+str(i+2)]=(price-previous)*100/previous
            if sheet['E'+str(i+2)].value!=0:    
                sheet['H'+str(i+2)]=(price-sheet['E'+str(i+2)].value)*100/sheet['E'+str(i+2)].value
            if sheet['J'+str(i+2)].value!=0:
                sheet['L'+str(i+2)]=(price-sheet['J'+str(i+2)].value)*100/sheet['J'+str(i+2)].value
        
            if(flag==1):
                sheet['J'+str(i+2)]=price
        except:
            print(sheet['A'+str(i+2)].value + " failed no updates done!!!")
    flag=0
    wb2.save('Data.xlsx')

