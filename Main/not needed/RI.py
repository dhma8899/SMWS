import openpyxl
from datetime import datetime

def Buy_RI():  
    wb_obj = openpyxl.load_workbook("Data.xlsx") 
    sheet_obj = wb_obj.active 
    RI_Current =sheet_obj['B3'].value
    RI_Change =sheet_obj['D3'].value
    RI_Purchase_Stock =sheet_obj['F3'].value
    RI_Money_left=sheet_obj['G3'].value
    RI_Purchase_Value=sheet_obj['E3'].value
    RI_Purchase_Change=sheet_obj['H3'].value
    Base_RI=sheet_obj['J3'].value
    flag=0
    if(RI_Current-Base_RI>5):
        flag=1
    if(RI_Current-Base_RI<0):
        sheet_obj['J3']=RI_Current


    now = datetime.now()
    current_time = now.strftime("%H:%M:%S")
    list_time=[]
    list_time.append(int(now.strftime('%H')))
    list_time.append(int(now.strftime('%M')))

    validation_time=list_time[0]*100+list_time[1]    
    if RI_Purchase_Change>=0.4 or RI_Purchase_Change<-3:
        Sell_RI()
        return 0
    elif RI_Change<=0 and RI_Change>-10 and (RI_Money_left>RI_Current*2) and validation_time<=1430 and flag==0: 
        file1 = open("logs.txt","a")
        New_Purchase_Stock=int(((RI_Money_left/RI_Current)/4))
        if New_Purchase_Stock==0 and RI_Money_left>RI_Current:
            New_Purchase_Stock=int(((RI_Money_left/RI_Current)/2))            
        cost=New_Purchase_Stock*RI_Current
        RI_Money_left=RI_Money_left-cost
        RI_Purchase_Stock+=New_Purchase_Stock
        if(RI_Purchase_Value!=0):
            sheet_obj['E3']=(RI_Current+RI_Purchase_Value)/2
        else:
            sheet_obj['E3']=RI_Current
        sheet_obj['F3']=RI_Purchase_Stock
        sheet_obj['G3']=RI_Money_left
        sheet_obj['K3']=(list_time[0]*100)+list_time[1]
        L=["\n",str(current_time),"\nBought RI Stocks:",str(New_Purchase_Stock),"\nPurchase Rate:",str(RI_Current),"\nCurrent Balance:",str(RI_Money_left)]
        file1.writelines(L) 
        file1.close()
        print("Bought {0} RI Stock at {1} and current balance:{2}".format(New_Purchase_Stock,RI_Current,RI_Money_left))
    elif((list_time[0]==15 and list_time[1]>=15)):
        Sell_RI()
        return 0
    wb_obj.save("Data.xlsx")

def Sell_RI():
    wb_obj = openpyxl.load_workbook("Data.xlsx") 
    sheet_obj = wb_obj.active 
    RI_Current =sheet_obj['B3'].value
    RI_Change =sheet_obj['D3'].value
    RI_Purchase_Stock =sheet_obj['F3'].value
    RI_Money_left=sheet_obj['G3'].value
    RI_Purchase_Value=sheet_obj['E3'].value
    RI_Purchase_Change=sheet_obj['H3'].value
    profit_loss=sheet_obj['I3'].value

    now = datetime.now()
    current_time = now.strftime("%H:%M:%S")

    Selling_Price = RI_Purchase_Stock*RI_Current
    profit_loss+=Selling_Price-(RI_Purchase_Value*RI_Purchase_Stock)
    RI_Money_left+=Selling_Price
    file1 = open("logs.txt","a")
    L=["\n",str(current_time),"\nSold RI Stocks:",str(RI_Purchase_Stock),"\nSelling Rate:",str(RI_Current),"\nProfit or Loss:",str(profit_loss),"\nCurrent Balance:",str(RI_Money_left)]
    file1.writelines(L) 
    file1.close()

    print("Sold {0} RI Stock at {1} now balance={2} thus profit/loss of{3}".format(RI_Purchase_Stock,RI_Current,RI_Money_left,profit_loss))

    RI_Purchase_Stock=0
    RI_Purchase_Value=0
    RI_Purchase_Change=0

    
    sheet_obj['E3']=RI_Purchase_Value
    sheet_obj['F3']=RI_Purchase_Stock
    sheet_obj['G3']=RI_Money_left
    sheet_obj['H3']=RI_Purchase_Change
    sheet_obj['I3']=profit_loss
    wb_obj.save("Data.xlsx")
    
