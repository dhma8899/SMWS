import openpyxl
from datetime import datetime

def Buy_BAF():  
    wb_obj = openpyxl.load_workbook("Data.xlsx") 
    sheet_obj = wb_obj.active 
    BAF_Current =sheet_obj['B4'].value
    BAF_Change =sheet_obj['D4'].value
    BAF_Purchase_Stock =sheet_obj['F4'].value
    BAF_Money_left=sheet_obj['G4'].value
    BAF_Purchase_Value=sheet_obj['E4'].value
    BAF_Purchase_Change=sheet_obj['H4'].value
    Base_BAF=sheet_obj['J4'].value
    flag=0
    if(BAF_Current-Base_BAF>5):
        flag=1
    if(BAF_Current-Base_BAF<0):
        sheet_obj['J4']=BAF_Current

    now = datetime.now()
    current_time = now.strftime("%H:%M:%S")
    list_time=[]
    list_time.append(int(now.strftime('%H')))
    list_time.append(int(now.strftime('%M')))

    validation_time=list_time[0]*100+list_time[1]
    if BAF_Purchase_Change>=0.4 or BAF_Purchase_Change<-3:
        Sell_BAF()
        return 0
    elif BAF_Change<=0 and BAF_Change>-10 and (BAF_Money_left>BAF_Current*2) and (validation_time<=1420) and flag==0: 
        file1 = open("logs.txt","a")
        New_Purchase_Stock=int(((BAF_Money_left/BAF_Current)/4))
        if New_Purchase_Stock==0 and BAF_Money_left>BAF_Current:
            New_Purchase_Stock=int(((BAF_Money_left/BAF_Current)/2))            
        cost=New_Purchase_Stock*BAF_Current
        BAF_Money_left=BAF_Money_left-cost
        BAF_Purchase_Stock+=New_Purchase_Stock
        if(BAF_Purchase_Value!=0):
            sheet_obj['E4']=(BAF_Current+BAF_Purchase_Value)/2
        else:
            sheet_obj['E4']=BAF_Current
        sheet_obj['F4']=BAF_Purchase_Stock
        sheet_obj['G4']=BAF_Money_left
        sheet_obj['K4']=(list_time[0]*100)+list_time[1]
        L=["\n",str(current_time),"\nBought BAF Stocks:",str(New_Purchase_Stock),"\nPurchase Rate:",str(BAF_Current),"\nCurrent Balance:",str(BAF_Money_left)]
        file1.writelines(L) 
        file1.close()
        print("Bought {0} BAF Stock at {1} and current balance:{2}".format(New_Purchase_Stock,BAF_Current,BAF_Money_left))
    elif((list_time[0]==15 and list_time[1]>=15)):
        Sell_BAF()
        return 0
    wb_obj.save("Data.xlsx")

def Sell_BAF():
    wb_obj = openpyxl.load_workbook("Data.xlsx") 
    sheet_obj = wb_obj.active 
    BAF_Current =sheet_obj['B4'].value
    BAF_Change =sheet_obj['D4'].value
    BAF_Purchase_Stock =sheet_obj['F4'].value
    BAF_Money_left=sheet_obj['G4'].value
    BAF_Purchase_Value=sheet_obj['E4'].value
    BAF_Purchase_Change=sheet_obj['H4'].value
    profit_loss=sheet_obj['I4'].value

    now = datetime.now()
    current_time = now.strftime("%H:%M:%S")

    Selling_Price = BAF_Purchase_Stock*BAF_Current
    profit_loss+=Selling_Price-(BAF_Purchase_Value*BAF_Purchase_Stock)
    BAF_Money_left+=Selling_Price
    file1 = open("logs.txt","a")
    L=["\n",str(current_time),"\nSold BAF Stocks:",str(BAF_Purchase_Stock),"\nSelling Rate:",str(BAF_Current),"\nProfit or Loss:",str(profit_loss),"\nCurrent Balance:",str(BAF_Money_left)]
    file1.writelines(L) 
    file1.close()

    print("Sold {0} BAF Stock at {1} now balance={2} thus profit/loss of{3}".format(BAF_Purchase_Stock,BAF_Current,BAF_Money_left,profit_loss))

    BAF_Purchase_Stock=0
    BAF_Purchase_Value=0
    BAF_Purchase_Change=0

    
    sheet_obj['E4']=BAF_Purchase_Value
    sheet_obj['F4']=BAF_Purchase_Stock
    sheet_obj['G4']=BAF_Money_left
    sheet_obj['H4']=BAF_Purchase_Change
    sheet_obj['I4']=profit_loss
    wb_obj.save("Data.xlsx")
    
