import openpyxl
from datetime import datetime

def Buy_IT():  
    wb_obj = openpyxl.load_workbook("Data.xlsx") 
    sheet_obj = wb_obj.active 
    IT_Current =sheet_obj['B6'].value
    IT_Change =sheet_obj['D6'].value
    IT_Purchase_Stock =sheet_obj['F6'].value
    IT_Money_left=sheet_obj['G6'].value
    IT_Purchase_Value=sheet_obj['E6'].value
    IT_Purchase_Change=sheet_obj['H6'].value
    Base_IT=sheet_obj['J6'].value
    flag=0
    if(IT_Current-Base_IT>5):
        flag=1
    if(IT_Current-Base_IT<0):
        sheet_obj['J6']=IT_Current

    now = datetime.now()
    current_time = now.strftime("%H:%M:%S")
    list_time=[]
    list_time.append(int(now.strftime('%H')))
    list_time.append(int(now.strftime('%M')))

    validation_time=list_time[0]*100+list_time[1]
    if IT_Purchase_Change>=0.4 or IT_Purchase_Change<-3:
        Sell_IT()
        return 0
    elif IT_Change<=0 and IT_Change>-10 and (IT_Money_left>IT_Current*2) and validation_time<=1430 and flag==0: 
        file1 = open("logs.txt","a")
        New_Purchase_Stock=int(((IT_Money_left/IT_Current)/4))
        if New_Purchase_Stock==0 and IT_Money_left>IT_Current:
            New_Purchase_Stock=int(((IT_Money_left/IT_Current)/2))            
        cost=New_Purchase_Stock*IT_Current
        IT_Money_left=IT_Money_left-cost
        IT_Purchase_Stock+=New_Purchase_Stock
        if(IT_Purchase_Value!=0):
            sheet_obj['E6']=(IT_Current+IT_Purchase_Value)/2
        else:
            sheet_obj['E6']=IT_Current
        sheet_obj['F6']=IT_Purchase_Stock
        sheet_obj['G6']=IT_Money_left
        sheet_obj['K6']=(list_time[0]*100)+list_time[1]
        L=["\n",str(current_time),"\nBought IT Stocks:",str(New_Purchase_Stock),"\nPurchase Rate:",str(IT_Current),"\nCurrent Balance:",str(IT_Money_left)]
        file1.writelines(L) 
        file1.close()
        print("Bought {0} IT Stock at {1} and current balance:{2}".format(New_Purchase_Stock,IT_Current,IT_Money_left))
    elif((list_time[0]==15 and list_time[1]>=15)):
        Sell_IT()
        return 0
    wb_obj.save("Data.xlsx")

def Sell_IT():
    wb_obj = openpyxl.load_workbook("Data.xlsx") 
    sheet_obj = wb_obj.active 
    IT_Current =sheet_obj['B6'].value
    IT_Change =sheet_obj['D6'].value
    IT_Purchase_Stock =sheet_obj['F6'].value
    IT_Money_left=sheet_obj['G6'].value
    IT_Purchase_Value=sheet_obj['E6'].value
    IT_Purchase_Change=sheet_obj['H6'].value
    profit_loss=sheet_obj['I6'].value

    now = datetime.now()
    current_time = now.strftime("%H:%M:%S")

    Selling_Price = IT_Purchase_Stock*IT_Current
    profit_loss+=Selling_Price-(IT_Purchase_Value*IT_Purchase_Stock)
    IT_Money_left+=Selling_Price
    file1 = open("logs.txt","a")
    L=["\n",str(current_time),"\nSold IT Stocks:",str(IT_Purchase_Stock),"\nSelling Rate:",str(IT_Current),"\nProfit or Loss:",str(profit_loss),"\nCurrent Balance:",str(IT_Money_left)]
    file1.writelines(L) 
    file1.close()

    print("Sold {0} IT Stock at {1} now balance={2} thus profit/loss of{3}".format(IT_Purchase_Stock,IT_Current,IT_Money_left,profit_loss))

    IT_Purchase_Stock=0
    IT_Purchase_Value=0
    IT_Purchase_Change=0

    
    sheet_obj['E6']=IT_Purchase_Value
    sheet_obj['F6']=IT_Purchase_Stock
    sheet_obj['G6']=IT_Money_left
    sheet_obj['H6']=IT_Purchase_Change
    sheet_obj['I6']=profit_loss
    wb_obj.save("Data.xlsx")
    
