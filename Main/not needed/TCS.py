import openpyxl
from datetime import datetime

def Buy_TCS():  
    wb_obj = openpyxl.load_workbook("Data.xlsx") 
    sheet_obj = wb_obj.active 
    TCS_Current =sheet_obj['B2'].value
    TCS_Change =sheet_obj['D2'].value
    TCS_Purchase_Stock =sheet_obj['F2'].value
    TCS_Money_left=sheet_obj['G2'].value
    TCS_Purchase_Value=sheet_obj['E2'].value
    TCS_Purchase_Change=sheet_obj['H2'].value
    Base_TCS=sheet_obj['J2'].value
    Base_Change=sheet_obj['L2'].value
    flag=0
    if(TCS_Current-Base_TCS>5):
    	flag=1
    if(TCS_Current-Base_TCS<0):
    	sheet_obj['J2']=TCS_Current

    now = datetime.now()
    current_time = now.strftime("%H:%M:%S")
    list_time=[]
    list_time.append(int(now.strftime('%H')))
    list_time.append(int(now.strftime('%M')))
    change_factor=False
    if(TCS_Purchase_Change>=0):
        pass
    elif(TCS_Purchase_Change<0 and Base_change>1):
        change_factor=True

    validation_time=list_time[0]*100+list_time[1]
    if change_factor or TCS_Purchase_Change<-3:
        Sell_TCS()
        #return 0
    elif TCS_Change<=0 and TCS_Change>-10 and (TCS_Money_left>TCS_Current*2) and validation_time<=1430 and flag==0:
        New_Purchase_Stock=int(((TCS_Money_left/TCS_Current)/4))
        if New_Purchase_Stock==0 and TCS_Money_left>TCS_Current:
            New_Purchase_Stock=int(((TCS_Money_left/TCS_Current)/2))            
        cost=New_Purchase_Stock*TCS_Current
        TCS_Money_left=TCS_Money_left-cost
        TCS_Purchase_Stock+=New_Purchase_Stock
        if(TCS_Purchase_Value!=0):
            sheet_obj['E2']=(TCS_Current+TCS_Purchase_Value)/2
        else:
            sheet_obj['E2']=TCS_Current
        sheet_obj['F2']=TCS_Purchase_Stock
        sheet_obj['G2']=TCS_Money_left
        sheet_obj['K2']=(list_time[0]*100)+list_time[1]
        file1 = open("logs.txt","a")
        L=["\n",str(current_time),"\nBought TCS Stocks:",str(New_Purchase_Stock),"\nPurchase Rate:",str(TCS_Current),"\nCurrent Balance:",str(TCS_Money_left)]
        file1.writelines(L) 
        file1.close()
        print("Bought {0} TCS Stock at {1} and current balance:{2}".format(New_Purchase_Stock,TCS_Current,TCS_Money_left))
    elif((list_time[0]==15 and list_time[1]>=15)):
        Sell_TCS()
        return 0
    wb_obj.save("Data.xlsx")

def Sell_TCS():
    wb_obj = openpyxl.load_workbook("Data.xlsx") 
    sheet_obj = wb_obj.active 
    TCS_Current =sheet_obj['B2'].value
    TCS_Change =sheet_obj['D2'].value
    TCS_Purchase_Stock =sheet_obj['F2'].value
    TCS_Money_left=sheet_obj['G2'].value
    TCS_Purchase_Value=sheet_obj['E2'].value
    TCS_Purchase_Change=sheet_obj['H2'].value
    profit_loss = sheet_obj['I2'].value

    now = datetime.now()
    current_time = now.strftime("%H:%M:%S")

    Selling_Price = TCS_Purchase_Stock*TCS_Current
    profit_loss+=Selling_Price-(TCS_Purchase_Value*TCS_Purchase_Stock)
    TCS_Money_left+=Selling_Price
    file1 = open("logs.txt","a")
    L=["\n",str(current_time),"\nSold TCS Stocks:",str(TCS_Purchase_Stock),"\nSelling Rate:",str(TCS_Current),"\nProfit or Loss:",str(profit_loss),"\nCurrent Balance:",str(TCS_Money_left)]
    file1.writelines(L) 
    file1.close()

    print("Sold {0} TCS Stock at {1} now balance={2} thus profit/loss of{3}".format(TCS_Purchase_Stock,TCS_Current,TCS_Money_left,profit_loss))

    TCS_Purchase_Stock=0
    TCS_Purchase_Value=0
    TCS_Purchase_Change=0

    
    sheet_obj['E2']=TCS_Purchase_Value
    sheet_obj['F2']=TCS_Purchase_Stock
    sheet_obj['G2']=TCS_Money_left
    sheet_obj['H2']=TCS_Purchase_Change
    sheet_obj['I2']=profit_loss
    wb_obj.save("Data.xlsx")
    
