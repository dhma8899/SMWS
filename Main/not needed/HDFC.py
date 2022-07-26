import openpyxl
from datetime import datetime

def Buy_HDFC():  
    wb_obj = openpyxl.load_workbook("Data.xlsx") 
    sheet_obj = wb_obj.active 
    HDFC_Current =sheet_obj['B5'].value
    HDFC_Change =sheet_obj['D5'].value
    HDFC_Purchase_Stock =sheet_obj['F5'].value
    HDFC_Money_left=sheet_obj['G5'].value
    HDFC_Purchase_Value=sheet_obj['E5'].value
    HDFC_Purchase_Change=sheet_obj['H5'].value
    Base_HDFC=sheet_obj['J5'].value
    flag=0
    if(HDFC_Current-Base_HDFC>5):
        flag=1
    if(HDFC_Current-Base_HDFC<0):
        sheet_obj['J5']=HDFC_Current


    now = datetime.now()
    current_time = now.strftime("%H:%M:%S")
    list_time=[]
    list_time.append(int(now.strftime('%H')))
    list_time.append(int(now.strftime('%M')))

    validation_time=list_time[0]*100+list_time[1]    
    if HDFC_Purchase_Change>=0.4 or HDFC_Purchase_Change<-3:
        Sell_HDFC()
        return 0
    elif HDFC_Change<=0 and HDFC_Change>-10 and (HDFC_Money_left>HDFC_Current*2) and (validation_time<=1430) and flag==0: 
        file1 = open("logs.txt","a")
        New_Purchase_Stock=int(((HDFC_Money_left/HDFC_Current)/4))
        if New_Purchase_Stock==0 and HDFC_Money_left>HDFC_Current:
            New_Purchase_Stock=int(((HDFC_Money_left/HDFC_Current)/2))            
        cost=New_Purchase_Stock*HDFC_Current
        HDFC_Money_left=HDFC_Money_left-cost
        HDFC_Purchase_Stock+=New_Purchase_Stock
        if(HDFC_Purchase_Value!=0):
            sheet_obj['E5']=(HDFC_Current+HDFC_Purchase_Value)/2
        else:
            sheet_obj['E5']=HDFC_Current
        sheet_obj['F5']=HDFC_Purchase_Stock
        sheet_obj['G5']=HDFC_Money_left
        sheet_obj['K5']=(list_time[0]*100)+list_time[1]
        L=["\n",str(current_time),"\nBought HDFC Stocks:",str(New_Purchase_Stock),"\nPurchase Rate:",str(HDFC_Current),"\nCurrent Balance:",str(HDFC_Money_left)]
        file1.writelines(L) 
        file1.close()
        print("Bought {0} HDFC Stock at {1} and current balance:{2}".format(New_Purchase_Stock,HDFC_Current,HDFC_Money_left))
    elif((list_time[0]==15 and list_time[1]>=15)):
        Sell_HDFC()
        return 0
    wb_obj.save("Data.xlsx")

def Sell_HDFC():
    wb_obj = openpyxl.load_workbook("Data.xlsx") 
    sheet_obj = wb_obj.active 
    HDFC_Current =sheet_obj['B5'].value
    HDFC_Change =sheet_obj['D5'].value
    HDFC_Purchase_Stock =sheet_obj['F5'].value
    HDFC_Money_left=sheet_obj['G5'].value
    HDFC_Purchase_Value=sheet_obj['E5'].value
    HDFC_Purchase_Change=sheet_obj['H5'].value
    profit_loss=sheet_obj['I5'].value

    now = datetime.now()
    current_time = now.strftime("%H:%M:%S")

    Selling_Price = HDFC_Purchase_Stock*HDFC_Current
    profit_loss+=Selling_Price-(HDFC_Purchase_Value*HDFC_Purchase_Stock)
    HDFC_Money_left+=Selling_Price
    file1 = open("logs.txt","a")
    L=["\n",str(current_time),"\nSold HDFC Stocks:",str(HDFC_Purchase_Stock),"\nSelling Rate:",str(HDFC_Current),"\nProfit or Loss:",str(profit_loss),"\nCurrent Balance:",str(HDFC_Money_left)]
    file1.writelines(L) 
    file1.close()

    print("Sold {0} HDFC Stock at {1} now balance={2} thus profit/loss of{3}".format(HDFC_Purchase_Stock,HDFC_Current,HDFC_Money_left,profit_loss))

    HDFC_Purchase_Stock=0
    HDFC_Purchase_Value=0
    HDFC_Purchase_Change=0

    
    sheet_obj['E5']=HDFC_Purchase_Value
    sheet_obj['F5']=HDFC_Purchase_Stock
    sheet_obj['G5']=HDFC_Money_left
    sheet_obj['H5']=HDFC_Purchase_Change
    sheet_obj['I5']=profit_loss
    wb_obj.save("Data.xlsx")
    
