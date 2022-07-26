import openpyxl
from datetime import datetime

def Buy():
    
    now = datetime.now()
    current_time = now.strftime("%H:%M:%S")
    list_time=[]
    list_time.append(int(now.strftime('%H')))
    list_time.append(int(now.strftime('%M')))
    validation_time=list_time[0]*100+list_time[1]
    for i in range(5):
        wb_obj = openpyxl.load_workbook("Data.xlsx") 
        sheet_obj = wb_obj.active
        Current =sheet_obj['B'+str(i+2)].value
        Change =sheet_obj['D'+str(i+2)].value
        Purchase_Stock =sheet_obj['F'+str(i+2)].value
        Money_left=sheet_obj['G'+str(i+2)].value
        Purchase_Value=sheet_obj['E'+str(i+2)].value
        Purchase_Change=sheet_obj['H'+str(i+2)].value
        Base=sheet_obj['J'+str(i+2)].value
        Base_Change=sheet_obj['L'+str(i+2)].value
        flag=0
        if(Current-Base>5):
            flag=1
        if(Current-Base<0):
            sheet_obj['J'+str(i+2)]=Current
       
        change_factor=False
        if(Purchase_Change>=0):
            pass
        elif(Purchase_Change<0 and Base_Change>0.5):
            change_factor=True
        
        if change_factor or Purchase_Change<-3:
            Sell(i)
        elif Change<=0 and Change>-10 and (Money_left>Current*2) and validation_time<=1430 and flag==0:
            
            New_Purchase_Stock=int(((Money_left/Current)/4))
            
            if New_Purchase_Stock==0 and Money_left>Current:
                
                New_Purchase_Stock=int(((Money_left/Current)/2))
                
            cost=New_Purchase_Stock*Current
            
            Money_left=Money_left-cost
            
            Purchase_Stock+=New_Purchase_Stock
            
            if(Purchase_Value!=0):
                sheet_obj['E'+str(i+2)]=(Current+Purchase_Value)/2
            else:
                sheet_obj['E'+str(i+2)]=Current
            sheet_obj['F'+str(i+2)]=Purchase_Stock
            sheet_obj['G'+str(i+2)]=Money_left
            sheet_obj['K'+str(i+2)]=(list_time[0]*100)+list_time[1]
            file1 = open("logs.txt","a")
            L=["\n",str(current_time),"\nBought " +sheet_obj['A'+str(i+2)].value+ "Stocks:",str(New_Purchase_Stock),"\nPurchase Rate:",str(Current),"\nCurrent Balance:",str(Money_left)]
            file1.writelines(L) 
            file1.close()
            print("Bought "+str(New_Purchase_Stock)+" "+sheet_obj['A'+str(i+2)].value + " Stock at "+str(Current)+"and current balance:"+str(Money_left))
        elif((list_time[0]==15 and list_time[1]>=15)):
            Sell(i)
            return 0
        wb_obj.save("Data.xlsx")

def Sell(i):
    wb_obj = openpyxl.load_workbook("Data.xlsx") 
    sheet_obj = wb_obj.active 
    Current =sheet_obj['B'+str(i+2)].value
    Change =sheet_obj['D'+str(i+2)].value
    Purchase_Stock =sheet_obj['F'+str(i+2)].value
    Money_left=sheet_obj['G'+str(i+2)].value
    Purchase_Value=sheet_obj['E'+str(i+2)].value
    Purchase_Change=sheet_obj['H'+str(i+2)].value
    profit_loss = sheet_obj['I'+str(i+2)].value

    now = datetime.now()
    current_time = now.strftime("%H:%M:%S")

    Selling_Price = Purchase_Stock*Current
    profit_loss+=Selling_Price-(Purchase_Value*Purchase_Stock)
    Money_left+=Selling_Price
    file1 = open("logs.txt","a")
    L=["\n",str(current_time),"\nSold " +sheet_obj['A'+str(i+2)].value+ " Stocks:",str(Purchase_Stock),"\nSelling Rate:",str(Current),"\nProfit or Loss:",str(profit_loss),"\nCurrent Balance:",str(Money_left)]
    file1.writelines(L) 
    file1.close()
    print("Sold "+str(Purchase_Stock)+" "+sheet_obj['A'+str(i+2)].value +" Stock at "+str(Current)+"now balance="+str(Money_left)+" thus profit/loss of "+str(profit_loss))
    
    sheet_obj['E'+str(i+2)]=0
    sheet_obj['F'+str(i+2)]=0
    sheet_obj['G'+str(i+2)]=Money_left
    sheet_obj['H'+str(i+2)]=0
    sheet_obj['I'+str(i+2)]=profit_loss
    wb_obj.save("Data.xlsx")
    
