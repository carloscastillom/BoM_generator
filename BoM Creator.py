from turtle import pd
import xlwings as xw
import numpy as np
import pandas as pd

#Create a Workbook-where the information will be saved"
WB_TEST= xw.Book()
WB_TEST.save(r'C:\Users\Carlos_Castillo\Test\1.BoM_Creation\test1.xlsx')
ws1=WB_TEST.sheets('Tabelle1')

#Open the WB project-where the info of the project is in"
WB_PROJECT= xw.Book(r'C:\Users\Carlos_Castillo\Test\1.BoM_Creation\Project.xlsx')
#Pick the worksheet"
ws2=WB_PROJECT.sheets('Tabelle1')

#Open the WB Database- where the db is saved. file to pull info into TEST
WB_DB=xw.Book(r'C:\Users\Carlos_Castillo\Test\1.BoM_Creation\Project_DB.xlsx')
ws3=WB_DB.sheets('Tabelle1')

#get Item codes names and quantities"
ItemCodes= ws2.range('F1:BK1').options(numbers=int).value
ItemNames= ws2.range('F2:BK2').value
ItemQuantity= ws2.range('F3:BK3').options(numbers=float).value   

#put Item codes, Names and Quantities" on a data frame
dfProject = pd.DataFrame(np.column_stack([ItemCodes, ItemNames, ItemQuantity]), columns=['Item_Codes', 'Item_Names', 'Item_Quantity'])
#Ensure the quantities are float
dfProject['Item_Quantity']= dfProject['Item_Quantity'].astype('float')
dfProject['Item_Codes']= dfProject['Item_Codes'].astype('int')
#Filter the Dataframe- these are the ones we only need to use 
filtProject = dfProject[dfProject['Item_Quantity']>0.0]

Item_list_project=filtProject['Item_Codes'].to_list()

print(filtProject)

"go look on a DataFrame to check what should be printed"
DF_DB=ws3.range('A1:D93').options(pd.DataFrame, header=1, index=False).value
DF_DB['Artikel']=DF_DB['Artikel'].astype('int')

#Filter the DF to include only items wiht DB
Filtered_DF_DB= DF_DB[DF_DB['Artikel'].isin(Item_list_project)]


#multiply the amount required by the project
for i in Item_list_project:

    #get the amount required to multiply the BoM- how many times the item is liste in the project
    temp=filtProject.loc[filtProject['Item_Codes'] == i, 'Item_Quantity'].values[0]
    temp=temp.astype('int')

    Filtered_DF_DB.loc[Filtered_DF_DB['Artikel'] == i, 'Mg.'] *= temp
    
print(Filtered_DF_DB)


"paste it in the test workbook"
#ws1.range("A1").value=dataToBePasted

"save Changes"
WB_TEST.save()
WB_PROJECT.save()
WB_DB.save()

"Close"
WB_TEST.close()
WB_PROJECT.close()
WB_DB.close()







