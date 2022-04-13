# BoM_Creator
The company requires to create a BoM for every single Sales order. I developed this project, which makes the whole process more efficient. 

Use of:
- xlwingz - integrates Excel and Python.
- pandas  - use of the Dataframes to manage data
- numpy - use of lists to manage data 

## Project File
It is the excel sheet which includes the sales order that we have received. It is a matrix where you can see the items on the column side and the sales order is the only row.

![example Project](https://user-images.githubusercontent.com/65776444/158789866-9497de59-74f7-43ae-9c80-b1dac636763d.PNG)


## Project_DB 
is the database with all the information. As Following

![example ProjectDB](https://user-images.githubusercontent.com/65776444/158790569-699af570-f183-4ca1-89bc-a56902a11cbe.PNG)

the code do as following:
- Takes information from the project file
- Examine which information is useful in the proejct file. Which item is more than one in inventory.
- Paste the information on the output file for just the information related to the order and all its items 


```
from turtle import pd
import xlwings as xw
import numpy as np
import pandas as pd

#Create a Workbook-where the information will be saved"
WB_Outp= xw.Book(r'C:\Users\Carlos_Castillo\Test\1.BoM_Creation\Output.xlsx')
ws1=WB_Outp.sheets('Tabelle1')
ws1.range('A1:G105').clear_contents()

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
DF_DB=ws3.range('A1:E126').options(pd.DataFrame,  index=False).value
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
ws1.range("A1").value=Filtered_DF_DB
ws1.range("A:A").delete()

Last_C = ws1.range(1,1).end('down')
Last_C=Last_C.end('right')

print(Last_C)

"save Changes"
WB_Outp.save()
WB_PROJECT.save()
WB_DB.save()

"Close"
WB_Outp.close()
WB_PROJECT.close()
WB_DB.close()
```

