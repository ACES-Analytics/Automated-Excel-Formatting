# -*- coding: UTF-8 -*-
"""
This script is to create dummy sales records for analytics purpose
Created on Tue Dec 13 17:38:00 2022
@author ACES ANALYTICS TEAM

"""

"""
Index
## 1. Import modules and packages
## 2. Import excel doc to be formated
## 3. Create workbook
## 4. Format sheet
## 5. Add title
## 6. Export data to excel file
"""

# Calculate start time of script running
from timeit import default_timer as timer
start = timer()

## 1. Import modules and packages
import pandas as pd
import xlwings as xw

## 2. Import excel doc to be formated
df = pd.read_excel(r'D:\Python\ACES_Analytics\ACES002\Input\ACES002_Sales Records.xlsx')

## 3. Creat workbook
# Open excel file
wb = xw.Book()

# Define name of worksheet
sheet = wb.sheets["Sheet1"]
sheet.name = "sales records"

## 4. Format sheet
# Assign values of df to worksheet
sheet.range("A1").options(index=False).value = df

# define range of whole worksheet
data_rng = sheet.range("A1").expand('table')

# define height and width of each cell
data_rng.row_height = 23
data_rng.column_width = 13

# Format border of cells
border_rng = sheet.range("A1").expand('table')
for bi in range(1,5):
    border_rng.api.Borders(bi).Weight = 2
    border_rng.api.Borders(bi).Color = 0x70ad47

# Format all range of the worksheet
data_rng.api.Font.Name = 'Verdana'
data_rng.api.Font.Size = 8
data_rng.api.VerticalAlignment = xw.constants.HAlign.xlHAlignCenter
data_rng.api.WrapText = True

# Format header
header_rng = sheet.range("A1").expand('right')
header_rng.color = ('#70ad47') #cca989 #2da8dc
header_rng.api.Font.Color = 0xffffff
header_rng.api.Font.Bold = True
header_rng.api.Font.Size = 9

# Format first column
id_column_rng = sheet.range("A2").expand('down')
id_column_rng.color = ('#c6e0b4') #6c8d84 #2da8dc #c6e0b4 green  #d8eaf1 light blue
id_column_rng.api.Font.Color = 0x000000 #000000 #0xffffff
id_column_rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignLeft
id_column_rng.column_width = 10

# Format B column , day
B_day_rng = sheet.range("B2").expand('down')
B_day_rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignLeft
B_day_rng.column_width = 5

# Format C column, customer
C_cust_rng = sheet.range("C2").expand('down')
C_cust_rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignLeft
C_cust_rng.column_width = 25

# Format D column, country
D_ctry_rng = sheet.range("D2").expand('down')
D_ctry_rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignLeft
D_ctry_rng.column_width = 13

# Format E column, salesman
E_slmn_rng = sheet.range("E2").expand('down')
E_slmn_rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignLeft
E_slmn_rng.column_width = 13

# Format F column, sales team
F_sltm_rng = sheet.range("F2").expand('down')
F_sltm_rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignLeft
F_sltm_rng.column_width = 13

# Format H column, material text
H_mtltext_rng = sheet.range("H2").expand('down')
H_mtltext_rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignLeft
H_mtltext_rng.column_width = 20

# Format I column, material code
I_mtlcode_rng = sheet.range("I2").expand('down')
I_mtlcode_rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignLeft
I_mtlcode_rng.column_width = 8

# Format J column, profit center
J_prfctr_rng = sheet.range("J2").expand('down')
J_prfctr_rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignLeft
J_prfctr_rng.column_width = 8

# Format K column, material group
K_mtlgrp_rng = sheet.range("K2").expand('down')
K_mtlgrp_rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignLeft
K_mtlgrp_rng.column_width = 8

# Format L colum, unit column
L_unit_rng = sheet.range("L2").expand('down')
L_unit_rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignLeft
L_unit_rng.column_width = 4

# Format G column, sales volume
G_volumn_rng = sheet.range("G2").expand('down')
G_volumn_rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignRight
G_volumn_rng.column_width = 8
G_volumn_rng.number_format = "#,###"

# Format M column, average price
M_aveprice_rng = sheet.range("M2").expand('down')
M_aveprice_rng.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignRight
M_aveprice_rng.column_width = 12
M_aveprice_rng.number_format = "#,###.0000"

# Format color of tab
sheet.api.Tab.Color = 0x70AD47

## 5. Add title
# Get length of rows and columns
rowl = str(len(df))
#clml = len(df.columns)

# Get time
from time import strftime, localtime
time_local = strftime ("%A, %d %b %Y, %H:%M")
print(time_local)

# Insert two rows above row 1
sheet.range("1:3").insert('down')
sheet.range("A1").value = "Sales Records"
sheet.range("I1").value = "The last update time :  " + time_local  +"."
sheet.range("I2").value = "Updated by:  ACES Analytics Team."
sheet.range("A2").value = "Length of Rows:"
sheet.range("C2").value = rowl

# Format title
title = sheet.range("A1")
title.api.Font.Name = 'Verdana'
title.api.Font.Size = 12

# Format cell of length of rows
rowlen = sheet.range("C2")
rowlen.api.HorizontalAlignment = xw.constants.HAlign.xlHAlignLeft
rowlen.number_format = "#,###"

## 6. Export data to excel file
wb.save(r'D:\Python\ACES_Analytics\ACES002\Output\ACES002_Sales Records_Formated.xlsx')
wb.close()

# Print the end for this script
print("The run of script is completed successfully.")

# Print time now
time_local = strftime ("%A, %d %b %Y, %H:%M")
print(time_local)

end = timer()
running_time = "{:,.2f}".format(end -start)
print (running_time)