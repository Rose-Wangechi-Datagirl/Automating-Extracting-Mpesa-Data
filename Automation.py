#[Convert](https://www.adobe.com/acrobat/online/pdf-to-excel.html) your Mpesa statement from pdf to excel

import pandas as pd
import numpy as np
import os

#loading your document
##The header for the detailed table is usually on the 17th row after converting the pdf doeument
data=pd.read_excel('your Mpesa excel', header=17) #replace with your file name/path 

#drop unwanted columns
data= data.drop(columns=['Unnamed: 2','Unnamed: 4','Unnamed: 5',
                         'Unnamed: 6','Unnamed: 7','Unnamed: 8','Unnamed: 9','Unnamed: 13'])

#  filter out the data you want based on your transactions
data_needed = data[data['Details'].str.contains('PICK UP', case=False, na=False)] #replace 'PICK UP' with the data/name you are looking for
data_needed_2 = data[data['Details'].str.contains('Patricia', case=False, na=False)]

#merge the two dataframes to a single dataframe
final_data = pd.concat([data_needed, data_needed_2], ignore_index=True)

#selecting the date period we want
##covert the data
final_data['Completion Time'] = pd.to_datetime(final_data['Completion Time'])

# Select the time period
start_date = '2023-12-24'
end_date = '2024-01-15'

final_data = final_data[final_data['Completion Time'].between(start_date, end_date)]

#dropping columns not needed in final spreadsheet
final_data= final_data.drop(columns=['Transaction Status', 'Balance'])

# Extract the relevant part of the "Details" column
final_data['Details'] = final_data['Details'].str.extract(r' - (.+)$')

#adding row to show the total of paid in and withdrawn
# Calculate the total for 'Paid In' and 'Withdrawn'
total_paid_in = final_data['Paid In'].sum()
total_withdrawn = final_data['Withdrawn'].sum()

# Create a new row with total values
total_row = pd.DataFrame({
    'Receipt No.': [''],
    'Completion Time': [''],
    'Details':[''],
    'Paid In': [total_paid_in],
    'Withdrawn': [total_withdrawn]
})

data_with_total= final_data.append(total_row, ignore_index=True)

#replacing then null values with ''
data_with_total['Paid In']= data_with_total['Paid In'].fillna('')
data_with_total['Withdrawn']= data_with_total['Withdrawn'].fillna('')

#renaming the colums
data_with_total= data_with_total.rename(columns={'Paid In':'Received',
                                                 'Details':'Particulars',
                                                 'Completion Time':'Date',
                                                 'Receipt No.':'Transaction Code',
                                                 'Withdrawn':'Out'
                                                 })

#saving the fnal document as .xlxs
# Specify the directory path
output_directory = 'C:/Users/hp/Desktop/'
os.makedirs(output_directory, exist_ok=True)  # Creates the directory if it doesn't exist

# Specify the path to save the Excel file
excel_path = os.path.join(output_directory, 'final_data.xlsx')

# Save the DataFrame as an Excel file
data_with_total.to_excel(excel_path, index=False)

print(f"Data has been saved to: {excel_path}")