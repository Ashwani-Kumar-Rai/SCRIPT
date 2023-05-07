import pandas as pd
import keyboard
import os
import time
from datetime import datetime, timedelta

# Load the Excel file into a DataFrame
df = pd.read_excel('AllPurchases.xlsx')

# Convert 'Expiry date' column to datetime format
df['Expiry date'] = pd.to_datetime(df['Expiry date'], errors='coerce')

# Get the unique values of the Corporate Unit column
corporate_units = df['Corporate unit'].unique()
df.drop(df[df['Status'] == "[Cancelled]"].index, inplace = True)


# Loop through the corporate units and create new Excel files
for unit in corporate_units:
    # Filter the DataFrame by the current corporate unit and Expiry date within next 30 days
    expiry_date = datetime.today() + timedelta(days=30)
    filtered_df = df[(df['Corporate unit'] == unit) & (df['Expiry date'] >= datetime.today()) & (df['Expiry date'] <= expiry_date)]
    
    # Check for and drop any duplicate rows
    filtered_df = filtered_df.drop_duplicates()

    if len(filtered_df) > 0 :

        # Create a new Excel file for the current corporate unit
        file_name = f'{unit.replace("GLOBAL/","").replace("INACTIVE/","")}.xlsx'
        writer = pd.ExcelWriter(file_name, engine='xlsxwriter')
        
        # Write the filtered data to the new file
        filtered_df.to_excel(writer, index=False)


        #Save and close the new file
        writer.close()
        
        # Open the new file in Excel
        os.startfile(file_name, 'open')
        time.sleep(10) ##################### WRITE A BETTER CONDITION HERE
        # # Wait for Excel to open
        # keyboard.wait('ctrl')
        
        keyboard.press_and_release('ctrl+a')
        
        # # Simulate keyboard shortcut for resizing rows
        keyboard.press_and_release('alt+h')
        keyboard.press_and_release('o')
        keyboard.press_and_release('a')
        
        # # Simulate keyboard shortcut for resizing columns
        keyboard.press_and_release('alt+h')
        keyboard.press_and_release('o')
        keyboard.press_and_release('i')
        
        # # Save and close the Excel file
        keyboard.press_and_release('ctrl+s')
        time.sleep(10)
        keyboard.press_and_release('alt+f4')




