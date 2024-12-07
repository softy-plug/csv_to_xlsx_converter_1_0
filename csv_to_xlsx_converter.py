import os

input("Нажмите Enter для запуска программы")

os.system("pip install openpyxl")

import csv
from openpyxl import Workbook

# Get the current directory
current_directory = os.getcwd()

# Iterate over all files in the directory
for filename in os.listdir(current_directory):
    if filename.endswith('.csv'):
        # Construct full file path
        csv_file_path = os.path.join(current_directory, filename)
        
        # Create a new workbook
        wb = Workbook()
        ws = wb.active

        # Read the CSV file and write to the Excel sheet
        with open(csv_file_path, newline='', encoding='utf-8') as csvfile:
            csvreader = csv.reader(csvfile, delimiter=';')  # Use ';' as delimiter
            for row in csvreader:
                # Replace ';;;' with ';' in each cell
                row = [cell.replace(';;;', ';') for cell in row]
                ws.append(row)  # Append row to the worksheet

        # Save the workbook with the same name but with .xlsx extension
        xlsx_file_path = os.path.join(current_directory, filename[:-4] + '.xlsx')
        wb.save(xlsx_file_path)

print("Конвертация файлов завершена успешно.")
 
input("Нажмите Enter для закрытия окна")

#softy_plug