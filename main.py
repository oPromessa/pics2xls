import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.drawing.image import Image

def insert_picture_to_excel(row, worksheet):
    img_path = row['picture']
    img = Image(img_path)
    img.width = 100  # Set the desired image width in the cell
    img.height = 100  # Set the desired image height in the cell
    worksheet.column_dimensions['F'].width = 25  # Adjust the width of the 'remarks' column
    worksheet.row_dimensions[row.name + 2].height = 100  # Adjust the row height to fit the image
    worksheet.add_image(img, f'F{row.name + 2}')  # Adjust the cell where the image will be inserted

def csv_to_excel_with_images(csv_file_path, excel_file_path):
    df = pd.read_csv(csv_file_path)

    # Create a new Excel workbook and add a worksheet
    workbook = Workbook()
    worksheet = workbook.active

    # Write the CSV data to the Excel worksheet
    for row in dataframe_to_rows(df, index=False, header=True):
        worksheet.append(row)

    # Insert the images into the worksheet
    for index, row in df.iterrows():
        insert_picture_to_excel(row, worksheet)

    # Save the Excel workbook to the specified file
    workbook.save(excel_file_path)

if __name__ == "__main__":
    csv_file_path = "input.csv"  # Replace with the path to your CSV file
    excel_file_path = "output.xlsx"  # Replace with the desired output Excel file path

    csv_to_excel_with_images(csv_file_path, excel_file_path)
