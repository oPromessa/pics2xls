import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.drawing.image import Image as ExcelImage
from PIL import Image as PILImage

def convert_to_jpeg_and_save(img_path):
    img = PILImage.open(img_path)
    img_format = img.format.lower()
    if img_format != 'jpeg':
        jpeg_path = img_path.replace(f'.{img_format}', '.jpeg')
        img = img.convert('RGB')  # Convert HEIC image to RGB before saving as JPEG
        img.save(jpeg_path, format='JPEG')
        return jpeg_path
    return img_path

def insert_picture_to_excel(row, worksheet):
    img_path = row['picture']
    img_path = convert_to_jpeg_and_save(img_path)  # Convert HEIC to JPEG if necessary
    img = ExcelImage(img_path)
    img.width = 100  # Set the desired image width in the cell
    img.height = 100  # Set the desired image height in the cell
    worksheet.column_dimensions['I'].width = 25  # Adjust the width of the 'remarks' column
    worksheet.row_dimensions[row.name + 2].height = 100  # Adjust the row height to fit the image
    worksheet.add_image(img, f'I{row.name + 2}')  # Adjust the cell where the image will be inserted

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
