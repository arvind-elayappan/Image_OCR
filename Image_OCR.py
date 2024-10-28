#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import os
import pandas as pd
from paddleocr import PaddleOCR, draw_ocr
from PIL import ImageFont, ImageDraw, Image
import matplotlib.pyplot as plt

# Initialize the PaddleOCR model
ocr_model = PaddleOCR(use_angle_cls=True, lang='en')

# Function to perform OCR and return text and tax references
def perform_ocr_and_extract_tax_info(image_path, fonts):
    # Perform OCR
    result = ocr_model.ocr(image_path)
    extracted_text = ' '.join([elements[1][0] for elements in result[0]])
    
    # Tax reference strings
    negative_tax_refs = ['non-taxable', 'non taxable', 'not taxed', 'not-taxed', 'no tax']
    positive_tax_refs = ['tax', 'taxed', 'VAT', 'GST', 'sales tax', 'TAX','%']

    # Initialize columns for negative and positive tax references
    neg_tax_found = [ref for ref in negative_tax_refs if ref.lower() in extracted_text.lower()]
    pos_tax_found = [ref for ref in positive_tax_refs if ref.lower() in extracted_text.lower()]

    # Count of positive tax references
    positive_tax_count = len(pos_tax_found)

    return extracted_text, neg_tax_found, pos_tax_found, positive_tax_count

# Function to dynamically use all Windows fonts
def apply_fonts_dynamically(image, result, fonts):
    for font in fonts:
        try:
            boxes = [elements[0] for elements in result[0]]
            txts = [elements[1][0] for elements in result[0]]
            scores = [elements[1][1] for elements in result[0]]
            im_show = draw_ocr(image, boxes, txts, scores, font_path=font)
            plt.imshow(im_show)
            plt.show()
            break  # If a font is successfully applied, exit the loop
        except Exception as e:
            print(f"Font {font} failed: {e}")
            continue  # Try the next font if there's an error

# Function to get all fonts available in the system
def get_system_fonts():
    windows_font_dir = 'C:/Windows/Fonts/'
    fonts = []
    for fontfile in os.listdir(windows_font_dir):
        if fontfile.endswith(".ttf"):
            fonts.append(os.path.join(windows_font_dir, fontfile))
    return fonts

# Function to process receipts and export OCR results and tax-related info to Excel
def process_receipts_and_export_to_excel(receipts_dir, output_excel, fonts):
    # List to store results
    results = []

    # Loop through all image files in the directory
    for filename in os.listdir(receipts_dir):
        if filename.endswith(".jpg") or filename.endswith(".png"):  # Add more extensions if needed
            image_path = os.path.join(receipts_dir, filename)
            print(f"Processing {filename}")
            
            # Load image
            image = Image.open(image_path)
            
            # Perform OCR and extract tax-related info
            extracted_text, neg_tax_found, pos_tax_found, positive_tax_count = perform_ocr_and_extract_tax_info(image_path, fonts)
            
            # Apply fonts dynamically and visualize (optional)
            result = ocr_model.ocr(image_path)
            apply_fonts_dynamically(image, result, fonts)

            # Append to results
            results.append({
                'File Name': filename,
                'Extracted Text': extracted_text
                # 'Negative Tax Ref': ', '.join(neg_tax_found),
                # 'Positive Tax Ref': ', '.join(pos_tax_found),
                # 'Positive Tax Count': positive_tax_count
            })

    # Create a DataFrame from the results
    df = pd.DataFrame(results)

# Function to process receipts and export OCR results to Excel
# def process_receipts_to_excel(receipts_dir, output_excel):
#     # List to store results
#     results = []

#     # Loop through all image files in the directory
#     for filename in os.listdir(receipts_dir):
#         if filename.endswith(".jpg") or filename.endswith(".png"):  # Add more extensions if needed
#             image_path = os.path.join(receipts_dir, filename)
#             print(f"Processing {filename}")
            
#             # Perform OCR
#             extracted_text = perform_ocr(image_path)
            
#             # Append to results
#             results.append({
#                 'File Name': filename,
#                 'Extracted Text': extracted_text
#             })

    
    
#     # Save the results to an Excel file
    df.to_excel(output_excel_path, index=False)

# Example usage:
receipts_directory = 'C:/Users/ArvindDS/Downloads/OCR_Receipts/images/img'  # Folder containing receipt images
output_excel_path = 'C:/Users/ArvindDS/Downloads/OCR_Receipts/Jupyter/OCR_Output/receipt_ocr_output_with_fonts10.xlsx'

# Get all system fonts
fonts_list = get_system_fonts()

# Process receipts and export to Excel with dynamic font handling
process_receipts_and_export_to_excel(receipts_directory, output_excel_path, fonts_list)
# process_receipts_to_excel(receipts_directory, output_excel_path)

# process_receipts_to_excel


# In[ ]:





# In[ ]:


import os
import pandas as pd
from paddleocr import PaddleOCR, draw_ocr
from PIL import Image, ImageFont
import matplotlib.pyplot as plt

# Initialize PaddleOCR
ocr_model = PaddleOCR(use_angle_cls=True, lang='en')

# Function to perform OCR and return extracted text
def perform_paddle_ocr(image_path):
    # Perform OCR on the image
    result = ocr_model.ocr(image_path)
    
    # Extract the text from OCR result
    extracted_text = ' '.join([elements[1][0] for elements in result[0]])
    
    return extracted_text

# Function to dynamically use Windows fonts and process multiple images
def process_images_to_excel(receipts_dir, output_excel):
    # List to store the OCR results
    results = []
    
    # Loop through all the files in the folder
    for filename in os.listdir(receipts_dir):
        if filename.endswith(".jpg") or filename.endswith(".png"):  # Add other formats as needed
            image_path = os.path.join(receipts_dir, filename)
            print(f"Processing {filename}...")

            # # Apply fonts dynamically and visualize (optional)
            # result = ocr_model.ocr(image_path)
            # apply_fonts_dynamically(image, result, fonts)
            
            # Perform OCR and get the extracted text
            extracted_text = perform_paddle_ocr(image_path)
            
            # Append the results to the list
            results.append({
                'File Name': filename,
                'Extracted Text': extracted_text
            })
    
    # Convert results to a DataFrame
    df = pd.DataFrame(results)
    
    # Save to Excel
    df.to_excel(output_excel, index=False)
    print(f"OCR results saved to {output_excel}")

# Specify paths
receipts_directory = 'C:/Users/ArvindDS/Downloads/OCR_Receipts/images/img'  # Update with your path
output_excel = 'C:/Users/ArvindDS/Downloads/OCR_Receipts/Jupyter/OCR_Output/receipt_text_output2.xlsx'

# Process the receipts and store OCR results in Excel
process_images_to_excel(receipts_directory, output_excel)


# In[ ]:





# In[ ]:




