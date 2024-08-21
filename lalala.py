import requests
from bs4 import BeautifulSoup
import os
import pandas as pd  # type: ignore
import re
import openpyxl
from openpyxl.styles import Alignment

def sanitize_filename(filename):
    return re.sub(r'[\\/*?:"<>|]', "_", filename)

def download_image(image_url, save_path):
    # Ensure the directory exists before saving the image
    os.makedirs(os.path.dirname(save_path), exist_ok=True)
    
    response = requests.get(image_url)
    if response.status_code == 200:
        with open(save_path, 'wb') as file:
            file.write(response.content)

def scrape_anphat_page(page_url, products_list, start_id):
    response = requests.get(page_url)
    soup = BeautifulSoup(response.text, 'html.parser')
    
    products = soup.find_all('div', class_='p-item')
    
    for product in products:
        name = product.find('a', class_='p-name').get_text(strip=True)
        price = product.find('span', class_='p-price').get_text(strip=True)
        product_detail_url = product.find('a', class_='p-name')['href']

        # Safe filename
        sanitized_name = sanitize_filename(name)

        # Get detailed images from the product page
        full_product_url = f"https://www.anphatpc.com.vn{product_detail_url}"
        response_detail = requests.get(full_product_url)
        soup_detail = BeautifulSoup(response_detail.text, 'html.parser')

        # Find the script containing the image URLs
        script_tags = soup_detail.find_all('script')
        image_urls = []

        for script in script_tags:
            if 'listImage' in script.text:
                # Extract image URLs using regex
                image_urls = re.findall(r'https://[^\s"]+\.jpg', script.text)
                break

        # Directory for this product's images
        image_directory = f"images/{start_id}"
        
        # Download and save the images to the specified directory
        for i, image_url in enumerate(image_urls[:3]):
            image_name = f"{image_directory}/{sanitized_name}_{i+1}.jpg"
            download_image(image_url, image_name)
        
        # Convert image URLs to the desired format and save product information to the list
        formatted_image_urls = '["' + '", "'.join(image_urls) + '"]'
        products_list.append({
            'ID': start_id,
            'Product Name': name,
            'Price': price,
            'Image URLs': formatted_image_urls
        })
        
        # Increment ID for the next product
        start_id += 1

    return start_id

def get_total_pages(soup):
    page_links = soup.select('div.paging a')
    pages = []
    for link in page_links:
        if link.get_text().strip().isdigit():
            pages.append(int(link.get_text().strip()))
    return max(pages) if pages else 1

def scrape_all_pages(start_url):
    products_list = []
    response = requests.get(start_url)
    soup = BeautifulSoup(response.text, 'html.parser')
    total_pages = get_total_pages(soup)

    # Create a new Excel workbook and worksheet
    workbook = openpyxl.Workbook()
    worksheet = workbook.active

    # Set column widths
    worksheet.column_dimensions['A'].width = 10
    worksheet.column_dimensions['B'].width = 50
    worksheet.column_dimensions['C'].width = 50
    worksheet.column_dimensions['D'].width = 10
    worksheet.column_dimensions['E'].width = 10
    worksheet.column_dimensions['F'].width = 10
    worksheet.column_dimensions['G'].width = 10
    worksheet.column_dimensions['H'].width = 10
    worksheet.column_dimensions['I'].width = 50
    worksheet.column_dimensions['J'].width = 10
    worksheet.column_dimensions['L'].width = 10
    worksheet.column_dimensions['M'].width = 20
    worksheet.column_dimensions['N'].width = 20

    # Write the header row
    worksheet['A1'] = 'ID'
    worksheet['B1'] = 'Product Name'
    worksheet['C1'] = 'Product Name'
    worksheet['D1'] = 'Category'
    worksheet['E1'] = 'Quantity'
    worksheet['F1'] = 'Quantity'
    worksheet['G1'] = 'Price'
    worksheet['H1'] = 'Discount'
    worksheet['I1'] = 'Image URLs'
    worksheet['J1'] = 'Post ID'
    worksheet['L1'] = 'Status'
    worksheet['M1'] = 'Created At'
    worksheet['N1'] = 'Updated At'

    # Align the header row
    header_row = worksheet[1]
    for cell in header_row:
        cell.alignment = Alignment(horizontal='center', vertical='center')

    row = 2
    id = 795
    category = 123
    post_id = 950
    current_time = '2024-06-21 13:09:37'

    for page in range(1, total_pages + 1):
        page_url = f"{start_url}?page={page}"
        print(f"Scraping page: {page_url}")
        id = scrape_anphat_page(page_url, products_list, id)

    for product_data in products_list:
        worksheet[f"A{row}"] = product_data['ID']
        worksheet[f"B{row}"] = product_data["Product Name"]
        worksheet[f"C{row}"] = product_data["Product Name"]
        worksheet[f"D{row}"] = category
        worksheet[f"E{row}"] = 0
        worksheet[f"F{row}"] = 0

        # Clean the price string: remove dots, remove currency symbol, and convert to integer
        cleaned_price = int(re.sub(r'\D', '', product_data["Price"]))
        worksheet[f"G{row}"] = cleaned_price

        worksheet[f"H{row}"] = 0
        worksheet[f"I{row}"] = product_data["Image URLs"]
        worksheet[f"J{row}"] = post_id
        worksheet[f"L{row}"] = 0
        worksheet[f"M{row}"] = current_time
        worksheet[f"N{row}"] = current_time
        post_id += 1
        row += 1

    # Save the Excel file
    workbook.save('anphat_products.xlsx')
    print("Product information has been saved to anphat_products.xlsx")

# Bắt đầu từ trang đầu tiên
start_url = 'https://www.anphatpc.com.vn/vo-may-tinh-case.html'
scrape_all_pages(start_url)
