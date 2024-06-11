from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
import time
from RPA.PDF import PDF
import os
import zipfile

    
def get_orders(): 
    """Read data from CSV and fill in the sales forms"""
    table = Tables()
    orders = table.read_table_from_csv("orders.csv", columns=["Order number", "Head", "Body","Legs","Address"], header=True) 
    return orders

def download_orders_csv_file():
    """Downloads CSV file from the given URL"""
    http = HTTP()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)
    print("Successfully downloaded the CSV file!")

def open_robot_order_website():
    """Navigates to the given URL"""
    browser.goto("https://robotsparebinindustries.com/#/robot-order")
    print("Successfully navigated to order page") 

def close_annoying_modal():
    """Closing the modal"""
    page = browser.page()
    page.locator("//button[@class='btn btn-dark']").click()
    print("Successfully closed the modal!")    

def click_order_button():
    """Clicking the Order button"""
    page = browser.page()
    page.locator("//button[@id='order']").click()
    print("Successfully clicked the Order button!")

def click_preview_button():
    """Clicking the Preview button"""
    page = browser.page()
    page.locator("//button[@id='preview']").click()
    print("Successfully clicked the Preview button!")

def order_another_robot():
    """Clicking the Order Another button"""
    page = browser.page()
    page.locator("//button[@id='order-another']").click()
    print("Order another button has been clicked!")

def take_a_preview_screenshot(fileName):
    """Take a screenshot of the preview"""
    robot_preview_locator = "//div[@id='robot-preview-image']"
    page = browser.page()
    page.locator(robot_preview_locator).screenshot(path=f"output/robot_screenshots/{fileName}.png")
    file = f"output/robot_screenshots/{fileName}.png"
    print(f"Screenshot {fileName} has been taken!")
    return file

def export_receipt_as_pdf(fileName):
    """Export the receipt to a pdf file"""
    page = browser.page()
    sales_results_html = page.locator("//*[@id='receipt']").inner_html()
    pdf = PDF()
    pdf.html_to_pdf(sales_results_html, f"output/receipts/{fileName}.pdf")
    file = f"output/receipts/{fileName}.pdf"
    print(f"PDF {fileName} has been export!")
    return file

def embed_preview_screenshot_to_receipt(pdf_file, screenshot, order):
    """Embedding Robot Preview to Receipt """
    pdf = PDF()
    pdf.add_watermark_image_to_pdf(
        image_path=screenshot,
        source_path=pdf_file,
        output_path=f"output/final_receipts/final_receipt_{order}.pdf")   
    print("Robot screenshot has been appended to PDF!") 

def archive_receipts():
    """Archiving Receipts"""
    pdf_folder = 'output/final_receipts'
    output_zip_path = 'output/archive.zip'

    if not os.path.exists(pdf_folder):
        print(f"Folder '{pdf_folder}' does not exist.")
        return
    
    """List all PDF files in the folder"""
    pdf_files = [f for f in os.listdir(pdf_folder) if f.endswith('.pdf')]

    if not pdf_files:
        print(f"No PDF files found in folder '{pdf_folder}'.")
        return

    with zipfile.ZipFile(output_zip_path, 'w') as zipf:
        for pdf_file in pdf_files:
            pdf_path = os.path.join(pdf_folder, pdf_file)
            zipf.write(pdf_path, os.path.basename(pdf_path))
            
    print(f"Robot Receipts have been archived to {output_zip_path}")


def fill_and_submit_order_form():
    orders = get_orders()
    page = browser.page()
    for rows in orders:
        print(f'Processing Order Number {rows["Order number"]}')
        print(rows)

        """Fills in the sales data and click the 'Submit' button"""
        head = rows["Head"]
        body = rows["Body"]
        legs = rows["Legs"]
        address = rows["Address"]

        '''Head'''
        page.locator("//select[@id='head']").select_option(value=head)
        '''Body'''
        body_locator = f"//input[@id='id-body-{body}']"
        page.locator(body_locator).click()
        '''Legs'''
        page.locator("//input[@placeholder='Enter the part number for the legs']").fill(value=legs)
        '''Shipping Address'''
        page.locator("//input[@id='address']").fill(value=address)

        '''Preview Order'''
        click_preview_button()

        time.sleep(2)
        
        '''Save Order Preview'''
        screenshot = take_a_preview_screenshot(f'robot_order_{rows["Order number"]}')

        '''Submit Order'''
        click_order_button()

        time.sleep(3)

        '''Check for Receipt'''
        receipt = page.locator("//h3[contains(text(),'Receipt')]")
        
        '''Adding a retry scope'''
        retry = 0
        while retry <= 4:
            if receipt.is_visible():
                print("No Error Detected!")
                break
            else:
                error = page.locator("//*[contains(@class,'alert-danger')]").text_content()
                print(f"{error} Error Detected! Clicking Order button again.") 
                click_order_button()
                retry = retry + 1
                print("Order has been resubmitted!")

        '''Creating PDF file from the receipt'''        
        pdf_file = export_receipt_as_pdf(f'receipt_order_{rows["Order number"]}')

        '''Embedding the robot preview to the receipt pdf file'''
        embed_preview_screenshot_to_receipt(pdf_file, screenshot, rows["Order number"])
        
        '''2 sec delay before ordering again'''
        time.sleep(2)
        order_another_robot()
        close_annoying_modal()
            
@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    open_robot_order_website()
    download_orders_csv_file()
    get_orders()
    close_annoying_modal()
    fill_and_submit_order_form()
    archive_receipts()