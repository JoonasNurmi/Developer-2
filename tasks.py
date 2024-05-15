from robocorp.tasks import task
from robocorp import browser
from RPA.HTTP import HTTP
from RPA.Tables import Tables
from RPA.PDF import PDF
from RPA.Archive import Archive

@task
def order_robots_from_RobotSpareBin():
    """
    Orders robots from RobotSpareBin Industries Inc.
    Saves the order HTML receipt as a PDF file.
    Saves the screenshot of the ordered robot.
    Embeds the screenshot of the robot to the PDF receipt.
    Creates ZIP archive of the receipts and the images.
    """
    #browser.configure(slowmo=300, )
    open_robot_order_website("https://robotsparebinindustries.com/#/robot-order")
    orders = get_orders()
    page = browser.page()
    for row in orders:
        close_annoying_modal()
        fill_the_form(row)
        file_path = store_receipt_as_pdf(row['Order number'])
        screenshot_path = screenshot_robot(row['Order number'])
        embed_screenshot_to_receipt(screenshot_path, file_path)
        page.click("#order-another")
    archive_receipts()


def open_robot_order_website(url):
    """Navigates to the given URL"""
    browser.goto(url)
    browser.browser()
    

def get_orders():
    """Downloads excel file from the given URL and sets it to a table variable"""
    http = HTTP()
    table = Tables()
    http.download(url="https://robotsparebinindustries.com/orders.csv", overwrite=True)
    orders = table.read_table_from_csv(path="orders.csv", header=True)
    return orders


def close_annoying_modal():
    page = browser.page()
    page.click("button:text('OK')")


def fill_the_form(row):
    """Fill the form using selectors by ID, attribute and text"""
    page = browser.page()
    head = str(row['Head'])
    body = "#id-body-" + str(row['Body'])   #Dynamcally create the HTML ID for each body type
    legs = str(row['Legs'])
    address = row['Address']
    page.select_option("#head", head)   #ID selector
    page.click(body)
    page.fill("input[placeholder='Enter the part number for the legs']", legs)  #Attribute selector
    page.fill("#address", address)
    page.click("button:text('Preview')")    #Text selector
    page.click("#order")
    
    """If there is an error ordering, try again 5 times before quitting"""
    retry_count = 0
    while page.is_visible("#order-another", timeout=1.0) == False:
        retry_count += 1
        if retry_count > 5:
            return
        page.click("#order")


def store_receipt_as_pdf(order_number):
    """Saves the receipt page as pdf with the given order number"""
    page = browser.page()
    receipt_html = page.locator("#receipt").inner_html()
    pdf = PDF()
    file_path = "output/robot_receipt_" + order_number + ".pdf"
    pdf.html_to_pdf(receipt_html, file_path)
    return file_path


def screenshot_robot(order_number):
    """Takes a screenshot of the whole page and saves it with the given order number as pdf"""
    page = browser.page()
    screenshot_path = "output/sales_summary_" + order_number + ".png"
    page.screenshot(path=screenshot_path)
    return screenshot_path


def embed_screenshot_to_receipt(screenshot, pdf_file):
    """Adds a list of files into a PDF. The list does not need to contain more than one item."""
    pdf = PDF()
    list_of_files = [
        #pdf_file,
        screenshot,
        ]
    pdf.add_files_to_pdf(list_of_files, pdf_file, append=True)


def archive_receipts():
    """Creates a zip file of the output folder while only including the receipts"""
    archive = Archive()
    archive.archive_folder_with_zip("output", "output/archive.zip", include="*robot_receipt*")