from selenium import webdriver
from bs4 import BeautifulSoup
import openpyxl
import time


# Controls
Time_For_Page_To_Load = 1
Display_Website = False



# use headless Edge
options = webdriver.EdgeOptions()
options.use_chromium = True
if not Display_Website:
    options.add_argument('headless')
# improve performance
options.add_argument('disable-gpu')
options.add_argument('disable-extensions')
options.add_argument('disable-dev-shm-usage')
options.add_argument('no-sandbox')
options.add_argument('disable-setuid-sandbox')
options.add_argument('disable-infobars')
driver = webdriver.Edge(options=options)

# Navigate to the webpage
print("Started")
print("--" * 40)

print("Loading webpage")
driver.get("https://ftuapps.dev/?1")
print("Webpage loaded")

print("--" * 40)

# Scroll down the webpage and wait for 5 seconds
driver.execute_script("\n"
                      "const resize_ob = new ResizeObserver(function(entries) {\n"
                      "	let rect = entries[0].contentRect;\n"
                      "\n"
                      "	let width = rect.width;\n"
                      "	let height = rect.height;\n"
                      "\n"
                      "	console.log('Current Width : ' + width);\n"
                      "	console.log('Current Height : ' + height);\n"
                      "window.scrollTo(0, document.body.scrollHeight - 700);"

                      "});\n"
                      "resize_ob.observe(document.body);\n"
                      "setInterval(()=>{window.scrollTo(0, document.body.scrollHeight - 800);}, 600);"
                      "setInterval(()=>{window.scrollTo(0, document.body.scrollHeight - 600);}, 560);"
                      "setInterval(()=>{window.scrollTo(0, document.body.scrollHeight - 400);}, 503);"
                      "setInterval(()=>{window.scrollTo(0, document.body.scrollHeight - 1000);}, 200);"
                      "setInterval(()=>{window.scrollTo(0, document.body.scrollHeight - 100);}, 310);"
                      "")

print(f"please wait for {Time_For_Page_To_Load} seconds")
for i in range(1, Time_For_Page_To_Load + 1):
    time.sleep(1)
    print(f"{i} ", end="")

# Get the HTML content of the webpage

print("--" * 40)

print("Fetching the HTML content of the webpage")
html = driver.page_source
soup = BeautifulSoup(html, "html.parser")

# Find all the h2 elements on the page
h2_elements = soup.find_all("h2")

print("Fetched All Data")

print("--" * 40)

# Create a new workbook and sheet
workbook = openpyxl.Workbook()
sheet = workbook.active

# Add a header row to the sheet
sheet.append(['Text', 'Link', 'Torrent'])

print("--" * 40)

print("Parsing The Content")

print("--" * 40)

print("Number of h2 elements found: ", len(h2_elements))


# Iterate through the h2 elements and add a row for each element
def ExplorePageAndExtractLinkThatContainsTorrent(link):
    driver.get(link)
    htmlElement = driver.page_source
    soupEle = BeautifulSoup(htmlElement, "html.parser")
    a_elements = soupEle.find_all('a', href=True)
    for a in a_elements:
        if "torrent" in a['href']:
            print("Found Torrent Link")
            return a['href']
        else:
            continue
    print("No Torrent Link Found")
    return ""


linksNumber = 0
for h2 in h2_elements:
    # Get the text content and link of the h2 element
    linksNumber += 1
    text = h2.text
    a_element = h2.find('a', href=True)
    if a_element is not None:
        link = a_element.get('href')
        print(f"Exploring #{linksNumber} - ", end=" ")
        torrent = ExplorePageAndExtractLinkThatContainsTorrent(link)
    else:
        link = ''
        torrent = ''
    # Add a row to the sheet
    sheet.append([text, link, torrent])

print("Parsing Completed")

print("--" * 40)

print("Saving the Excel File")
# Save the workbook to a file
workbook.save('Software-List.xlsx')
print("Saved the Excel File")

print("--" * 40)

print("Closing the Webdriver")

# Close the webdriver
driver.close()

print("Closed the Webdriver")

print("--" * 40)

print("Finished")
