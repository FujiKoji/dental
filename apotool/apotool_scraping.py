from apotool_method import Scraping

id = ""
pw = ""
excel_path = ""

scraping = Scraping(id,pw,excel_path)

page_num = 10
urls, num = scraping.get_url(page_num)

start = 0
end = 10
scraping.get_treatmentdata(0,10)