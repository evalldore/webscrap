import requests
import pandas
import sys
import threading
from bs4 import BeautifulSoup

class page():
	def __init__(self, link):
		self.link = link
		self.thread = None
		self.paragraphs = []

headers = {
	'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/107.0.0.0 Safari/537.36'
}

def page_process(page):
	print("loading page: " + page.link)
	response = requests.get(page.link, headers=headers)
	if (response.status_code == 200):
		soup = BeautifulSoup(response.text, 'html.parser')
		paragraphs_list = soup.find_all('p')
		for paragraph in paragraphs_list:
			text = paragraph.get_text()
			if (text != None and len(text) > 0):
				page.paragraphs.append(text)
def main():
	if (len(sys.argv) < 4):
		print("Invalid amount of arguments [Excel Path, Sheet Name, Output Path]")
		return

	try:
		excel_file = pandas.ExcelFile(sys.argv[1])
		data_frame = pandas.read_excel(excel_file, sys.argv[2])
	except FileNotFoundError:
		print("Excel file not found: " + sys.argv[1])
		return
	except ValueError:
		print("Excel sheet not found: " + sys.argv[2])
		return

	with open(sys.argv[3], "w", encoding="utf-8") as output_file:
		
		pages_array = []
     
		for row in data_frame.itertuples(index=False):
			link = row._0
			try:
				page_instance = page(link)
				page_instance.thread = threading.Thread(target = page_process, args=[page_instance])
				page_instance.thread.start()
				pages_array.append(page_instance)
			except Exception as e:
				print(e + "error processing link: " + link)
		for page_element in pages_array:
			page_element.thread.join()
			for text in page_element.paragraphs:
				output_file.write('• ' + text + '\n')
		output_file.close()


if __name__ == "__main__":
    main()