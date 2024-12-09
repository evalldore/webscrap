import requests
import pandas
import sys
import threading
import tkinter as tk
from tkinter import filedialog
import subprocess
from bs4 import BeautifulSoup

class page():
	def __init__(self, link):
		self.link = link
		self.thread = None
		self.paragraphs = []

headers = {
	'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/107.0.0.0 Safari/537.36'
}

# Création de la fenêtre principale
root = tk.Tk()
root.title("Exécution du script")

# Variables
excel_path = tk.StringVar()
sheet_name = tk.StringVar()
output_path = tk.StringVar()
status = tk.StringVar()

def select_excel_file():
    excel_path.set(filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")]))

def select_output_file():
    output_path.set(filedialog.asksaveasfilename(defaultextension=".txt"))

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

def extract():

	if not excel_path.get() or not sheet_name.get() or not output_path.get():
		status.set("Veuillez remplir tous les champs")
		return
 
	excel_file = pandas.ExcelFile(excel_path.get())
	data_frame = pandas.read_excel(excel_file, sheet_name.get(), header=None)
 
	with open(output_path.get(), "w", encoding="utf-8") as output_file:
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

def main():
	tk.Label(root, text="Fichier Excel:").grid(row=0, column=0, sticky="e")
	tk.Entry(root, textvariable=excel_path, width=50).grid(row=0, column=1)
	tk.Button(root, text="Parcourir", command=select_excel_file).grid(row=0, column=2)

	tk.Label(root, text="Nom de la feuille:").grid(row=1, column=0, sticky="e")
	tk.Entry(root, textvariable=sheet_name).grid(row=1, column=1)

	tk.Label(root, text="Fichier de sortie:").grid(row=2, column=0, sticky="e")
	tk.Entry(root, textvariable=output_path, width=50).grid(row=2, column=1)
	tk.Button(root, text="Parcourir", command=select_output_file).grid(row=2, column=2)

	tk.Button(root, text="Exécuter", command=extract).grid(row=3, column=1)

	tk.Label(root, textvariable=status).grid(row=4, column=1)

	root.mainloop()


if __name__ == "__main__":
    main()