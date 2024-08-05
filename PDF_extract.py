
import os
import re
import sys

import tkinter as tk
from tkinter import filedialog
from tkinter import PhotoImage
from tkinter import Canvas
from tkinter import ttk

from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter


def split_pdf_by_employee(input_pdf_path, output_folder, status_label, progressbar):

    # Verificando se o arquivo de saída existe,
    status_label.config(text='Verificando a pasta de saída...')
    status_label.place(x=200, y=400, width=400)
    progressbar['value'] = 5
    root.update_idletasks()
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Lendo o documento PDF
    status_label.config(text='Lendo arquivo PDF...')
    status_label.place(x=200, y=400, width=400)
    progressbar['value'] = 10
    root.update_idletasks()
    reader = PdfReader(input_pdf_path)
    total_pages = len(reader.pages)

    # Armazenando as páginas por empregado
    employees = {}

    # Regex para encontrar os nomes do empregados
    name_pattern = re.compile(r"Empregado\s+\d{6}\s+([a-zA-Z ]+)")

    for page_num in range(total_pages):
        status_label.config(text=f'Processando página {
                            page_num+1} de {total_pages}...')
        status_label.place(x=200, y=400)
        progressbar['value'] = 10 + (page_num / total_pages) * 80
        root.update_idletasks()
        page = reader.pages[page_num]
        text = page.extract_text()

        # Procura pelo nome do empregado
        match = name_pattern.search(text)
        if match:
            employee_name = match.group(1)
            if employee_name not in employees:
                employees[employee_name] = []
            employees[employee_name].append(page)

    # Criando um novo PDf para cada empregado
    for employee_name, pages in employees.items():
        status_label.config(text='Criando arquivos PDF para cada empregado...')
        status_label.place(x=200, y=400)
        progressbar['value'] = 95
        root.update_idletasks()
        output_pdf_path = os.path.join(output_folder, f'{employee_name}.pdf')
        writer = PdfWriter()

        for page in pages:
            writer.add_page(page)

        # Salvando o PDF no arquivo de saída
        with open(output_pdf_path, 'wb') as output_pdf_file:
            writer.write(output_pdf_file)

    status_label.config(text='Arquivos extraídos com sucesso')
    progressbar['value'] = 100
    root.update_idletasks()


def select_file():
    global input_pdf_path
    input_pdf_path = filedialog.askopenfilename(title='Selecione o arquivo')
    file_entry.delete(0, tk.END)
    file_entry.insert(0, input_pdf_path)


def select_folder():
    global output_folder
    output_folder = filedialog.askdirectory(title='Selecione a pasta')
    output_entry.delete(0, tk.END)
    output_entry.insert(0, output_folder)


def extract_receipts():
    global output_folder
    output_folder = output_entry.get()
    try:
        split_pdf_by_employee(input_pdf_path, output_folder,
                              status_label, progressbar)
    except NameError:
        status_label.config(text='Por favor, escolha um arquivo!', font='Verdana')
        status_label.place(x=220, y=110)
    except FileNotFoundError:
        status_label.config(
            text='Por favor, escolha uma pasta para salvar os arquivos!', font='Verdana')
    else:
        status_label.config(text='Arquivos extraídos com sucesso!', font='Verdana')
        status_label.place(x=200, y=400)


root = tk.Tk()
root.iconbitmap('logo_conectrom.ico')
root.title("Extrator de Recibos")
root.geometry('780x439')
root.resizable(False, False)
root.wm_attributes('-transparent')
bg = PhotoImage(file='2.png')
bg_image = tk.Label(root, image=bg)
bg_image.pack()


file_label = tk.Label(root, text="Selecione o arquivo PDF:", font='Verdana', bg='#363636', fg='#fff')
file_label.place(x=25, y=80, height=30)
file_entry = tk.Entry(root)
file_entry.place(x=236.2, y=80, width=403, height=30)
file_button = tk.Button(root, text="Procurar", width=10, font='Verdana', fg='#fff', command=select_file)
file_button.place(x=640, y=80)
file_button.config(background='#D2691E')

output_label = tk.Label(root, text="Local para salvar arquivo:", font='Verdana', bg='#363636', fg='#fff')
output_label.place(x=25, y=125, height=30)
output_entry = tk.Entry(root)
output_entry.place(x=239, y=125, width=400, height=30)
output_button = tk.Button(root, text='Selecionar', width=10, font='Verdana', fg='#fff', command=select_folder)
output_button.place(x=640, y=125)
output_button.config(background='#D2691E')


extract_button = tk.Button(
    root, text="Extrair arquivo", font='Verdana', fg='#fff', command=extract_receipts)
extract_button.place(x=330, y=300, width=130, height=50)
extract_button.config(background='#D2691E')

status_label = tk.Label(
    root, text="Escolha o arquivo e a pasta para salvar o PDF", font='Verdana', bg='#363636', fg='#fff')
status_label.place(x=220, y=20)


progressbar = ttk.Progressbar(root, orient='horizontal')

root.mainloop()
