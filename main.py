from os import path, listdir
from docx2pdf import convert
from tkinter import filedialog
import PySimpleGUI as sg

layout = [
    [sg.Text('Convert word to pdf')],
    [sg.Button('Convert')]
]

janela = sg.Window('Converter', layout)

while True:
    event, values = janela.read()
    if event == sg.WIN_CLOSED:
        break
    try:
        #função que irá pedir pasta
        pasta = filedialog.askopenfilename(filetypes=[('Documentos Word', '*.docx')])


        if pasta and path.isfile(pasta):
            convert(pasta)
            sg.popup('Converted with success')
        else:
            sg.popup('Cannot find the file')
    except FileNotFoundError:
        pass