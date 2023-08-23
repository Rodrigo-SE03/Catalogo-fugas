import docx2txt
import PySimpleGUI as sg

layout = [[sg.Text('Selecione o arquivo com as tabelas')],
          [sg.Text('Arquivo', size=(8, 1)), sg.Input(), sg.FileBrowse()],
          [sg.Text('Selecione a pasta onde serão salvas as fotos')],
          [sg.Text('Pasta', size=(8, 1)), sg.Input(), sg.FolderBrowse()],
          [sg.Submit(), sg.Cancel()]]

window = sg.Window('Salvamento de imagens', layout)
event, values = window.read()

arquivo = values[0]
pasta = values[1]

text = docx2txt.process(arquivo,pasta)