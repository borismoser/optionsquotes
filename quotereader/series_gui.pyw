import PySimpleGUI as sg

layout = [[sg.Button('Carregar Dados', key='-LOAD-')],
          [sg.Button('Gerar arquivo', key='-TXT-')],
          [sg.Button('Exit', key='-EXIT-')]]

window = sg.Window('QuotesReader', layout)

while True:  # The Event Looop
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == '-EXIT-':
        break

    if event == '-LOAD-':
        # generate_datasets()
        sg.popup('Dados carregados.')
    elif event == '-TXT-':
        # txt_from_tabs()
        sg.popup('Arquivo gerado.')

window.close()
