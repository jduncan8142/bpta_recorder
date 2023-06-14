import PySimpleGUI as sg
from utils import THEME

sg.theme(THEME)


def select_sap_window(windows: list) -> str:
    layout = [
        [sg.Text('Select the SAP window to record.')], 
        [sg.Listbox(
            [f"{w.text} | {w.id}" for w in windows], 
            key='-window-', 
            s=(60,20), 
            select_mode=sg.LISTBOX_SELECT_MODE_EXTENDED)
        ], 
        [sg.Button('Ok'), sg.Button('Cancel')]
    ]
    window = sg.Window('SAP GUI Window', layout, finalize=True)
    window['-window-'].bind("<Return>", "_Enter")
    event, values = window.read(close=True)
    if event == sg.WIN_CLOSED or event == 'Cancel':
        return
    return str(values["-window-"])


def recording_popup(self) -> sg.Window:
    layout = [
        [sg.Text('Recording - Please wait...')],
    ]
    window = sg.Window('Recording - Please wait...', layout, size=(350,10), finalize=True)
    return window
