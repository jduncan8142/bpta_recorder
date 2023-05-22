import pystray
from PIL import Image
import PySimpleGUI as sg
import win32com.client
import json

#Constants 
TAB = "...."
THEME = "gray gray gray"
IMAGE = Image.open("src\sap_screen.png")
APP_NAME = "BPTA"
APP_TOOLTIP = "BPTA Recorder"

sg.theme(THEME)


class SAPEventHandler:
    def Change(self, session, component, function_code) -> None:
        print("Called on_change event handler")


class SAPWindows:
    def __init__(self) -> None:
        self.gui = win32com.client.GetObject("SAPGUI")
        self.app = self.gui.GetScriptingEngine
        self.connections: list = self.app.connections
        self.sessions: list = None
        self.windows: list = []
        self.connection = None
        self.session = None
        self.main_window = None
        self.user_area = None
        self.window_selection: str = None
        self.counter: int = 0
        self.elements: list[dict] = []
        self.recording: list[dict] = []
        self.record_flag: bool = False
        if len(self.connections) != 0:
            for conn in self.connections:
                self.sessions = conn.sessions
                if len(self.sessions) != 0:
                    for ses in self.sessions:
                        __windows = ses.children
                        for win in __windows:
                            self.windows.append(win)
        self.event = None

    def get_windows(self) -> None:
        layout = [
            [sg.Text('Select the SAP window to record.')], 
            [sg.Listbox(
                [f"{w.text} | {w.Id}" for w in self.windows], 
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
        try:
            self.selection = str(values["-window-"]).split(" | ")[1]
            __parts = self.selection.split("/")
            self.connection = self.connections[int(__parts[2][-2])]
            self.session = self.sessions[int(__parts[3][-2])]
            self.main_window = self.windows[int(__parts[4][-4])]
            self.user_area = self.session.FindById(f"{self.main_window.Id}/usr")
            try:
                self.session.Record = True
            except Exception as e1:
                print(f"ERROR<get_windows::Record> {e1}")
            try:
                self.session.TestToolMode = 1
            except Exception as e1:
                print(f"ERROR<get_windows::TestToolMode> {e1}")
            # try:
            #     self.session.SuppressBackendPopups = True
            # except Exception as e1:
            #     print(f"ERROR<get_windows::SuppressBackendPopups> {e1}")
        except IndexError:
            return
        except Exception as e:
            print(f"ERROR<get_windows> {e}")
    
    def get_session_info(self) -> dict:
        try:
            if not self.session.Busy:
                s_info = self.session.Info
                return {
                    "ApplicationServer": s_info.ApplicationServer, 
                    "Client": s_info.Client, 
                    "Codepage": s_info.Codepage, 
                    "Flushes": s_info.Flushes, 
                    "Group": s_info.Group, 
                    "GuiCodepage": s_info.GuiCodepage, 
                    "I18NMode": s_info.I18NMode, 
                    "InterpretationTime": s_info.InterpretationTime, 
                    "IsLowSpeedConnection": s_info.IsLowSpeedConnection, 
                    "Language": s_info.Language, 
                    "MessageServer": s_info.MessageServer, 
                    "Program": s_info.Program, 
                    "ResponseTime": s_info.ResponseTime, 
                    "RoundTrips": s_info.RoundTrips, 
                    "ScreenNumber": s_info.ScreenNumber, 
                    "ScriptingModeReadOnly": s_info.ScriptingModeReadOnly, 
                    "ScriptingModeRecordingDisabled": s_info.ScriptingModeRecordingDisabled, 
                    "SessionNumber": s_info.SessionNumber, 
                    "SystemName": s_info.SystemName, 
                    "SystemNumber": s_info.SystemNumber, 
                    "SystemSessionId": s_info.SystemSessionId, 
                    "Transaction": s_info.Transaction, 
                    "UI_GUIDELINE": s_info.UI_GUIDELINE, 
                    "User": s_info.User, 
                }
        except Exception as e:
            print(f"ERROR<get_session_info> {e}")
            return {}

    def parse_element_id(self, obj: dict, lvl: int = 0) -> None:
        __lvl = lvl
        if isinstance(obj, dict):
            if 'properties' in obj.keys():
                self.elements.append({'lvl': __lvl, 'element': obj.get('properties')['Id']})
            if 'children' in obj.keys():
                for i in obj.get('children'):
                    self.parse_element_id(obj=i, lvl=__lvl + 1)
        elif isinstance(obj, list):
            for i in obj.get('children'):
                self.parse_element_id(obj=i, lvl=__lvl + 1)
        else:
            self.elements.append({'lvl': __lvl, 'element': obj})

    def dump_elements(self) -> None:
        try:
            obj_tree = json.loads(self.session.GetObjectTree(self.main_window.Id))
            self.elements.clear()
            self.parse_element_id(obj_tree)
        except Exception as e:
            pass

    def record_gui(self) -> None:
        self.counter = 1
        last_rec: list = None
        self.recording.clear()
        while self.record_flag:
            if self.main_window.Modified:
                __rec = []
                self.dump_elements()
                for i in self.elements:
                    try:
                        __rec.append({'lvl': i['lvl'], 'element': i['element'], 'value': self.session.FindById(i['element']).text})
                    except Exception as e:
                        pass
                if last_rec is None:
                    last_rec = __rec
                    self.recording.append({'frame': self.counter, 'data': __rec})
                    self.counter += 1
                elif last_rec is not None and last_rec != __rec:
                    last_rec = __rec
                    self.recording.append({'frame': self.counter, 'data': __rec})
                    self.counter += 1
                else:
                    pass


class App:
    def __init__(self) -> None:
        self.sap = SAPWindows()
        self.icon = pystray.Icon(
            APP_NAME, 
            IMAGE, 
            APP_TOOLTIP,
            menu=pystray.Menu(
                pystray.MenuItem("Select SAP session", self.after_click),
                pystray.MenuItem("Dump GUI elements", self.after_click),
                pystray.MenuItem("Start recording", self.after_click),
                pystray.MenuItem("Stop recording", self.after_click),
                pystray.MenuItem("Save recording", self.after_click),
                pystray.MenuItem("Dump recording", self.after_click),
                pystray.MenuItem("Exit", self.after_click),
            )
        )

    def run(self) -> None:
        self.icon.run()

    def after_click(self, icon, query):
        if str(query) == "Select SAP session":
            self.sap.get_windows()
        elif str(query) == "Dump GUI elements":
            self.sap.dump_elements()
            for i in self.sap.elements:
                try:
                    print(str(TAB * i['lvl']) + i['element'] + " => " + self.sap.session.FindById(i['element']).text)
                except Exception as e:
                    pass
            # try:
            #     result = self.sap.session.FocusChanged
            #     print(f"Result: {result} -- Type: {type(result)}")
            # except Exception as e:
            #     print(f"ERROR<after_click::FocusChanged> {e}")
        elif str(query) == "Start recording":
            self.sap.record_flag = True
            print("SAP GUI recording started...")
            self.sap.record_gui()
            # self.sap.session.TestToolMode = 1
            # self.sap.session.Record = True
            # self.sap.session.RecordFile = "mu_auto_recording.txt"
        elif str(query) == "Stop recording":
            self.sap.record_flag = False
            self.sap.session.Record = False
            print("SAP GUI recording stopped.")
        elif str(query) == "Save recording":
            self.sap.record_flag = False
            print("Saving recording to file...")
            with open("my_recording.txt", "w") as f:
                for i in self.sap.recording:
                    f.write(f"Frame {i['frame']}")
                    f.write("\n")
                    for j in i['data']:
                        f.write(f"Lvl: {j['lvl']} | Element: {j['element']} | Value: {j['value']}")
                        f.write("\n")
        elif str(query) == "Dump recording":
            self.sap.record_flag = False
            print("Dumping recording...")
            for i in self.sap.recording:
                print(f"Frame {i['frame']}")
                for j in i['data']:
                    print(f"Lvl: {j['lvl']} | Element: {j['element']} | Value: {j['value']}")
        elif str(query) == "Exit":
            try:
                self.icon.stop()
                exit()
            except SystemExit:
                pass
            except Exception as e:
                print(f"ERROR<after_click> {e}")


if __name__ == "__main__":
    app = App()
    app.run()
