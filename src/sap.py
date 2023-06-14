from deepdiff import DeepDiff
import hashlib
import json
import win32com.client
import threading
from typing import Optional
from gui import select_sap_window


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
        self.program: str = None
        self.screen_number: int = None
        self.transaction: str = None
        self.titl: str = None
        self.current_element = None
        self.obj_tree_hash = None
        self.obj_tree = None
        self.last_diff = None
        self.counter: int = 0
        self.elements: list[dict] = []
        self.recording: list[dict] = []
        self.record_flag: bool = False
        self.frame_count: int = 0
        if len(self.connections) != 0:
            for conn in self.connections:
                self.sessions = conn.sessions
                if len(self.sessions) != 0:
                    for ses in self.sessions:
                        __windows = ses.children
                        for win in __windows:
                            self.windows.append(win)
    
    def visualize_element(self, element, interval: Optional[float]=3.0) -> None:
        self.current_element = element
        self.enable_visualize()
        timer = threading.Timer(interval, self.disable_visualize)
        timer.start()

    def enable_visualize(self) -> None:
        try:
            _ = self.current_element.Visualize(on=True)
        except:
            pass
    
    def disable_visualize(self) -> None:
        try:
            _ = self.current_element.Visualize(on=False)
        except:
            pass
    
    def screenshot(self, filename: str, element_cords: Optional[tuple[int]] = None) -> str:
        if element_cords:
            if len(element_cords) == 4:
                return self.main_window.HardCopy(
                    filename,
                    element_cords[0],
                    element_cords[1],
                    element_cords[2],
                    element_cords[3]
                )
        else:
            return self.main_window.HardCopy(filename)
    
    def capture_screen(self, event) -> None:
        if self.session:
            if self.screen_diff():
                try:
                    self.screenshot(f"screenshots\\frame{self.frame_count}")
                    self.dump_elements()
                    self.recording.append({'frame': self.frame_count, 'data': self.elements})
                    self.frame_count += 1
                except Exception as e:
                    print(f"ERROR<capture_screen::LockSessionUI> {e}")

    def window_selection(self) -> None:
        __selection: str = select_sap_window(windows=self.windows)
        try:
            __selection_id: str = __selection.split(" | ")[1]
            __parts: list = __selection_id.split("/")
            self.connection = self.connections[int(__parts[2][-2])]
            self.session = self.sessions[int(__parts[3][-2])]
            self.main_window = self.windows[int(__parts[4][-4])]
            self.user_area = self.session.FindById(f"{self.main_window.Id}/usr")
            try:
                self.session.Record = True
            except Exception as e1:
                print(f"ERROR<window_selection::Record> {e1}")
            try:
                self.session.TestToolMode = 1
            except Exception as e1:
                print(f"ERROR<window_selection::TestToolMode> {e1}")
            try:
                _ = self.updated_screen()
            except Exception as e1:
                print(f"ERROR<window_selection::updated_screen> {e1}")
        except IndexError:
            return
        except Exception as e:
            print(f"ERROR<window_selection> {e}")
    
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
                self.elements.append({'lvl': __lvl, 'element': obj.get('properties')})
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
            obj_tree = json.loads(self.obj_tree)
            self.elements.clear()
            self.parse_element_id(obj_tree)
        except Exception as e:
            print(f"ERROR<dump_elements> {e}")
    
    def updated_screen(self) -> bool:
        __captured = False
        while not __captured:
            try:
                __obj_tree_hash = hashlib.md5(
                    self.session.GetObjectTree(
                        self.main_window.Id, 
                        [
                            "Id", 
                            "Text", 
                            "Type", 
                            "Name", 
                            "ScreenLeft", 
                            "ScreenTop", 
                            "Left", 
                            "Top", 
                            "Height", 
                            "Width", 
                            "Changeable",
                            "ContainerType",
                            "ToolTip",
                            "DefaultToolTip",
                            "IconName",
                            "Key",
                            "Handle"
                        ]
                    ).encode()
                ).hexdigest()
                __captured = True
                if __obj_tree_hash != self.obj_tree_hash:
                    self.obj_tree_hash = __obj_tree_hash
                    return True
            except:
                continue
        return False

    def get_obj_tree(self) -> str:
        __captured = False
        __count = 0
        while not __captured and __count < 100:
            __count += 1
            try:
                return self.session.GetObjectTree(
                    self.main_window.Id, 
                    [
                        "Id", 
                        "Text", 
                        "Type", 
                        "Name", 
                        "ScreenLeft", 
                        "ScreenTop", 
                        "Left", 
                        "Top", 
                        "Height", 
                        "Width", 
                        "Changeable",
                        "ContainerType",
                        "ToolTip",
                        "DefaultToolTip",
                        "IconName",
                        "Key",
                        "Handle"
                    ]
                )
            except:
                continue
        return None
    
    def screen_diff(self) -> bool:
        __obj_tree = self.get_obj_tree()
        if __obj_tree is not None:
            ddiff = DeepDiff(self.obj_tree, __obj_tree)
            if ddiff != {}:
                self.obj_tree = __obj_tree
                self.last_diff = ddiff
                return True
        return False

    def session_close(self) -> None:
        try:
            self.session.UnlockSessionUI()
        except:
            pass
        try:
            self.session.Record = False
        except:
            pass
        try:
            self.session.TestToolMode = 0
        except:
            pass
