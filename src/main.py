from deepdiff import DeepDiff
from pynput.keyboard import Key, KeyCode
from pynput.mouse import Button
import pystray
from typing import Optional
from input_controller import KMInput, MouseEvent, KeyboardEvent, MouseEventType, KeyEventType
from sap import SAPWindows
from utils import TAB, IMAGE, APP_NAME, APP_TOOLTIP


class App:
    def __init__(self) -> None:
        self.sap = SAPWindows()
        self.kmi: KMInput = KMInput(on_move=False, on_scroll=False, on_press=False)
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
    
    def start_recording(self) -> None:
        self.kmi.start()
        self.sap.record_flag = True
        self.sap.recording.clear()
        try:
            self.sap.session.Record = True
        except Exception as e:
            print(f"ERROR<start_recording::sap.session.Record> {e}")
        print("SAP GUI recording started...")
        while self.sap.record_flag:
            if self.kmi.events.has_new:
                __event = self.kmi.events.last()
                if isinstance(__event, MouseEvent):
                    if __event.event_type == MouseEventType.CLICK:
                        self.sap.capture_screen(__event)
                elif isinstance(__event, KeyboardEvent):
                    if __event.event_type == KeyEventType.RELEASE:
                        if __event.key in (Key.enter, Key.f8, Key.f3, Key.esc):
                            self.sap.capture_screen(__event)
    
    def stop_recording(self) -> None:
        self.kmi.stop()
        self.sap.record_flag = False
        try:
            self.sap.session.Record = False
        except Exception as e:
            print(f"ERROR<stop_recording::sap.session.Record> {e}")
        print("SAP GUI recording stopped.")
    
    def on_exit(self) -> None:
        try:
            if self.sap.record_flag:
                self.stop_recording()
        except Exception as e:
            print(f"ERROR<on_exit::stop_recording> {e}")
        try:
            self.kmi.stop()
        except Exception as e:
            print(f"ERROR<on_exit::recorder::kmi::stop> {e}")
        try:
            self.sap.session_close()
        except Exception as e:
            print(f"ERROR<on_exit::sap::session_close> {e}")
        try:
            self.icon.stop()
        except Exception as e:
            print(f"ERROR<on_exit::icon::stop> {e}")

    def save_recording_file(self, filename: Optional[str] = "my_recording.txt") -> str:
        recording_filename = filename
        with open(recording_filename, "w+") as f:
            for i in self.sap.recording:
                f.write(f"Frame {i['frame']}\n")
                if isinstance(i['data'], list):
                    for j in i['data']:
                        f.write(f"Lvl: {j['lvl']} | Element: {j['element']}\n")
                if isinstance(i['data'], DeepDiff):
                    f.write(f"{i['data'].to_json()}\n")
        return recording_filename

    def after_click(self, icon, query):
        if str(query) == "Select SAP session":
            self.sap.window_selection()
        elif str(query) == "Dump GUI elements":
            self.sap.dump_elements()
            for i in self.sap.elements:
                try:
                    print(str(TAB * i['lvl']), i['element'])
                except Exception as e:
                    print(f"ERROR<after_click::dump_elements> {e}")
        elif str(query) == "Start recording":
            self.start_recording()
        elif str(query) == "Stop recording":
            self.stop_recording()
        elif str(query) == "Save recording":
            if self.sap.record_flag:
                self.stop_recording()
            if self.sap.recording:
                print("Saving recording to file...")
                rc_file = self.save_recording_file()
                print(f"Recording saved to file: {rc_file}")
            else:
                print(f"Number of recorded steps found: {len(self.sap.recording)}, nothing to save.")
        elif str(query) == "Dump recording":
            if self.sap.recording:
                print("Dumping recording...")
                for i in self.sap.recording:
                    print(f"Frame {i['frame']}")
                    for j in i['data']:
                        print(f"Lvl: {j['lvl']} | Element: {j['element']}")
            else:
                print(f"Number of recorded steps found: {len(self.sap.recording)} nothing to dump.")
        elif str(query) == "Exit":
            try:
                self.on_exit()  
                exit()
            except SystemExit:
                pass
            except Exception as e:
                print(f"ERROR<after_click> {e}")
            finally:
                try:
                    self.on_exit()
                    exit()
                except:
                    pass


if __name__ == "__main__":
    app = App()
    app.run()
