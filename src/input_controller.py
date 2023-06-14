from enum import Enum, auto
from pynput import keyboard, mouse
from pynput.keyboard import Key, KeyCode
from pynput.mouse import Button
from typing import Callable
from utils import MQueue


class MouseEventType(Enum):
    MOVE = auto()
    CLICK = auto()
    SCROLL = auto()


class KeyEventType(Enum):
    PRESS = auto()
    RELEASE = auto()


class MouseEvent:
    def __init__(self, event_type: MouseEventType, xy: tuple = None, dxdy: tuple = None, btn: Button = None, pressed: bool = False, suppress: bool = False) -> None:
        self.event_type: MouseEventType = event_type
        self.x: int = xy[0] if xy is not None else None
        self.y: int = xy[1] if xy is not None else None
        self.dx: int = dxdy[0] if dxdy is not None else None
        self.dy: int = dxdy[1] if dxdy is not None else None
        self.button: Button = btn
        self.pressed: bool = pressed
        self.suppress: bool = suppress

    def __str__(self) -> str:
        return f"<MouseEvent: type={self.event_type}, x={self.x}, y={self.y}, dx={self.dx}, dy={self.dy}, button={self.button}, pressed={self.button}, suppress={self.suppress}>"

    def __repr__(self) -> str:
        return self.__str__()


class KeyboardEvent:
    def __init__(self, event_type: KeyEventType, key: Key | KeyCode = None, suppress: bool = False, event_filter: Callable = None) -> None:
        self.event_type: KeyEventType = event_type
        self.key: Key | KeyCode = key
        self.suppress: bool = suppress
        # A callable taking the arguments (msg, data) If this callback returns False, the event will not be propagated to the listener callback.
        self.win32_event_filter: Callable = event_filter

    def __str__(self) -> str:
        return f"<KeyboardEvent: type={self.event_type}, key={self.key}, suppress={self.suppress}, event_filter={self.win32_event_filter}>"

    def __repr__(self) -> str:
        return self.__str__()


class KMInput:
    def __init__(self, on_move: bool = True, on_click: bool = True, on_scroll: bool = True, on_press: bool = True, 
                 on_release: bool = True, max_event_queue: int = 1000, **kwargs) -> None:
        self.on_move_filter = on_move
        self.on_click_filter = on_click
        self.on_scroll_filter = on_scroll
        self.on_press_filter = on_press
        self.on_release_filter = on_release
        self.max_event_queue: int = max_event_queue
        self.events: MQueue[MouseEvent | KeyboardEvent] = MQueue(self.max_event_queue)
        self.mouse_listener: mouse.Listener = mouse.Listener(on_move=self.on_move, on_click=self.on_click, on_scroll=self.on_scroll)
        self.keyboard_listener: keyboard.Listener = keyboard.Listener(on_press=self.on_press, on_release=self.on_release)
    
    def start(self) -> None:
        self.mouse_listener.start()
        self.keyboard_listener.start()
    
    def stop(self) -> None:
        self.mouse_listener.stop()
        self.keyboard_listener.stop()
    
    def __exit__(self) -> None:
        self.stop()
    
    def on_move(self, x: int, y: int) -> None:
        if self.on_move_filter:
            self.events.append(MouseEvent(event_type=MouseEventType.MOVE, xy=(x, y)))

    def on_click(self, x: int, y: int, button: Button, pressed: bool) -> None:
        if self.on_click_filter:
            self.events.append(MouseEvent(event_type=MouseEventType.CLICK, xy=(x, y), btn=button, pressed=pressed))

    def on_scroll(self, x: int, y: int, dx: int, dy: int) -> None:
        if self.on_scroll_filter:
            self.events.append(MouseEvent(event_type=MouseEventType.SCROLL, xy=(x, y), dxdy=(dx, dy)))
    
    def on_press(self, key: Key | KeyCode) -> None:
        if self.on_press_filter:
            self.events.append(KeyboardEvent(event_type=KeyEventType.PRESS, key=key))

    def on_release(self, key: Key | KeyCode) -> None:
        if self.on_release_filter:
            self.events.append(KeyboardEvent(event_type=KeyEventType.RELEASE, key=key))
