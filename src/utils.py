from collections import deque
from PIL import Image
from typing import Any


#Constants 
TAB = "...."
IMAGE = Image.open("src\sap_screen.png")
THEME = "gray gray gray"
APP_NAME = "BPTA"
APP_TOOLTIP = "BPTA Recorder"


class MQueue(deque):
    """
    inherits deque, pops the oldest data to make room
    for the newest data when size is reached
    """
    def __init__(self, size: int) -> None:
        deque.__init__(self)
        self.size: int = size
        self.new_item: bool = False

    def full_append(self, item: Any) -> None:
        # full, pop the oldest item, left most item
        self.popleft()
        deque.append(self, item)
        self.new_item = True

    def append(self, item: Any) -> None:
        # max size reached, append becomes full_append
        if len(self) == self.size:
            self.append = self.full_append(item=item)
        else:
            deque.append(self, item)
            self.new_item = True

    def get(self) -> list:
        """returns a list of size items (newest items)"""
        return list(self)
    
    def last(self) -> Any:
        if len(self) == 0:
            return None
        else:
            return self[-1]
    
    def first(self) -> Any:
        if len(self) == 0:
            return None
        else:
            return self[0]
    
    def has_new(self) -> bool:
        if len(self) == 0:
            return False
        if self.new_item and len(self) > 0:
            self.new_item = False
            return True
        return False
