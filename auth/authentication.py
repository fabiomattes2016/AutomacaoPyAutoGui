import pyautogui
import pyperclip


class Authentication:
    def __init__(self, username, password):
        self.username = username
        self.password = password

    def login(self) -> None:
        pyautogui.PAUSE = 1
        pyautogui.click(x=624, y=379)
        pyperclip.copy(self.username)
        pyautogui.hotkey("ctrl", "v")
        pyautogui.PAUSE = 1
        pyautogui.click(x=629, y=410)
        pyperclip.copy(self.password)
        pyautogui.hotkey("ctrl", "v")

        pyautogui.PAUSE = 1
        pyautogui.click(x=710, y=478)
