import pyautogui
import pyperclip


class Program:
    def __init__(self, nome_sistema):
        self.nome_sistema = nome_sistema

    def open_system(self) -> None:
        pyautogui.PAUSE = 1
        pyautogui.press("win")
        pyperclip.copy(self.nome_sistema)
        pyautogui.hotkey("ctrl", "v")
        pyautogui.press("enter")

    @staticmethod
    def encerrar() -> None:
        with pyautogui.hold('ctrl'):
            pyautogui.press('f')

        pyautogui.press("left")
        pyautogui.press("enter")
        pyautogui.press("right")
        pyautogui.press("right")
        pyautogui.press("enter")
