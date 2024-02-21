import pyautogui


class Cadastro:


    def open_products(self) -> None:
        pyautogui.PAUSE = 3
        with pyautogui.hold('alt'):
            pyautogui.press(['c', 'p', 's'])
