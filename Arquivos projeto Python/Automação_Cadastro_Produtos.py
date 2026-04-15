import openpyxl
import pyautogui as pg
import pyperclip
import time
import keyboard

parar = False

def parar_automacao():
    global parar
    parar = True
    print("🛑 Automação parada!")

# tecla para parar
keyboard.add_hotkey('esc', parar_automacao)

workbook = openpyxl.load_workbook('vendas_de_produtos.xlsx')
vendas_sheet = workbook['vendas']

time.sleep(3)

for linha in vendas_sheet.iter_rows(min_row=2):

    if parar:
        break

    pg.click(924, 518, duration=0.1)

    if linha[0].value:
        pyperclip.copy(str(linha[0].value))
        pg.hotkey("ctrl", "v")

    pg.press('tab')

    if linha[1].value:
        pyperclip.copy(str(linha[1].value))
        pg.hotkey("ctrl", "v")

    pg.press('tab')

    if linha[2].value:
        pyperclip.copy(str(linha[2].value))
        pg.hotkey("ctrl", "v")

    pg.press('tab')

    if linha[3].value:
        pyperclip.copy(str(linha[3].value))
        pg.hotkey("ctrl", "v")

    pg.click(x=865, y=618, duration=0.1)
    pg.click(x=945, y=578, duration=0.1)

    time.sleep(0.1)