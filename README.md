# Beatriz-Ramos


import xlwings as xw
import pyautogui as pygui
import autoit
import os


def Bloquear(identificador, dias, motivo):

    if identificador == 'None':
        os.system('wmic process where name="excel.exe" delete')
        exit()
    autoit.run(r"caminho do programa")
    autoit.win_wait("Login CSLOG", 4)
    autoit.control_send("Login CSLOG", "TEdit1", login)
    autoit.control_send("Login CSLOG", "TEdit2", senha)
    autoit.control_click("Login CSLOG", "TBitBtn21")
    pygui.sleep(3)
    janela = autoit.win_get_title("[ACTIVE]")

    # começar bloqueio
    pygui.press('F2')
    pygui.sleep(1)
    janela2 = autoit.win_get_title('[ACTIVE]')
    pygui.sleep(1)
    autoit.control_click(janela2, 'TMaskEdit1')
    pygui.sleep(1)
    autoit.control_send(janela2, 'TMaskEdit1', identificador)
    pygui.sleep(1)
    autoit.control_click(janela2, 'TComboBox1')
    pygui.sleep(1)
    autoit.control_send(janela2, 'TComboBox1', 'identificador')
    pygui.sleep(1)
    pygui.press('enter')
    pygui.sleep(1)
    autoit.control_click(janela2, 'TBitBtn8')
    pygui.sleep(1)

    # se os dias for igual a 0 faça isso
    if dias == 0:
        autoit.control_click(janela2, 'TBitBtn7')
        pygui.sleep(1)
        pygui.press('enter')
        pygui.sleep(1)
        janela3 = autoit.win_get_title('[ACTIVE]')
        pygui.sleep(1)
        autoit.control_click(janela3, 'TBitBtn25')
        pygui.sleep(1)
        janela4 = autoit.win_get_title('[ACTIVE]')
        pygui.sleep(1)
        autoit.control_click(janela4, 'TComboBox2')
        pygui.sleep(1)
        pygui.press('down')
        pygui.sleep(1)
        pygui.press('enter')
        pygui.sleep(1)
        autoit.control_click(janela4, 'TBitBtn23')
        pygui.sleep(1)
        janela5 = autoit.win_get_title('[ACTIVE]')
        autoit.control_click(janela5, 'TEdit2')
        autoit.control_set_text(janela5, 'TEdit2', motivo)
        autoit.control_click(janela5, 'TBitBtn22')

    # se não faça isso
    else:
        autoit.control_click(janela5, 'TBitBtn4')
        pygui.sleep(1)
        janela6 = autoit.win_get_title('[ACTIVE]')
        pygui.sleep(1)
        autoit.control_click(janela6, 'TBitBtn1')
        pygui.sleep(1)
        janela7 = autoit.win_get_title('[ACTIVE]')
        pygui.sleep(1)
        autoit.control_click(janela7, 'TDBLookupComboBox1')
        pygui.press('down', presses=3)
        pygui.sleep(1)
        pygui.press('enter')
        pygui.sleep(1)
        autoit.control_click(janela7, 'TEditButton1')
        pygui.sleep(2)

    # se o dia for igual a 6
    if dias == 6:
        pygui.press('left', presses=4)
    elif dias == 8:
        pygui.press('left', presses=2)
    elif dias == 16:
        pygui.press('right', presses=6)

    pygui.sleep(1)
    pygui.press('space')
    pygui.sleep(1)
    pygui.press('enter')
    autoit.control_click(janela7, 'TBitBtn21')

    # fechar siscob
    os.system('wmic process where name="siscob.exe" delete')

#fim da função----------------------------------------------------------

#acesso
login = 'usuario'
senha = '****'

#abrir planilha e atualizar
xb = xw.Book('caminho do arquivo')
janela = autoit.win_get_title("[ACTIVE]")
pygui.sleep(1)
pygui.press("alt")
pygui.sleep(1)
pygui.press("s")
pygui.sleep(1)
pygui.press("g")
pygui.sleep(1)
pygui.press("a")

#rodar loop sobre as celulas para pegar a informação

for i in range(2,1000):
    i = str(i)
    identificador = xw.Range('A' + i).value
    identificador = str(identificador)
    dias = xw.Range('G' + i).value
    dias = str(dias)
    motivo = xw.Range('E' + i).value
    motivo = str(motivo)
    print(identificador)
    Bloquear(identificador,dias,motivo)
    if identificador == 'None':
        break

