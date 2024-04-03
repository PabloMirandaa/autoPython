import openpyxl
import pyperclip 
import pyautogui
from time import sleep

# Acessando a planilha
workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos'] #selcionando qual pagina da planilha sera utilizada

# Copiar as informações de um campo para colar no seu campo correspondente
for linha in sheet_produtos.iter_rows(min_row = 2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(1170,416,duration=1) #usado para as coordenadas do mouse
    pyautogui.hotkey('ctrl','v')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(1182,570,duration=1) 
    pyautogui.hotkey('ctrl','v')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(1181,716,duration=1) 
    pyautogui.hotkey('ctrl','v')

    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(1181,866,duration=1) 
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(1163,950,duration=1)
    sleep(5)

    preco = linha[4].value
    pyperclip.copy(preco)
    pyautogui.click(1179,468,duration=1) 
    pyautogui.hotkey('ctrl','v')

    quantidade_estoque = linha[5].value
    pyperclip.copy(quantidade_estoque)
    pyautogui.click(1173,620,duration=1) 
    pyautogui.hotkey('ctrl','v')

    data_validade = linha[6].value
    pyperclip.copy(data_validade)
    pyautogui.click(1187,766,duration=1) 
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(1229,928,duration=1) 
    tamanho = linha[7].value
    if tamanho == 'Pequeno':
        pyautogui.click(1255,870,duration=1)
    elif tamanho == 'Médio':
        pyautogui.click(1197,926,duration=1)
    else:
        pyautogui.click(1188,979,duration=1)

    pyautogui.click(1255,1010,duration=1)
    sleep(5)
    
    fabricante = linha[8].value
    pyperclip.copy(fabricante)
    pyautogui.click(1186,470,duration=1) 
    pyautogui.hotkey('ctrl','v')

    pais_origem = linha[9].value
    pyperclip.copy(pais_origem)
    pyautogui.click(1179,621,duration=1) 
    pyautogui.hotkey('ctrl','v')

    codigo_barras = linha[10].value
    pyperclip.copy(codigo_barras)
    pyautogui.click(1189,773,duration=1) 
    pyautogui.hotkey('ctrl','v')

    localizacao_armazem = linha[11].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(1190,924,duration=1) 
    pyautogui.hotkey('ctrl','v')

    pyautogui.click(1258,1001,duration=1)
    sleep(5)

    pyautogui.click(1214,314,duration=1)
    sleep(5)