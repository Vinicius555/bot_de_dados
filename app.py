# Importando Bibliotecas

import pyautogui
import openpyxl
import pyperclip
from time import sleep

# Entrar na planilha

workbook = openpyxl.load_workbook('produtos_ficticios.xlsx')
sheet_produtos = workbook['Produtos']

# Copiar informções de campo e colar no seu campo correspondente
for linha in sheet_produtos.iter_rows(min_row=2):
    nome_produto = linha[0].value
    pyperclip.copy(nome_produto)
    pyautogui.click(304, 174, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    descricao = linha[1].value
    pyperclip.copy(descricao)
    pyautogui.click(422, 280, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    categoria = linha[2].value
    pyperclip.copy(categoria)
    pyautogui.click(366, 388, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    codigo_produto = linha[3].value
    pyperclip.copy(codigo_produto)
    pyautogui.click(315, 476, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    peso = linha[4].value
    pyperclip.copy(peso)
    pyautogui.click(310, 564, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    dimensoes = linha[5].value
    pyperclip.copy(dimensoes)
    pyautogui.click(295, 648, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Clicando no botão para a próxima página
    pyautogui.click(149, 702, duration=1)
    sleep(5)

    preco = linha[6].value
    pyperclip.copy(preco)
    pyautogui.click(255, 201, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    quantidade_em_estoque = linha[7].value
    pyperclip.copy(quantidade_em_estoque)
    pyautogui.click(256, 269, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    data_de_validade = linha[8].value
    pyperclip.copy(data_de_validade)
    pyautogui.click(232, 368, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    cor = linha[9].value
    pyperclip.copy(cor)
    pyautogui.click(226, 451)
    pyautogui.hotkey('ctrl', 'v')

    tamanho = linha[10].value
    pyautogui.click(338, 535, duration=1)
    if tamanho == "Pequeno":
        pyautogui.click(285, 577, duration=1)
    if tamanho == "Médio":
        pyautogui.click(284, 594, duration=1)
    if tamanho == "Grande":
        pyautogui.click(279, 612, duration=1)

    material = linha[11].value
    pyperclip.copy(material)
    pyautogui.click(367, 619, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Clicando para a próxima página
    pyautogui.click(159, 670, duration=1)
    sleep(5)

    fabricante = linha[12].value
    pyperclip.copy(fabricante)
    pyautogui.click(279, 216, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    pais_origem = linha[13].value
    pyperclip.copy(pais_origem)
    pyautogui.click(273, 291, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    obsevacoes = linha[14].value
    pyperclip.copy(obsevacoes)
    pyautogui.click(282, 385, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    codigo_de_barras = linha[15].value
    pyperclip.copy(codigo_de_barras)
    pyautogui.click(275, 524, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    localizacao_armazem = linha[16].value
    pyperclip.copy(localizacao_armazem)
    pyautogui.click(276, 600, duration=1)
    pyautogui.hotkey('ctrl', 'v')

    # Clicando no botão concluit

    pyautogui.click(151, 658, duration=1)
    sleep(5)

    # Clicando no botão OK

    pyautogui.click(866, 157, duration=1)
    sleep(5)

    # Clicando no botão de Adicionar mais um produto
    pyautogui.click(717, 445, duration=1)
    sleep(5)
print("Planilha Adicionada com Sucesso!")
