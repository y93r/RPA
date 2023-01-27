#bibliotecas
import pyautogui
import pyperclip
import subprocess
import time
import pandas as pd

#ERP utilizado é o Fakturama
pyautogui.FAILSAFE = True

def encontrar_imagem(imagem):
    while not pyautogui.locateOnScreen(imagem, grayscale=True, confidence=0.9):  # reconhecimento de imagem
        time.sleep(1)
    encontrou = pyautogui.locateOnScreen(imagem, grayscale=True, confidence=0.9)
    return encontrou

def escrever_texto(texto):
    pyperclip.copy(texto)
    pyautogui.hotkey("ctrl", "v")

def direita(posicoes_imagem):
    return posicoes_imagem[0] + posicoes_imagem[2], posicoes_imagem[1] + posicoes_imagem[3]/2

# Abrir o ERP 
subprocess.Popen([r"C:\Program Files\Fakturama2\Fakturama.exe"])
encontrou = encontrar_imagem("fakturama.png")

# fakturama está aberto

# vários produtos
tabela_produtos = pd.read_excel("Produtos.xlsx")
print(tabela_produtos)
for linha in tabela_produtos.index:
    nome = tabela_produtos.loc[linha, "Nome"]
    id = tabela_produtos.loc[linha, "ID"]
    categoria = tabela_produtos.loc[linha, "Categoria"]
    gtin = tabela_produtos.loc[linha, "GTIN"]
    supplier = tabela_produtos.loc[linha, "Supplier"]
    descricao = tabela_produtos.loc[linha, "Descrição"]
    imagem = tabela_produtos.loc[linha, "Imagem"]
    preco = tabela_produtos.loc[linha, "Preço"]
    custo = tabela_produtos.loc[linha, "Custo"]
    estoque = tabela_produtos.loc[linha, "Estoque"]

    # Clicar em New Product
    encontrou = encontrar_imagem("new_product.png")
    pyautogui.click(pyautogui.center(encontrou))

    # preencher todos os campos
    encontrou = encontrar_imagem("item_number.png")
    pyautogui.click(direita(encontrou))
    escrever_texto(str(id))

    encontrou = encontrar_imagem("name_produto.png")
    pyautogui.click(direita(encontrou))
    escrever_texto(str(nome))

    encontrou = encontrar_imagem("category_produto.png")
    pyautogui.click(direita(encontrou))
    escrever_texto(str(categoria))

    encontrou = encontrar_imagem("GTIN_produto.png")
    pyautogui.click(direita(encontrou))
    escrever_texto(str(gtin))

    encontrou = encontrar_imagem("supplier_produto.png")
    pyautogui.click(direita(encontrou))
    escrever_texto(str(supplier))

    encontrou = encontrar_imagem("description_produto.png")
    pyautogui.click(direita(encontrou))
    escrever_texto(str(descricao))

    encontrou = encontrar_imagem("preco_produto.png")
    pyautogui.click(direita(encontrou))
    preco_texto = f"{preco:.2f}".replace(".", ",")
    escrever_texto(str(preco_texto))

    encontrou = encontrar_imagem("custo_produto.png")
    pyautogui.click(direita(encontrou))
    custo_texto = f"{custo:.2f}".replace(".", ",")
    escrever_texto(str(custo_texto))

    encontrou = encontrar_imagem("stock_produto.png")
    pyautogui.click(direita(encontrou))
    estoque_texto = f"{estoque:.2f}".replace(".", ",")
    escrever_texto(str(estoque_texto))

    # selecionar a imagem
    encontrou = encontrar_imagem("select_picture.png")
    pyautogui.click(pyautogui.center(encontrou))
    
    encontrou = encontrar_imagem("nome_arquivo.png")
    escrever_texto(rf"C:\Users\Usuário\Documents\ImagensProdutos\{str(imagem)}")
    pyautogui.press("enter")

    # clicar em salvar
    encontrou = encontrar_imagem("botaosave.png")
    pyautogui.click(pyautogui.center(encontrou))
