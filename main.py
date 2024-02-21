import time

import pyautogui
import pyttsx3
import pandas as pd

from auth.authentication import Authentication
from produtos.program import Program


caminho_arquivo = "data/produtos.xlsx"
df = pd.read_excel(caminho_arquivo, dtype={
    'cod_icms': str,
    'cod_pis_saida': str,
    'cod_cofins_saida': str
})
lista_produtos = df.values.tolist()

nome_sistema = "Eagle Gestão"
usuario = ""
senha = ""

engine = pyttsx3.init()
engine.setProperty('rate', 150)

engine.say("Iniciando processo de cadastro de produtos! Afaste-se do teclado e do mouse!")
engine.runAndWait()


def open_products() -> None:
    pyautogui.PAUSE = 3
    with pyautogui.hold('alt'):
        pyautogui.press(['c', 'p', 's'])


def new_product() -> None:
    pyautogui.PAUSE = 1
    pyautogui.click(x=834, y=628)


def geral(
        nome_produto: str,
        cod_departamento: str,
        cod_subgrupo: str,
        custo: str,
        venda: str
) -> None:
    pyautogui.PAUSE = 1
    pyautogui.click(x=341, y=217)
    pyautogui.write(nome_produto)
    pyautogui.press('enter')
    pyautogui.write(cod_departamento)
    pyautogui.press('enter')
    pyautogui.press('enter')
    pyautogui.write(cod_subgrupo)
    pyautogui.press('enter')
    pyautogui.press('enter')
    pyautogui.press('enter')
    pyautogui.press('enter')
    pyautogui.press('enter')
    pyautogui.press('enter')
    pyautogui.press('enter')
    pyautogui.press('enter')
    pyautogui.write(custo)
    pyautogui.press('enter')
    pyautogui.press('enter')
    pyautogui.write(venda)


def estoque(cod_unidade: str) -> None:
    pyautogui.click(x=259, y=100)
    pyautogui.click(x=433, y=140)
    pyautogui.write('1')
    pyautogui.press('enter')


def tributacao(
        cod_ncm: str,
        cod_icms: str,
        cod_ipi: str,
        cod_pis_saida: str,
        cod_pis_entrada: str,
        cod_natureza_pis: str,
        cod_cofins_saida: str,
        cod_cofins_entrada: str,
        cod_natureza_cofins: str
) -> None:
    pyautogui.click(x=423, y=99)
    # Aba Geral
    # NCM
    pyautogui.click(x=281, y=174)
    pyautogui.write(str(cod_ncm))
    pyautogui.press('enter')

    # Aba ICMS
    pyautogui.click(x=271, y=135)
    # Saida
    pyautogui.click(x=561, y=192)
    pyautogui.click(x=425, y=146)
    pyautogui.press('up')
    pyautogui.click(x=690, y=146)
    pyautogui.write(cod_icms)
    pyautogui.click(x=509, y=224)
    pyautogui.click(x=941, y=624)

    # Aba IPI
    pyautogui.click(x=318, y=136)
    # Saida
    pyautogui.click(x=445, y=173)
    pyautogui.write(cod_ipi)
    pyautogui.press('enter')

    # Aba PIS
    pyautogui.click(x=373, y=133)
    # Saida
    pyautogui.click(x=494, y=170)
    pyautogui.write(cod_pis_saida)
    pyautogui.press('enter')
    # Entrada
    pyautogui.click(x=483, y=243)
    pyautogui.write(cod_pis_entrada)
    pyautogui.press('enter')
    # Natureza
    pyautogui.click(x=487, y=318)
    pyautogui.write(cod_natureza_pis)
    pyautogui.press('enter')

    # Aba COFINS
    pyautogui.click(x=432, y=133)
    # Saida
    pyautogui.click(x=487, y=171)
    pyautogui.write(cod_cofins_saida)
    pyautogui.press('enter')
    # Entrada
    pyautogui.click(x=486, y=243)
    pyautogui.write(cod_cofins_entrada)
    pyautogui.press('enter')
    # Natureza
    pyautogui.click(x=487, y=317)
    pyautogui.write(cod_natureza_cofins)
    pyautogui.press('enter')


def salvar() -> None:
    pyautogui.click(x=209, y=136)
    pyautogui.click(x=834, y=628)
    pyautogui.click(x=201, y=101)


def sair() -> None:
    pyautogui.click(x=1152, y=631)
    pyautogui.press('left')
    pyautogui.press('enter')
    engine.say("Cadastro de produtos finalizado!")
    engine.runAndWait()


# Abrir Sistema
program = Program(nome_sistema)
program.open_system()

# Efetuar o login
autenticacao = Authentication(usuario, senha)
autenticacao.login()

# Abrir Cadastro de Produtos
open_products()

for linha in lista_produtos:
    descricao = linha[0]
    departamento = str(linha[1])
    subgrupo = str(linha[2])
    preco_custo = str(linha[3])
    preco_venda = str(linha[4])
    unidade = str(linha[5])
    ncm = str(linha[6])
    icms = str(linha[7])
    ipi = str(linha[8])
    pis_saida = str(linha[9])
    pis_entrada = str(linha[10])
    natureza_pis = str(linha[11])
    cofins_saida = str(linha[12])
    cofins_entrada = str(linha[13])
    natureza_cofins = str(linha[14])

    # Adicionar um novo Produto
    new_product()

    # Aba Geral
    geral(
        descricao,
        departamento,
        subgrupo,
        preco_custo,
        preco_venda
    )

    # Aba Estoque
    estoque(unidade)

    # Aba Tributação
    tributacao(
        ncm,
        icms,
        ipi,
        pis_saida,
        pis_entrada,
        natureza_pis,
        cofins_saida,
        cofins_entrada,
        natureza_cofins
    )

    # Salvar Produto
    salvar()

# Sair do cadastro de Produtos
sair()

# Encerrar Aplicação
program.encerrar()

# time.sleep(60)
# position = pyautogui.position()
# print(f"Posição atual: {position}")
