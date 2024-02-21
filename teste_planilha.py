import pandas as pd


caminho_arquivo = "data/produtos.xlsx"

df = pd.read_excel(caminho_arquivo, dtype={
    'cod_icms': str
})
lista_produtos = df.values.tolist()

print("Iniciando...")

for linha in lista_produtos:
    descricao = linha[0]
    departamento = str(linha[1])
    subgrupo = str(linha[2])
    preco_custo = str(linha[3])
    preco_venda = str(linha[4])
    cod_unidade = str(linha[5])
    cod_ncm = str(linha[6])
    cod_icms = str(linha[7])
    cod_ipi = str(linha[8])
    cod_pis_saida = str(linha[9])
    cod_pis_entrada = str(linha[10])
    cod_natureza_pis = str(linha[11])
    cod_cofins_saida = str(linha[12])
    cod_cofins_entrada = str(linha[13])
    cod_natureza_cofins = str(linha[13])

    print(
        f"Descrição: {descricao} - Departamento: {departamento} - " +
        f"SubGrupo: {subgrupo} - Preço de Custo: {preco_custo} - " +
        f"Preço de Venda {preco_venda} - Cod. Unidade: {cod_unidade} - " +
        f"Cod. NCM: {cod_ncm} - Cod. ICMS: {cod_icms} - Cod. IPI: {cod_ipi} - " +
        f"Cod. PIS Saída: {cod_pis_saida} - Cod. PIS Entrada: {cod_pis_entrada}" +
        f"Cod. Natureza PIS: {cod_natureza_pis} - " +
        f"Cod. COFINS Saída: {cod_cofins_saida} - Cod. COFINS Entrada: {cod_cofins_entrada} - " +
        f"Cod. Natureza COFINS: {cod_natureza_cofins}"
    )

print("Encerrou!")
