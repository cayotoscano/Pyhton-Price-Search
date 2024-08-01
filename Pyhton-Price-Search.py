# criar um navegador e importar bibliotecas
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import time
from selenium.webdriver.common.keys import Keys
import pandas as pd
import os
import smtplib
import email.message

servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome() #service=servico

# importar/visualizar base de dados
caminho_arquivo = os.getcwd()
buscas_df = pd.read_excel(caminho_arquivo + "/buscas.xlsx")

# criação da função responsavel pela pesquisa dos produtos no google shopping
def busca_google_shopping(produto, termos_banidos, preco_minimo, preco_maximo):

    # tratamento das variaveis recebidas
    produto = produto.lower()
    termos_banidos = termos_banidos.lower()
    lista_termos_banidos = termos_banidos.split(" ")
    lista_termos_nome_produto = produto.split(" ")
    preco_minimo = float(preco_minimo)
    preco_maximo = float(preco_maximo)

    # criação da lista de ofertas vazia
    lista_ofertas = []

    # entrar no google shopping
    navegador.get(r"https://shopping.google.com.br/")

    # pesquisar pelo produto
    navegador.find_element(By.XPATH,'//*[@id="REsRA"]').send_keys(produto, "\n")

    # esperar pagina carregar completamente e pegar informações dos produtos
    time.sleep(2)
    lista_resultados = navegador.find_elements(By.CLASS_NAME, 'i0X6df')

    for resultado in lista_resultados:
    # pegar nome
        nome = resultado.find_element(By.CLASS_NAME, 'tAxDx').text
        nome = nome.lower()
    # nome correspondente e termos banidos no nome do produto
        tem_termos_banidos = False
        for palavra in lista_termos_banidos:
            if palavra in nome:
                tem_termos_banidos = True
        tem_todos_termos_produto = True
        for palavra in lista_termos_nome_produto:
            if palavra not in nome:
                tem_todos_termos_produto = False
    # selecionar os nomes que preenchem os requisitos
        if tem_termos_banidos == False and tem_todos_termos_produto == True:

    # pegar preço
            try:
                preco = resultado.find_element(By.CLASS_NAME, 'a8Pemb').text
                preco = float(preco.replace("R$", "").replace(" ", "").replace(".", "").replace(",", "."))
            except:
                continue
    # preço entre o valor estipulado
            if preco_minimo <= preco <= preco_maximo:
                
    # pegar link
                elemento_referencia = resultado.find_element(By.CLASS_NAME, 'bONr3b')
                elemento_pai = elemento_referencia.find_element(By.XPATH,'..')
                link = elemento_pai.get_attribute("href")

    # adicionar item que passou pelo filtro na lista de ofertas
                lista_ofertas.append((nome, preco, link))
    return lista_ofertas

# criação da função responsavel pela pesquisa dos produtos no bucapé
def busca_buscapé(produto, termos_banidos, preco_minimo, preco_maximo):

    # tratamento das variaveis recebidas
    produto = produto.lower()
    termos_banidos = termos_banidos.lower()
    lista_termos_banidos = termos_banidos.split(" ")
    lista_termos_nome_produto = produto.split(" ")
    preco_minimo = float(preco_minimo)
    preco_maximo = float(preco_maximo)

    # criação da lista de ofertas vazia
    lista_ofertas = []

    # entrar no buscapé
    navegador.get(r"https://www.buscape.com.br/")

    # pesquisar pelo produto
    navegador.find_element(By.XPATH,'//*[@id="new-header"]/div[1]/div/div/div[3]/div/div/div[2]/div/div[1]/input').send_keys(produto, "\n")

    # esperar pagina carregar completamente e pegar informações dos produtos
    time.sleep(2)
    lista_resultados = navegador.find_elements(By.CLASS_NAME, 'ProductCard_ProductCard_Inner__gapsh')

    for resultado in lista_resultados:
    # pegar nome
        nome = resultado.find_element(By.CLASS_NAME, 'ProductCard_ProductCard_Name__U_mUQ').text
        nome = nome.lower()
    # nome correspondente e termos banidos no nome do produto
        tem_termos_banidos = False
        for palavra in lista_termos_banidos:
            if palavra in nome:
                tem_termos_banidos = True
        tem_todos_termos_produto = True
        for palavra in lista_termos_nome_produto:
            if palavra not in nome:
                tem_todos_termos_produto = False
    # selecionar os nomes que preenchem os requisitos
        if tem_termos_banidos == False and tem_todos_termos_produto == True:

    # pegar preço
            try:
                preco = resultado.find_element(By.CLASS_NAME, 'Text_MobileHeadingS__HEz7L').text
                preco = float(preco.replace("R$", "").replace(" ", "").replace(".", "").replace(",", "."))
            except:
                continue
    # preço entre o valor estipulado
            if preco_minimo <= preco <= preco_maximo:

    # pegar link
                link = resultado.get_attribute("href")

    # adicionar item que passou pelo filtro na lista de ofertas
                lista_ofertas.append((nome, preco, link))
    return lista_ofertas

# criar tabela vazia
tabela_ofertas = pd.DataFrame()

# salver ofertas na "tabela_ofertas"
for linha in buscas_df.index:
# pesquisar pelo produto
    produto = buscas_df.loc[linha,"Nome"]
    termos_banidos = buscas_df.loc[linha,"Termos banidos"]
    preco_minimo = buscas_df.loc[linha,"Preço mínimo"]
    preco_maximo = buscas_df.loc[linha,"Preço máximo"]

# aplicar funções em cada produto da base de dados
    lista_ofertas_google_shopping = busca_google_shopping(produto, termos_banidos, preco_minimo, preco_maximo)
    if lista_ofertas_google_shopping:
        tabela_google_shopping = pd.DataFrame(lista_ofertas_google_shopping, columns=["Produto", "Preço", "Link"])
        tabela_ofertas = pd.concat([tabela_ofertas, tabela_google_shopping])

    lista_ofertas_buscape = busca_buscapé(produto, termos_banidos, preco_minimo, preco_maximo)
    if lista_ofertas_buscape:
        tabela_ofertas_buscape = pd.DataFrame(lista_ofertas_buscape, columns=["Produto", "Preço", "Link"])
        tabela_ofertas = pd.concat([tabela_ofertas, tabela_ofertas_buscape])

# exportar para o excel
tabela_ofertas.to_excel("Ofertas.xlsx", index=False)

# enviar por e-mail os resultados

corpo_email = f"""
RESULTADO DA PESQUISA DE PREÇOS DO(S) PRODUTO(S)

{tabela_ofertas.to_html(index=False)}
"""

msg = email.message.Message()
msg['Subject'] = "PESQUISA DE PREÇOS"
msg['From'] = 'cayocraft@gmail.com'
msg['To'] = 'cayocraft@gmail.com'
password = 'senha de app' 
msg.add_header('Content-Type', 'text/html')
msg.set_payload(corpo_email )

s = smtplib.SMTP('smtp.gmail.com: 587')
s.starttls()
s.login(msg['From'], password)
s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
print('Email enviado')

navegador.quit()