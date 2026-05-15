import glob
import os
import re
import openpyxl
from bs4 import BeautifulSoup


def mostraestoque(dados: list, ultvenda=False, precificado=False):
    somatotal = 0

    if not isinstance(dados, list):
        return "Produto não encontrado!"
    
    nome_produto = dados[0]
    marca = dados[1]
    vvenda = dados[2]
    qtdest = dados[3]
    qtdocian = dados[4]
    qtdboq = dados[5]
    qtdpet = dados[6]
    secao = dados[7]
    estmin = dados[8]
    codbarra = dados[9]
    valor_ipi = dados[10]
    valor_icms = dados[11]
    valor_ivast = dados[12]
    valor_total = dados[13]
    valor_mkup = dados[14]
    num_nfe = dados[15]
    complemento = dados[16]
    datacompra = dados[17]
    valor_unit = dados[18]
    qtdcompra = dados[19]
    datavendabq = dados[20]
    datavendaoci = dados[21]
    datavendapet = dados[22]
    if len(dados) > 23:
        icms_difal = dados[23]
    else:
        icms_difal = "0,00"

    if qtdocian < estmin:
        qtdocian = f"Ocian {qtdocian} 🔴"
    else:
        qtdocian = f"Ocian {qtdocian}"

    if qtdboq < estmin:
        qtdboq = f"Boq {qtdboq} 🔴"
    else:
        qtdboq = f"Boq {qtdboq}"

    if qtdpet is not None:
        if qtdpet < estmin:
            qtdpet = f"Pet {qtdpet} 🔴"
        else:
            qtdpet = f"Pet {qtdpet}"

    encontrados = glob.glob('listas\\**\\*- ESTOQUE*.txt', recursive=True)
    if encontrados:
        for arquivo in encontrados:
            if os.path.exists(arquivo):
                with open(arquivo, 'r', encoding="utf-8") as arq:
                    conteudo = arq.read()
                    padrao = codbarra + r';(\d+)'
                    correspondencias = re.finditer(padrao, conteudo)
                    soma_numeros = 0
                    for x in correspondencias:
                        quantidade = int(x.group(1))
                        soma_numeros += quantidade
                    if soma_numeros > 0:
                        if "|" not in nome_produto:
                            nome_produto += " |"
                        nome_produto += f" E📋({soma_numeros})"

    encontrados = glob.glob('listas\\**\\*- BOQUEIRAO*.txt', recursive=True)
    if encontrados:
        for arquivo in encontrados:
            if os.path.exists(arquivo):
                with open(arquivo, 'r', encoding="utf-8") as arq:
                    conteudo = arq.read()
                    padrao = codbarra + r';(\d+)'
                    correspondencias = re.finditer(padrao, conteudo)
                    soma_numeros = 0
                    for x in correspondencias:
                        quantidade = int(x.group(1))
                        soma_numeros += quantidade
                    if soma_numeros > 0:
                        if "|" not in nome_produto:
                            nome_produto += " |"
                        nome_produto += f" B📋({soma_numeros})"
                        somatotal += soma_numeros

    encontrados = glob.glob('listas\\**\\*- OCIAN*.txt', recursive=True)
    for arquivo in encontrados:
        if os.path.exists(arquivo):
            with open(arquivo, 'r', encoding="utf-8") as arq:
                conteudo = arq.read()
                padrao = codbarra + r';(\d+)'
                correspondencias = re.finditer(padrao, conteudo)
                soma_numeros = 0
                for x in correspondencias:
                    quantidade = int(x.group(1))
                    soma_numeros += quantidade
                if soma_numeros > 0:
                    if "|" not in nome_produto:
                        nome_produto += " |"
                    nome_produto += f" O📋({soma_numeros})"
                    somatotal += soma_numeros


    encontrados = glob.glob('listas\\**\\*- HAPPY PET*.txt', recursive=True)
    for arquivo in encontrados:
        if os.path.exists(arquivo):
            with open(arquivo, 'r', encoding="utf-8") as arq:
                conteudo = arq.read()
                padrao = codbarra + r';(\d+)'
                correspondencias = re.finditer(padrao, conteudo)
                soma_numeros = 0
                for x in correspondencias:
                    quantidade = int(x.group(1))
                    soma_numeros += quantidade
                if soma_numeros > 0:
                    if "|" not in nome_produto:
                        nome_produto += " |"
                    nome_produto += f" P📋({soma_numeros})"
                    somatotal += soma_numeros

    mudando = glob.glob('listas\\**\\Mudança*.txt', recursive=True)
    for arquivo in mudando:
        if os.path.exists(arquivo):
            with open(arquivo, 'r', encoding="utf-8") as arq:
                conteudo = arq.read()
                # Modificar o padrão para encontrar o número e o que vem depois de " - "
                padrao = re.escape(str(codbarra )) + r'\s-\s(.+)'
                correspondencias = re.finditer(padrao, conteudo)
                texto_encontrado = None
                for correspondencia in correspondencias:
                    # O grupo 1 contém o que vem depois de " - "
                    texto_encontrado = correspondencia.group(1)

                if texto_encontrado:
                    if f" ➡ Indo para {texto_encontrado}" not in nome_produto:
                        nome_produto += f" ➡ Indo para {texto_encontrado}"

    encontradossaidaestoque = glob.glob('listas\\**\\ESTOQUE -*.txt', recursive=True)
    encontradossaidaestoque += glob.glob('listas\\**\\SAIDA ESTOQUE*.txt', recursive=True)
    soma_numeros = 0
    if encontradossaidaestoque:
        for arquivo in encontradossaidaestoque:
            if os.path.exists(arquivo):
                with open(arquivo, 'r', encoding="utf-8") as arq:
                    conteudo = arq.read()
                    padrao = codbarra + r';(\d+)'
                    correspondencias = re.finditer(padrao, conteudo)
                    for x in correspondencias:
                        quantidade = int(x.group(1))
                        soma_numeros += quantidade
        if soma_numeros > 0:
            nome_produto += f" \nSaindo {soma_numeros} Estoque ⏬"
    
    encontradossaidaocian = glob.glob('listas\\**\\OCIAN -*.txt', recursive=True)
    encontradossaidaocian += glob.glob('listas\\**\\SAIDA OCIAN*.txt', recursive=True)
    soma_numeros = 0
    if encontradossaidaocian:
        for arquivo in encontradossaidaocian:
            if os.path.exists(arquivo):
                with open(arquivo, 'r', encoding="utf-8") as arq:
                    conteudo = arq.read()
                    padrao = codbarra + r';(\d+)'
                    correspondencias = re.finditer(padrao, conteudo)
                    for x in correspondencias:
                        quantidade = int(x.group(1))
                        soma_numeros += quantidade
        if soma_numeros > 0:
            nome_produto += f" \nSaindo {soma_numeros} Ocian ⏬"

    encontradossaidaboq = glob.glob('listas\\**\\BOQUEIRAO -*.txt', recursive=True)
    encontradossaidaboq += glob.glob('listas\\**\\SAIDA BOQ*.txt', recursive=True)
    soma_numeros = 0
    if encontradossaidaboq:
        for arquivo in encontradossaidaboq:
            if os.path.exists(arquivo):
                with open(arquivo, 'r', encoding="utf-8") as arq:
                    conteudo = arq.read()
                    padrao = codbarra + r';(\d+)'
                    correspondencias = re.finditer(padrao, conteudo)
                    for x in correspondencias:
                        quantidade = int(x.group(1))
                        soma_numeros += quantidade
        if soma_numeros > 0:
            nome_produto += f" \nSaindo {soma_numeros} Boq ⏬"

    encontradossaidaboq = glob.glob('listas\\**\\HAPPY PET -*.txt', recursive=True)
    encontradossaidaboq += glob.glob('listas\\**\\SAIDA HAPPY PET*.txt', recursive=True)
    soma_numeros = 0
    if encontradossaidaboq:
        for arquivo in encontradossaidaboq:
            if os.path.exists(arquivo):
                with open(arquivo, 'r', encoding="utf-8") as arq:
                    conteudo = arq.read()
                    padrao = codbarra + r';(\d+)'
                    correspondencias = re.finditer(padrao, conteudo)
                    for x in correspondencias:
                        quantidade = int(x.group(1))
                        soma_numeros += quantidade
        if soma_numeros > 0:
            nome_produto += f" \nSaindo {soma_numeros} Pet ⏬"

    # Formatando para mostrar mensagem
    nome_produto = f"<b>{nome_produto}</b>"
    codbarra = f"<code>{codbarra}</code>"
    qtdest = f"Estoque {qtdest} {f'| <b>{complemento}</b>' if complemento != '' else ''}"
    secao = f"<b>{secao}</b>"
    estmin = f"Est Min {estmin}"
    marca = f"Fornecedor: {marca}"
    vvenda = f"R$ {vvenda}"

    if ultvenda is True or precificado is True:
        qtdocian = f"{qtdocian} | Ult. Venda {datavendaoci}"
        qtdboq = f"{qtdboq} | Ult Venda {datavendabq}"
        if qtdpet is not None:
            qtdpet = f"{qtdpet} | Ult Venda {datavendapet}"

    if precificado is True:
        vvenda = f"Custo: R${valor_unit} | {f'ICMS: R${valor_icms} | ' if valor_icms != '0,00' else ''}{f'Difal: R${icms_difal} | ' if icms_difal != '0,00' else ''}{f'IPI: R${valor_ipi} | ' if valor_ipi != '0,00' else ''}{f'ST: R${valor_ivast} | ' if valor_ivast != '0,00' else ''}CustoTotal: <b>R${valor_total}</b> | MkUp: {valor_mkup}%\n<b>{vvenda}</b>"


    # Codigo de barra tem que ser o ultimo elemento
    if qtdpet:
        mensagem = f"{nome_produto}\n{marca}\n{vvenda}\n{qtdest}\n{qtdocian}\n{qtdboq}\n{qtdpet}\n{secao}\n{estmin}\n{codbarra}"
    else:
        mensagem = f"{nome_produto}\n{marca}\n{vvenda}\n{qtdest}\n{qtdocian}\n{qtdboq}\n{secao}\n{estmin}\n{codbarra}"
    return mensagem

def insereTabelaBusca(mensagem, unidade):
    if not unidade:
        return "Unidade não encontrada"
    mensagemSplit = mensagem.split("\n")

    estoque = re.search(r"Estoque\s+(-?\d+)", mensagem).group(1)
    ocian = re.search(r"Ocian\s+(-?\d+)", mensagem).group(1)
    boq = re.search(r"Boq\s+(-?\d+)", mensagem).group(1)
    secao = mensagemSplit[-3]
    codigo = strip_html_tags(mensagemSplit[-1])
    est_min = re.search(r"Est\s+Min\s+(\d+)", mensagem).group(1)

    if unidade.lower() == "boqueirão":
        file = "planilhassecao/pedidosBoq.xlsx"
    if unidade.lower() == "ocian":
        file = "planilhassecao/pedidosOci.xlsx"
    if unidade.lower() == "pet":
        file = "planilhassecao/pedidosPet.xlsx"

    if os.path.exists(file):
        wb = openpyxl.load_workbook(file)
        sheet = wb.active
        sheet.append([mensagemSplit[0], codigo, estoque, boq, ocian, est_min, secao])
    else:
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.append(["Produto", "Codigo", "Estoque", "Boq", "Ocian", "EstMin", "Seção"])
        sheet.append([mensagemSplit[0], codigo, estoque, boq, ocian, est_min, secao])

    wb.save(file)

def preparalista(dados, ordernar=False, checalistas = True):
    try:
        table = ""
        mensagens = []
        if ordernar is True:
            dados.sort(key=lambda row: row[0])
        for row in dados:
            row[0] = f"<b>{row[0]}</b>"
            codbarra = row[1]
            row[1] = f"/{codbarra}"
            if checalistas is True:
                row[2] = f"{row[2]} E"
            else:
                num = float(row[2])
                row[2] = f"QTD: <b>{row[2]}</b>"
                if num < 0:
                    row[2] += " ⛔"

            if len(row) > 6:
                row[-1] = f"<b>{row[-1]}</b>"

            codbarra = str(codbarra)

            if checalistas is True:
                encontrados = glob.glob('listas\\**\\*- ESTOQUE*.txt', recursive=True)
                if encontrados:
                    for arquivo in encontrados:
                        if os.path.exists(arquivo):
                            with open(arquivo, 'r', encoding="utf-8") as arq:
                                conteudo = arq.read()
                                padrao = codbarra + r';(\d+)'
                                correspondencias = re.finditer(padrao, conteudo)
                                soma_numeros = 0
                                for x in correspondencias:
                                    quantidade = int(x.group(1))
                                    soma_numeros += quantidade
                                if soma_numeros > 0:
                                    row[2] = f"{row[2]} 📋({soma_numeros})"

                if int(row[3]) <= int(row[5]):
                    row[3] = f"🔴 <b>{row[3]} B</b>"  # Adicionar "!" apenas antes do número na coluna D se for menor que estoque minimo
                else:
                    row[3] = f"{row[3]} B"

                encontrados = glob.glob('listas\\**\\*- BOQUEIRAO*.txt', recursive=True)
                if encontrados:
                    for arquivo in encontrados:
                        if os.path.exists(arquivo):
                            with open(arquivo, 'r', encoding="utf-8") as arq:
                                conteudo = arq.read()
                                padrao = codbarra + r';(\d+)'
                                correspondencias = re.finditer(padrao, conteudo)
                                soma_numeros = 0
                                for x in correspondencias:
                                    quantidade = int(x.group(1))
                                    soma_numeros += quantidade
                                if soma_numeros > 0:
                                    row[3] = f"{row[3]} 📋({soma_numeros})"
                
                if int(row[4]) <= int(row[5]):
                    row[4] = f"🔴 <b>{row[4]} O</b>"  # Adicionar "!" apenas antes do número na coluna D se for menor que estoque minimo
                else:
                    row[4] = f"{row[4]} O"

                encontrados = glob.glob('listas\\**\\*- OCIAN*.txt', recursive=True)
                for arquivo in encontrados:
                    if os.path.exists(arquivo):
                        with open(arquivo, 'r', encoding="utf-8") as arq:
                            conteudo = arq.read()
                            padrao = codbarra + r';(\d+)'
                            correspondencias = re.finditer(padrao, conteudo)
                            soma_numeros = 0
                            for x in correspondencias:
                                quantidade = int(x.group(1))
                                soma_numeros += quantidade
                            if soma_numeros > 0:
                                row[4] = f"{row[4]} 📋({soma_numeros})"

                row[5] = f"{row[5]} EMin"

            line_divider = "-" * 50
            linha = f"{row[0]}\n{' | '.join(map(str, row[1:]))}\n{line_divider}\n"

            if len(linha) + len(table) > 4096:
                mensagens.append(table)
                table = linha
            else:
                table += linha

        mensagens.append(table)
        return mensagens
    except Exception as e:
        print(e)
        traceback.print_exc()

def strip_html_tags(texto: str) -> str:
    soup = BeautifulSoup(texto, "html.parser")
    return soup.get_text()
