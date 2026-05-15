import calendar
from collections import defaultdict
from datetime import datetime, timedelta, date
from decimal import Decimal
import os
from dotenv import load_dotenv
import traceback
import fdb
import locale
import html
import openpyxl
import pandas as pd
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.styles import numbers


COLUNAS_INTERESSE = ['PRODUTO', 'VENDA', 'ESTOQUEFISICO', 'COD1', 'DEPT_NM', 'ESTOQUEMINIMO', 'TAM',
                    'COMPRA', 'COMPRADT', 'QUATDE_NEW', 'VENDIDO', 'SMARCA', 'ESTOQDEP', 'SGRUPO',
                    'IPIE', 'ICMSV', 'IVAST', 'UNITARIO', 'MARKUP', 'NF', 'CODFORNECEDOR', 'ICMSDIFAL']

def get_today_formatted():
    today = datetime.today()
    day_of_year = today.timetuple().tm_yday
    last_two_digits_of_year = today.year % 100
    formatted_date = f"{day_of_year:03}{last_two_digits_of_year:02}"
    return formatted_date

def ordenarplanilha(arquivo, col):
    wb = openpyxl.load_workbook(arquivo)
    sheet = wb.active
    data = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        data.append(tuple(row))
    data.sort(key=lambda x: x[col])
    sheet.delete_rows(2, sheet.max_row)
    for row_data in data:
        sheet.append(row_data)
    sheet.title = "Planilha1"
    wb.save(arquivo)


load_dotenv()

def consulta_estoque(consultasql, params = None):
    with fdb.connect(dsn=f"{os.getenv('FB_HOST')}:{os.getenv('FB_DB_ESTOQUE')}",
                       user=os.getenv("FB_USER"),
                       password=os.getenv("FB_PASSWORD")
                       ) as con:
        cur = con.cursor()
        cur.execute(consultasql, params or ())
        resultado = cur.fetchall()
        return resultado
    

def consulta_boq(consultasql, params = None):
    with fdb.connect(dsn=f'{os.getenv("FB_HOST")}:{os.getenv("FB_DB_BOQ")}',
                       user=os.getenv("FB_USER"),
                       password=os.getenv("FB_PASSWORD")
                       ) as con:
        cur = con.cursor()
        cur.execute(consultasql, params or ())
        resultado = cur.fetchall()
        cur.close()
        return resultado

def consulta_oci(consultasql, params = None):
    with fdb.connect(dsn=f'{os.getenv("FB_HOST")}:{os.getenv("FB_DB_OCI")}',
                       user=os.getenv("FB_USER"),
                       password=os.getenv("FB_PASSWORD")
                       ) as con:
        cur = con.cursor()
        cur.execute(consultasql, params or ())
        resultado = cur.fetchall()
        cur.close()
        return resultado

def consulta_pet(consultasql, params = None):
    with fdb.connect(dsn=f'{os.getenv("FB_HOST")}:{os.getenv("FB_DB_PET")}',
                       user=os.getenv("FB_USER"),
                       password=os.getenv("FB_PASSWORD")
                       ) as con:
        cur = con.cursor()
        cur.execute(consultasql, params or ())
        resultado = cur.fetchall()
        cur.close()
        return resultado

def get_ultnotas():
    colunas_interesse = ['FORNECEDORNOME', 'NUMERO']
    resultado = consulta_estoque(f"SELECT {', '.join(colunas_interesse)} FROM NFCOMPRA")
    return resultado[-5:]

def checavalidadepet(diasvalidade :int):
    "Checa produtos que estão para vencer, passe como argumento em quantos dias para vencer"
    produtosvalidade = []
    colunas_produto_identifica = ['ID_PRODUTO', 'NUMERO_SERIE', 'DATA_VALIDADE']
    consulta_sql = f"SELECT {', '.join(colunas_produto_identifica)} FROM PRODUTO_IDENTIFICA"

    resultados = consulta_pet(consulta_sql)

    for linha in resultados:
        codigoid = linha[0]
        resultadosid = consulta_pet(f"SELECT {', '.join(COLUNAS_INTERESSE)} FROM PRODUTO WHERE ID = ?",
                                    (codigoid,))
        if resultadosid:
            if int(resultadosid[0][2]) <= 0:
                continue
            else:
                produto = resultadosid[0][0]
                validade = linha[2]

                diferenca = validade - datetime.now()
                dias = diferenca.days
                if dias <= diasvalidade and dias > 0:
                    produtosvalidade.append(f"⚠ <b>Atenção!</b>\n{produto}\n<code>{resultadosid[0][3]}</code>\n<b>Vence em {dias} dia(s)!\nValidade: {validade.strftime('%d/%m/%Y')}</b>")

    return produtosvalidade

def inseretroca(codigo, trocaqtd=1):
    if len(codigo) > 255:
        return "Código comprido demais!"
    try:
        codigolimpo = html.escape(codigo)
        resultadosest = consulta_estoque(f"SELECT {', '.join(COLUNAS_INTERESSE)} FROM PRODUTO WHERE COD1 = ?",
                                        (codigolimpo,))
        nome_arquivo_troca = "Trocas.xlsx"
        wb = openpyxl.load_workbook(nome_arquivo_troca)
        nome_produto = resultadosest[0][0]
        codigo_produto = resultadosest[0][3]
        marca = resultadosest[0][11]
        dataog = resultadosest[0][8]
        custo = resultadosest[0][7]
        ref = resultadosest[0][20]
        nnfe = resultadosest[0][19]
        numatual = 0

        if dataog is not None:
            data_formatada = dataog.strftime("%d/%m/%Y %H:%M:%S")
        else:
            data_formatada = ""
        
        sheet = wb["Trocas"]
        marcasparanaosalvar = ["PADRAO", "MARCA"]
        codplan = datetime.now().strftime("%d%m%y")

        for posicao, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
            coluna_codigoplan = row[7]
            coluna_gtin = row[3]
            if coluna_codigoplan is not None and coluna_codigoplan == codplan and coluna_gtin == codigo_produto:
                # Incrementar o valor na coluna 5
                celula_col5 = sheet.cell(row=posicao, column=5)
                numatual = celula_col5.value
                break

        if marca not in marcasparanaosalvar:

            if numatual > 0:
                sheet[f"A{posicao}"].value = marca
                sheet[f"B{posicao}"].value = ref
                sheet[f"C{posicao}"].value = nome_produto
                sheet[f"D{posicao}"].value = codigo_produto
                sheet[f"E{posicao}"].value = int(trocaqtd) + int(numatual)
                sheet[f"F{posicao}"].value = nnfe
                sheet[f"G{posicao}"].value = data_formatada
                sheet[f"H{posicao}"].value = codplan

            else:
                last_row = sheet.max_row + 1
                sheet[f"A{last_row}"].value = marca
                sheet[f"B{last_row}"].value = ref
                sheet[f"C{last_row}"].value = nome_produto
                sheet[f"D{last_row}"].value = codigo_produto
                sheet[f"E{last_row}"].value = int(trocaqtd)
                sheet[f"F{last_row}"].value = nnfe
                sheet[f"G{last_row}"].value = data_formatada
                sheet[f"H{last_row}"].value = codplan

            result = [f"Produto: {nome_produto} - Salvo!\nMarca: {marca}\nQTD: {trocaqtd}\nCodigo: <code>{codigo_produto}</code>",
                        codplan]
        else:
            result = [f"Produto: {nome_produto} \nMarca: {marca}\nCodigo: <code>{codigo_produto}</code>\n🔴 Produto não possivel para troca"]

        wb.save(nome_arquivo_troca)
        return result
    except Exception:
        traceback.print_exc()

def historicoRetirada(codigo, trocaqtd, unidade):
        
    if len(codigo) > 255:
        return "Código comprido demais!"
    try:
        codigolimpo = html.escape(codigo)
        resultadosest = consulta_estoque(f"SELECT {', '.join(COLUNAS_INTERESSE)} FROM PRODUTO WHERE COD1 = ?",
                                        (codigolimpo,))
        nome_arquivo_troca = "Trocas.xlsx"
        wb = openpyxl.load_workbook(nome_arquivo_troca)
        nome_produto = resultadosest[0][0]
        codigo_produto = resultadosest[0][3]
        marca = resultadosest[0][11]
        dataog = resultadosest[0][8]
        custo = resultadosest[0][7]
        ref = resultadosest[0][20]
        nnfe = resultadosest[0][19]
        numatual = 0

        if dataog is not None:
            data_formatada = dataog.strftime("%d/%m/%Y %H:%M:%S")
        else:
            data_formatada = ""

        posicao = 0
        sheet = wb["HistoricoRetiradas"]
        codplan = datetime.now().strftime("%d%m%y")
        for posicao, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
            coluna_codigoplan = row[7]
            coluna_gtin = row[3]
            coluna_unidade = row[8]
            if coluna_codigoplan is not None and coluna_codigoplan == codplan and coluna_gtin == codigo_produto and coluna_unidade == unidade:
                # Incrementar o valor na coluna 5
                celula_col5 = sheet.cell(row=posicao, column=5)
                numatual = celula_col5.value
                break

        if numatual > 0:
            sheet[f"A{posicao}"].value = marca
            sheet[f"B{posicao}"].value = ref
            sheet[f"C{posicao}"].value = nome_produto
            sheet[f"D{posicao}"].value = codigo_produto
            sheet[f"E{posicao}"].value = int(trocaqtd) + int(numatual)
            sheet[f"F{posicao}"].value = nnfe
            sheet[f"G{posicao}"].value = data_formatada
            sheet[f"H{posicao}"].value = codplan
            sheet[f"I{posicao}"].value = unidade
            sheet[f"J{posicao}"].value = custo
            
        else:
            last_row = sheet.max_row + 1
            sheet[f"A{last_row}"].value = marca
            sheet[f"B{last_row}"].value = ref
            sheet[f"C{last_row}"].value = nome_produto
            sheet[f"D{last_row}"].value = codigo_produto
            sheet[f"E{last_row}"].value = int(trocaqtd) + int(numatual)
            sheet[f"F{last_row}"].value = nnfe
            sheet[f"G{last_row}"].value = data_formatada
            sheet[f"H{last_row}"].value = codplan
            sheet[f"I{last_row}"].value = unidade
            sheet[f"J{last_row}"].value = custo

        wb.save(nome_arquivo_troca)
    except Exception:
        traceback.print_exc()


def todos_nome(termos, todos=False, somente_pet=False):
    "Pega todos os produtos pelo nome"

    if len(termos) > 50:
        return ValueError

    consulta_sql = f"SELECT {', '.join(COLUNAS_INTERESSE)} FROM PRODUTO WHERE ATIVO = 1"
    consulta_sql += " AND " + " AND ".join(["PRODUTO CONTAINING ?" for _ in termos])

    if todos is False:
        consulta_sql += " AND ESTOQUEFISICO >= 1"

    dados = []

    if somente_pet is True:
        resultadospet = consulta_pet(consulta_sql, termos)
        for rowest in resultadospet:
            produto = rowest[0]
            qtdest = rowest[2]
            codigoatual = rowest[3]
            dados.append([produto, codigoatual, qtdest])
        return dados

    resultadosest = consulta_estoque(consulta_sql, termos)

    for rowest in resultadosest:
        produto = rowest[0]
        qtdest = rowest[2]
        codigoatual = rowest[3]
        secao = rowest[4]
        estmin = rowest[5]

        consulta_sql = f"SELECT {', '.join(COLUNAS_INTERESSE)} FROM PRODUTO WHERE COD1 = ?"

        try:
            resultadosboq = consulta_boq(consulta_sql, (codigoatual,))
            resultadosoci = consulta_oci(consulta_sql, (codigoatual,))
        except Exception as e:
            print(e)
            continue
        
        if resultadosboq:
            qtdbq = resultadosboq[0][2]
        else:
            qtdbq = 0

        if resultadosoci:
            qtdoci = resultadosoci[0][2]
        else:
            qtdoci = 0

        dados.append([produto, codigoatual, qtdest, qtdbq, qtdoci,
                        estmin, secao])
    return dados


def todos_marca(marca, todos=False):
    "Pega todos os produtos pela marca"

    if len(marca) > 50:
        return ValueError

    consulta_sql = f"SELECT {', '.join(COLUNAS_INTERESSE)} FROM PRODUTO WHERE SMARCA CONTAINING ?"

    if todos is False:
        consulta_sql += " AND ESTOQUEFISICO >= 1"

    dados = []
    resultadosest = consulta_estoque(consulta_sql, (marca,))

    for rowest in resultadosest:
        produto = rowest[0]
        qtdest = rowest[2]
        codigoatual = rowest[3]
        secao = rowest[4]
        estmin = rowest[5]

        consulta_sql = f"SELECT {', '.join(COLUNAS_INTERESSE)} FROM PRODUTO WHERE COD1 = ?"

        try:
            resultadosboq = consulta_boq(consulta_sql, (codigoatual,))
            resultadosoci = consulta_oci(consulta_sql, (codigoatual,))
        except Exception as e:
            print(e)
            continue

        if not resultadosboq:
            continue
        if not resultadosoci:
            continue
        qtdbq = resultadosboq[0][2]
        qtdoci = resultadosoci[0][2]

        dados.append([produto, codigoatual, qtdest, qtdbq, qtdoci,
                      estmin, secao])
    return dados

def todos_negativos(unidade:str, marca="FALSE"):
    "Pega todos os produtos negativos"
    marca = f"%{marca.upper()}%"
    if marca == "%FALSE%":
        consulta_sql = f"SELECT {', '.join(COLUNAS_INTERESSE)} FROM PRODUTO WHERE ATIVO = 1 AND ESTOQUEFISICO < 0"
    else:
        consulta_sql = f"SELECT {', '.join(COLUNAS_INTERESSE)} FROM PRODUTO WHERE ATIVO = 1 AND ESTOQUEFISICO < 0 AND SMARCA LIKE ?"

    dados = []

    if unidade == "estoque":
        resultadosest = consulta_estoque(consulta_sql, (marca,))
    elif unidade == "boqueirao":
        resultadosest = consulta_boq(consulta_sql, (marca,))
    elif unidade == "ocian":
        resultadosest = consulta_oci(consulta_sql, (marca,))
    elif unidade == "pet":
        resultadosest = consulta_pet(consulta_sql, (marca,))
    else:
        raise ValueError

    for rowest in resultadosest:
        produto = rowest[0]
        qtdest = rowest[2]
        codigoatual = rowest[3]

        dados.append([produto, codigoatual, qtdest])
        
    return dados
                                                                                                                                                                                                                                                                                                                                                                            
def consulta(codigo, somente_pet=False):

    qtdoci = None
    qtdest = None
    qtdbq = None
    qtdpet = None
    resultadospet = None
    resultadosest = None
    vendabq = None
    datavendabq = None
    datavendaoci = None
    vendapet = None
    datavendapet = None
    complemento = None

    if len(str(codigo)) < 255:
        codigolimpo = html.escape(codigo)
        icms_difal = "0,00"
        difalformat = "0,00"

        if somente_pet is True:
            consulta_sql = f"SELECT {', '.join(COLUNAS_INTERESSE)} FROM PRODUTO WHERE ATIVO = 1 AND COD1 = ?"
            resultadospet = consulta_pet(consulta_sql, (codigolimpo,))
            print(codigolimpo)

            nomeprod = resultadospet[0][0]
            preco = resultadospet[0][1]
            qtdpet = resultadospet[0][2]
            codprod = resultadospet[0][3]
            secao = resultadospet[0][4]
            estmin = resultadospet[0][5]
            complemento = resultadospet[0][6]
            ptotal = resultadospet[0][7]
            compradata = resultadospet[0][8]
            qtdcompra = resultadospet[0][9]
            vendapet = resultadospet[0][10]
            fornecedor = resultadospet[0][11]
            ipie = resultadospet[0][14]
            icmsValor = resultadospet[0][15]
            ivast = resultadospet[0][16]
            punitario = resultadospet[0][17]
            mkup = resultadospet[0][18]
            nnfe = resultadospet[0][19]
            icms_difal = resultadospet[0][21]

            print(resultadospet)
            if complemento is None:
                complemento = ""
            datavendapet = ""
            if vendapet is not None:
                datavendapet = datetime.strptime(str(vendapet), "%Y-%m-%d").strftime("%d-%m-%Y")

            vvenda = locale.format_string("%.2f", preco, grouping=True)
            punitarioformat = locale.format_string("%.2f", punitario, grouping=True)
            ipieformat = locale.format_string("%.2f", ipie, grouping=True)
            icmsValorformat = locale.format_string("%.2f", icmsValor, grouping=True)
            ivastformat = locale.format_string("%.2f", ivast, grouping=True)
            mkupformat = locale.format_string("%.2f", mkup, grouping=True)
            ptotalformat = locale.format_string("%.2f", ptotal, grouping=True)
            if icms_difal is not None:
                difalformat = locale.format_string("%.2f", icms_difal, grouping=True)

            dados = [nomeprod, fornecedor, vvenda, "", "",
                    "", qtdpet, "", "", codprod, ipieformat, icmsValorformat,
                    ivastformat, ptotalformat, mkupformat, nnfe, complemento,
                    compradata, punitarioformat, qtdcompra, "", "", datavendapet, difalformat]
            
            return dados

        else:
            consulta_sql = f"SELECT {', '.join(COLUNAS_INTERESSE)} FROM PRODUTO WHERE COD1 = ?"

            resultadosest = consulta_estoque(consulta_sql, (codigolimpo,))
            resultadosboq = consulta_boq(consulta_sql, (codigolimpo,))
            resultadosoci = consulta_oci(consulta_sql, (codigolimpo,))
            resultadospet = consulta_pet(consulta_sql, (codigolimpo,))

            if resultadosest:
                for row in resultadosest:
                    nomeprod = row[0]
                    preco = row[1]
                    qtdest = row[2]
                    codprod = row[3]
                    secao = row[4]
                    estmin = row[5]
                    complemento = row[6]
                    ptotal = row[7]
                    compradata = row[8]
                    qtdcompra = row[9]
                    fornecedor = row[11]
                    ipie = row[14]
                    icmsValor = row[15]
                    ivast = row[16]
                    punitario = row[17]
                    mkup = row[18]
                    nnfe = row[19]
                    icms_difal = row[21]

                if complemento is None:
                    complemento = ""

                if ivast is None:
                    ivast = 0

                if mkup is None:
                    mkup = 0

                if resultadosboq:
                    qtdbq = resultadosboq[0][2]
                    vendabq = resultadosboq[0][10]
                    datavendabq = ""
                    if vendabq is not None:
                        datavendabq = datetime.strptime(str(vendabq), "%Y-%m-%d").strftime("%d-%m-%Y")

                if resultadosoci:
                    qtdoci = resultadosoci[0][2]
                    vendaoci = resultadosoci[0][10]
                    datavendaoci = ""
                    if vendaoci is not None:
                        datavendaoci = datetime.strptime(str(vendaoci), "%Y-%m-%d").strftime("%d-%m-%Y")

                if resultadospet:
                    qtdpet = resultadospet[0][2]
                    vendapet = resultadospet[0][10]
                    datavendapet = ""
                    if vendapet is not None:
                        datavendapet = datetime.strptime(str(vendapet), "%Y-%m-%d").strftime("%d-%m-%Y")

                vvenda = locale.format_string("%.2f", preco, grouping=True)
                punitarioformat = locale.format_string("%.2f", punitario, grouping=True)
                ipieformat = locale.format_string("%.2f", ipie, grouping=True)
                icmsValorformat = locale.format_string("%.2f", icmsValor, grouping=True)
                ivastformat = locale.format_string("%.2f", ivast, grouping=True)
                mkupformat = locale.format_string("%.2f", mkup, grouping=True)
                ptotalformat = locale.format_string("%.2f", ptotal, grouping=True)
                if icms_difal is not None:
                    difalformat = locale.format_string("%.2f", icms_difal, grouping=True)
                else:
                    difalformat = "0,00"
                
                dados = [nomeprod, fornecedor, vvenda, qtdest, qtdoci,
                        qtdbq, qtdpet, secao, estmin, codprod, ipieformat, icmsValorformat,
                        ivastformat, ptotalformat, mkupformat, nnfe, complemento,
                        compradata, punitarioformat, qtdcompra, datavendabq, datavendaoci, datavendapet, difalformat]
                
                return dados
                
            else:
                raise ValueError
    
def consultaref(codigo, somente_pet=False):

    qtdoci = None
    qtdest = None
    qtdbq = None
    qtdpet = None
    resultadospet = None
    resultadosest = None
    vendabq = None
    datavendabq = None
    datavendaoci = None

    if len(str(codigo)) < 255:
        codigolimpo = html.escape(codigo)
        try:
            palavras = codigolimpo.upper().split()
            filtros_like = " AND ".join([f"PRODUTO LIKE ?" for _ in palavras])
            consulta_sql = f"SELECT {', '.join(COLUNAS_INTERESSE)} FROM PRODUTO WHERE ATIVO = 1 AND {filtros_like}"
            parametros = [f"%{palavra}%" for palavra in palavras]

            resultadosest = consulta_estoque(consulta_sql, parametros)
            if resultadosest:
                for row in resultadosest:
                    nomeprod = row[0]
                    preco = row[1]
                    qtdest = row[2]
                    codprod = row[3]
                    secao = row[4]
                    estmin = row[5]
                    complemento = row[6]
                    ptotal = row[7]
                    compradata = row[8]
                    qtdcompra = row[9]
                    fornecedor = row[11]
                    ipie = row[14]
                    icmsValor = row[15]
                    ivast = row[16]
                    punitario = row[17]
                    mkup = row[18]
                    nnfe = row[19]
                    icms_difal = row[21]

                if complemento is None:
                    complemento = ""

                consulta_sql = f"SELECT {', '.join(COLUNAS_INTERESSE)} FROM PRODUTO WHERE COD1 = ?"
                resultadosboq = consulta_boq(consulta_sql, (codprod,))
                resultadosoci = consulta_oci(consulta_sql, (codprod,))
                resultadospet = consulta_pet(consulta_sql, (codprod,))


                if resultadosboq:
                    qtdbq = resultadosboq[0][2]
                    vendabq = resultadosboq[0][10]
                    datavendabq = ""
                    if vendabq is not None:
                        datavendabq = datetime.strptime(str(vendabq), "%Y-%m-%d").strftime("%d-%m-%Y")

                if resultadosoci:
                    qtdoci = resultadosoci[0][2]
                    vendaoci = resultadosoci[0][10]
                    datavendaoci = ""
                    if vendaoci is not None:
                        datavendaoci = datetime.strptime(str(vendaoci), "%Y-%m-%d").strftime("%d-%m-%Y")

                if resultadospet:
                    qtdpet = resultadospet[0][2]

                vvenda = locale.format_string("%.2f", preco, grouping=True)
                punitarioformat = locale.format_string("%.2f", punitario, grouping=True)
                ipieformat = locale.format_string("%.2f", ipie, grouping=True)
                icmsValorformat = locale.format_string("%.2f", icmsValor, grouping=True)
                ivastformat = locale.format_string("%.2f", ivast, grouping=True)
                mkupformat = locale.format_string("%.2f", mkup, grouping=True)
                ptotalformat = locale.format_string("%.2f", ptotal, grouping=True)
                difalformat = locale.format_string("%.2f", icms_difal, grouping=True)
                
                dados = [nomeprod, fornecedor, vvenda, qtdest, qtdoci,
                        qtdbq, qtdpet, secao, estmin, codprod, ipieformat, icmsValorformat,
                        ivastformat, ptotalformat, mkupformat, nnfe, complemento,
                        compradata, punitarioformat, qtdcompra, datavendabq, datavendaoci, difalformat]
            
            elif somente_pet:
                palavras = codigolimpo.upper().split()
                filtros_like = " AND ".join([f"PRODUTO LIKE ?" for _ in palavras])
                consulta_sql = f"SELECT {', '.join(COLUNAS_INTERESSE)} FROM PRODUTO WHERE ATIVO = 1 AND {filtros_like}"
                parametros = [f"%{palavra}%" for palavra in palavras]

                resultadospet = consulta_pet(consulta_sql, parametros)
                for row in resultadospet:
                    nomeprod = row[0]
                    preco = row[1]
                    qtdpet = row[2]
                    codprod = row[3]
                    secao = row[4]
                    estmin = row[5]
                    complemento = row[6]
                    ptotal = row[7]
                    compradata = row[8]
                    qtdcompra = row[9]
                    fornecedor = row[11]
                    ipie = row[14]
                    icmsValor = row[15]
                    ivast = row[16]
                    punitario = row[17]
                    mkup = row[18]
                    nnfe = row[19]
                    icms_difal = row[21]

                if complemento is None:
                    complemento = ""
                vendapet = resultadospet[0][10]
                datavendapet = ""
                if vendapet is not None:
                    datavendapet = datetime.strptime(str(vendapet), "%Y-%m-%d").strftime("%d-%m-%Y")

                vvenda = locale.format_string("%.2f", preco, grouping=True)
                punitarioformat = locale.format_string("%.2f", punitario, grouping=True)
                ipieformat = locale.format_string("%.2f", ipie, grouping=True)
                icmsValorformat = locale.format_string("%.2f", icmsValor, grouping=True)
                ivastformat = locale.format_string("%.2f", ivast, grouping=True)
                mkupformat = locale.format_string("%.2f", mkup, grouping=True)
                ptotalformat = locale.format_string("%.2f", ptotal, grouping=True)
                difalformat = locale.format_string("%.2f", icms_difal, grouping=True)

                dados = [nomeprod, fornecedor, vvenda, "", "",
                        "", qtdpet, "", "", codprod, ipieformat, icmsValorformat,
                        ivastformat, ptotalformat, mkupformat, nnfe, complemento,
                        compradata, punitarioformat, qtdcompra, "", "", datavendapet, difalformat]
                
            else:
                raise ValueError

        except Exception as e:
            traceback.print_exc()

        return dados


def gerar_lista_secoes():
    """Gera planilhas de excel das secoes"""
    dirsecao = "planilhassecao/"
    secoes = ['A', 'A2', 'B', 'C', 'D', 'D2', 'D3', 'D4', 'E', 'F', 'G', 'H',
              'I', 'J', 'K', 'L', 'L2', 'L3', 'M', 'N', 'O', 'P', 'Q', 'R', 
              'S', 'T', 'U', 'U2', 'V', 'W', 'X', 'Y', 'Z', 'CORREDOR3', 
              'BOMBONIERE', 'SEPARACAO', 'EXTERNO', 'COZINHA', '3 ANDAR', 'GARAGEM',
              'ESCRITORIO']

    for nomesecao in secoes:
        if len(nomesecao) < 3:
            nomesecao = "SECAO " + nomesecao
        print(nomesecao, "...", end="")
        colunas_interesse = ['PRODUTO', 'VENDA', 'ESTOQUEFISICO', 'COD1',
                             'DEPT_NM', 'ESTOQUEMINIMO', 'TAM', 'COMPRA', 
                             'COMPRADT', 'QUATDE_NEW', 'VENDIDO', 'SMARCA', 
                             'ESTOQDEP', 'SGRUPO']

        resultadosest = consulta_estoque(f"""SELECT {', '.join(colunas_interesse)}
                                         FROM PRODUTO WHERE DEPT_NM = ?""",
                                         (nomesecao,))

        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.append(["Produto", "Codigo", "Estoque", "Boq", "Ocian", "Est Min", "Seção"])

        for rowest in resultadosest:
            nomeprod = rowest[0]
            qtdest = rowest[2]
            codigobarra = rowest[3]
            secao = rowest[4]
            estmin = rowest[5]

            codigoatual = rowest[3]
            resultadosboq = consulta_boq(f"""SELECT {', '.join(colunas_interesse)}
                                         FROM PRODUTO WHERE COD1 = ?""",
                                         (codigoatual,))

            resultadosoci = consulta_oci(f"""SELECT {', '.join(colunas_interesse)}
                                         FROM PRODUTO WHERE COD1 = ?""",
                                         (codigoatual,))

            if resultadosboq:
                qtdbq = resultadosboq[0][2]
                if qtdbq is None:
                    qtdbq = 0

            if resultadosoci:
                qtdoci = resultadosoci[0][2]
                if qtdoci is None:
                    qtdoci = 0

            sheet.append([nomeprod, codigobarra, qtdest, qtdbq, qtdoci, estmin, secao])

        wb.save(f"{dirsecao}{nomesecao}.xlsx")

        ordenarplanilha(f"{dirsecao}{nomesecao}.xlsx", 0)
        print(" Feito")

def produtos_por_nota(num_nota):
    if len(num_nota) > 8:
        raise ValueError("Numero de nota muito longo")
    consulta_sql = f"SELECT {', '.join(COLUNAS_INTERESSE)} FROM PRODUTO WHERE NF = ?"

    resultadosest = consulta_estoque(consulta_sql, (num_nota,))
    dados = []
    for rowest in resultadosest:
        nome_produto = rowest[0]
        codigo_produto = rowest[3]
        qtdestoque = rowest[2]
        secao = rowest[4]
        estmin = rowest[5]
        qtdcompra = str(rowest[9]) + 'UN'
        qtdboq = 0
        qtdoci = 0

        consulta_sql = "SELECT ESTOQUEFISICO FROM PRODUTO WHERE COD1 = ?"

        resultadosboq = consulta_boq(consulta_sql, (codigo_produto,))
        resultadosoci = consulta_oci(consulta_sql, (codigo_produto,))

        if resultadosboq:
            qtdboq = resultadosboq[0][0]
        else:
            qtdboq = 0

        if resultadosoci:
            qtdoci = resultadosoci[0][0]
        else:
            qtdoci = 0

        dados.append([nome_produto, codigo_produto, qtdestoque, qtdboq,
                      qtdoci, estmin, secao, qtdcompra])

    return dados

def gera_relatoriopet():
    consulta_sql = f"""SELECT {', '.join(COLUNAS_INTERESSE)} FROM PRODUTO
    WHERE ATIVO = 1 
    AND COMPRADT IS NOT NULL"""

    resultadospet = consulta_pet(consulta_sql)

    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.append(["Marca", "Produto", "Codigo", "Estoque", "Boq",
                  "Ocian", "Pet", "EstMin", "Seção", "DataCompra", "QtdCompra", 
                  "UltVendaEst", "UltVendaBoq", "UltVendaOci", "PCTVendidoEst",
                  "PCTVendidoBQ", "PCTVendidoOci", "PCTVendidoPET", "PCTVendidoTotal", "TotalVendaEst",
                  "TotalVendaBoq", "TotalVendaOci", "TotalVendaPet", "TotalVendaUltCompra", "VendaP30Dias",
                  "QTDMesPassado", "QTDMesRetrasado", "VendaUltSemana(PET)", "Grupo", "Venda"])

    consulta_sql = "SELECT COD1, DIA, QUANTIDADE FROM PRODUTOS_VENDIDOS WHERE STATUS = 1"

    print("Consulta Estoque Vendidos")
    vendasest = consulta_estoque(consulta_sql)
    vendas_por_produto_est = defaultdict(list)
    for cod, dia, qtd in vendasest:
        if isinstance(dia, str):
            dia = datetime.strptime(dia, "%Y-%m-%d").date()
        vendas_por_produto_est[cod].append((dia, qtd))

    print("Consulta Boq Vendidos")
    vendasboq = consulta_boq(consulta_sql)
    vendas_por_produto_boq = defaultdict(list)
    for cod, dia, qtd in vendasboq:
        if isinstance(dia, str):
            dia = datetime.strptime(dia, "%Y-%m-%d").date()
        vendas_por_produto_boq[cod].append((dia, qtd))

    print("Consulta Oci Vendidos")
    vendasoci = consulta_oci(consulta_sql)
    vendas_por_produto_oci = defaultdict(list)
    for cod, dia, qtd in vendasoci:
        if isinstance(dia, str):
            dia = datetime.strptime(dia, "%Y-%m-%d").date()
        vendas_por_produto_oci[cod].append((dia, qtd))

    print("Consulta Pet Vendidos")
    vendasoci = consulta_pet(consulta_sql)
    vendas_por_produto_pet = defaultdict(list)
    for cod, dia, qtd in vendasoci:
        if isinstance(dia, str):
            dia = datetime.strptime(dia, "%Y-%m-%d").date()
        vendas_por_produto_pet[cod].append((dia, qtd))

    # Consulta quantidade de devoluções por base
    consulta_sql = "SELECT PRODUTOCODBAR, DIA, QUANTIDADE FROM PRODUTOS_DEVOLUCAO"
    print("Consulta Boq Devolução")
    devolucaoboq = consulta_boq(consulta_sql)
    devolucao_por_produto_bq = defaultdict(list)
    for cod, dia, qtd in devolucaoboq:
        if isinstance(dia, str):
            dia = datetime.strptime(dia, "%Y-%m-%d").date()
        devolucao_por_produto_bq[cod].append((dia, qtd))

    print("Consulta Oci Devolução")
    devolucaooci = consulta_oci(consulta_sql)
    devolucao_por_produto_oci = defaultdict(list)
    for cod, dia, qtd in devolucaooci:
        if isinstance(dia, str):
            dia = datetime.strptime(dia, "%Y-%m-%d").date()
        devolucao_por_produto_oci[cod].append((dia, qtd))

    print("Consulta Pet Devolução")
    devolucaopet = consulta_pet(consulta_sql)
    devolucao_por_produto_pet = defaultdict(list)
    for cod, dia, qtd in devolucaopet:
        if isinstance(dia, str):
            dia = datetime.strptime(dia, "%Y-%m-%d").date()
        devolucao_por_produto_pet[cod].append((dia, qtd))

    #Consulta produto por base
    consulta_sql = f"SELECT {', '.join(COLUNAS_INTERESSE)} FROM PRODUTO"
    print("Consulta base Boq")
    resultadosboq = consulta_boq(consulta_sql)
    resultadosboq_df = pd.DataFrame(resultadosboq, columns=COLUNAS_INTERESSE)

    print("Consulta base Ocian")
    resultadosoci = consulta_oci(consulta_sql)
    resultadosoci_df = pd.DataFrame(resultadosoci, columns=COLUNAS_INTERESSE)

    print("Consulta base Est")
    resultadosest = consulta_estoque(consulta_sql)
    resultadosest_df = pd.DataFrame(resultadosest, columns=COLUNAS_INTERESSE)

    print("processando dados...")
    for rowest in resultadospet:
        produto = rowest[0]
        qtdpet = rowest[2]
        codigobarra = rowest[3]
        secao = rowest[4]
        estmin = rowest[5]
        timestampcompra = rowest[8]
        if timestampcompra is None:
            continue
        else:
            datacompra = timestampcompra.strftime("%d/%m/%Y")
            datacompra = datetime.strptime(datacompra, '%d/%m/%Y')
        quantidadecompra = rowest[9]
        valor_venda = rowest[1]
        if quantidadecompra in (None, Decimal('0')):
            continue

        if rowest[10]:
            dataultvenda_pet = rowest[10].strftime("%d/%m/%Y")
            dataultvenda_pet = datetime.strptime(dataultvenda_pet, '%d/%m/%Y')
        else:
            dataultvenda_est = None
        marca = rowest[11]

        # #Duplicar quantidade compra caso necessario
        # if marca in ("RAINHA",):
        #     quantidadecompra *= 2
            
        grupo = rowest[13]

        totalvenda_est = sum(qtd for dia, qtd in vendas_por_produto_est.get(codigobarra, []) if dia >= timestampcompra)
        totalvenda_bq = sum(qtd for dia, qtd in vendas_por_produto_boq.get(codigobarra, []) if dia >= timestampcompra)
        totalvenda_oci = sum(qtd for dia, qtd in vendas_por_produto_oci.get(codigobarra, []) if dia >= timestampcompra)
        totalvenda_pet = sum(qtd for dia, qtd in vendas_por_produto_pet.get(codigobarra, []) if dia >= timestampcompra)

        totaldevolucao_boq = sum(qtd for dia, qtd in devolucao_por_produto_bq.get(codigobarra, []) if dia >= timestampcompra)
        totaldevolucao_oci = sum(qtd for dia, qtd in devolucao_por_produto_oci.get(codigobarra, []) if dia >= timestampcompra)
        totaldevolucao_pet = sum(qtd for dia, qtd in devolucao_por_produto_pet.get(codigobarra, []) if dia >= timestampcompra)        

        #Vendas por dias estoque
        venda_7dias_est = sum(qtd for dia, qtd in vendas_por_produto_est.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=8))
        venda_15dias_est = sum(qtd for dia, qtd in vendas_por_produto_est.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=15))
        venda_30dias_est = sum(qtd for dia, qtd in vendas_por_produto_est.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=30))
        venda_60dias_est = sum(qtd for dia, qtd in vendas_por_produto_est.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=60))


        #Vendas por dias boqueirão
        venda_7dias_boq = sum(qtd for dia, qtd in vendas_por_produto_boq.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=8))
        devolucao_7dias_boq = sum(qtd for dia, qtd in devolucao_por_produto_bq.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=8))
        venda_7dias_boq = venda_7dias_boq - devolucao_7dias_boq

        venda_15dias_boq = sum(qtd for dia, qtd in vendas_por_produto_boq.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=15))
        devolucao_15dias_boq = sum(qtd for dia, qtd in devolucao_por_produto_bq.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=15))
        venda_15dias_boq = venda_15dias_boq - devolucao_15dias_boq

        venda_30dias_boq = sum(qtd for dia, qtd in vendas_por_produto_boq.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=30))
        devolucao_30dias_boq = sum(qtd for dia, qtd in devolucao_por_produto_bq.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=30))
        venda_30dias_boq = venda_30dias_boq - devolucao_30dias_boq

        venda_60dias_boq = sum(qtd for dia, qtd in vendas_por_produto_boq.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=60))
        devolucao_60dias_boq = sum(qtd for dia, qtd in devolucao_por_produto_bq.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=60))
        venda_60dias_boq = venda_60dias_boq - devolucao_60dias_boq

        #Vendas por dias ocian
        venda_7dias_oci = sum(qtd for dia, qtd in vendas_por_produto_oci.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=8))
        devolucao_7dias_oci = sum(qtd for dia, qtd in devolucao_por_produto_oci.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=8))
        venda_7dias_oci = venda_7dias_oci - devolucao_7dias_oci

        venda_15dias_oci = sum(qtd for dia, qtd in vendas_por_produto_oci.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=15))
        devolucao_15dias_oci = sum(qtd for dia, qtd in devolucao_por_produto_oci.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=15))
        venda_15dias_oci = venda_15dias_oci - devolucao_15dias_oci

        venda_30dias_oci = sum(qtd for dia, qtd in vendas_por_produto_oci.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=30))
        devolucao_30dias_oci = sum(qtd for dia, qtd in devolucao_por_produto_oci.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=30))
        venda_30dias_oci = venda_30dias_oci - devolucao_30dias_oci

        venda_60dias_oci = sum(qtd for dia, qtd in vendas_por_produto_oci.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=60))
        devolucao_60dias_oci = sum(qtd for dia, qtd in devolucao_por_produto_oci.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=60))
        venda_60dias_oci = venda_60dias_oci - devolucao_60dias_oci

        #Vendas por dias pet
        venda_7dias_pet = sum(qtd for dia, qtd in vendas_por_produto_pet.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=8))
        devolucao_7dias_pet = sum(qtd for dia, qtd in devolucao_por_produto_pet.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=8))
        venda_7dias_pet = venda_7dias_pet - devolucao_7dias_pet

        venda_15dias_pet = sum(qtd for dia, qtd in vendas_por_produto_pet.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=15))
        devolucao_15dias_pet = sum(qtd for dia, qtd in devolucao_por_produto_pet.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=15))
        venda_15dias_pet = venda_15dias_pet - devolucao_15dias_pet

        venda_30dias_pet = sum(qtd for dia, qtd in vendas_por_produto_pet.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=30))
        devolucao_30dias_pet = sum(qtd for dia, qtd in devolucao_por_produto_pet.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=30))
        venda_30dias_pet = venda_30dias_pet - devolucao_30dias_pet

        venda_60dias_pet = sum(qtd for dia, qtd in vendas_por_produto_pet.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=60))
        devolucao_60dias_pet = sum(qtd for dia, qtd in devolucao_por_produto_pet.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=60))
        venda_60dias_pet = venda_60dias_pet - devolucao_60dias_pet

        totalvenda_bq  = totalvenda_bq - totaldevolucao_boq
        totalvenda_oci = totalvenda_oci - totaldevolucao_oci
        totalvenda_pet = totalvenda_pet - totaldevolucao_pet

        # Calcula porcentagem de venda de cada base
        pctest = totalvenda_est / quantidadecompra
        pctbq = totalvenda_bq / quantidadecompra
        pctoci = totalvenda_oci / quantidadecompra
        pctpet = totalvenda_pet / quantidadecompra

        pct_7dias_total = (venda_7dias_est + venda_7dias_boq + venda_7dias_oci + venda_7dias_pet) / quantidadecompra
        pct_15dias_total = (venda_15dias_est + venda_15dias_boq + venda_15dias_oci + venda_15dias_pet) / quantidadecompra
        pct_30dias_total = (venda_30dias_est + venda_30dias_boq + venda_30dias_oci + venda_30dias_pet) / quantidadecompra
        pct_60dias_total = (venda_60dias_est + venda_60dias_boq + venda_60dias_oci + venda_60dias_pet) / quantidadecompra

        #Calcula porcentagem de venda da ultsemana(pet)
        timestamphoje = datetime.now()
        vendaultsemanapet = sum(qtd for dia, qtd in vendas_por_produto_pet.get(codigobarra, []) if dia >= timestamphoje - timedelta(days=7))
        qtdvendaultsemanapet = vendaultsemanapet

        #Calcular venda ultimos meses
        hoje = date.today()
        if hoje.month == 1:
            mespassado = 12
            mesretrasado = 11
            ano_anterior = hoje.year - 1
        else:
            mespassado = hoje.month - 1
            mesretrasado = hoje.month - 2
            ano_anterior = hoje.year

        if hoje.month == 2:
            mesretrasado = 12
        
        ultimodiamespassado = calendar.monthrange(ano_anterior, mespassado)[1]
        ultimodiamesretrasado = calendar.monthrange(ano_anterior, mesretrasado)[1]

        datamespassado_inicio = datetime.strptime(f"01/{mespassado}/{ano_anterior}", "%d/%m/%Y")
        datamespassado_fim = datetime.strptime(f"{ultimodiamespassado}/{mespassado}/{ano_anterior}", "%d/%m/%Y")
        datamesretrasado_inicio = datetime.strptime(f"01/{mesretrasado}/{ano_anterior}", "%d/%m/%Y")
        datamesretrasado_fim = datetime.strptime(f"{ultimodiamesretrasado}/{mesretrasado}/{ano_anterior}", "%d/%m/%Y")

        vendamespassado_est = sum(qtd for dia, qtd in vendas_por_produto_est.get(codigobarra, []) if dia >= datamespassado_inicio and dia <= datamespassado_fim)
        vendamesretrasado_est = sum(qtd for dia, qtd in vendas_por_produto_est.get(codigobarra, []) if dia >= datamesretrasado_inicio and dia <= datamesretrasado_fim)        
        vendamespassado_boq = sum(qtd for dia, qtd in vendas_por_produto_boq.get(codigobarra, []) if dia >= datamespassado_inicio and dia <= datamespassado_fim)
        vendamesretrasado_boq = sum(qtd for dia, qtd in vendas_por_produto_boq.get(codigobarra, []) if dia >= datamesretrasado_inicio and dia <= datamesretrasado_fim)
        vendamespassado_ocian = sum(qtd for dia, qtd in vendas_por_produto_oci.get(codigobarra, []) if dia >= datamespassado_inicio and dia <= datamespassado_fim)
        vendamesretrasado_ocian = sum(qtd for dia, qtd in vendas_por_produto_oci.get(codigobarra, []) if dia >= datamesretrasado_inicio and dia <= datamesretrasado_fim)
        vendamespassado_pet = sum(qtd for dia, qtd in vendas_por_produto_pet.get(codigobarra, []) if dia >= datamespassado_inicio and dia <= datamespassado_fim)
        vendamesretrasado_pet = sum(qtd for dia, qtd in vendas_por_produto_pet.get(codigobarra, []) if dia >= datamesretrasado_inicio and dia <= datamesretrasado_fim)

        vendamespassado_total =  vendamespassado_est + vendamespassado_boq + vendamespassado_ocian + vendamespassado_pet
        vendamesretrasado_total = vendamesretrasado_est + vendamesretrasado_boq + vendamesretrasado_ocian + vendamesretrasado_pet

        totalvendidos = totalvenda_est + totalvenda_bq + totalvenda_oci + totalvenda_pet

        try:
            pctvendido = totalvendidos/quantidadecompra
        except Exception:
            pctvendido = None

        filtroboq = resultadosboq_df[resultadosboq_df['COD1'] == codigobarra]
        filtrooci = resultadosoci_df[resultadosoci_df['COD1'] == codigobarra]
        filtroest = resultadosest_df[resultadosest_df['COD1'] == codigobarra]

        if not filtroboq.empty:
            qtdbq = filtroboq['ESTOQUEFISICO'].iloc[0] if pd.notna(filtroboq['ESTOQUEFISICO'].iloc[0]) else 0
            timestampultvendaboq = filtroboq['VENDIDO'].iloc[0] if pd.notna(filtroboq['VENDIDO'].iloc[0]) else 0
            if timestampultvendaboq != 0:
                dataultvenda_bq = timestampultvendaboq.strftime("%d/%m/%Y")
            else:
                dataultvenda_bq = None
        else:
            qtdbq = 0
            dataultvenda_bq = None
                
        if not filtrooci.empty:
            qtdoci = filtrooci['ESTOQUEFISICO'].iloc[0] if pd.notna(filtrooci['ESTOQUEFISICO'].iloc[0]) else 0
            timestampultvendaoci = filtrooci['VENDIDO'].iloc[0] if pd.notna(filtrooci['VENDIDO'].iloc[0]) else 0
            if timestampultvendaoci != 0:
                dataultvenda_ocian = timestampultvendaoci.strftime("%d/%m/%Y")
            else:
                dataultvenda_ocian = None
        else:
            qtdoci = 0
            dataultvenda_ocian = None

        if not filtroest.empty:
            qtdest = filtroest['ESTOQUEFISICO'].iloc[0] if pd.notna(filtroest['ESTOQUEFISICO'].iloc[0]) else 0
            timestampultvendaest = filtroest['VENDIDO'].iloc[0] if pd.notna(filtroest['VENDIDO'].iloc[0]) else 0
            if timestampultvendaest != 0:
                dataultvenda_est = timestampultvendaest.strftime("%d/%m/%Y")
            else:
                dataultvenda_est = None
        else:
            qtdest = 0
            dataultvenda_est = None

        sheet.append([marca, produto, codigobarra, qtdest, qtdbq, qtdoci, qtdpet, estmin, secao,
                    datacompra, quantidadecompra, dataultvenda_est, dataultvenda_bq,
                    dataultvenda_ocian, pctest, pctbq, pctoci, pctpet, pctvendido, totalvenda_est,
                    totalvenda_bq, totalvenda_oci, totalvenda_pet, totalvendidos, pct_30dias_total, vendamespassado_total,
                    vendamesretrasado_total, qtdvendaultsemanapet, grupo, valor_venda])


    color_scale_rule = ColorScaleRule(
        start_type="min", start_color="FFF8696B",  # Vermelho
        mid_type="percentile", mid_value=50, mid_color="FFFFEB84",  # Amarelo (percentil 50)
        end_type="max", end_color="FF63BE7B"  # Verde
    )
    sheet.conditional_formatting.add(f"O1:O{sheet.max_row}", color_scale_rule)
    sheet.conditional_formatting.add(f"P1:P{sheet.max_row}", color_scale_rule)
    sheet.conditional_formatting.add(f"Q1:Q{sheet.max_row}", color_scale_rule)
    sheet.conditional_formatting.add(f"R1:R{sheet.max_row}", color_scale_rule)
    sheet.conditional_formatting.add(f"S1:S{sheet.max_row}", color_scale_rule)

    for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=15, max_col=19):
        for cell in row:
            cell.number_format = numbers.FORMAT_PERCENTAGE_00

    colunas_data = [10, 12, 13, 14]

    for col in colunas_data:
        for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=col, max_col=col):
            for cell in row:
                cell.number_format = "DD/MM/YYYY"

    try:
        wb.save("relatorioestoquePET.xlsx")
    except PermissionError:
        wb.save("relatorioestoquePETNOVO.xlsx")
    finally:
        print("Concluido!")

def gera_relatorio():
    consulta_sql = f"""SELECT {', '.join(COLUNAS_INTERESSE)} FROM PRODUTO
    WHERE ATIVO = 1 
    AND COMPRADT IS NOT NULL"""

    resultadosest = consulta_estoque(consulta_sql)

    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.append(["Marca", "Produto", "Codigo", "Estoque", "Boq",
                  "Ocian", "EstMin", "Seção", "DataCompra", "QtdCompra", 
                  "UltVendaEst", "UltVendaBoq", "UltVendaOci", "PCTVendidoEst",
                  "PCTVendidoBQ", "PCTVendidoOci", "PCTVendidoTotal", "TotalVendaEst",
                  "TotalVendaBoq", "TotalVendaOci", "TotalVendaUltCompra", "Venda 7 dias",
                  "Venda 15 dias", "Venda 30 dias", "Venda 60 dias", "QTDMesPassado", "QTDMesRetrasado",
                  "VendaUltSemana(BOQ)", "VelVendaEst(30)", "VelVendaBq(30)", "VelVendaOci(30)", "Grupo", "Venda"])

    consulta_sql = "SELECT COD1, DIA, QUANTIDADE FROM PRODUTOS_VENDIDOS WHERE STATUS = 1"

    print("Consulta Estoque Vendidos")
    vendasest = consulta_estoque(consulta_sql)
    vendas_por_produto_est = defaultdict(list)
    for cod, dia, qtd in vendasest:
        if isinstance(dia, str):
            dia = datetime.strptime(dia, "%Y-%m-%d").date()
        vendas_por_produto_est[cod].append((dia, qtd))

    print("Consulta Boq Vendidos")
    vendasboq = consulta_boq(consulta_sql)
    vendas_por_produto_boq = defaultdict(list)
    for cod, dia, qtd in vendasboq:
        if isinstance(dia, str):
            dia = datetime.strptime(dia, "%Y-%m-%d").date()
        vendas_por_produto_boq[cod].append((dia, qtd))

    print("Consulta Oci Vendidos")
    vendasoci = consulta_oci(consulta_sql)
    vendas_por_produto_oci = defaultdict(list)
    for cod, dia, qtd in vendasoci:
        if isinstance(dia, str):
            dia = datetime.strptime(dia, "%Y-%m-%d").date()
        vendas_por_produto_oci[cod].append((dia, qtd))

    consulta_sql = "SELECT PRODUTOCODBAR, DIA, QUANTIDADE FROM PRODUTOS_DEVOLUCAO"

    print("Consulta Boq Devolução")
    devolucaoboq = consulta_boq(consulta_sql)
    devolucao_por_produto_bq = defaultdict(list)
    for cod, dia, qtd in devolucaoboq:
        if isinstance(dia, str):
            dia = datetime.strptime(dia, "%Y-%m-%d").date()
        devolucao_por_produto_bq[cod].append((dia, qtd))

    print("Consulta Boq Devolução")
    devolucaooci = consulta_oci(consulta_sql)
    devolucao_por_produto_oci = defaultdict(list)
    for cod, dia, qtd in devolucaooci:
        if isinstance(dia, str):
            dia = datetime.strptime(dia, "%Y-%m-%d").date()
        devolucao_por_produto_oci[cod].append((dia, qtd))

    print("Consulta base Boq")
    consulta_sql = f"SELECT {', '.join(COLUNAS_INTERESSE)} FROM PRODUTO"
    resultadosboq = consulta_boq(consulta_sql)
    resultadosboq_df = pd.DataFrame(resultadosboq, columns=COLUNAS_INTERESSE)

    print("Consulta base Ocian")
    consulta_sql = f"SELECT {', '.join(COLUNAS_INTERESSE)} FROM PRODUTO"
    resultadosoci = consulta_oci(consulta_sql)
    resultadosoci_df = pd.DataFrame(resultadosoci, columns=COLUNAS_INTERESSE)

    print("processando dados...")
    for rowest in resultadosest:
        produto = rowest[0]
        qtdest = rowest[2]
        codigobarra = rowest[3]
        secao = rowest[4]
        estmin = rowest[5]
        timestampcompra = rowest[8]
        datacompra = timestampcompra.strftime("%d/%m/%Y")
        datacompra = datetime.strptime(datacompra, '%d/%m/%Y')
        quantidadecompra = rowest[9]
        valor_venda = rowest[1]
        if quantidadecompra in (None, Decimal('0')):
            continue

        if rowest[10]:
            dataultvenda_est = rowest[10].strftime("%d/%m/%Y")
            dataultvenda_est = datetime.strptime(dataultvenda_est, '%d/%m/%Y')
        else:
            dataultvenda_est = None
        marca = rowest[11]

        # #Duplicar quantidade compra caso necessario
        # if marca in ("RAINHA",):
        #     quantidadecompra *= 2
            
        grupo = rowest[13]

        totalvenda_est = sum(qtd for dia, qtd in vendas_por_produto_est.get(codigobarra, []) if dia >= timestampcompra)
        totalvenda_bq = sum(qtd for dia, qtd in vendas_por_produto_boq.get(codigobarra, []) if dia >= timestampcompra)
        totalvenda_oci = sum(qtd for dia, qtd in vendas_por_produto_oci.get(codigobarra, []) if dia >= timestampcompra)

        totaldevolucao_boq = sum(qtd for dia, qtd in devolucao_por_produto_bq.get(codigobarra, []) if dia >= timestampcompra)
        totaldevolucao_oci = sum(qtd for dia, qtd in devolucao_por_produto_oci.get(codigobarra, []) if dia >= timestampcompra)

        #Vendas por dias estoque
        venda_7dias_est = sum(qtd for dia, qtd in vendas_por_produto_est.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=8))
        venda_15dias_est = sum(qtd for dia, qtd in vendas_por_produto_est.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=15))
        venda_30dias_est = sum(qtd for dia, qtd in vendas_por_produto_est.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=30))
        venda_60dias_est = sum(qtd for dia, qtd in vendas_por_produto_est.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=60))


        #Vendas por dias boqueirão
        venda_7dias_boq = sum(qtd for dia, qtd in vendas_por_produto_boq.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=8))
        devolucao_7dias_boq = sum(qtd for dia, qtd in devolucao_por_produto_bq.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=8))
        venda_7dias_boq = venda_7dias_boq - devolucao_7dias_boq

        venda_15dias_boq = sum(qtd for dia, qtd in vendas_por_produto_boq.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=15))
        devolucao_15dias_boq = sum(qtd for dia, qtd in devolucao_por_produto_bq.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=15))
        venda_15dias_boq = venda_15dias_boq - devolucao_15dias_boq

        venda_30dias_boq = sum(qtd for dia, qtd in vendas_por_produto_boq.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=30))
        devolucao_30dias_boq = sum(qtd for dia, qtd in devolucao_por_produto_bq.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=30))
        venda_30dias_boq = venda_30dias_boq - devolucao_30dias_boq

        venda_60dias_boq = sum(qtd for dia, qtd in vendas_por_produto_boq.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=60))
        devolucao_60dias_boq = sum(qtd for dia, qtd in devolucao_por_produto_bq.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=60))
        venda_60dias_boq = venda_60dias_boq - devolucao_60dias_boq

        #Vendas por dias ocian
        venda_7dias_oci = sum(qtd for dia, qtd in vendas_por_produto_oci.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=8))
        devolucao_7dias_oci = sum(qtd for dia, qtd in devolucao_por_produto_oci.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=8))
        venda_7dias_oci = venda_7dias_oci - devolucao_7dias_oci

        venda_15dias_oci = sum(qtd for dia, qtd in vendas_por_produto_oci.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=15))
        devolucao_15dias_oci = sum(qtd for dia, qtd in devolucao_por_produto_oci.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=15))
        venda_15dias_oci = venda_15dias_oci - devolucao_15dias_oci

        venda_30dias_oci = sum(qtd for dia, qtd in vendas_por_produto_oci.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=30))
        devolucao_30dias_oci = sum(qtd for dia, qtd in devolucao_por_produto_oci.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=30))
        venda_30dias_oci = venda_30dias_oci - devolucao_30dias_oci

        venda_60dias_oci = sum(qtd for dia, qtd in vendas_por_produto_oci.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=60))
        devolucao_60dias_oci = sum(qtd for dia, qtd in devolucao_por_produto_oci.get(codigobarra, []) if dia >= timestampcompra and dia < timestampcompra + timedelta(days=60))
        venda_60dias_oci = venda_60dias_oci - devolucao_60dias_oci

        totalvenda_bq  = totalvenda_bq - totaldevolucao_boq
        totalvenda_oci = totalvenda_oci - totaldevolucao_oci

        # Calcula porcentagem de venda de cada base
        pctest = totalvenda_est / quantidadecompra
        pctbq = totalvenda_bq / quantidadecompra
        pctoci = totalvenda_oci / quantidadecompra

        pct_7dias_total = (venda_7dias_est + venda_7dias_boq + venda_7dias_oci) / quantidadecompra
        pct_15dias_total = (venda_15dias_est + venda_15dias_boq + venda_15dias_oci) / quantidadecompra
        pct_30dias_total = (venda_30dias_est + venda_30dias_boq + venda_30dias_oci) / quantidadecompra
        pct_60dias_total = (venda_60dias_est + venda_60dias_boq + venda_60dias_oci) / quantidadecompra

        #Calcula porcentagem de venda da ultsemana(boq)
        timestamphoje = datetime.now()
        vendaultsemanabq = sum(qtd for dia, qtd in vendas_por_produto_boq.get(codigobarra, []) if dia >= timestamphoje - timedelta(days=7))
        qtdvendaultsemanaboq = vendaultsemanabq

        #Calcular venda ultimos meses
        hoje = date.today()
        if hoje.month == 1:
            mespassado = 12
            mesretrasado = 11
            ano_anterior = hoje.year - 1
        else:
            mespassado = hoje.month - 1
            mesretrasado = hoje.month - 2
            ano_anterior = hoje.year

        if hoje.month == 2:
            mesretrasado = 12
        
        ultimodiamespassado = calendar.monthrange(ano_anterior, mespassado)[1]
        ultimodiamesretrasado = calendar.monthrange(ano_anterior, mesretrasado)[1]

        datamespassado_inicio = datetime.strptime(f"01/{mespassado}/{ano_anterior}", "%d/%m/%Y")
        datamespassado_fim = datetime.strptime(f"{ultimodiamespassado}/{mespassado}/{ano_anterior}", "%d/%m/%Y")
        datamesretrasado_inicio = datetime.strptime(f"01/{mesretrasado}/{ano_anterior}", "%d/%m/%Y")
        datamesretrasado_fim = datetime.strptime(f"{ultimodiamesretrasado}/{mesretrasado}/{ano_anterior}", "%d/%m/%Y")        

        vendamespassado_boq = sum(qtd for dia, qtd in vendas_por_produto_boq.get(codigobarra, []) if dia >= datamespassado_inicio and dia <= datamespassado_fim)
        vendamesretrasado_boq = sum(qtd for dia, qtd in vendas_por_produto_boq.get(codigobarra, []) if dia >= datamesretrasado_inicio and dia <= datamesretrasado_fim)
        vendamespassado_ocian = sum(qtd for dia, qtd in vendas_por_produto_oci.get(codigobarra, []) if dia >= datamespassado_inicio and dia <= datamespassado_fim)
        vendamesretrasado_ocian = sum(qtd for dia, qtd in vendas_por_produto_oci.get(codigobarra, []) if dia >= datamesretrasado_inicio and dia <= datamesretrasado_fim)

        vendamespassado_total = vendamespassado_boq + vendamespassado_ocian
        vendamesretrasado_total = vendamesretrasado_boq + vendamesretrasado_ocian

        totalvendidos = totalvenda_est + totalvenda_bq + totalvenda_oci

        try:
            pctvendido = totalvendidos/quantidadecompra
        except Exception:
            pctvendido = None

        filtroboq = resultadosboq_df[resultadosboq_df['COD1'] == codigobarra]
        filtrooci = resultadosoci_df[resultadosoci_df['COD1'] == codigobarra]

        if not filtroboq.empty:
            qtdbq = filtroboq['ESTOQUEFISICO'].iloc[0] if pd.notna(filtroboq['ESTOQUEFISICO'].iloc[0]) else 0
            timestampultvendaboq = filtroboq['VENDIDO'].iloc[0] if pd.notna(filtroboq['VENDIDO'].iloc[0]) else 0
            if timestampultvendaboq != 0:
                dataultvenda_bq = timestampultvendaboq.strftime("%d/%m/%Y")
            else:
                dataultvenda_bq = None
                
        if not filtrooci.empty:
            qtdoci = filtrooci['ESTOQUEFISICO'].iloc[0] if pd.notna(filtrooci['ESTOQUEFISICO'].iloc[0]) else 0
            timestampultvendaoci = filtrooci['VENDIDO'].iloc[0] if pd.notna(filtrooci['VENDIDO'].iloc[0]) else 0
            if timestampultvendaoci != 0:
                dataultvenda_ocian = timestampultvendaoci.strftime("%d/%m/%Y")
            else:
                dataultvenda_ocian = None

        velvendaest = (venda_30dias_est / quantidadecompra) / 30
        velvendabq = (venda_30dias_boq / quantidadecompra) / 30
        velvendaoci = (venda_30dias_oci / quantidadecompra) / 30
    
        sheet.append([marca, produto, codigobarra, qtdest, qtdbq, qtdoci, estmin, secao,
                    datacompra, quantidadecompra, dataultvenda_est, dataultvenda_bq,
                    dataultvenda_ocian, pctest, pctbq, pctoci, pctvendido, totalvenda_est,
                    totalvenda_bq, totalvenda_oci, totalvendidos, pct_7dias_total,
                    pct_15dias_total, pct_30dias_total, pct_60dias_total, vendamespassado_total, vendamesretrasado_total,
                    qtdvendaultsemanaboq, velvendaest, velvendabq, velvendaoci, grupo, valor_venda])


    color_scale_rule = ColorScaleRule(
        start_type="min", start_color="FFF8696B",  # Vermelho
        mid_type="percentile", mid_value=50, mid_color="FFFFEB84",  # Amarelo (percentil 50)
        end_type="max", end_color="FF63BE7B"  # Verde
    )
    sheet.conditional_formatting.add(f"N1:N{sheet.max_row}", color_scale_rule)
    sheet.conditional_formatting.add(f"O1:O{sheet.max_row}", color_scale_rule)
    sheet.conditional_formatting.add(f"P1:P{sheet.max_row}", color_scale_rule)
    sheet.conditional_formatting.add(f"Q1:Q{sheet.max_row}", color_scale_rule)

    for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=14, max_col=17):
        for cell in row:
            cell.number_format = numbers.FORMAT_PERCENTAGE_00

    for row in sheet.iter_rows(min_row=1, max_row=sheet.max_row, min_col=22, max_col=25):
        for cell in row:
            cell.number_format = "0.000%"

    colunas_data = [9, 11, 12, 13]

    for col in colunas_data:
        for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=col, max_col=col):
            for cell in row:
                cell.number_format = "DD/MM/YYYY"

    try:
        wb.save("relatorioestoque.xlsx")
    except PermissionError:
        wb.save("relatorioestoqueNOVO.xlsx")
    finally:
        print("Concluido!")

def melhoresfimano():
    consulta_sql = """SELECT
    MAX(NOME) AS NOME,
    COD1,
    SUM(QUANTIDADE) AS TOTAL
FROM PRODUTOS_VENDIDOS
WHERE STATUS = 1 AND
  DIA >= CAST((EXTRACT(YEAR FROM CURRENT_DATE) - 1) || '-12-01' AS DATE)
  AND DIA <  CAST(EXTRACT(YEAR FROM CURRENT_DATE) || '-02-01' AS DATE)
GROUP BY COD1
ORDER BY TOTAL DESC;"""

    resultadosest = consulta_boq(consulta_sql)

    # 1) Gera o arquivo vendasfimano.xlsx
    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = "Vendas"
    sheet.append(["Produto", "COD", "TOTAL"])

    for rowest in resultadosest:
        qtdvenda = rowest[2]
        sheet.append(rowest)

    wb.save("vendasfimano.xlsx")

    # 2) Lê as planilhas
    df_est = pd.read_excel("relatorioestoque.xlsx")
    df_vendas = pd.read_excel("vendasfimano.xlsx")

    palavras_excluir = ["TOALHA", "REVISTA", "FLOR ARTIFICIAL"]
    padrao = "|".join(palavras_excluir)

    df_vendas = df_vendas[
        ~df_vendas["Produto"]
            .str.contains(padrao, case=False, na=False)
    ]

    # 3) Ajusta nomes e tipos das colunas
    # vendas: COD -> COD1 (pra casar com estoque)
    df_vendas = df_vendas.rename(columns={"COD": "COD1"})
    df_est = df_est.rename(columns={"Codigo": "COD1"})

    # Normaliza COD1
    df_est["COD1"] = df_est["COD1"].astype(str).str.strip()
    df_vendas["COD1"] = df_vendas["COD1"].astype(str).str.strip()

    # Garante numéricos
    for col in ["Estoque", "Boq", "Ocian"]:
        df_est[col] = pd.to_numeric(df_est[col], errors="coerce").fillna(0)

    df_vendas["TOTAL"] = pd.to_numeric(df_vendas["TOTAL"], errors="coerce").fillna(0)

    # 4) Soma estoque
    df_est["ESTOQUE_SOMADO"] = df_est["Estoque"] + df_est["Boq"] + df_est["Ocian"]

    # 5) Merge (JOIN) pelo COD1
    df_final = df_vendas.merge(
        df_est[["COD1", "Marca", "Estoque", "Boq", "Ocian", "ESTOQUE_SOMADO"]],
        on="COD1",
        how="left"
    )

    # Se não achou o código no relatório de estoque, vira 0 (ou você pode marcar como "NA")
    df_final[["Estoque", "Boq", "Ocian", "ESTOQUE_SOMADO"]] = (
        df_final[["Estoque", "Boq", "Ocian", "ESTOQUE_SOMADO"]].fillna(0)
    )

    # 6) Comparação
    df_final["OK"] = df_final["ESTOQUE_SOMADO"] >= df_final["TOTAL"]
    df_final["FALTA"] = (df_final["TOTAL"] - df_final["ESTOQUE_SOMADO"]).clip(lower=0)

    # 7) Ordena (maior falta primeiro, depois maior venda)
    df_final = df_final.sort_values(["FALTA", "TOTAL"], ascending=[False, False])

    # 8) Salva relatórios
    df_final.to_excel("comparativo_vendas_x_estoque.xlsx", index=False)

    df_faltando = df_final.loc[df_final["FALTA"] > 0].copy()
    df_faltando.to_excel("faltando_estoque.xlsx", index=False)

    return df_final

def listacods(lista):
    dados = []
    for codigo in lista:
        consulta_sql = f"SELECT {', '.join(COLUNAS_INTERESSE)} FROM PRODUTO WHERE COD1 = ?"
        resultadosest = consulta_estoque(consulta_sql, (codigo,))
        for rowest in resultadosest:
            nome_produto = rowest[0]
            codigo_produto = rowest[3]
            qtdestoque = rowest[2]
            secao = rowest[4]
            estmin = rowest[5]
            qtdcompra = str(rowest[9]) + 'UN'
            qtdboq = 0
            qtdoci = 0

            consulta_sql = "SELECT ESTOQUEFISICO FROM PRODUTO WHERE COD1 = ?"

            resultadosboq = consulta_boq(consulta_sql, (codigo_produto,))
            resultadosoci = consulta_oci(consulta_sql, (codigo_produto,))

            if resultadosboq:
                qtdboq = resultadosboq[0][0]
            else:
                qtdboq = 0

            if resultadosoci:
                qtdoci = resultadosoci[0][0]
            else:
                qtdoci = 0

            dados.append([nome_produto, codigo_produto, qtdestoque, qtdboq,
                        qtdoci, estmin, secao, qtdcompra])
            
    return dados