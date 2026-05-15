import decimal
import shutil
import fdb
from pathlib import Path
import os
from dotenv import load_dotenv
import traceback
from decimal import Decimal
from datetime import datetime, timedelta, date
from time import sleep
import openpyxl
import modconV2 as mc

load_dotenv()  # Carrega o .env

def con_db(path_env_var):
    dsn = f"{os.getenv('FB_HOST')}:{os.getenv(path_env_var)}"
    return fdb.connect(
        dsn=dsn,
        user=os.getenv("FB_USER"),
        password=os.getenv("FB_PASSWORD")
    )

def conestoque():
    return con_db("FB_DB_ESTOQUE")

def conboq():
    return con_db("FB_DB_BOQ")

def conoci():
    return con_db("FB_DB_OCI")

def conpet():
    return con_db("FB_DB_PET")

PASTA = Path(r"listas/")
MUDANCA_SECAO_DIR = Path(os.getenv("MUDANCA_SECAO_DIR", "data/mudanca-secao"))
HISTORICO_TRANSFERENCIA_DIR = Path(os.getenv("HISTORICO_TRANSFERENCIA_DIR", "data/historico-transferencia"))

def processaLista(arquivo):
    totais = {}
    with open(arquivo, "r", encoding="utf-8") as f:
        for linha in f:
            codigo, qtd = linha.strip().split(";")
            if "," in qtd:
                qtd = qtd.replace(",", ".")
            qtd = Decimal(qtd)

            if codigo in totais:
                totais[codigo] += qtd
            else:
                totais[codigo] = qtd

    return totais

def ajustaEstoque(cur, cod, num, ajuste=False, saida=False, arquivo_origem = None, user="Padrao"):
    # 1. Buscar produto
    sql = "SELECT ID, ESTOQUEFISICO FROM PRODUTO WHERE COD1 = ?"
    cur.execute(sql, (cod,))
    resultado = cur.fetchone()

    if not resultado:
        raise TypeError("Produto não encontrado!")

    if ajuste is False:
        if saida is True:
            estoquenovo = resultado[1] - num
            descricao = f"Saída - {arquivo_origem}"
        else:
            estoquenovo = resultado[1] + num
            descricao = f"Entrada - {arquivo_origem}"
        descricao = descricao[:50]
    else:
        estoquenovo = num
        descricao = f"Ajuste - {user}"

    # 3. Inserir no Log de Ajuste
    sql = 'SELECT GEN_ID("PRODUTO_ESTOQUE_AJUSTADO", 1) FROM RDB$DATABASE'
    cur.execute(sql)
    row = cur.fetchone()
    novo_id = row[0]
    sql = """
    INSERT INTO PRODUTO_ESTOQUE_AJUSTADO
        (ID, PRODUTO, DATA, USUARIO, ESTOQUE_OLD, ESTOQUE_NEW, DESCRICAO)
    VALUES (?, ?, ?, ?, ?, ?, ?)
    """
    agora = datetime.now()
    usuario = 1  # como você já vinha usando

    cur.execute(sql, (
        novo_id,
        resultado[0],      # ID do produto
        agora,
        usuario,
        resultado[1],      # estoque antigo
        estoquenovo,       # estoque novo
        descricao
    ))

    # 4. Atualizar Estoque
    sql = "UPDATE PRODUTO SET ESTOQUEFISICO = ? WHERE COD1 = ?"
    cur.execute(sql, (estoquenovo, cod))

def alterasecao(codigo, secao):
    try:
        with conestoque() as con:
            cur = con.cursor()
            sqldep = "SELECT ID FROM DEPARTAMENTO WHERE DEPARTAMENTO = ?"
            cur.execute(sqldep, (secao,))
            depid = cur.fetchone()

            if depid is None:
                raise ValueError(f"Seção '{secao}' não encontrada!")

            depid = depid[0]

            sql = "UPDATE PRODUTO SET DEPT_NM = ?, DEPT_ID = ? WHERE COD1 = ?"
            cur.execute(sql, (secao, depid, codigo))

            if cur.rowcount == 0:
                raise ValueError(f"{codigo} não encontrado!")
        
            con.commit()

    except Exception as e:
        newdir = r"\\DESKTOP-UA78AAV\Desktop\\"
        now = datetime.now()
        nome_arquivo = f"Erro Alterar Secao {now.day}-{now.month}.txt"
        caminho_arquivo = os.path.join(newdir, nome_arquivo)
        with open(caminho_arquivo, 'a', encoding="utf-8") as arquivo:
            arquivo.write(f"Erro ao alterar seção | COD1={codigo} | SECAO={secao} | ERRO={e}\n")

def alteraestmin(codigo, estmin):
    try:
        with conestoque() as con:
            cur = con.cursor()
            sql = "UPDATE PRODUTO SET ESTOQUEMINIMO = ? WHERE COD1 = ?"
            cur.execute(sql, (estmin, codigo))

            if cur.rowcount == 0:
                raise ValueError(f"{codigo} não encontrado!")
            
            con.commit()

    except Exception as e:
        try:
            newdir = r"\\DESKTOP-UA78AAV\Desktop\\"
            now = datetime.now()
            nome_arquivo = f"Erro Alterar Secao {now.day}-{now.month}.txt"
            caminho_arquivo = os.path.join(newdir, nome_arquivo)
            with open(caminho_arquivo, 'a', encoding="utf-8") as arquivo:
                arquivo.write(f"Erro ao alterar seção | COD1={codigo} | EstMin={estmin} | ERRO={e}\n")
        except Exception as e:
            print("Falha ao criar o arquivo de erro!")

def passarEstoque(padrao="*.txt", tmin=0):
    # SAIDAS
    for arquivo in PASTA.rglob(padrao):
        agora = datetime.now()
        tempo_minimo = timedelta(minutes=tmin)
        mtime = datetime.fromtimestamp(arquivo.stat().st_mtime)

        if agora - mtime < tempo_minimo:
            continue # Arquivo novo ignora

        if "Mudança" in arquivo.name:
            with open(arquivo, 'r', encoding='utf-8') as arq:
                for linha in arq:
                    elementos = linha.split("-")
                    codigo = elementos[0].strip()
                    secao = elementos[1].upper().replace("Ã", "A").replace("Ç", "C").strip()
                    alterasecao(codigo, secao)

            dest_dir = MUDANCA_SECAO_DIR
            destino = dest_dir / arquivo.name
            contador = 1

            while True:
                try:
                    shutil.move(str(arquivo.resolve()), str(destino))
                    break
                except shutil.Error as e:
                    if "already exists" in str(e):
                        destino = dest_dir / f"{arquivo.stem}_{contador}.txt"
                        contador += 1
                    else:
                        raise
            continue

        if "Ajuste EstMin" in arquivo.name:
            with open(arquivo, 'r', encoding='utf-8') as arq:
                for linha in arq:
                    elementos = linha.split("-")
                    codigo = elementos[0].strip()
                    estmin = float(elementos[1].strip())
                    alteraestmin(codigo, estmin)
            shutil.move(str(arquivo.resolve()), str(MUDANCA_SECAO_DIR))
            continue

        saida = False
        if " - saiu" in arquivo.name or "ENTRADA" in arquivo.name:
            continue

        try:
            if "SAIDA" in arquivo.name:
                nome_saida = arquivo.stem.split(" ")[1]
            else:
                nomes = arquivo.stem.split("-")
                nome_saida = nomes[0].strip()

            if nome_saida == "ESTOQUE":
                dados = processaLista(arquivo)
                with conestoque() as con:
                    cur = con.cursor()
                    for codigo, quantidade in dados.items():
                        ajustaEstoque(cur, codigo, quantidade, saida=True, arquivo_origem=arquivo.stem)
                    con.commit()
                    saida = True

            elif nome_saida == "BOQUEIRAO" or nome_saida == "BOQ":
                dados = processaLista(arquivo)
                with conboq() as con:
                    cur = con.cursor()
                    for codigo, quantidade in dados.items():
                        ajustaEstoque(cur, codigo, quantidade, saida=True, arquivo_origem=arquivo.stem)
                    con.commit()
                    saida = True

            elif nome_saida == "OCIAN":
                dados = processaLista(arquivo)
                with conoci() as con:
                    cur = con.cursor()
                    for codigo, quantidade in dados.items():
                        ajustaEstoque(cur, codigo, quantidade, saida=True, arquivo_origem=arquivo.stem)
                    con.commit()
                    saida = True

            elif nome_saida == "HAPPY" or nome_saida == "HAPPY PET":
                dados = processaLista(arquivo)
                with conpet() as con:
                    cur = con.cursor()
                    for codigo, quantidade in dados.items():
                        ajustaEstoque(cur, codigo, quantidade, saida=True, arquivo_origem=arquivo.stem)
                    con.commit()
                    saida = True

            if saida is True:
                # Renomeia o arquivo para marcar que já teve saída ou entrada
                print(f"{arquivo.name} RETIRADO COM SUCESSO--")
                base_nome = arquivo.stem + " - saiu"
                sufixo = ".txt"

                contador = 0
                while True:
                    if contador == 0:
                        novo_nome = f"{base_nome}{sufixo}"
                    else:
                        novo_nome = f"{base_nome}_{contador}{sufixo}"

                    arquivo_renomeado = arquivo.parent / novo_nome

                    if not arquivo_renomeado.exists():
                        break

                    contador += 1

                arquivo.rename(arquivo_renomeado)
        except decimal.InvalidOperation:
            try:
                print("Quantidade inválida, movendo para o desktop")
                shutil.move(str(arquivo.resolve()), fr"\\DESKTOP-UA78AAV\Desktop\\{arquivo.stem} (quantidade inválida).txt")
            except Exception as e:
                print("Erro ao mover o arquivo: ", e)
        except TypeError:
            try:
                print("Produto não encontrado, movendo para o desktop")
                shutil.move(str(arquivo.resolve()), fr"\\DESKTOP-UA78AAV\Desktop\\{arquivo.stem} (produto não encontrado).txt")
            except Exception as e:
                print("Erro ao mover o arquivo: ", e)
        except Exception as e:
            traceback.print_exc()

    # ENTRADAS
    for arquivo in PASTA.rglob(padrao):
        agora = datetime.now()
        tempo_minimo = timedelta(minutes=tmin)
        mtime = datetime.fromtimestamp(arquivo.stat().st_mtime)

        if agora - mtime < tempo_minimo:
            continue # Arquivo novo ignora

        entrada = False
        try:
            if " - saiuentrou" in arquivo.name or " - entrou" in arquivo.name or "SAIDA" in arquivo.name:
                continue

            if "ENTRADA" in arquivo.name:
                nome_entrada = arquivo.stem.split(" ")[1]
            else:
                nomes = arquivo.stem.split("-")
                nome_entrada = nomes[1].split(" ")[1].strip()

            if nome_entrada == "ESTOQUE":
                dados = processaLista(arquivo)
                with conestoque() as con:
                    cur = con.cursor()
                    for codigo, quantidade in dados.items():
                        ajustaEstoque(cur, codigo, quantidade, arquivo_origem=arquivo.stem)
                    con.commit()
                    entrada = True

            elif nome_entrada == "BOQUEIRAO" or nome_entrada == "BOQ":
                dados = processaLista(arquivo)
                with conboq() as con:
                    cur = con.cursor()
                    for codigo, quantidade in dados.items():
                        ajustaEstoque(cur, codigo, quantidade, arquivo_origem=arquivo.stem)
                    con.commit()
                    entrada = True

            elif nome_entrada == "OCIAN":
                dados = processaLista(arquivo)
                with conoci() as con:
                    cur = con.cursor()
                    for codigo, quantidade in dados.items():
                        ajustaEstoque(cur, codigo, quantidade, arquivo_origem=arquivo.stem)
                    con.commit()
                    entrada = True

            elif nome_entrada == "HAPPY":
                dados = processaLista(arquivo)
                with conpet() as con:
                    cur = con.cursor()
                    for codigo, quantidade in dados.items():
                        ajustaEstoque(cur, codigo, quantidade, arquivo_origem=arquivo.stem)
                    con.commit()
                    entrada = True

            if entrada is True:
                print(f"{arquivo.stem} - ENTRADO COM SUCESSO++")
                nomearquivo = arquivo.stem

                if " - saiu" in arquivo.name:
                    novo_nome = nomearquivo + "entrou" + ".txt"
                else:
                    novo_nome = nomearquivo + " - entrou" + ".txt"

                # Renomeia o arquivo para marcar que já teve entrada

                contador = 0
                while True:
                    if contador > 0:
                        novo_nome = f"{novo_nome}_{contador}.txt"

                    arquivo_renomeado = arquivo.parent / novo_nome

                    if not arquivo_renomeado.exists():
                        break

                    contador += 1

                arquivo.rename(arquivo_renomeado)
        except decimal.InvalidOperation:
            print("Quantidade inválida, movendo para o desktop")
            try:    
                shutil.move(str(arquivo.resolve()), fr"\\DESKTOP-UA78AAV\Desktop\\{arquivo.stem} (quantidade inválida).txt")
            except Exception as e:
                print("Erro ao mover o arquivo: ", e)
        except TypeError:
            print("Produto não encontrado, movendo para o desktop")
            try:
                shutil.move(str(arquivo.resolve()), fr"\\DESKTOP-UA78AAV\Desktop\\{arquivo.stem} (produto não encontrado).txt")
            except Exception as e:
                print("Erro ao mover o arquivo: ", e)
        except Exception as e:
            traceback.print_exc()
        
    destino = HISTORICO_TRANSFERENCIA_DIR
    for arquivo in PASTA.rglob(padrao):
        if " - saiuentrou" in arquivo.name or ("ENTRADA" in arquivo.name and " - entrou" in arquivo.name) or ("SAIDA" in arquivo.name and " - saiu" in arquivo.name):
            destino_final = destino / arquivo.name
            contador = 1

            while destino_final.exists():
                destino_final = destino / f"{arquivo.stem}_{contador}{arquivo.suffix}"
                contador += 1

            shutil.move(str(arquivo), str(destino_final))
            print(arquivo.stem, " movido")

if __name__ == "__main__":
    while True:
        now = datetime.now()
        hoje = date.today()
        if now.hour < 19:
            # passarEstoque("*Mateus*.txt")
            passarEstoque("*SAIDA HAPPY PET*.txt", 2)
            passarEstoque("*ENTRADA HAPPY PET*.txt", 2)
            passarEstoque("*SAIDA*.txt", 5)
            passarEstoque("*ENTRADA*.txt", 5)
            passarEstoque("*TROCAS*.txt", 5)
            passarEstoque("* - ESTOQUE*.txt", 5)
        elif now.hour >= 19:
            passarEstoque()

        for pasta in Path(r"listas").iterdir():
            if not pasta.is_dir():
                continue
            try:
                data_pasta = datetime.strptime(pasta.name, "%d-%m-%Y").date()
            except ValueError:
                continue

            if data_pasta >= hoje:
                continue
            
            PASTA_OLD = PASTA
            try:
                PASTA = pasta
                passarEstoque("*.txt")
            finally:
                PASTA = PASTA_OLD

            # tenta apagar subpastas vazias primeiro (de baixo para cima)
            subdirs = [p for p in pasta.rglob("*") if p.is_dir()]
            subdirs.sort(key=lambda p: len(p.parts), reverse=True)
            for d in subdirs:
                try:
                    if not any(d.iterdir()):
                        d.rmdir()
                except OSError:
                    pass  # não estava vazia, ou sem permissão

            # apaga a pasta do dia se estiver vazia
            try:
                if not any(pasta.iterdir()):
                    pasta.rmdir()
                    print(f"Removida pasta vazia: {pasta}")
            except OSError:
                pass

        sleep(10)
