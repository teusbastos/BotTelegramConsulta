import traceback
import math
import pytz
import requests
import locale
import os
import json
import pandas as pd
import subprocess
from PIL import Image
from pyzbar.pyzbar import decode
import openpyxl
import time
from time import sleep, strftime, strptime
import datetime
import re
import threading
import glob
from telegram.ext import Updater, MessageHandler, Filters, CallbackContext, CommandHandler, CallbackQueryHandler, JobQueue
from telegram import Update, ParseMode, Bot, InlineKeyboardButton, InlineKeyboardMarkup, ReplyKeyboardRemove
from openpyxl.styles import PatternFill
from dotenv import load_dotenv
from utils.bot_funcs import insereTabelaBusca, mostraestoque, preparalista, strip_html_tags
import modconV2 as mc
import ajustaestoque as ae

locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
brasilia_tz = pytz.timezone("America/Sao_Paulo")

TESTE = True
load_dotenv()

def _env_int(name, default=0):
    value = os.getenv(name)
    return int(value) if value else default

def _env_int_tuple(name):
    value = os.getenv(name, "")
    return tuple(int(item.strip()) for item in value.split(",") if item.strip())

if TESTE is False:
    TOKEN = os.getenv("TOKEN_TESTE")
    print(('-=-'*5) + 'Bot Consulta Estoque -- MODO TESTE' + ('-=-'*5))
else:
    TOKEN = os.getenv("TOKEN_PRINCIPAL")
    print(('-=-'*5) + 'Bot Consulta Estoque -- Main' + ('-=-'*5))

barcode_data = [0]
lock = threading.Lock()
ordenacao = 'Boqueirao'

consultabot = Bot(token=TOKEN)
SENHA = os.getenv("SENHA")
arqblock = os.path.join(os.getcwd(), 'block.txt')
ID_ADMIN = _env_int("ID_ADMIN")
IDGRUPOBUSCA = _env_int("ID_GRUPO_BUSCA")
IDGRUPOPRODUTOS = _env_int("ID_GRUPO_PRODUTOS")
JSON_PATH = "topicos.json"
IDS_USERS_BUSCA = _env_int_tuple("IDS_USERS_BUSCA")
IDS_USERS_COMPRA = _env_int_tuple("IDS_USERS_COMPRA")

# Os últimos 3 elementos da lista serão apagados para mostrar no comando MUDAR SECAO
SECOES = ['Seção A', 'Seção B', 'Seção C', 'Seção D', 'Seção D2', 'Seção D3',
          'Seção D4', 'Seção E', 'Seção F', 'Seção G', 'Seção H', 'Seção I', 'Seção J',
          'Seção K', 'Seção L', 'Seção L2', 'Seção L3', 'Seção M', 'Seção N',
          'Seção O', 'Seção P', 'Seção Q', 'Seção R', 'Seção S', 'Seção T',  'Seção U',
          'Seção V', 'Seção W', 'Seção X', 'Seção Y', 'Seção Z', 'Corredor3',
          'Quartinho', 'Bomboniere', 'Cozinha', 'Externo', '3 Andar', 'Separação', 
          'Garagem', 'Escritorio', 'Loja', 'OCIANBOQ', 'BOQOCIAN', 'Relatorio']

def filetotuple(arquivo):
    dados = []
    with open(arquivo, 'r') as f:
        for linha in f:
            linha = linha.strip()
            if not linha.startswith('#') and ':' in linha:
                partes = linha.split(':')
                dados_tupla = (int(partes[0]), int(partes[1]), int(partes[2]))
                dados.append(dados_tupla)

    return dados

def get_color_from_value(qtdvendida, min_venda=0, max_venda=100):
    # Limita os valores dentro da faixa
    qtdvendida = max(min_venda, min(qtdvendida, max_venda))
    
    # Calcula intensidade: 0 (claro) → 255 (escuro)
    # Aqui usamos vermelho fixo e variamos o verde e azul
    intensity = int(255 - (qtdvendida - min_venda) / (max_venda - min_venda) * 200)  # escurece até 55

    red = 255
    green = intensity
    blue = intensity

    # Converte para string hexadecimal
    color_hex = f"{red:02X}{green:02X}{blue:02X}"
    return PatternFill(start_color=color_hex, end_color=color_hex, fill_type="solid")

class MessageSender:
    def __init__(self, update, context):
        self.update = update
        self.context = context
        self.chat_id = update.effective_chat.id
        self.username = (update.message.from_user.first_name if update.message else update.callback_query.from_user.first_name)

    def send(self, text, botao=False, keyboard=None, log=True, replyid=None):
        try:
            # Monta teclado padrão se necessário
            if keyboard is None and botao:
                keyboard = [
                    [InlineKeyboardButton("Mudar Seção", callback_data="mudasecao")],
                    [InlineKeyboardButton("Pedir", callback_data="pedir"),
                     InlineKeyboardButton("Lista Impressão", callback_data="listaimpressao")]
                ]
            reply_markup = InlineKeyboardMarkup(keyboard) if keyboard else None
            # Log opcional da mensagem
            if log:
                messlog = text.replace('\n', ' ')
                messlog = strip_html_tags(messlog)
                with open(f'logs/ConsultaEstoque/ce.chat_log-{self.username}.txt', 'a', encoding='utf-8') as log_file:
                    log_file.write(f'{datetime.datetime.now():%d/%m/%y - %H:%M:%S} Bot: {messlog}\n')

            if replyid:
                self.update.message.reply_text(
                    text,
                    reply_to_message_id=replyid,
                    parse_mode="html",
                    reply_markup=reply_markup
                )
            else:
                self.context.bot.send_message(
                    chat_id=self.chat_id,
                    text=text,
                    parse_mode="html",
                    reply_markup=reply_markup
                )

        except Exception as e:
            print(f"Erro ao enviar mensagem: {e}")
            traceback.print_exc()

def button_callback(update, context):
    username = update.callback_query.from_user.first_name

    query = update.callback_query

    if query.data == 'mudasecao':
        bot = MessageSender(update, context)
        mensagem = query.message.text.split('\n')
        context.user_data['codigo'] = mensagem[-1]
        context.user_data['nomeprod'] = mensagem[0]
        keyboard = []
        secoes = SECOES[0:-3]
        for secao in secoes:
            button = InlineKeyboardButton(secao, callback_data=f'{secao}')

            # Organize os botões em três linhas
            if len(keyboard) == 0 or len(keyboard[-1]) == 4:
                    keyboard.append([button])
            else:
                keyboard[-1].append(button)

        keyboard.append([InlineKeyboardButton("Voltar", callback_data='voltar')])

        reply_markup = InlineKeyboardMarkup(keyboard)

        query.edit_message_reply_markup(reply_markup=reply_markup)

    elif query.data.startswith('atualizarnota'):
        partes = query.data.split(' ', 1)
        todosnota(update, context, partes[1])
        return

    elif query.data == "listaimpressao":
        username = context.user_data['username']
        mensagem = query.message.text.split('\n')
        nomeproduto = mensagem[0]
        codigoproduto = mensagem[-1]

        posicao = nomeproduto.find("|")

        # Eliminar tudo após "|" e um caractere antes dela
        if posicao != -1:  # Certificar-se de que "|" existe na string
            nomeproduto = nomeproduto[:posicao - 1]  # Tudo antes do caractere anterior ao símbolo

        posicaosecao = nomeproduto.find("➡")
        # Eliminar tudo após "➡" e um caractere antes dela
        if posicaosecao != -1:  # Certificar-se de que "➡" existe na string
            nomeproduto = nomeproduto[:posicaosecao - 1]  # Tudo antes do caractere anterior ao símbolo

        context.user_data["listaimpressao"].append(nomeproduto.strip())

        with open(f'etiquetas/etiqueta - {username}.txt', 'a', encoding='utf-8') as file:
            file.write(f"""^XA
^FO600,50
^A0R,140,90
^FD{nomeproduto[0:26]}
^FS

^FO460,50
^A0R,140,90
^FD{nomeproduto[26:]}
^FS

^FO220,160
^BY5,3
^BCR,180
^FD{codigoproduto}
^FS

^XZ

""")

        context.bot.send_message(chat_id=update.effective_chat.id,
                                 text=f'{nomeproduto}\nsalvo na lista de impressão',
                                 parse_mode="html")

    elif query.data == 'ajuste':

        mensagem = query.message.text.split('\n')
        context.user_data['codigo'] = mensagem[-1]
        context.user_data['nomeprod'] = mensagem[0]

        tecladozerar = [[InlineKeyboardButton("Zerar Estoque", callback_data='zerar Estoque'),
                    InlineKeyboardButton("Zerar Ocian", callback_data='zerar Ocian'),
                    InlineKeyboardButton("Zerar Boq", callback_data='zerar Boqueirao')],
                    [InlineKeyboardButton("Voltar", callback_data='voltar')]]

        teclado = InlineKeyboardMarkup(tecladozerar)

        query.edit_message_reply_markup(reply_markup=teclado)

    elif query.data == 'voltar':
        keyboard = [[InlineKeyboardButton("Mudar Seção", callback_data="mudasecao")],
            [InlineKeyboardButton("Pedir", callback_data="pedir"), InlineKeyboardButton("Lista Impressão", callback_data="listaimpressao")]]
        teclado = InlineKeyboardMarkup(keyboard)
        query.edit_message_reply_markup(reply_markup=teclado)
        return

    elif query.data.startswith('zerar'):

        partes = query.data.split(' ')
        if len(partes) > 2:
            codigo = partes[2]
            unidade = partes[1]

            if unidade == "Ocian":
                with ae.conoci() as con:
                    cur = con.cursor()
                    ae.ajustaEstoque(cur, codigo, 0, ajuste = True, user=username)
                    con.commit()
                    query.edit_message_text(text='Estoque Ocian Negativo Zerado ✅')

            if unidade == "Boqueirao":
                with ae.conboq() as con:
                    cur = con.cursor()
                    ae.ajustaEstoque(cur, codigo, 0, ajuste = True, user=username)
                    con.commit()
                    query.edit_message_text(text='Estoque Boqueirão Negativo Zerado ✅')

            if unidade == "Pet":
                with ae.conpet() as con:
                    cur = con.cursor()
                    ae.ajustaEstoque(cur, codigo, 0, ajuste = True, user=username)
                    con.commit()
                    query.edit_message_text(text='Estoque Happy Pet Negativo Zerado ✅')

            if unidade == "Estoque":
                with ae.conestoque() as con:
                    cur = con.cursor()
                    ae.ajustaEstoque(cur, codigo, 0, ajuste = True, user=username)
                    con.commit()
                    query.edit_message_text(text='Estoque Happy Pet Negativo Zerado ✅')                                   


    elif query.data == "cancelar":
        query.edit_message_text(text='Cancelado ❌')

    elif query.data == "pedir":
        mensagem = query.message.text
        linhas = mensagem.split('\n')
        linhas_menos_ultima = linhas[:-1]
        nova_string = "\n".join(linhas_menos_ultima)
        mensagemfinal = f"{nova_string}\n<code>{linhas[-1]}</code>"
        
        context.user_data["pedido"] = mensagemfinal

        tecladopedir = [[InlineKeyboardButton("Boqueirão", callback_data='pedirBoq'),
                    InlineKeyboardButton("Ocian", callback_data='pedirOci'),
                    InlineKeyboardButton("Pet", callback_data='pedirPet')],
                    [InlineKeyboardButton("Voltar", callback_data='voltar')]]
        
        teclado = InlineKeyboardMarkup(tecladopedir)

        query.edit_message_reply_markup(reply_markup=teclado)

    elif query.data == "pedirBoq":
        insereTabelaBusca(context.user_data["pedido"], "boqueirão")
        # consultabot.send_message(chat_id=IDGRUPOBUSCA, text=context.user_data["pedido"], parse_mode="html")
        # consultabot.send_message(chat_id=IDGRUPOBUSCA, text=f"_Pedido por: {context.user_data['username']}, Para: Boqueirão_",
        #                          parse_mode="html")
        
        tecladopedir = [[InlineKeyboardButton("Pedido Feito!", callback_data='voltar')]]
        teclado = InlineKeyboardMarkup(tecladopedir)
        query.edit_message_reply_markup(reply_markup=teclado)

        for id in IDS_USERS_BUSCA:
            try:
                context.bot.send_message(chat_id=id, text="Novo pedido para BOQUEIRÃO! Digite /pedidos para vizualizar. 📋",
                                parse_mode="html")
            except Exception as e:
                print(f"[Erro notificar] chatid={id} -> {e}")


    elif query.data == "pedirOci":
        insereTabelaBusca(context.user_data["pedido"], "ocian")
        # consultabot.send_message(chat_id=IDGRUPOBUSCA, text=context.user_data["pedido"], parse_mode="html")
        # consultabot.send_message(chat_id=IDGRUPOBUSCA, text=f"_Pedido por: {context.user_data['username']} Para: Ocian_",
        #                          parse_mode="html")
        
        tecladopedir = [[InlineKeyboardButton("Pedido Feito!", callback_data='voltar')]]
        teclado = InlineKeyboardMarkup(tecladopedir)
        query.edit_message_reply_markup(reply_markup=teclado)

        for id in IDS_USERS_BUSCA:
            try:
                context.bot.send_message(chat_id=id, text="Novo pedido para OCIAN! Digite /pedidos para vizualizar. 📋",
                                parse_mode="html")
            except Exception as e:
                print(f"[Erro notificar] chatid={id} -> {e}")

    elif query.data == "pedirPet":
        insereTabelaBusca(context.user_data["pedido"], "pet")
        # consultabot.send_message(chat_id=IDGRUPOBUSCA, text=context.user_data["pedido"], parse_mode="html")
        # consultabot.send_message(chat_id=IDGRUPOBUSCA, text=f"_Pedido por: {context.user_data['username']} Para: Pet_",
        #                          parse_mode="html")
        
        tecladopedir = [[InlineKeyboardButton("Pedido Feito!", callback_data='voltar')]]
        teclado = InlineKeyboardMarkup(tecladopedir)
        query.edit_message_reply_markup(reply_markup=teclado)

        for id in IDS_USERS_BUSCA:
            try:
                context.bot.send_message(chat_id=id, text="Novo pedido para PET! Digite /pedidos para vizualizar. 📋",
                                parse_mode="html")
            except Exception as e:
                print(f"[Erro notificar] chatid={id} -> {e}")

    elif "avulso" in query.data:
        context.bot.send_message(chat_id=update.effective_chat.id, text="Redirecionando para o codigo...",
                                 parse_mode="html")
        cod = query.data.split(' ')
        mc.consulta(cod[1])
        estoque = FileWaitRead('consultatemp.txt')
        texto = mostraestoque(estoque)[0]
        resultado = re.search(r"Estoque\s*(-?\d+)", estoque)
        if resultado:
            context.user_data["qtestoque"] = resultado.group(1)

        mensagem = texto.split('\n')
        codnovo = strip_html_tags(mensagem[-1])
        context.user_data["codigo"] = codnovo
        context.user_data['nomeprod'] = mensagem[0]
        keyboard = [[InlineKeyboardButton("Mudar Seção", callback_data=f"mudasecao")],
                    [InlineKeyboardButton("Pedir", callback_data=f"pedir")]]
        reply_markup = InlineKeyboardMarkup(keyboard)

        context.bot.send_message(chat_id=update.effective_chat.id, text=texto,
                                 parse_mode="html", reply_markup=reply_markup)

    elif query.data == 'desfazer':
        mensagem = query.message.text
        texto = mensagem.split('\n')
        padraoarq = r"Salvo no arquivo (.+\.txt)"
        corresarq = re.search(padraoarq, mensagem)
        if corresarq:
            nome_arquivo = corresarq.group(1)
        else:
            nome_arquivo = "erro"
        nomeprod = texto[0].strip()
        codapag = texto[1].strip()
        qtdapag = texto[2].split(": ")[1].strip()
        now = datetime.datetime.now()
        caminho = 'listas/'+ now.strftime("%d-%m-%Y") + '/' + nome_arquivo
        if not os.path.exists(caminho):
            context.bot.send_message(chat_id=update.effective_chat.id, text="Lista já foi apagada!")
        if nome_arquivo != "erro" and codapag != "erro" and qtdapag != "erro":
            linhacompleta = f"{codapag};{qtdapag}"
            with open(caminho, 'r', encoding="utf-8") as arquivo:
                linhas = arquivo.readlines()
                encontrou_ocorrencia = False
                novas_linhas = []
                for linha in linhas:
                    if linhacompleta in linha and encontrou_ocorrencia is False:
                        encontrou_ocorrencia = True
                    else:
                        novas_linhas.append(linha)
            with open(caminho, 'w', encoding="utf-8") as arquivo:
                arquivo.writelines(novas_linhas)
            context.bot.send_message(chat_id=update.effective_chat.id, text=f"{nomeprod} QTD: {qtdapag} deletado da lista {nome_arquivo}")
            query.edit_message_text(text=f"{mensagem} Desfeito ❌")
            if os.path.exists(caminho) and os.path.getsize(caminho) == 0:
                os.remove(caminho)
                context.bot.send_message(chat_id=update.effective_chat.id, text=f"Arquivo {nome_arquivo} apagado!")
        else:
            context.bot.send_message(chat_id=update.effective_chat.id, text="Codigo não encontrado em sua lista!")

        for raiz, subpastas, arquivos in os.walk("/listas"):
            for nomearquivo in arquivos:
                if nomearquivo.startswith("Mudança Seção"):
                    caminhocompleto = os.path.join(raiz, nomearquivo)
                    with open(caminhocompleto, 'r', encoding='utf-8') as arquivo:
                        linhas = arquivo.readlines()
                        novas_linhas = []
                        for linha in linhas:
                            if codapag in linha and "LOJA" in linha:
                                continue
                            else:
                                novas_linhas.append(linha)
                    with open(caminhocompleto, 'w', encoding='utf-8') as arquivo:
                        arquivo.writelines(novas_linhas)
                    if os.path.exists(caminhocompleto) and os.path.getsize(caminhocompleto) == 0:
                        os.remove(caminhocompleto)
                        context.bot.send_message(chat_id=update.effective_chat.id, text=f"Arquivo mudança seção apagado!")
        return
    
    elif query.data == "gerapedido":
        partes = update.callback_query.message.text.split(" ")
        marca = ""
        for parte in partes[1:]:
            marca += parte + " "
        marca = marca.replace("🔥", "")
        marca = marca.strip()
        comprar(update, context, marca)

    elif query.data == "gerapedidonatal":
        partes = update.callback_query.message.text.split(" ")
        marca = ""
        for parte in partes[1:]:
            marca += parte + " "
        marca = marca.strip()
        marca += " natal"
        comprar(update, context, marca)
    elif query.data == "ignorar":
        context.user_data['mensagemorignal'] = query.data
        partes = update.callback_query.message.text.split(" ")
        marca = ""
        for parte in partes[1:]:
            marca += parte + " "
        marca = marca.replace('🔥', '').upper().strip()
        context.user_data['marcaignorar'] = marca
        keyboard = [[InlineKeyboardButton("1 Mês", callback_data="ignorar30"),
                     InlineKeyboardButton("2 Meses", callback_data="ignorar60"),
                     InlineKeyboardButton("3 Meses", callback_data="ignorar90")],
                [InlineKeyboardButton("4 Meses", callback_data="ignorar120"),
                 InlineKeyboardButton("5 Meses", callback_data="ignorar120"),
                 InlineKeyboardButton("6 Meses", callback_data="ignorar120")],
                 [InlineKeyboardButton("1 Ano", callback_data="ignorar365"),
                 InlineKeyboardButton("Sempre", callback_data="ignorarsempre")],
                 [InlineKeyboardButton("Cancelar ↩", callback_data="ignorarcancelar")]]
        reply_markup = InlineKeyboardMarkup(keyboard)
        query.edit_message_text(text=f"{update.callback_query.message.text}\nPor quanto tempo deseja ignorar?",
                                reply_markup=reply_markup)
        return

    elif query.data.startswith("ignorar"):
        mensagemoriginal = update.callback_query.message.text.split("\n")[0]
        keyboard = [[InlineKeyboardButton("Ignorar", callback_data="ignorar"),
                    InlineKeyboardButton("Gerar pedido", callback_data="gerapedido")]]
        reply_markup = InlineKeyboardMarkup(keyboard)
        dias_ignorar = query.data.replace("ignorar", "")
        if dias_ignorar == 'sempre':
            with open('marcasignoradassempre.txt', 'a', encoding="utf-8") as file:
                file.write(f'{context.user_data["marcaignorar"]}\n')
            query.edit_message_text(text=mensagemoriginal, reply_markup=reply_markup)
            context.bot.send_message(chat_id=update.effective_chat.id, text=f"Marca {context.user_data['marcaignorar']} ignorada sempre")
        elif dias_ignorar == 'cancelar':
            query.edit_message_text(text=mensagemoriginal, reply_markup=reply_markup)
        else:
            datafinal = (datetime.datetime.now() + datetime.timedelta(days=int(dias_ignorar))).date()
            with open('marcasignoradas.txt', 'a', encoding="utf-8") as file:
                file.write(f'{context.user_data["marcaignorar"]} - {datafinal}\n')
            query.edit_message_text(text=mensagemoriginal, reply_markup=reply_markup)
            context.bot.send_message(chat_id=update.effective_chat.id, text=f"Marca {context.user_data['marcaignorar']} ignorada por {dias_ignorar} dias")
        return

    elif query.data.startswith("listar"):
        nomeplan = query.data.replace('listar', '').strip()
        nomeplan = nomeplan.replace('ç', 'c')
        nomeplan = nomeplan.replace('ã','a')
        nomeplan = nomeplan.lower()
        caminhoplanilha = "planilhassecao/" + nomeplan + ".xlsx"
        if os.path.exists(caminhoplanilha):

            # Carregando o arquivo Excel
            wb = openpyxl.load_workbook(caminhoplanilha)
            planilha = wb.active

            data = []
            for row in planilha.iter_rows(min_row=2, values_only=True):
                data.append(list(row))
            
            if nomeplan == 'relatorio':
                ordernar = False
            else:
                ordernar = True
            blocos = preparalista(data, ordernar=ordernar)
            newlist = []
            context.user_data["mensagem"] = []

            for bloco in blocos:
                message = context.bot.send_message(chat_id=update.effective_chat.id, text=bloco, parse_mode="html")
                pattern = r'/\d+'
                matches = re.findall(pattern, bloco)
                for i, match in enumerate(matches):
                    newlist.append(match)
                    newlist.append(message.message_id)

            context.user_data["mensagem"] = newlist
            query.edit_message_text(text= '-=' * 5 + f'<b>{nomeplan.upper()} 🗒️</b>' + '=-' * 5, parse_mode='html')
        
        else:
            query.edit_message_text(text='Lista não encontrada!')

        return



    else:
        section_name = query.data

        now = datetime.datetime.now()
        newdir = f'listas/{now.strftime("%d-%m-%Y")}/'
        os.makedirs(newdir, exist_ok=True)
        username = context.user_data['username']
        file_name = f'Mudança Seção - {username} - {now.strftime("%d-%m-%Y")}.txt'
        file_path = f'{newdir}{file_name}'

        # Adicione o texto que você deseja salvar no bloco de notas
        note_text = f"{context.user_data['codigo']} - {section_name}"

        with open(file_path, 'a', encoding="utf-8") as file:
            file.write(note_text + '\n')

        tecladopedir = [[InlineKeyboardButton(f"{context.user_data['nomeprod']} ➡ {section_name}", callback_data='voltar')]]
        teclado = InlineKeyboardMarkup(tecladopedir)
        query.edit_message_reply_markup(reply_markup=teclado)

        with open(f'logs/ConsultaEstoque/ce.chat_log-{username}.txt', 'a', encoding='utf-8') as log_file:
            log_file.write(f'{datetime.datetime.now().strftime("""%d/%m/%y - %H:%M:%S""")} Bot: {context.user_data["nomeprod"]} {context.user_data["codigo"]} mudando para Seção {section_name} salvo no bloco {file_name}\n')
        


def FileWaitRead(path, seconds_to_wait = 5):
    """Espera o arquivo existir e depois o deleta"""
    a = 1
    try:
        while not os.path.exists(path):
            time.sleep(1)
            a += 1
            if a > seconds_to_wait:
                return 'noprod'
        with open(path, 'r') as arquivo:
            content = arquivo.read()

        return content
    finally:
        os.remove(path)

def bloquearuso():
    """Cria arquivo de bloqueio que é removido depois da execução de outro bloco"""
    if os.path.exists(arqblock):
        while os.path.exists(arqblock):
            sleep(1)
    else:
        open(arqblock, "w", encoding="utf-8").close()

def process_image(update, context):
    """Processamento das imagens enviadas no telegram"""
    username = update.message.from_user.first_name
    userid = update.message.from_user.id
    bot = MessageSender(update, context)

    # Verifique se o usuário já está autenticado
    if os.path.exists('autenticados.txt'):
        with open('autenticados.txt', 'r') as file:
            autenticados = file.read()
    else:
        with open('autenticados.txt', 'w') as file:
            autenticados = ""

    if str(userid) in autenticados:
        try:
            with lock:
                bloquearuso()
                # Se o usuário está autenticado, execute a função
                # Baixa a imagem enviada pelo usuário
                photo_file = update.message.photo[-1].get_file()
                photo_file.download('image.jpg')
                img = Image.open('image.jpg')

                # Decodifica o código de barras na imagem usando a biblioteca pyzbar
                decoded_objects = decode(img)

                # Extrai o código de barras
                if decoded_objects:
                    context.user_data["codigo"] = decoded_objects[0].data.decode()
                    codigobarra = decoded_objects[0].data.decode()
                    dados = mc.consulta(codigobarra)
                    bot.send("Processando...")
                    if userid in IDS_USERS_COMPRA:
                        mensagem = mostraestoque(dados, precificado=True)
                    elif userid in IDS_USERS_BUSCA:
                        mensagem = mostraestoque(dados, ultvenda=True)
                    else:
                        mensagem = mostraestoque(dados)

                    bot.send(mensagem, botao=True)

                    context.user_data["nomeprod"] = dados[0]
                    qtdestoque = dados[3]
                    context.user_data["qtestoque"] = qtdestoque
                    context.user_data["codigo"] = dados[9]
                    context.user_data["estmin"] = dados[8]

                    os.remove('image.jpg')
                else:
                    bot.send("Não foi possível ler o código de barras.")
                    os.remove('image.jpg')
        except Exception:
            traceback.print_exc()
            bot.send("Não foi possível ler o código de barras.")
        finally:
            sleep(0.5)
            os.remove(arqblock)
    else:
        # Se o usuário não está autenticado, solicite a senha
        bot.send("Digite a senha para acessar esta função:")

def texto(update: Update, context: CallbackContext) -> None:
    """Processamento de qualquer mensagem enviado no telegram"""
    def produto_para_context(mensagem):
        mensagem = strip_html_tags(mensagem)
        estoqueqtd = 0
        ocianqtd = 0
        boqqtd = 0
        petqtd = 0

        resultado = re.search(r"Estoque\s*(-?\d+)", mensagem)
        if resultado:
            estoqueqtd = int(resultado.group(1))

        resultado = re.search(r"Ocian\s*(-?\d+)", mensagem)
        if resultado:
            ocianqtd = int(resultado.group(1))

        resultado = re.search(r"Boq\s*(-?\d+)", mensagem)
        if resultado:
            boqqtd = int(resultado.group(1))

        resultado = re.search(r"Pet\s*(-?\d+)", mensagem)
        if resultado:
            petqtd = int(resultado.group(1))

        resultado = re.search(r"Est Min\s*(-?\d+)", mensagem)
        if resultado:
            estmin = int(resultado.group(1))

        resultado = re.search(r"SECAO\s*([A-Za-z]+)", mensagem)
        if resultado:
            secao = resultado.group(1)
        else:
            secao = mensagem.split("\n")[-3]


        mensagem = mensagem.split("\n")
        context.user_data["nomeprod"] = mensagem[0]
        context.user_data["qtestoque"] = estoqueqtd
        context.user_data["qtboq"] = boqqtd
        context.user_data["qtocian"] = ocianqtd
        context.user_data["qtpet"] = petqtd
        context.user_data["codigo"] = mensagem[-1]
        context.user_data["estmin"] = estmin
        context.user_data["secao"] = secao


    bot = MessageSender(update, context)
    

    if 'mensagem' not in context.user_data:
        context.user_data['mensagem'] = []
    if 'listaimpressao' not in context.user_data:
        context.user_data['listaimpressao'] = []
    if 'pedidosBoq' not in context.user_data:
        context.user_data['pedidosBoq'] = []
    if 'pedidosOci' not in context.user_data:
        context.user_data['pedidosOci'] = []
    if 'pedidosPet' not in context.user_data:
        context.user_data['pedidosPet'] = []

    estoque = ''
    replyid = 0
    username = update.message.from_user.first_name
    context.user_data['username'] = username
    userid = update.message.from_user.id

    with open(f'logs/ConsultaEstoque/ce.chat_log-{username}.txt', 'a', encoding='utf-8') as log_file:
        log_file.write(f'{datetime.datetime.now():%d/%m/%y - %H:%M:%S} {username}: {update.message.text}\n')

    with open('usersid.txt', 'r+', encoding='utf-8') as idfile:
        frase = f'{username} - ID: {userid}\n'
        if frase not in idfile.read():
            idfile.write(frase)

    def criabloco(nomebloco, codigo, nomeprod, secao, estmin=-1, numero=0):
        if codigo == "":
            bot.send("Primeiro mande um codigo válido!")
            return
        if numero == 0:
            numero = ''.join(filter(str.isdigit, update.message.text))
        if '|' in nomeprod:
            blocos = nomeprod.split('|')
            nomeprod = blocos[0]
        if numero:
            nint = int(numero)
            now = datetime.datetime.now().strftime("%d-%m-%Y")
            newdir = 'listas/'+ now + '/'
            os.makedirs(newdir, exist_ok=True)
            if float(estmin) == 0:
                estminnovo = (nint/2)
                with open(newdir + "Ajuste EstMin " + now + " - " + username + ".txt", "a", encoding="utf-8") as arquivo:
                    arquivo.write(codigo + " - " + str(estminnovo) + "\n")
                
            nome_arquivo = nomebloco + now + " - " + username + ".txt"
            try:
                with open(newdir + nome_arquivo, "a", encoding="utf-8") as arquivo:
                    arquivo.write(codigo + ";" + str(numero) + "\n")
                    keyboard = [[InlineKeyboardButton("Desfazer ↩", callback_data="desfazer")]]
                    reply_markup = InlineKeyboardMarkup(keyboard)
                    mensagem = f"<b>{nomeprod}</b>\n{codigo}\n<b>QTD: {numero}</b> \nSalvo no arquivo <b>{nome_arquivo} 📋</b>"
                    context.bot.send_message(chat_id=update.effective_chat.id, text=mensagem, parse_mode='html', reply_markup=reply_markup)
                    if nint > 50:
                        mensagem += "\n⚠ <b>Aviso: QTD acima de 50</b> ⚠"
                        if update.effective_chat.id == ID_ADMIN: # Mandar mensagem para Admin se o valor for mais que 50
                            context.bot.send_message(chat_id=update.effective_chat.id, text=mensagem, parse_mode="html", reply_markup=reply_markup)
                        else:
                            context.bot.send_message(chat_id=ID_ADMIN, text=mensagem, parse_mode="html", reply_markup=reply_markup)
                            context.bot.send_message(chat_id=update.effective_chat.id, text=mensagem, parse_mode="html", reply_markup=reply_markup)

                    context.user_data["a_lastpath"] = newdir + nome_arquivo

                # Verficar existencia do codigo nas planilhas de pedidos
                caminhoplanilhaBoq = 'planilhassecao/pedidosBoq.xlsx'
                caminhoplanilhaOci = 'planilhassecao/pedidosOci.xlsx'
                caminhoplanilhaPet = 'planilhassecao/pedidosPet.xlsx'

                if os.path.exists(caminhoplanilhaBoq) and "- BOQUEIRAO" in nomebloco:
                    wb = openpyxl.load_workbook(caminhoplanilhaBoq)
                    planilha = wb.active
                    for row in planilha.iter_rows(min_row=2, values_only=False):
                        if codigo == row[1].value:
                            planilha.delete_rows(row[0].row)
                            wb.save(caminhoplanilhaBoq)
                            break 
                    if planilha.max_row == 1:
                        os.remove(caminhoplanilhaBoq)
                
                if os.path.exists(caminhoplanilhaOci) and "- OCIAN" in nomebloco:
                    wb = openpyxl.load_workbook(caminhoplanilhaOci)
                    planilha = wb.active
                    for row in planilha.iter_rows(min_row=2, values_only=False):
                        if codigo == row[1].value:
                            planilha.delete_rows(row[0].row)
                            wb.save(caminhoplanilhaOci)
                            break 
                    if planilha.max_row == 1:
                        os.remove(caminhoplanilhaOci)

                if os.path.exists(caminhoplanilhaPet) and "- HAPPY PET" in nomebloco:
                    wb = openpyxl.load_workbook(caminhoplanilhaPet)
                    planilha = wb.active
                    for row in planilha.iter_rows(min_row=2, values_only=False):
                        if codigo == row[1].value:
                            planilha.delete_rows(row[0].row)
                            wb.save(caminhoplanilhaPet)
                            break 
                    if planilha.max_row == 1:
                        os.remove(caminhoplanilhaPet)
                # Fim pedidos

                somatotal = 0
                encontrados = glob.glob('listas\\**\\ESTOQUE*.txt', recursive=True)
                if encontrados:
                    for arquivo in encontrados:
                        if os.path.exists(arquivo):
                            with open(arquivo, 'r', encoding="utf-8") as arq:
                                conteudo = arq.read()
                                padrao = codigo + r';(\d+)'
                                correspondencias = re.finditer(padrao, conteudo)
                                soma_numeros = 0
                                for x in correspondencias:
                                    quantidade = int(x.group(1))
                                    soma_numeros += quantidade
                                if soma_numeros > 0:
                                    somatotal += soma_numeros
                if int(context.user_data["qtestoque"]) > 0:
                    if (int(context.user_data["qtestoque"]) - somatotal) <= 0:
                        if secao != "LOJA":
                            with open(f"{newdir}Mudança Seção - {username} - {now}.txt", "a", encoding="utf-8") as arquivo:
                                arquivo.write(codigo + " - LOJA\n")
                        bot.send("🔵 Estoque zerando 🔵")

            except Exception as e:
                traceback.print_exc()
                bot.send("Não foi possivel salvar no arquivo, tente novamente.")

    #define o texto de resposta padrão do bot
    boasvindas = f"""Olá {username}, mande uma foto para conferir estoque
e depois mande:
B (qtd) para mandar para Boqueirão
O (qtd) para mandar para Ocian
T (qtd) para transferir de ocian para boqueirão
P (qtd) para mandar para Pet
Mande "Ref" para procurar um produto pela referência ou nome"""

    # Bloco para cuidar caso o usuario digite a senha incorretamenta já autenticado
    if os.path.exists('autenticados.txt'):
        with open('autenticados.txt', 'r') as file:
            autenticados = file.read()
    else:
        with open('autenticados.txt', 'w') as file:
            autenticados = ""

    if update.message.text.lower() == SENHA:
        with open('autenticados.txt', 'r+') as file:
            if username not in file.read():
                file.seek(0, 2)
                file.write('\n'+username)
        bot.send("Senha correta!")
        bot.send(boasvindas)
        with open('autenticados.txt', 'r', encoding='utf-8') as file:
            autenticados = file.read()

    elif str(userid) not in autenticados:
        bot.send("Digite a senha para usar este bot")

    if str(userid) in autenticados:
        if update.message.text.lower() == "/start":
            bot.send(boasvindas)
            return
        
        if "statelista" in context.user_data:
            try:
                if update.message.text in SECOES:
                    nomeplan = update.message.text.replace('ç', 'c')
                    nomeplan = nomeplan.replace('ã','a')
                    nomeplan = nomeplan.lower()
                    caminhoplanilha = "planilhassecao/" + nomeplan + ".xlsx"
                    if os.path.exists(caminhoplanilha):

                        # Carregando o arquivo Excel
                        wb = openpyxl.load_workbook(caminhoplanilha)
                        planilha = wb.active

                        data = []
                        for row in planilha.iter_rows(min_row=2, values_only=True):
                            data.append(list(row))
                        
                        if nomeplan == 'relatorio':
                            ordernar = False
                        else:
                            ordernar = True
                        blocos = preparalista(data, ordernar=ordernar)
                        newlist = []
                        context.user_data["mensagem"] = []

                        for bloco in blocos:
                            message = context.bot.send_message(chat_id=update.effective_chat.id, text=bloco, parse_mode="html")
                            pattern = r'/\d+'
                            matches = re.findall(pattern, bloco)
                            for i, match in enumerate(matches):
                                newlist.append(match)
                                newlist.append(message.message_id)

                        context.user_data["mensagem"] = newlist

                else:
                    bot.send("Seção Inválida!")
            finally:
                context.user_data.pop('statelista')
            return

        if update.message.reply_to_message and update.message.reply_to_message.from_user.id == context.bot.id:

            if 'r$' not in update.message.reply_to_message.text.lower():
                bot.send("⚠ Responda apenas para um produto ⚠")
                return

            if any(palavra in update.message.text.lower() for palavra in ("ajuste", "ajustar")):
                qtdnum = 0
                numero = None
                mensagem_original = update.message.reply_to_message.text.split('\n')
                codigo = mensagem_original[-1]
                partes_mensagem = update.message.text.lower().split()
                for elemento in partes_mensagem:
                    if elemento.isnumeric():
                        qtdnum += 1
                        numero = elemento
                    else:
                        unidade = elemento
                if qtdnum > 1 or numero is None:
                    bot.send("⚠ Quantidade inválida ⚠")
                    return
                
                if int(numero) > 1000:
                    bot.send("⚠ Quantidade muito alta ⚠")
                    return

                if unidade.startswith('est'):
                    unidade = "ESTOQUE"
                    with ae.conestoque() as con:
                        cur = con.cursor()
                        ae.ajustaEstoque(cur, codigo, numero, ajuste = True, user=username)
                        con.commit()
                        if int(numero) == 0:
                            now = datetime.datetime.now().strftime("%d-%m-%Y")
                            newdir = 'listas/'+ now + '/'
                            with open(f"{newdir}Mudança Seção - {username} - {now}.txt", "a", encoding="utf-8") as arquivo:
                                arquivo.write(codigo + " - LOJA\n")
                            bot.send("🔵 Estoque zerando 🔵")

                elif unidade.startswith("oci"):
                    unidade = "OCIAN"
                    with ae.conoci() as con:
                        cur = con.cursor()
                        ae.ajustaEstoque(cur, codigo, numero, ajuste = True, user=username)
                        con.commit()

                elif unidade.startswith("boq"):
                    unidade = "BOQUEIRÃO"
                    with ae.conboq() as con:
                        cur = con.cursor()
                        ae.ajustaEstoque(cur, codigo, numero, ajuste = True, user=username)
                        con.commit()

                elif unidade.startswith('pet'):
                    unidade = "HAPPY PET"
                    with ae.conpet() as con:
                        cur = con.cursor()
                        ae.ajustaEstoque(cur, codigo, numero, ajuste = True, user=username)
                        con.commit()

                else:
                    bot.send("Unidade não encontrada! Tente novamente")
                    return
                
                recebido = f"{mensagem_original[0]}\n<code>{codigo}</code>\nAjustado para <b>{numero}</b> na unidade <b>{unidade.upper()}</b> ✏️"

                if update.effective_chat.id == ID_ADMIN: # Mandar mensagem para Admin
                    context.bot.send_message(chat_id=update.effective_chat.id, text=recebido, parse_mode="html")
                else:
                    context.bot.send_message(chat_id=update.effective_chat.id, text=recebido, parse_mode="html")
                    recebido += f" por {username}"
                    context.bot.send_message(chat_id=ID_ADMIN, text=recebido, parse_mode="html")
                return

            #Salva no bloco respondendo algum item já decodificado
            if any(caractere.isdigit() for caractere in update.message.text.lower()):
                mensagem = strip_html_tags(update.message.reply_to_message.text).split('\n')
                estmin = mensagem[-2]
                n_estmin = int(estmin[8:])
                codigo = strip_html_tags(mensagem[-1])
                estoqueqtd = 0
                ocianqtd = 0
                boqqtd = 0
                petqtd = 0

                resultado = re.search(r"Estoque\s*(-?\d+)", update.message.reply_to_message.text)
                if resultado:
                    estoqueqtd = int(resultado.group(1))

                resultado = re.search(r"Ocian\s*(-?\d+)", update.message.reply_to_message.text)
                if resultado:
                    ocianqtd = int(resultado.group(1))

                resultado = re.search(r"Boq\s*(-?\d+)", update.message.reply_to_message.text)
                if resultado:
                    boqqtd = int(resultado.group(1))

                resultado = re.search(r"Pet\s*(-?\d+)", update.message.reply_to_message.text)
                if resultado:
                    petqtd = int(resultado.group(1))


                resultado = re.search(r"Estoque\s*(-?\d+)", update.message.reply_to_message.text)
                secao = mensagem[-3]
                if resultado:
                    context.user_data["qtestoque"] = resultado.group(1)

                if 'boc' in update.message.text.lower():
                    criabloco('BOQUEIRAO - OCIAN ', codigo, mensagem[0], secao=secao)
                    if ocianqtd < 0:
                        keyboard = [[InlineKeyboardButton("Sim", callback_data=f"zerar Ocian {codigo}"), InlineKeyboardButton("Não", callback_data="cancelar")]]
                        bot.send("Deseja zerar o estoque negativo da Ocian antes de enviar?", botao=True, keyboard=keyboard)
                    return
                
                elif 'poc' in update.message.text.lower():
                    criabloco('HAPPY PET - OCIAN ', codigo, mensagem[0], secao=secao)
                    if ocianqtd < 0:
                        keyboard = [[InlineKeyboardButton("Sim", callback_data=f"zerar Ocian {codigo}"), InlineKeyboardButton("Não", callback_data="cancelar")]]
                        bot.send("Deseja zerar o estoque negativo da Ocian antes de enviar?", botao=True, keyboard=keyboard)
                    return
                
                elif 'best' in update.message.text.lower():
                    criabloco('BOQUEIRAO - ESTOQUE ', codigo, mensagem[0], secao=secao)
                    if estoqueqtd < 0:
                        keyboard = [[InlineKeyboardButton("Sim", callback_data=f"zerar Estoque {codigo}"), InlineKeyboardButton("Não", callback_data="cancelar")]]
                        bot.send("Deseja zerar o estoque negativo do Estoque antes de enviar?", botao=True, keyboard=keyboard)
                    return
                
                elif 'oest' in update.message.text.lower():
                    criabloco('OCIAN - ESTOQUE ', codigo, mensagem[0], secao=secao)
                    if estoqueqtd < 0:
                        keyboard = [[InlineKeyboardButton("Sim", callback_data=f"zerar Estoque {codigo}"), InlineKeyboardButton("Não", callback_data="cancelar")]]
                        bot.send("Deseja zerar o estoque negativo do Estoque antes de enviar?", botao=True, keyboard=keyboard)

                    return
                
                elif 'pest' in update.message.text.lower():
                    criabloco('HAPPY PET - ESTOQUE ', codigo, mensagem[0], secao=secao)
                    if estoqueqtd < 0:
                        keyboard = [[InlineKeyboardButton("Sim", callback_data=f"zerar Estoque {codigo}"), InlineKeyboardButton("Não", callback_data="cancelar")]]
                        bot.send("Deseja zerar o estoque negativo do Estoque antes de enviar?", botao=True, keyboard=keyboard)
                    return
                
                elif 'b' in update.message.text.lower():
                    criabloco('ESTOQUE - BOQUEIRAO ', codigo, mensagem[0], secao=secao, estmin=n_estmin)
                    if boqqtd < 0:
                        keyboard = [[InlineKeyboardButton("Sim", callback_data=f"zerar Boqueirao {codigo}"), InlineKeyboardButton("Não", callback_data="cancelar")]]
                        bot.send("Deseja zerar o estoque negativo do Boqueirão antes de enviar?", botao=True, keyboard=keyboard)
                    return
                
                elif 'o' in update.message.text.lower():
                    criabloco('ESTOQUE - OCIAN ', codigo, mensagem[0], secao=secao, estmin=n_estmin)
                    if ocianqtd < 0:
                        keyboard = [[InlineKeyboardButton("Sim", callback_data=f"zerar Ocian {codigo}"), InlineKeyboardButton("Não", callback_data="cancelar")]]
                        bot.send("Deseja zerar o estoque negativo da Ocian antes de enviar?", botao=True, keyboard=keyboard)
                    return
                
                elif 'p' in update.message.text.lower():
                    criabloco('ESTOQUE - HAPPY PET ', codigo, mensagem[0], secao=secao, estmin=n_estmin)
                    if petqtd < 0:
                        keyboard = [[InlineKeyboardButton("Sim", callback_data=f"zerar Pet {codigo}"), InlineKeyboardButton("Não", callback_data="cancelar")]]
                        bot.send("Deseja zerar o estoque negativo do Pet antes de enviar?", botao=True, keyboard=keyboard)     
                    return
                
                elif 't' in update.message.text.lower() and 's' not in update.message.text.lower():
                    criabloco('OCIAN - BOQUEIRAO ', codigo, mensagem[0], secao=secao)
                    if boqqtd < 0:
                        keyboard = [[InlineKeyboardButton("Sim", callback_data=f"zerar Boqueirao {codigo}"), InlineKeyboardButton("Não", callback_data="cancelar")]]
                        bot.send("Deseja zerar o estoque negativo do Boqueirão antes de enviar?", botao=True, keyboard=keyboard)
                    return

        msg = update.message.text.lower().split()
        if re.search(r'\bred\b', msg[0]):
            bot.send('⚠ Você escreveu "RED" ao invés de "REF" ⚠')
            return
        if 'ref' in update.message.text.lower():
            if update.message.text.lower() == 'ref':
                bot.send("Mande também algo junto com 'ref' para eu procurar!")
                return
            try:
                bot.send("Processando...")
                refLimpa = update.message.text.lower().replace("ref", "").strip()
                termos = refLimpa.split(" ")
                data = mc.todos_nome(termos, todos=True)
                data.sort(key=lambda row: row[0])
                mensagem = preparalista(data)
                tamanho = len(mensagem[0].split("\n"))
                if tamanho == 4:
                    partido = mensagem[0].split("\n")
                    cod = partido[1][1:14]
                    dados = mc.consulta(cod)
                    if userid in IDS_USERS_COMPRA:
                        mensagem = mostraestoque(dados, precificado=True)
                    elif userid in IDS_USERS_BUSCA:
                        mensagem = mostraestoque(dados, ultvenda=True)
                    else:
                        mensagem = mostraestoque(dados)

                    bot.send(mensagem, botao=True, replyid=replyid)

                    produto_para_context(mensagem)

                else:
                    bot.send("Mostrando Planilha")
                    newlist = []
                    context.user_data["mensagem"] = []
                    for bloco in mensagem:
                        message = context.bot.send_message(chat_id=update.effective_chat.id, text=bloco, parse_mode="html")
                        pattern = r'/\d+'
                        matches = re.findall(pattern, bloco)
                        for i, match in enumerate(matches):
                            newlist.append(match)
                            newlist.append(message.message_id)

                        context.user_data["mensagem"] = newlist

            except Exception as e:
                traceback.print_exc()
            finally:
                sleep(0.5)
                os.remove(arqblock)
                return
            
        elif update.message.text.startswith('/nota'):
            num_nota = update.message.text.replace('/nota',"").strip()
            todosnota(update, context, num_nota)
            return

        elif len(update.message.text) >= 7 and "/s" in update.message.text:
            try:
                with lock:
                    bloquearuso()
                    bot.send("Processando...")
                    mensagemsplit = update.message.text.split()
                    codigo = mensagemsplit[0]
                    qtd = int(mensagemsplit[2]) if len(mensagemsplit) >= 3 else 1
                    dados = mc.consulta(codigo)
                    if dados is None:
                        bot.send("Produto não encontrado!")
                        return
                    criabloco("SAIDA ESTOQUE ", codigo, dados[0], secao=dados[-3], numero=qtd)
            except Exception:
                traceback.print_exc()
                bot.send("Produto não encontrado!")
                context.user_data["codigo"] = ""

            finally:
                sleep(0.5)
                os.remove(arqblock)
                return
            
        elif len(update.message.text) >= 7 and update.message.text.isnumeric() or update.message.text.startswith('/'):
            try:
                with lock:
                    bloquearuso()
                    bot.send("Processando...")
                    if update.message.text.startswith('/'):
                        if update.message.text in context.user_data["mensagem"]:
                            posicaocodigo = context.user_data["mensagem"].index(update.message.text)
                            try:
                                replyid = context.user_data["mensagem"][posicaocodigo + 1]
                            except Exception:
                                replyid = ""
                        codigobarra = update.message.text[1:]
                    else:
                        codigobarra = update.message.text

                    dados = mc.consulta(codigobarra)
                    if dados is None:
                        bot.send("Produto não encontrado!")
                        context.user_data["codigo"] = ""
                        return
                    
                    if userid in IDS_USERS_COMPRA:
                        mensagem = mostraestoque(dados, precificado=True)
                    elif userid in IDS_USERS_BUSCA:
                        mensagem = mostraestoque(dados, ultvenda=True)
                    else:
                        mensagem = mostraestoque(dados)

                    bot.send(mensagem, botao=True, replyid=replyid)

                    produto_para_context(mensagem)

            except Exception:
                traceback.print_exc()
                bot.send("Produto não encontrado!")
                context.user_data["codigo"] = ""

            finally:
                sleep(0.5)
                os.remove(arqblock)
                return

        elif any(caractere.isdigit() for caractere in update.message.text.lower()):
            if len(update.message.text) <= 8:
                try:
                    estoqueqtd = 0
                    ocianqtd = 0
                    boqqtd = 0
                    petqtd = 0

                    estoqueqtd = int(context.user_data['qtestoque'])
                    ocianqtd = int(context.user_data['qtocian'])
                    boqqtd = int(context.user_data['qtboq'])
                    petqtd = int(context.user_data['qtpet'])
                    codigo = context.user_data['codigo']
                    with lock:
                        bloquearuso()
                        # Salva Coletagem para boqueirão em arquivo texto
                        if 'boc' in update.message.text.lower():
                            criabloco('BOQUEIRAO - OCIAN ', context.user_data["codigo"], context.user_data["nomeprod"], secao=context.user_data["secao"])
                            if ocianqtd < 0:
                                keyboard = [[InlineKeyboardButton("Sim", callback_data=f"zerar Ocian {codigo}"), InlineKeyboardButton("Não", callback_data="cancelar")]]
                                bot.send("Deseja zerar o estoque negativo da Ocian antes de enviar?", botao=True, keyboard=keyboard)                            
                            return
                        
                        elif 'poc' in update.message.text.lower():
                            criabloco('HAPPY PET - OCIAN ', context.user_data["codigo"], context.user_data["nomeprod"], secao=context.user_data["secao"])
                            if ocianqtd < 0:
                                keyboard = [[InlineKeyboardButton("Sim", callback_data=f"zerar Ocian {codigo}"), InlineKeyboardButton("Não", callback_data="cancelar")]]
                                bot.send("Deseja zerar o estoque negativo da Ocian antes de enviar?", botao=True, keyboard=keyboard)                            
                            return
                        
                        elif 'best' in update.message.text.lower():
                            criabloco('BOQUEIRAO - ESTOQUE ', context.user_data["codigo"], context.user_data["nomeprod"], secao=context.user_data["secao"])
                            if estoqueqtd < 0:
                                keyboard = [[InlineKeyboardButton("Sim", callback_data=f"zerar Estoque {codigo}"), InlineKeyboardButton("Não", callback_data="cancelar")]]
                                bot.send("Deseja zerar o estoque negativo do Estoque antes de enviar?", botao=True, keyboard=keyboard)                            
                            return
                        
                        elif 'oest' in update.message.text.lower():
                            criabloco('OCIAN - ESTOQUE ', context.user_data["codigo"], context.user_data["nomeprod"], secao=context.user_data["secao"])
                            if estoqueqtd < 0:
                                keyboard = [[InlineKeyboardButton("Sim", callback_data=f"zerar Estoque {codigo}"), InlineKeyboardButton("Não", callback_data="cancelar")]]
                                bot.send("Deseja zerar o estoque negativo do Estoque antes de enviar?", botao=True, keyboard=keyboard)                            
                            return
                        
                        elif 'pest' in update.message.text.lower():
                            criabloco('HAPPY PET - ESTOQUE ', context.user_data["codigo"], context.user_data["nomeprod"], secao=context.user_data["secao"])
                            if estoqueqtd < 0:
                                keyboard = [[InlineKeyboardButton("Sim", callback_data=f"zerar Estoque {codigo}"), InlineKeyboardButton("Não", callback_data="cancelar")]]
                                bot.send("Deseja zerar o estoque negativo do Estoque antes de enviar?", botao=True, keyboard=keyboard)                            
                            return

                        if 'b' in update.message.text.lower() and update.message.text.lower() != SENHA and 's' not in update.message.text.lower():
                            criabloco('ESTOQUE - BOQUEIRAO ', context.user_data["codigo"], context.user_data["nomeprod"], secao=context.user_data["secao"], estmin=context.user_data["estmin"])
                            context.bot.edit_message_text(mc.consulta(context.user_data["codigo"]))
                            if boqqtd < 0:
                                keyboard = [[InlineKeyboardButton("Sim", callback_data=f"zerar Boqueirao {codigo}"), InlineKeyboardButton("Não", callback_data="cancelar")]]
                                bot.send("Deseja zerar o estoque negativo do Boqueirão antes de enviar?", botao=True, keyboard=keyboard)

                        # Salva Coletagem para ocian em arquivo texto
                        elif 'o' in update.message.text.lower() and 's' not in update.message.text.lower():
                            criabloco('ESTOQUE - OCIAN ', context.user_data["codigo"], context.user_data["nomeprod"], secao=context.user_data["secao"], estmin=context.user_data["estmin"])
                            if ocianqtd < 0:
                                keyboard = [[InlineKeyboardButton("Sim", callback_data=f"zerar Ocian {codigo}"), InlineKeyboardButton("Não", callback_data="cancelar")]]
                                bot.send("Deseja zerar o estoque negativo da Ocian antes de enviar?", botao=True, keyboard=keyboard)                            

                        
                        # Salva Coletagem para transferencia de Ocian para Boq
                        elif 't' in update.message.text.lower() and 'tp' not in update.message.text.lower() and 's' not in update.message.text.lower():
                            criabloco('OCIAN - BOQUEIRAO ', context.user_data["codigo"], context.user_data["nomeprod"], secao=context.user_data["secao"])
                            if boqqtd < 0:
                                keyboard = [[InlineKeyboardButton("Sim", callback_data=f"zerar Boqueirao {codigo}"), InlineKeyboardButton("Não", callback_data="cancelar")]]
                                bot.send("Deseja zerar o estoque negativo do Boqueirão antes de enviar?", botao=True, keyboard=keyboard)                            

                        # Salva Coletagem para transferencia de Ocian para Pet
                        elif 'tp' in update.message.text.lower() and 's' not in update.message.text.lower():
                            criabloco('OCIAN - HAPPY PET ', context.user_data["codigo"], context.user_data["nomeprod"], secao=context.user_data["secao"])
                            if petqtd < 0:
                                keyboard = [[InlineKeyboardButton("Sim", callback_data=f"zerar Pet {codigo}"), InlineKeyboardButton("Não", callback_data="cancelar")]]
                                bot.send("Deseja zerar o estoque negativo do Pet antes de enviar?", botao=True, keyboard=keyboard)                            

                        # Salva Coletagem para transferencia de Estoque para pet
                        elif 'p' in update.message.text.lower() and 's' not in update.message.text.lower():
                            criabloco('ESTOQUE - HAPPY PET ', context.user_data["codigo"], context.user_data["nomeprod"], secao=context.user_data["secao"], estmin=context.user_data["estmin"])
                            if petqtd < 0:
                                keyboard = [[InlineKeyboardButton("Sim", callback_data=f"zerar Pet {codigo}"), InlineKeyboardButton("Não", callback_data="cancelar")]]
                                bot.send("Deseja zerar o estoque negativo do Pet antes de enviar?", botao=True, keyboard=keyboard)                            

                finally:
                    sleep(0.5)
                    os.remove(arqblock)
                    return
            else:
                bot.send("⚠ Quantidade de caracteres maior que a permitida ⚠")

        elif 'bom dia' in update.message.text.lower():
            bot.send(f"Bom dia, {username}!")
        elif 'boa tarde' in update.message.text.lower():
            bot.send(f"Boa tarde, {username}!")
        elif 'boa noite' in update.message.text.lower():
            bot.send(f"Boa noite, {username}!")
        else:
            if update.message.chat.type == 'private':
                bot.send(boasvindas)

def listar(update, context):
    if context.args:
        nomeplan = " ".join(context.args)
        nomeplan = nomeplan.replace('ç', 'c')
        nomeplan = nomeplan.replace('ã','a')
        nomeplan = nomeplan.lower()
        caminhoplanilha = "planilhassecao/" + nomeplan + ".xlsx"
        if os.path.exists(caminhoplanilha):

            # Carregando o arquivo Excel
            wb = openpyxl.load_workbook(caminhoplanilha)
            planilha = wb.active

            data = []
            for row in planilha.iter_rows(min_row=2, values_only=True):
                data.append(list(row))
            
            if nomeplan == 'relatorio':
                ordernar = False
            else:
                ordernar = True
            blocos = preparalista(data, ordernar=ordernar)
            newlist = []
            context.user_data["mensagem"] = []
            context.bot.send_message(chat_id=update.effective_chat.id, text= '-=' * 5 + f'<b> {nomeplan.upper()} 🗒️ </b>' + '=-' * 5, parse_mode="html")

            for bloco in blocos:
                message = context.bot.send_message(chat_id=update.effective_chat.id, text=bloco, parse_mode="html")
                pattern = r'/\d+'
                matches = re.findall(pattern, bloco)
                for i, match in enumerate(matches):
                    newlist.append(match)
                    newlist.append(message.message_id)

            context.user_data["mensagem"] = newlist
        
        else:
            context.bot.send_message(chat_id=update.effective_chat.id, text=bloco, parse_mode="html")
        return
    
    bot = MessageSender(update, context)
    keyboard = []
    secoes = SECOES[0:-3]
    for secao in secoes:
        button = InlineKeyboardButton(secao, callback_data=f'listar {secao}')

        # Organize os botões em três linhas
        if len(keyboard) == 0 or len(keyboard[-1]) == 4:
                keyboard.append([button])
        else:
            keyboard[-1].append(button)
            
    keyboard.append([InlineKeyboardButton("Relatório 🔥", callback_data='listar relatorio')])
    keyboard.append([InlineKeyboardButton("Cancelar", callback_data='cancelar')])

    reply_markup = InlineKeyboardMarkup(keyboard)
    context.bot.send_message(chat_id=update.effective_chat.id, text='Escolha uma seção para listar', reply_markup=reply_markup)
  
def todos(update, context):
    """Comando bot que cria lista de todos os produtos com o mesmo termo fornecido pelo usuario"""
    bot = MessageSender(update, context)
    statustodos = False
    args = context.args
    if len(args) > 0:
        bloquearuso()
        try:
            bot.send("Processando...")
            if 'todos' in args:
                statustodos = True
                args.remove('todos')
            data = mc.todos_nome(args, todos=statustodos)
            data.sort(key=lambda row: row[0])
            mensagem = preparalista(data)
            bot.send("Mostrando Planilha")
            newlist = []
            context.user_data["mensagem"] = []
            for bloco in mensagem:
                message = context.bot.send_message(chat_id=update.effective_chat.id, text=bloco, parse_mode="html")
                pattern = r'/\d+'
                matches = re.findall(pattern, bloco)
                for i, match in enumerate(matches):
                    newlist.append(match)
                    newlist.append(message.message_id)

                context.user_data["mensagem"] = newlist

        except Exception as e:
            traceback.print_exc()
        finally:
            sleep(0.5)
            os.remove(arqblock)
            return
    bot.send('Mande o termo que deseja buscar')

def comando_comprar(update, context):
    bot = MessageSender(update, context)
    if len(context.args) == 0:
        bot.send("Envie a marca junto com o comando /comprar! (Ex: /comprar plasutil)")
        return
    marca = ' '.join(context.args).lower()
    comprar(update, context, marca)

def comprar(update, context, args):
    """Comando bot para retornar uma lista com todos os produtos da marca"""
    bot = MessageSender(update, context)
    if os.path.exists('relatorioestoque.xlsx'):
        marcaproc = args.lower()
        bot.send("Processando...")
        wb = openpyxl.load_workbook('relatorioestoque.xlsx')
        sheet = wb.active
        novoworkbook = openpyxl.Workbook()
        novosheet = novoworkbook.active
        dados = []
        produtos = []
        table = ""
        vendidosecom = []
        novosheet.append(("Produto", "Codigo", "QTD"))

        for row in sheet.iter_rows(min_row=2, values_only=True):
            marca = row[0]
            if not marca:
                continue
            marca = str(marca).lower()
            marca = row[0]
            nomeprod = row[1]
            codprod = row[2]
            qtdest = row[3]
            qtdboq = row[4]
            qtdoci = row[5]
            qtdcompra = row[9]
            pctvendido = row[16]
            qtdvendidoest = row[17]
            pct7dias = row[21]
            pct15dias = row[22]
            pct30dias = row[23]
            pct60dias = row[24]
            valor_venda = row[-1]
            grupo = row[-2]
            qtdvendido30dias = (pct30dias * qtdcompra)
            qtdvendido60dias = (pct60dias * qtdcompra)
            estoque_total = qtdest + qtdoci + qtdboq

            if qtdvendido60dias > 6:
                qtd_para_comprar_60 = math.ceil(qtdvendido60dias / 12) * 12
            else:
                qtd_para_comprar_60 = qtdcompra

            linha = (nomeprod, codprod, 1)

            if "natal" in marcaproc:
                grupoproc = "NATAL"
                marcaproc = marcaproc.replace("natal", '')
            elif "carnaval" in marcaproc:
                grupoproc = "CARNAVAL"
                marcaproc = marcaproc.replace("carnaval", '')
            elif "junina" in marcaproc:
                grupoproc = "JUNINA"
                marcaproc = marcaproc.replace("junina", '')
            else:
                grupoproc = "GRUPO"

            # Colocar na lista para comprar
            if marcaproc == marca.lower():
                if (pct30dias >= 0.5 or qtdvendido30dias > 20) and (estoque_total < qtdvendido30dias or estoque_total < qtdcompra * 0.2) and grupo == grupoproc and qtdest == 0:
                    novosheet.append(linha)

            if int(qtdvendidoest) > 0 and marca.lower() == marcaproc:
                produtos.append([row[1], row[2], qtdvendidoest])
        
        if novosheet.max_row < 2:
            bot.send("Lista sem produtos!")
            return

        produtos.sort(key=lambda x: x[-1], reverse=True)
        for produto in produtos:
            line_divider = "-" * 50
            linha = f"<b>{produto[0]}</b> | /{produto[1]} | <b>Vendidos ECOM: {produto[2]}</b>\n{line_divider}\n"
            if len(linha) + len(table) > 4096:
                vendidosecom.append(table)
                table = ""
            else:
                table += linha
        vendidosecom.append(table)
    
        novosheet.column_dimensions["A"].width = 53
        novosheet.column_dimensions["B"].width = 14
        novosheet.column_dimensions["C"].width = 4

        datahoje = datetime.datetime.today().strftime('%d-%m-%Y')

        nomearquivo = f"Pedido {marcaproc.upper()} - {datahoje}.xlsx"
        novoworkbook.save(nomearquivo)
        with open(nomearquivo, "rb") as arquivo:
            context.bot.send_document(chat_id=update.effective_chat.id, document=arquivo)
        # for bloco in vendidosecom:
        #     bot.send(bloco)
        os.remove(nomearquivo)
    else:
        bot.send("Relatório estoque inexistente")
        bot.send("Gerando novo relatório")
        mc.gera_relatorio()
        bot.send("Relatorio gerado com sucesso!")

def comando_nota(update, context):
    bot = MessageSender(update, context)
    if len(context.args) == 0:
        bot.send("Envie o numero da nota junto com o comando /nota! (Ex: /nota 1234)")
        return
    num_nota = context.args[0]
    todosnota(update, context, num_nota)
    
def todosnota(update, context, num_nota):
    """Comando bot que cria lista de todos os produtos da mesma nota"""
    try:
        bot = MessageSender(update, context)
        resultado = mc.produtos_por_nota(num_nota)
        mensagens = preparalista(resultado)
        bot.send("Mostrando Planilha")
        newlist = []
        context.user_data["mensagem"] = []
        for o, bloco in enumerate(mensagens):
            if o == len(mensagens) - 1:
                tecladoatualizar = [[InlineKeyboardButton("Atualizar Nota", callback_data=f'atualizarnota {num_nota}')]]
                teclado = InlineKeyboardMarkup(tecladoatualizar)
                message = context.bot.send_message(chat_id=update.effective_chat.id, text=bloco, parse_mode="html", reply_markup=teclado)
            else:
                message = context.bot.send_message(chat_id=update.effective_chat.id, text=bloco, parse_mode="html")
            pattern = r'/\d+'
            matches = re.findall(pattern, bloco)
            for i, match in enumerate(matches):
                newlist.append(match)
                newlist.append(message.message_id)

            context.user_data["mensagem"] = newlist

    except Exception as e:
        traceback.print_exc()

def todosdamarca(update, context):
    """Comando bot para retornar uma lista com todos os produtos da marca"""
    bot = MessageSender(update, context)
    args = context.args
    if len(args) > 0:
        bloquearuso()
        try:
            bot.send("Processando...")
            data = mc.todos_marca(args[0])
            data.sort(key=lambda row: row[0])
            mensagem = preparalista(data)
            bot.send("Mostrando Planilha")
            newlist = []
            context.user_data["mensagem"] = []
            for bloco in mensagem:
                message = context.bot.send_message(chat_id=update.effective_chat.id, text=bloco, parse_mode="html")
                pattern = r'/\d+'
                matches = re.findall(pattern, bloco)
                for i, match in enumerate(matches):
                    newlist.append(match)
                    newlist.append(message.message_id)

                context.user_data["mensagem"] = newlist

        except Exception as e:
            traceback.print_exc()
        finally:
            sleep(0.5)
            os.remove(arqblock)
            return
    bot.send('Mande o termo que deseja buscar')


def ultnotas(update, context):
    bot = MessageSender(update, context)
    dados = mc.get_ultnotas()
    resultado = "\n".join(f"{nome} - /nota{codigo}" for nome, codigo in dados)
    bot.send(resultado)

def comandos(update, context):
    username = update.message.from_user.first_name
    def sendmessage(mensagem):
        messlog = mensagem.replace('\n', ' ')
        with open(f'logs/ConsultaEstoque/ce.chat_log-{username}.txt', 'a', encoding="utf-8") as log_file:
            log_file.write(f'{datetime.datetime.now().strftime("""%d/%m/%y - %H:%M:%S""")} Bot: {messlog}\n')
        context.bot.send_message(chat_id=update.effective_chat.id, text=mensagem, parse_mode="html")

    sendmessage("""<b>B</b> -- para mandar do <b>Estoque</b> para <b>Boqueirão</b>
<b>O</b> -- para mandar do <b>Estoque</b> para <b>Ocian</b>
<b>P</b> -- para mandar do <b>Estoque</b> para <b>Pet</b>
-----------------------------------------------------------------
<b>best</b> -- para mandar do <b>Boqueirão</b> para <b>Estoque</b>
<b>oest</b> -- para mandar da <b>Ocian</b> para <b>Estoque</b>
<b>pest</b> -- para mandar do <b>Pet</b> para <b>Estoque</b>
-----------------------------------------------------------------
<b>T</b> -- para mandar da <b>Ocian</b> para <b>Boqueirão</b> 
<b>TP</b> -- para mandar da <b>Ocian</b> para <b>Pet</b>                          
<b>boc</b> - para mandar do <b>Boqueirão</b> para <b>Ocian</b>
<b>poc</b> - para mandar do <b>Pet</b> para <b>Ocian</b>
""")

def imprimir(update: Update, context: CallbackContext) -> None:
    username = update.message.from_user.first_name
    printer_path = r"\\DESKTOP-7K72SMI\ELGINL42ProMAYA"
    file_path = fr"etiquetas\etiqueta - {username}.txt"
    try:
        os.system(f'print /D:"{printer_path}" "{file_path}"')
        os.remove(file_path)
        context.user_data['listaimpressao'] = []
    except Exception as e:
        context.bot.send_message(chat_id=update.effective_chat.id, text="Não foi possivel imprimir", parse_mode="html")


def mostralistaimpressao(update: Update, context: CallbackContext) -> None:
    username = update.message.from_user.first_name
    arquivolista = f"etiquetas/etiqueta - {username}.txt"
    lista = ""
    
    for prod in context.user_data['listaimpressao']:
        lista += prod + "\n"

    if os.path.exists(arquivolista):
        context.bot.send_message(chat_id=update.effective_chat.id, text=lista, parse_mode="html")
    else:
        context.bot.send_message(chat_id=update.effective_chat.id, text="Não há nenhuma lista no momento", parse_mode="html")

def apagarlista(update: Update, context: CallbackContext) -> None:
    username = update.message.from_user.first_name
    arquivoetiqueta = f"etiquetas/etiqueta - {username}.txt"
    if os.path.exists(arquivoetiqueta):
        os.remove(arquivoetiqueta)
        context.bot.send_message(chat_id=update.effective_chat.id, text="Lista apagada com sucesso!", parse_mode="html")
    else:
        context.bot.send_message(chat_id=update.effective_chat.id, text="Não há nenhuma lista no momento", parse_mode="html")

    context.user_data['listaimpressao'] = []

def handle_document(update: Update, context: CallbackContext) -> None:
    document = update.message.document
    allowed_types = ["application/pdf", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "text/plain"]

    if document.mime_type in allowed_types:
        if document.mime_type == "application/pdf":
            file = update.message.document.get_file()
            file_path = "arquivodown.pdf"  # Defina um caminho temporário para salvar o arquivo
            file.download(file_path)

            try:
                impressora = os.getenv("PRINTER_NAME", "HP LaserJet Pro MFP M127fn")
                sumatra_path = os.getenv("SUMATRA_PATH", "SumatraPDF.exe")
                comando = f'"{sumatra_path}" -print-to "{impressora}" -silent "{file_path}"'

                try:
                    subprocess.run(comando, shell=True, check=True)
                    update.message.reply_text("Impressão enviada com sucesso!")
                except subprocess.CalledProcessError as e:
                    update.message.reply_text("Erro ao enviar para a impressora")
                    print(f"Erro ao imprimir: {e}")
            except Exception as e:
                update.message.reply_text(f"Erro ao enviar para a impressora: {e}")

            # Excluir o arquivo temporário após a impressão
            os.remove(file_path)
    else:
        update.message.reply_text("Este tipo de arquivo não é permitido.")

def coletados(update, context):
    userid = update.message.from_user.id
    line_divider = "-" * 50
    table = ""
    try:
        dados = []
        achado = False
        username = update.message.from_user.first_name
        padrao = re.compile(r'^[A-Z]+ - [A-Z]+ \d{2}-\d{2}-\d{4} - ' + re.escape(username) + r'\.txt$')
        if userid == ID_ADMIN:
            padrao = re.compile(r'^[A-Z]+ - [A-Z]+ \d{2}-\d{2}-\d{4} - (.*)$')
        for raiz, subpastas, arquivos in os.walk("listas"):
            for nome_arquivo in arquivos:
                dados.clear()
                table = ""
                linha = ""
                if padrao.match(nome_arquivo):
                    achado = True
                    nomelista = nome_arquivo.replace(".txt", "")
                    caminho = os.path.join(raiz, nome_arquivo)
                    with open(caminho, 'r') as arquivo:
                        conteudo = arquivo.readlines()
                        for item in conteudo:
                            cod, qtd = item.strip().split(';')
                            produto = mc.consulta(cod)
                            linha = f"<b>{produto[0]}</b> | /{produto[9]} | QTD: <b>{qtd}</b>\n{line_divider}\n"
                            if len(linha) + len(table) > 4096:
                                dados.append(table)
                                table = linha
                            else:
                                table += linha

                        dados.append(table)
                    context.bot.send_message(chat_id=update.effective_chat.id, text=f"📋 <b>{nomelista.upper()}</b> 📋", parse_mode="html")
                    newlist = []
                    context.user_data["mensagem"] = []
                    for bloco in dados:
                        message = context.bot.send_message(chat_id=update.effective_chat.id, text=bloco, parse_mode="html")
                        pattern = r'/\d+'
                        matches = re.findall(pattern, bloco)
                        for i, match in enumerate(matches):
                            newlist.append(match)
                            newlist.append(message.message_id)

                        context.user_data["mensagem"] = newlist
        if achado is False:
            context.bot.send_message(chat_id=update.effective_chat.id, text="Nenhuma lista encontrada", parse_mode="html")            
    except Exception:
        traceback.print_exc()

def checavendarapida(context: CallbackContext):
    ids_usuarios = (ID_ADMIN,)
    print("Executando checapedidos")
    if os.path.exists('relatorioestoque.xlsx'):
        df = pd.read_excel('relatorioestoque.xlsx')
        agrupado = df.groupby("Marca")
        df["DataCompra"] = pd.to_datetime(df["DataCompra"], errors="coerce")
        hoje = datetime.datetime.today()
        data_limite30 = hoje - datetime.timedelta(days=30)
        data_limite60 = hoje - datetime.timedelta(days=60)

        keyboard = [[InlineKeyboardButton("Ignorar", callback_data="ignorar"),
                     InlineKeyboardButton("Gerar pedido", callback_data="gerapedido")]]
        reply_markup = InlineKeyboardMarkup(keyboard)

        with open("marcasignoradassempre.txt", 'r') as arquivo:
            marcas_nao_procurar = arquivo.read().splitlines()

        with open("marcasignoradas.txt", 'r+') as arquivo:
            lista = arquivo.read().splitlines()
            listanova = []
            for linha in lista:
                linhasplit = linha.split(" - ")
                marcaignorar = linhasplit[0]
                dataignorar = datetime.datetime.strptime(linhasplit[1], "%Y-%m-%d")
                if dataignorar >= hoje:
                    listanova.append(f"{linha}\n")
                    marcas_nao_procurar.append(marcaignorar)
            arquivo.seek(0)
            arquivo.writelines(listanova)
            arquivo.truncate()
        
        # Percorrer cada grupo
        for marca, grupo in agrupado:
            if marca in marcas_nao_procurar:
                continue

            # Filtrar linhas onde todas as 3 colunas são 0
            filtrovendarapida = (
                (grupo["PCTVendidoTotal"] >= 0.8) &
                (grupo["DataCompra"] >= data_limite60) &
                (grupo["UltVendaBoq"].notna()) &
                (grupo["Grupo"] == "GRUPO")
            )
            vendarapida = grupo[filtrovendarapida]

            qtd = len(grupo[(grupo["DataCompra"] >= data_limite30)])

            if len(vendarapida) >= 5:
                qtdfogo = "🔥" * (len(vendarapida) // 5)
            else:
                qtdfogo = ""
            if len(qtdfogo) > 0:
                for user in IDS_USERS_COMPRA:
                    try:
                        context.bot.send_message(chat_id=user, text=f"Pedir {marca} {qtdfogo}", parse_mode="html", reply_markup=reply_markup)
                    except Exception as e:
                        print(f"[Erro notificar] chatid={user} -> {e}")


def checacoletados(context: CallbackContext):
    if os.path.exists('HistoricoColetasFun.xlsx'):
        wb = openpyxl.load_workbook('HistoricoColetasFun.xlsx', data_only=True)
        data_procurada = datetime.datetime.now().date()

        for user in wb.sheetnames:
            ws = wb[user]
            somapreco = 0
            somaelementos = 0

            for row in ws.iter_rows(min_row=2):  # pula cabeçalho
                valor_coluna_b = row[1].value   # coluna B (índice começa em 0)
                valor_coluna_a = row[0].value
                valor_coluna_d = row[3].value   # coluna D

                if valor_coluna_d and valor_coluna_d.date() == data_procurada:
                    if isinstance(valor_coluna_b, (int, float)):
                        somapreco += valor_coluna_b
                    if isinstance(valor_coluna_a, (int, float)):
                        somaelementos += valor_coluna_a

            context.bot.send_message(chat_id=ID_ADMIN, text=f"{user}\nCods: {somaelementos} VendaTotal: R$ {somapreco:.2f}", parse_mode="html")

def MandaProdutos(update, context):
    # Carrega JSON se existir
    if os.path.exists(JSON_PATH):
        with open(JSON_PATH, "r", encoding="utf-8") as f:
            try:
                topicos = json.load(f)
            except json.JSONDecodeError:
                topicos = {}
    else:
        topicos = {}

    chat_key = str(IDGRUPOPRODUTOS)
    if chat_key not in topicos:
        topicos[chat_key] = {}

    if not os.path.exists("relatorioestoque.xlsx"):
        update.message.reply_text("❌ Arquivo 'relatorioestoque.xlsx' não encontrado.")
        return

    df = pd.read_excel("relatorioestoque.xlsx")

    for index, row in df.iterrows():
        marca = row["Marca"]
        produto = row["Produto"]
        codigo = row["Codigo"]
        quantidaderestante = row["Estoque"] + row["Boq"] + row["Ocian"]
        quantidadecompra = row["QtdCompra"]
        datacompra = row['DataCompra']
        totalvendidos = row['TotalVendaUltCompra']
        pctvendidototal = row["PCTVendidoTotal"]
        pct30dias = row["Venda 30 dias"]
        estmin = row["EstMin"]
        qtdmesretrasado = row["QTDMesRetrasado"]
        qtdmespassado = row["QTDMesPassado"]
        pctmespassado = qtdmespassado/quantidadecompra

        if estmin == 0:
            estmin = quantidadecompra / 2

        ts = pd.Timestamp(datacompra)
        datacompra = ts.strftime("%d/%m/%Y")

        # Condição de envio
        if (qtdmespassado >= estmin * 3 or pctmespassado >= 0.6) and (quantidaderestante < qtdmespassado or quantidaderestante <= 3):
            nome_topico = str(marca)

            # Verifica se já temos thread_id no JSON
            if nome_topico in topicos[chat_key]:
                thread_id = topicos[chat_key][nome_topico]
            else:
                # Cria o tópico
                payload = {
                    "chat_id": IDGRUPOPRODUTOS,
                    "name": nome_topico[:128]
                }
                try:
                    res = requests.post(
                        f"https://api.telegram.org/bot{TOKEN}/createForumTopic",
                        json=payload,
                        timeout=30
                    ).json()
                except requests.exceptions.Timeout:
                    print(f"[TIMEOUT] criando tópico '{nome_topico}'")
                    continue

                time.sleep(1)
                if not res.get("ok"):
                    print(f"[ERRO] Criando tópico '{nome_topico}': {res}")
                    continue

                thread_id = res["result"]["message_thread_id"]
                topicos[chat_key][nome_topico] = thread_id

                # Salva no JSON
                with open(JSON_PATH, "w", encoding="utf-8") as f:
                    json.dump(topicos, f, ensure_ascii=False, indent=2)

            # Monta a mensagem
            try:
                locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
            except locale.Error:
                try:
                    locale.setlocale(locale.LC_TIME, 'Portuguese_Brazil.1252')
                except locale.Error:
                    pass

            hoje = datetime.date.today()
            mespassado = datetime.date(hoje.year, hoje.month, 1) - datetime.timedelta(days=1)
            ano = mespassado.year
            nome_mes = mespassado.strftime('%B').capitalize()

            if quantidaderestante < estmin * 2:
                quantidaderestante = str(quantidaderestante) + " 🟢"



            texto = (
                f"<b>Produto:</b> {produto}\n"
                f"<b>Código:</b> <code>{codigo}</code>\n"
                f"<b>Estoque:</b> {row['Estoque']:.0f}\n"
                f"<b>Boq:</b> {row['Boq']:.0f}\n"
                f"<b>Ocian:</b> {row['Ocian']:.0f}\n\n"
                f"<b>Data compra:</b> {datacompra}\n"
                f"<b>Qtd compra:</b> {quantidadecompra}\n"
                f"<b>Total Vendidos:</b> {totalvendidos}\n"
                f"<b>Qtd restante:</b> {quantidaderestante}\n"
                f"<b>Vendidos Mes Passado:</b> {qtdmespassado:.0f} {'🔥' * math.ceil(qtdmespassado/estmin)}\n"
                f"{nome_mes}/{ano}"
            )

            # Envia para o tópico
            msg_payload = {
                "chat_id": IDGRUPOPRODUTOS,
                "text": texto,
                "parse_mode": "html",
                "message_thread_id": thread_id
            }
            try:
                send_res = requests.post(
                    f"https://api.telegram.org/bot{TOKEN}/sendMessage",
                    json=msg_payload,
                    timeout=30
                ).json()
            except requests.exceptions.Timeout:
                print(f"[TIMEOUT] mandando mensagem '{texto}'")


            if not send_res.get("ok"):
                if send_res.get("error_code") == 429:
                    wait_time = send_res["parameters"]["retry_after"]
                    print(f"[RATE LIMIT] Aguardando {wait_time} segundos...")
                    time.sleep(wait_time)
                    # tenta enviar de novo
                    send_res = requests.post(
                        f"https://api.telegram.org/bot{TOKEN}/sendMessage",
                        json={
                            "chat_id": IDGRUPOPRODUTOS,
                            "text": texto,
                            "parse_mode": "html",
                            "message_thread_id": thread_id
                        },
                        timeout=30
                    ).json()
                else:
                    print(f"[ERRO] Enviando para tópico '{nome_topico}': {send_res}")

    print("Concluido")

def listahoje(update, context):
    caminhoplanilha = 'planilhassecao/cacarhoje.xlsx'
    if os.path.exists(caminhoplanilha):

        # Carregando o arquivo Excel
        wb = openpyxl.load_workbook(caminhoplanilha)
        planilha = wb.active

        data = []
        for row in planilha.iter_rows(min_row=2, values_only=True):
            data.append(list(row))

        blocos = preparalista(data)
    
        newlist = []
        context.user_data["mensagem"] = []
        context.bot.send_message(chat_id=update.effective_chat.id, text= '-=' * 5 + "<b> Lista Hoje 🗒️</b>" + '=-' * 5, parse_mode="html")

        for bloco in blocos:
            message = context.bot.send_message(chat_id=update.effective_chat.id, text=bloco, parse_mode="html")
            pattern = r'/\d+'
            matches = re.findall(pattern, bloco)
            for i, match in enumerate(matches):
                newlist.append(match)
                newlist.append(message.message_id)
        
        context.user_data["mensagem"] = newlist

def Negativos(update, context):
    """Comando bot que cria lista de todos os produtos com o mesmo termo fornecido pelo usuario"""
    bot = MessageSender(update, context)
    args = context.args
    marca = None
    if len(args) > 0:
        unidade = args[0].lower()
    else:
        bot.send('Uso: /negativos "unidade(obrigatorio)" "marca(opcional)"')
        return        
    if len(args) > 1:
        marca = " ".join(args[1:])
    
    if unidade.startswith("boq"):
        unidade = "boqueirao"
    elif unidade.startswith("oci"):
        unidade = "ocian"
    elif unidade.startswith("pet"):
        unidade = "pet"
    elif unidade.startswith("est"):
        unidade = "estoque"
    else:
        bot.send("Unidade inválida!")
        return
    
    if marca is not None:
        dados = mc.todos_negativos(unidade, marca)
    else:
        dados = mc.todos_negativos(unidade)

    if len(dados) == 0:
        bot.send("Não há produtos na lista!")
        return
    
    mensagem = preparalista(dados, ordernar=True, checalistas=False)
    bot.send("Mostrando Planilha")
    newlist = []
    context.user_data["mensagem"] = []
    for bloco in mensagem:
        message = context.bot.send_message(chat_id=update.effective_chat.id, text=bloco, parse_mode="html")
        pattern = r'/\d+'
        matches = re.findall(pattern, bloco)
        for i, match in enumerate(matches):
            newlist.append(match)
            newlist.append(message.message_id)

        context.user_data["mensagem"] = newlist

def anuncio(update, context):
    bot = MessageSender(update, context)
    args = context.args
    mensagem = f'<b>{" ".join(args)}</b>'
    for chatid in (IDS_USERS_BUSCA):
        try:
            context.bot.send_message(chat_id=chatid, text=f"{update.message.from_user.first_name}:\n{mensagem}", parse_mode="html")
        except Exception as e:
            print(f"[Erro notificar] chatid={chatid} -> {e}")            

def job_notificar(context: CallbackContext):
    tz = pytz.timezone("America/Sao_Paulo")
    now = datetime.datetime.now(tz)
    if now.weekday() == 6: # Não notifica se for domingo
        return

    if not (9 <= now.hour < 17):
        return

    caminhoplanilhaBoq = 'planilhassecao/pedidosBoq.xlsx'
    caminhoplanilhaOci = 'planilhassecao/pedidosOci.xlsx'
    caminhoplanilhaPet = 'planilhassecao/pedidosPet.xlsx'

    if os.path.exists(caminhoplanilhaBoq) or os.path.exists(caminhoplanilhaOci) or os.path.exists(caminhoplanilhaPet):
        for chatid in IDS_USERS_BUSCA:
            try:
                context.bot.send_message(chat_id=chatid, text="Existem pedidos pendentes! Digite /pedidos 📋")
            except Exception as e:
                print(f"[Erro notificar] chatid={chatid} -> {e}")


def pedidos(update, context):
    bot = MessageSender(update, context)
    caminhoplanilhaBoq = 'planilhassecao/pedidosBoq.xlsx'
    caminhoplanilhaOci = 'planilhassecao/pedidosOci.xlsx'
    caminhoplanilhaPet = 'planilhassecao/pedidosPet.xlsx'

    if os.path.exists(caminhoplanilhaBoq):
        # Carregando o arquivo Excel
        wb = openpyxl.load_workbook(caminhoplanilhaBoq)
        planilha = wb.active

        data = []
        for row in planilha.iter_rows(min_row=2, values_only=True):
            data.append(list(row))

        blocos = preparalista(data)

        newlist = []
        context.user_data["mensagem"] = []
        context.bot.send_message(chat_id=update.effective_chat.id, text= '-=' * 5 + "<b> Pedidos Boqueirão 🗒️</b>" + '=-' * 5, parse_mode="html")

        for bloco in blocos:
            message = context.bot.send_message(chat_id=update.effective_chat.id, text=bloco, parse_mode="html")
            pattern = r'/\d+'
            matches = re.findall(pattern, bloco)
            for i, match in enumerate(matches):
                newlist.append(match)
                newlist.append(message.message_id)
    
    if os.path.exists(caminhoplanilhaOci):
        # Carregando o arquivo Excel
        wb = openpyxl.load_workbook(caminhoplanilhaOci)
        planilha = wb.active

        data = []
        for row in planilha.iter_rows(min_row=2, values_only=True):
            data.append(list(row))

        blocos = preparalista(data)

        newlist = []
        context.user_data["mensagem"] = []
        context.bot.send_message(chat_id=update.effective_chat.id, text= '-=' * 5 + "<b> Pedidos Ocian 🗒️</b>" + '=-' * 5, parse_mode="html")

        for bloco in blocos:
            message = context.bot.send_message(chat_id=update.effective_chat.id, text=bloco, parse_mode="html")
            pattern = r'/\d+'
            matches = re.findall(pattern, bloco)
            for i, match in enumerate(matches):
                newlist.append(match)
                newlist.append(message.message_id)

    if os.path.exists(caminhoplanilhaPet):
        # Carregando o arquivo Excel
        wb = openpyxl.load_workbook(caminhoplanilhaPet)
        planilha = wb.active

        data = []
        for row in planilha.iter_rows(min_row=2, values_only=True):
            data.append(list(row))

        blocos = preparalista(data)

        newlist = []
        context.user_data["mensagem"] = []
        context.bot.send_message(chat_id=update.effective_chat.id, text= '-=' * 5 + "<b> Pedidos Pet 🗒️</b>" + '=-' * 5, parse_mode="html")

        for bloco in blocos:
            message = context.bot.send_message(chat_id=update.effective_chat.id, text=bloco, parse_mode="html")
            pattern = r'/\d+'
            matches = re.findall(pattern, bloco)
            for i, match in enumerate(matches):
                newlist.append(match)
                newlist.append(message.message_id)


def main():
    updater = Updater(TOKEN, use_context=True)
    dp = updater.dispatcher
    job_queue = updater.job_queue
    job_queue.run_daily(checavendarapida, datetime.time(hour=11, minute=30, second=5, tzinfo=brasilia_tz), job_kwargs={"misfire_grace_time": 60})
    # job_queue.run_daily(checacoletados, datetime.time(hour=19, minute=10, second=5, tzinfo=brasilia_tz), job_kwargs={"misfire_grace_time": 60})
    job_queue.run_repeating(job_notificar, 3600)
    dp.add_handler(CommandHandler('f', anuncio))
    dp.add_handler(CommandHandler('pedidos', pedidos))
    dp.add_handler(CommandHandler('negativos', Negativos))
    dp.add_handler(CommandHandler('manda', MandaProdutos))
    dp.add_handler(CommandHandler('hoje', listahoje))
    dp.add_handler(CommandHandler('ultnotas', ultnotas))
    dp.add_handler(CommandHandler('comprar', comando_comprar))
    dp.add_handler(CommandHandler('listar', listar))
    dp.add_handler(CommandHandler('marca', todosdamarca))
    dp.add_handler(CommandHandler('nota', comando_nota))
    dp.add_handler(CommandHandler('todos', todos))
    dp.add_handler(CommandHandler('comandos', comandos))
    dp.add_handler(CommandHandler('imprimir', imprimir))
    dp.add_handler(CommandHandler('coletados', coletados))
    dp.add_handler(CommandHandler('verlista', mostralistaimpressao))
    dp.add_handler(CommandHandler('apagarlista', apagarlista))
    dp.add_handler(MessageHandler(Filters.photo, process_image))
    dp.add_handler(MessageHandler(Filters.text, texto))
    dp.add_handler(CallbackQueryHandler(button_callback))
    dp.add_handler(MessageHandler(Filters.document, handle_document))
    updater.start_polling()
    updater.idle()

if __name__ == '__main__':
    main()
