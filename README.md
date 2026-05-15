# 📦 ConsultaEstoqueBot

Um **bot de Telegram** desenvolvido em Python para auxiliar na **gestao de estoque, pedidos e consultas de produtos**, integrado a um **banco Firebird SQL** e a rotinas de geracao de planilhas.

> ⚠️ Este repositorio foi preparado para portfolio.  
> Tokens, credenciais, logs, planilhas reais e dados sensiveis nao estao incluidos.

---

## 🚀 O que o bot faz

- 🔍 Consulta produtos por **codigo de barras, referencia ou nome**
- 🖼️ Reconhece **codigos de barras em imagens** enviadas pelo Telegram
- 📊 Gera e consulta relatorios em **Excel**
- 📝 Cria listas de pedidos por secao, nota fiscal ou marca
- 📦 Apoia movimentacoes entre estoque e unidades
- 🖨️ Gera listas para impressao de etiquetas
- 🔒 Autentica usuarios e controla acesso ao bot
- 🔔 Envia alertas automaticos para produtos com alta saida
- 🗄️ Integra consultas com **Firebird SQL**

---

## 🛠️ Tecnologias

- **Python 3**
- **python-telegram-bot**
- **Firebird SQL / fdb**
- **openpyxl / pandas**
- **Pillow / pyzbar**
- **requests**
- **python-dotenv**
- **threading / regex / glob**

---

## ⚙️ Como foi construido

- Separacao entre comandos do bot, consultas ao banco e funcoes auxiliares
- Uso de variaveis de ambiente para tokens, senhas e configuracoes locais
- Processamento de imagens para leitura de codigo de barras
- Geracao e leitura de planilhas para apoiar processos ja existentes
- Rotinas agendadas com `JobQueue` para verificacoes automaticas
- Persistencia auxiliar em arquivos locais para integracao operacional

---

## 🔐 Observação sobre execução

Este projeto depende de configurações privadas, como tokens, credenciais, banco de dados, arquivos locais ou integrações da operação real.

Por segurança, esses dados não foram incluídos no repositório. O objetivo aqui é demonstrar a arquitetura, as tecnologias utilizadas e as soluções implementadas.

---

## ✨ Aprendizados do projeto

- Criacao de fluxos conversacionais complexos no Telegram
- Integracao entre Python, banco SQL e planilhas Excel
- Automacao de processos reais de estoque
- Tratamento de imagens para leitura de codigos de barras
- Organizacao de configuracoes sensiveis com `.env`

---

## 👨‍💻 Autor

Desenvolvido por **Mateus Bastos**  
🔗 [LinkedIn](https://www.linkedin.com/in/mateus-bastos-825572139/) | [GitHub](https://github.com/teusbastos)
