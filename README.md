# 🚀 Robô de Disparo de E-mails - Benefícios

Um aplicativo Desktop desenvolvido em Python para automatizar o envio de comunicados aos colaboradores (como atualizações de rastreio de cartões de benefícios), integrando leitura de planilhas Excel e disparo de e-mails via Outlook.

## 🎯 Por que este projeto existe?
Anteriormente, o processo de disparo de e-mails era realizado através de uma macro em **VBA** que dependia de abas ocultas ("Layout") e manipulação direta da interface do Excel, o que tornava a execução lenta e sujeita a falhas humanas (como quebra de formatação ou travamentos).

Este projeto moderniza a operação migrando a lógica para **Python**. O resultado é uma ferramenta independente, mais rápida, à prova de falhas de formatação (utilizando HTML/CSS puro inserido direto no código) e com uma interface gráfica amigável para a equipe de Operações de Benefícios.

## ✨ Funcionalidades
* **Interface Gráfica (GUI):** Tela moderna e intuitiva construída com `ttkbootstrap`.
* **Leitura Inteligente:** Processa planilhas Excel automaticamente usando `pandas`, ignorando linhas já enviadas.
* **Template HTML/CSS:** E-mails gerados com a identidade visual da empresa de forma blindada (sem depender de células do Excel).
* **Integração com Outlook:** Disparo em segundo plano via protocolo COM (`pywin32`), sem travar o computador do usuário.
* **Feedback em Tempo Real:** Pop-ups informando o progresso, quantidade de e-mails enviados e eventuais falhas.
* **Memória de Sessão:** O aplicativo lembra automaticamente a última planilha utilizada.

## 🛠️ Tecnologias Utilizadas
* **Python 3**
* **Pandas** (Manipulação e filtragem de dados)
* **OpenPyXL** (Atualização do status na planilha original)
* **PyWin32** (Comunicação com o aplicativo do Outlook)
* **ttkbootstrap / Tkinter** (Interface gráfica)

## 📋 Estrutura Esperada da Planilha
O robô lê a aba chamada `Capa`. O cabeçalho dos dados deve estar localizado na **linha 8** e conter as seguintes colunas essenciais:
* `Matricula`
* `Nome`
* `Cargo`
* `Código de Rastreio`
* `Data de Postagem`
* `Status` (Onde o robô escreverá "Enviado")
* `Email`
* `Enviar` (Deve conter a letra "x" para engatilhar o disparo)

---

## 🚀 Como executar o projeto (Ambiente de Desenvolvimento)

1. **Clone o repositório:**
   ```bash
   git clone [https://github.com/SEU_USUARIO/SEU_REPOSITORIO.git](https://github.com/SEU_USUARIO/SEU_REPOSITORIO.git)
   cd SEU_REPOSITORIO
   ```
Crie e ative o ambiente virtual:

```bash
python -m venv .venv
```

# No Windows:
```bash
.venv\Scripts\activate
```
Instale as dependências:

```bash
pip install pandas openpyxl pywin32 ttkbootstrap
```
Execute o aplicativo:

```bash
python app_disparo.py
```

Nota: É necessário ter o aplicativo Desktop do Outlook instalado e configurado na máquina para que os disparos ocorram.
