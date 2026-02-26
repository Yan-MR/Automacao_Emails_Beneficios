import win32com.client as win32
import pythoncom  
from openpyxl import load_workbook
import pandas as pd
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs import Messagebox
from tkinter import filedialog
import threading
import os
import json

import sys

def resolver_caminho(caminho_relativo):
    """ Retorna o caminho absoluto, funcionando tanto no VS Code quanto no .exe gerado """
    try:
        # O PyInstaller cria uma pasta temporária em sys._MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, caminho_relativo)

CONFIG_FILE = "config_beneficios.json"

# --- LÓGICA DE BACK-END (O DISPARO) ---
def processar_disparos(caminho_arquivo, label_status, btn_iniciar, master):
    # Avisa ao Windows que esta Thread vai usar ferramentas COM (Outlook)
    pythoncom.CoInitialize() 
    
    try:
        master.after(0, lambda: label_status.config(text="Lendo planilha e conectando ao Outlook...", bootstyle=INFO))
        
        wb = load_workbook(caminho_arquivo)
        sh_capa = wb["Capa"]
        
        assunto = sh_capa["C4"].value
        bcc = sh_capa["C6"].value
        remetente_oficial = "rodriguesyan143@gmail.com" # Seu e-mail de teste
        
        outlook = win32.Dispatch('outlook.application')
        emails_enviados = 0
        erros = 0
        
        # CORREÇÃO: header=7 (Linha 8 do Excel)
        df_completo = pd.read_excel(caminho_arquivo, sheet_name="Capa", header=7)
        # Limpa espaços invisíveis dos cabeçalhos para evitar o KeyError
        df_completo.columns = df_completo.columns.str.strip()
        
        df_pendentes = df_completo[
            (df_completo['Status'] != 'Enviado') & 
            (df_completo['Enviar'] == 'x')
        ]
        
        grupos = df_pendentes.groupby('Nome') 
        
        for nome_agrupado, dados_responsavel in grupos:
            try:
                destino = dados_responsavel['Email'].iloc[0] 
                matricula = dados_responsavel['Matricula'].iloc[0]
                nome_colaborador = dados_responsavel['Nome'].iloc[0]
                cargo = dados_responsavel['Cargo'].iloc[0]
                rastreio = dados_responsavel['Código de Rastreio'].iloc[0]
                
                data_bruta = pd.to_datetime(dados_responsavel['Data de Postagem'].iloc[0], format='%d/%m/%Y')
                data_postagem = data_bruta.strftime('%d/%m/%Y')
                
                master.after(0, lambda n=nome_agrupado: label_status.config(text=f"Enviando para: {n}...", bootstyle=PRIMARY))
                
                corpo_html = f"""
                <html>
                    <body style="font-family: Calibri, Arial, sans-serif; color: #1F3864; font-size: 11pt; line-height: 1.5;">
                        <p>Boa tarde.</p>
                        <p>Tudo bem?</p>
                        <p>É um prazer tê-lo em nossa CIA como um de nossos (as) colaboradores (as), nós do time de benefícios temos uma excelente noticia para você!<br>
                        Informamos que o seu cartão <strong>Alelo</strong>, foi entregue em nosso HUB SCS e iremos redireciona-lo via SEDEX ao endereço cadastrado em sistema.<br>
                        Segue abaixo código de rastreio, para acessar, basta entrar no link: <a href="https://www.correios.com.br/" style="color: #0563C1; text-decoration: underline;">https://www.correios.com.br/</a></p>

                        <table style="border-collapse: collapse; width: 100%; max-width: 900px; font-family: Calibri, Arial, sans-serif; font-size: 11pt; color: #000000; text-align: center; margin-top: 20px; margin-bottom: 20px;">
                            <thead>
                                <tr>
                                    <th style="border: 1px solid black; background-color: #9BC2E6; padding: 8px; font-style: italic; font-weight: bold;">Matricula</th>
                                    <th style="border: 1px solid black; background-color: #9BC2E6; padding: 8px; font-style: italic; font-weight: bold;">Nome</th>
                                    <th style="border: 1px solid black; background-color: #9BC2E6; padding: 8px; font-style: italic; font-weight: bold;">Cargo</th>
                                    <th style="border: 1px solid black; background-color: #9BC2E6; padding: 8px; font-style: italic; font-weight: bold;">Código de Rastreio</th>
                                    <th style="border: 1px solid black; background-color: #9BC2E6; padding: 8px; font-style: italic; font-weight: bold;">Data de Postagem</th>
                                </tr>
                            </thead>
                            <tbody>
                                <tr>
                                    <td style="border: 1px solid black; padding: 8px;">{matricula}</td>
                                    <td style="border: 1px solid black; padding: 8px;">{nome_colaborador}</td>
                                    <td style="border: 1px solid black; padding: 8px;">{cargo}</td>
                                    <td style="border: 1px solid black; padding: 8px;">{rastreio}</td>
                                    <td style="border: 1px solid black; padding: 8px;">{data_postagem}</td>
                                </tr>
                            </tbody>
                        </table>

                        <p style="color: #1F3864;">Obs.: Esta é uma mensagem automática. Por favor não responda este e-mail</p>

                        <p style="color: #1F3864; margin-top: 30px;">Atenciosamente,<br>
                        Grupo Casas Bahia - Gente, Gestão e Sustentabilidade<br>
                        Operações de Benefícios</p>
                    </body>
                </html>
                """
                
                mail = outlook.CreateItem(0)
                mail.SentOnBehalfOfName = remetente_oficial
                mail.Subject = assunto
                mail.To = destino
                if bcc:
                    mail.BCC = bcc
                mail.HTMLBody = corpo_html
                
                mail.Send()
                emails_enviados += 1
                
                indices_originais = dados_responsavel.index
                for idx in indices_originais:
                    # CORREÇÃO: Escreve a partir da linha 9
                    linha_excel = idx + 9 
                    sh_capa.cell(row=linha_excel, column=7).value = "Enviado" 
                    
            except Exception as e:
                erros += 1
                print(f"Erro ao enviar para {nome_agrupado}: {str(e)}")

        wb.save(caminho_arquivo)
        
        def mostrar_sucesso():
            if erros == 0 and emails_enviados > 0:
                Messagebox.show_info(f"Processo finalizado perfeitamente!\n\n✔️ E-mails enviados: {emails_enviados}\n❌ Falhas: {erros}", "Sucesso")
            elif erros > 0:
                Messagebox.show_warning(f"Processo finalizado com ressalvas.\n\n✔️ E-mails enviados: {emails_enviados}\n❌ Falhas: {erros}", "Aviso")
            else:
                Messagebox.show_info("Nenhum e-mail pendente de envio encontrado na planilha.", "Concluído")
            label_status.config(text="Aguardando nova operação...", bootstyle=SECONDARY)
            btn_iniciar.config(state=NORMAL)
            
        master.after(0, mostrar_sucesso)

    except Exception as e:
        def mostrar_erro(mensagem=str(e)):
            Messagebox.show_error(f"Erro crítico ao processar o arquivo:\n{mensagem}", "Erro Fatal")
            label_status.config(text="Falha no processamento.", bootstyle=DANGER)
            btn_iniciar.config(state=NORMAL)
            
        master.after(0, mostrar_erro)
        
    finally:
        pythoncom.CoUninitialize()

# --- FRONT-END (INTERFACE GRÁFICA) ---
class AppDisparoEmails:
    def __init__(self, master):
        self.master = master
        self.master.title("Robô de E-mails - Benefícios")
        self.master.geometry("500x350")
        
        self.caminho_planilha = None
        
        try:
            # Agora ele acha a imagem embutida!
            caminho_logo = resolver_caminho("logo.png")
            icone = tb.PhotoImage(file=caminho_logo)
            self.master.iconphoto(False, icone)
        except Exception as e:
            print("Logo não encontrada na pasta, seguindo sem ícone.")
        
        lbl_titulo = tb.Label(master, text="Disparo de Benefícios", font=("Helvetica", 18, "bold"))
        lbl_titulo.pack(pady=20)
        
        frame_arquivo = tb.Frame(master)
        frame_arquivo.pack(pady=10, fill=X, padx=20)
        
        self.btn_selecionar = tb.Button(frame_arquivo, text="📂 Selecionar Planilha", command=self.selecionar_arquivo, bootstyle=INFO)
        self.btn_selecionar.pack(side=LEFT, padx=10)
        
        self.lbl_arquivo = tb.Label(frame_arquivo, text="Nenhum arquivo selecionado", foreground="gray")
        self.lbl_arquivo.pack(side=LEFT, padx=10)
        
        self.lbl_status = tb.Label(master, text="Aguardando seleção de arquivo...", font=("Helvetica", 10), bootstyle=SECONDARY)
        self.lbl_status.pack(pady=20)
        
        self.btn_iniciar = tb.Button(master, text="🚀 Iniciar Disparos", command=self.iniciar_processo, bootstyle=SUCCESS, state=DISABLED, width=20)
        self.btn_iniciar.pack(pady=10)

        self.carregar_config()

    def carregar_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    dados = json.load(f)
                    caminho_salvo = dados.get("caminho", "")
                    
                    if os.path.exists(caminho_salvo):
                        self.caminho_planilha = caminho_salvo
                        nome_arquivo = os.path.basename(caminho_salvo)
                        self.lbl_arquivo.config(text=nome_arquivo, foreground="black")
                        self.lbl_status.config(text="Última planilha carregada automaticamente.", bootstyle=PRIMARY)
                        self.btn_iniciar.config(state=NORMAL)
            except Exception as e:
                print("Não foi possível carregar as configurações antigas.")

    def salvar_config(self, caminho):
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump({"caminho": caminho}, f)
        except Exception as e:
            print("Não foi possível salvar a configuração.")

    def selecionar_arquivo(self):
        caminho = filedialog.askopenfilename(
            title="Selecione a base do Excel",
            filetypes=(("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*"))
        )
        if caminho:
            self.caminho_planilha = caminho
            nome_arquivo = os.path.basename(caminho)
            self.lbl_arquivo.config(text=nome_arquivo, foreground="black")
            self.lbl_status.config(text="Planilha carregada. Pronto para iniciar.", bootstyle=PRIMARY)
            self.btn_iniciar.config(state=NORMAL)
            self.salvar_config(caminho)

    def iniciar_processo(self):
        if not self.caminho_planilha:
            return
            
        self.btn_iniciar.config(state=DISABLED)
        thread = threading.Thread(target=processar_disparos, args=(self.caminho_planilha, self.lbl_status, self.btn_iniciar, self.master))
        thread.start()

if __name__ == "__main__":
    app = tb.Window(themename="cosmo") 
    AppDisparoEmails(app)
    app.mainloop()