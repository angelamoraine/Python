import win32com.client
import openpyxl
import os
import re

# --- CONFIGURAÇÕES GLOBAIS ---
EMAIL_ACCOUNT = "EMAIL ONDE ESTÁ O RETORNO AQUI" # Certifique-se de que esta é a conta de e-mail correta no seu Outlook

# --- CONFIGURAÇÕES PARA GOL ---
GOL_FOLDER_NAME = "PASTA ONDE ESTÃO AS INFORMAÇÕES" # <--- Certifique-se que essa é a pasta exata no seu Outlook
GOL_EXCEL_PATH = r"CAMINHO DA PLANILHA DESTINO AQUI" # Caminho completo para a planilha da GOL

# --- FUNÇÕES DE EXTRAÇÃO GOL ---
def extrair_bilhete_gol(corpo):
    # Procura por números de bilhete que começam com "127" seguidos por 10 dígitos.
    match = re.search(r'ticketNumber:\s*(127\d{10})', corpo)
    return match.group(1) if match else ""

def extrair_trecho_gol(corpo):
    # Lista de padrões de início e fim para capturar trechos relevantes.
    # A ordem dos padrões pode influenciar qual trecho será capturado primeiro.
    padroes = [
        ('Estimo que esteja bem!', '.'),
        ('Espero que esteja bem :)', 'Agradecemos o contato e estamos à disposição!'),
        ('Espero que este e-mail o(a) encontre bem.', '.'),
        ('Olá, bom dia!', '.'),
        ('Prezados,', 'utilizada.'),
        ('Prezado(a) Agente', 'programado.'),
        ('Prezado(a) Agente', 'ação.'),
        ('Estimo que esteja bem!', 'Data do processamento'),
        ('Prezado agente,', 'Caso a solicitação '),
        ('Prezada (o)', 'Para reembolsos solicitados'),
        ('Prezado agente,', 'Para reembolsos solicitados'),
        ('Prezado agente,', 'Para reembolsos solicitados'),
        ('Prezado(a) Agente', 'Em caso de dúvidas'),
        ('Prezado agente,', 'por favor.'),
        ('Estimo que esteja bem!', 'Para reembolsos solicitados'),
        ('Estimo que esteja bem!', 'desconsiderada.'),
        ('Prezado Agente,', 'Em caso de dúvidas'),
        ('Prezado Agente,', 'processo.'),
        ('Prezado Agente,', 'Atendimento.'),
        ('Protocolo deste', 'ação.'),
        ('Protocolo deste', 'Para reembolsos solicitados'),
        ('O bilhete', 'sempre inicial'),
    ]
    for inicio, fim in padroes:
        inicio_esc = re.escape(inicio)
        fim_esc = re.escape(fim)
        # Regex para capturar o texto entre 'inicio' e 'fim'. re.DOTALL permite que '.' inclua quebras de linha.
        regex = re.compile(rf'{inicio_esc}\s*(.*?)\s*{fim_esc}', re.DOTALL)
        match = regex.search(corpo)
        if match:
            trecho = match.group(1).strip()
            # Normaliza os espaços para evitar múltiplos espaços ou quebras de linha indesejadas.
            trecho = ' '.join(trecho.split())
            return trecho
    return "" # Retorna vazio se nenhum padrão for encontrado

# --- CONEXÃO COM OUTLOOK ---
try:
    outlook_app = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
except Exception as e:
    print(f"Erro ao conectar ao Outlook: {e}")
    print("Certifique-se de que o Outlook está aberto e a biblioteca pywin32 está instalada e configurada.")
    exit()

current_account = None
for acc in outlook_app.Folders:
    if acc.Name.lower() == EMAIL_ACCOUNT.lower():
        current_account = acc
        break

if not current_account:
    raise Exception(f"Conta de e-mail '{EMAIL_ACCOUNT}' não encontrada no Outlook!")

inbox = current_account.Folders["Caixa de Entrada"]

# --- PROCESSAMENTO PARA GOL ---
print("Iniciando processamento para GOL...")
gol_folder = inbox.Folders.Item(GOL_FOLDER_NAME) # Usar .Item() para acessar a pasta
if not gol_folder:
    print(f"Pasta '{GOL_FOLDER_NAME}' não encontrada na Caixa de Entrada da sua conta.")
else:
    # Abre ou cria a planilha da GOL
    if os.path.exists(GOL_EXCEL_PATH):
        wb_gol = openpyxl.load_workbook(GOL_EXCEL_PATH)
        ws_gol = wb_gol.active
        print(f"Planilha existente carregada: {GOL_EXCEL_PATH}")
    else:
        wb_gol = openpyxl.Workbook()
        ws_gol = wb_gol.active
        ws_gol.append(["Bilhete", "Trecho"]) # Adiciona cabeçalhos se for uma nova planilha
        print(f"Nova planilha criada: {GOL_EXCEL_PATH}")

    # Itera sobre os itens (e-mails) da pasta GOL
    print(f"Verificando e-mails na pasta '{GOL_FOLDER_NAME}'...")
    emails_processados = 0
    for message in list(gol_folder.Items): # Converte para lista para evitar problemas se itens forem modificados durante o loop
        if message.UnRead:
            print(f"Processando e-mail: '{message.Subject}'")
            corpo = message.Body
            bilhete = extrair_bilhete_gol(corpo)
            trecho = extrair_trecho_gol(corpo)

            if bilhete or trecho: # Só adiciona e marca como lido se algo útil for extraído
                ws_gol.append([bilhete, trecho])
                message.UnRead = False # Marca o e-mail como lido
                emails_processados += 1
                print(f"  -> Adicionado: Bilhete='{bilhete}', Trecho='{trecho}'")
            else:
                print(f"  -> Nenhum bilhete ou trecho reconhecido no e-mail '{message.Subject}'. Mantendo como não lido.")
        else:
            # print(f"E-mail já lido ignorado: '{message.Subject}'") # Opcional: descomente para ver e-mails ignorados
            pass # Ignora e-mails já lidos

    wb_gol.save(GOL_EXCEL_PATH) # Salva as alterações na planilha
    print(f"\nDados da GOL salvos em: {GOL_EXCEL_PATH}")
    print(f"{emails_processados} e-mails da GOL processados e marcados como lidos.")

print("\nAutomação de e-mails da GOL concluída.")