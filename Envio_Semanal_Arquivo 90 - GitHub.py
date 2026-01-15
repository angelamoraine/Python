import os
import smtplib
import shutil
import re
import logging
import traceback
import tempfile
from email.message import EmailMessage
from dotenv import load_dotenv
from datetime import datetime

# Tenta importar Outlook COM (fallback se SMTP falhar)
try:
    import win32com.client
except Exception:
    win32com = None

# Configuração de logs
logging.basicConfig(
    filename=f'envio_arquivos_{datetime.now().strftime("%Y%m%d")}.log',
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# Carrega variáveis do arquivo .env
load_dotenv(dotenv_path=r"CAMINHO DO ARQUIVO ENV")  # altere o caminho conforme necessário

# Debug mínimo para confirmar leitura
print("=== Verificação de Configurações ===")
print(f".env existe: {os.path.exists(r'ARQUIVO ENVI')}")
print(f"EMAIL_REMETENTE: {os.getenv('EMAIL_REMETENTE')}")
print(f"EMAIL_SENHA carregada: {'Sim' if os.getenv('EMAIL_SENHA') else 'Não'}")
print("====================================")

# Diretórios
base_dir = r"CAMINHO DA PASTA ONDE ESTÃO OS ARQUIVOS"
arquivo_morto_dir = os.path.join(base_dir, "Arquivo morto")

# Lista de arquivos e destinatários
arquivos_info = [
    {"keyword": "ZUPPER VIAGENS", "email": "EMAIL DE QUEM VAI RECEBER ", "cc": ["PESSOAS EM CÓPIA"]},
    {"keyword": "KONTRIP VIAGENS", "email": "EMAIL DE QUEM VAI RECEBER ", "cc": ["PESSOAS EM CÓPIA"]},
    {"keyword": "K-CLUB", "email": "EMAIL DE QUEM VAI RECEBER ", "cc": ["PESSOAS EM CÓPIA"]},
    {"keyword": "TOOU", "email": "EMAIL DE QUEM VAI RECEBER ", "cc": ["PESSOAS EM CÓPIA"]},
    {"keyword": "KONTIK BUSINESS TRAVEL", "email": "EMAIL DE QUEM VAI RECEBER ", "cc": ["PESSOAS EM CÓPIA"]},
    {"keyword": "DNIT", "email": "EMAIL DE QUEM VAI RECEBER ", "cc": ["PESSOAS EM CÓPIA"]},
]

# Configurações do e-mail (lidas do .env)
remetente = os.getenv('EMAIL_REMETENTE')
senha = os.getenv('EMAIL_SENHA')
smtp_server = "smtp.office365.com"
smtp_port = 587

mensagem_padrao = """TEXTO PADRÃO QUE SERÁ ENVIADO JUNTO COM O ARQUIVO
"""

def verificar_diretorios():
    if not os.path.exists(base_dir):
        raise FileNotFoundError(f"Diretório base não encontrado: {base_dir}")
    if not os.path.exists(arquivo_morto_dir):
        os.makedirs(arquivo_morto_dir)
        logging.info(f"Diretório arquivo morto criado: {arquivo_morto_dir}")

def enviar_email_smtp(destinatario, assunto, corpo, caminho_arquivo, cc=None, timeout=30):
    try:
        if not os.path.exists(caminho_arquivo):
            raise FileNotFoundError(f"Arquivo não encontrado: {caminho_arquivo}")

        msg = EmailMessage()
        msg['Subject'] = assunto
        msg['From'] = remetente
        msg['To'] = destinatario

        if cc and isinstance(cc, list) and len(cc) > 0:
            cc_str = ', '.join(cc)
            msg['Cc'] = cc_str
            print(f"Adicionando destinatários em cópia: {cc_str}")
            logging.info(f"Destinatários em cópia: {cc_str}")

        msg.set_content(corpo)

        with open(caminho_arquivo, 'rb') as f:
            file_data = f.read()
            file_name = os.path.basename(caminho_arquivo)
            msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)

        with smtplib.SMTP(smtp_server, smtp_port, timeout=timeout) as smtp:
            smtp.ehlo()
            smtp.starttls()
            smtp.ehlo()
            smtp.login(remetente, senha)

            recipients = []
            if destinatario:
                if isinstance(destinatario, str) and ',' in destinatario:
                    recipients.extend([a.strip() for a in destinatario.split(',') if a.strip()])
                else:
                    recipients.append(destinatario)
            if cc:
                if isinstance(cc, list):
                    recipients.extend([a.strip() for a in cc if a and a.strip()])
                elif isinstance(cc, str):
                    recipients.extend([a.strip() for a in cc.split(',') if a.strip()])

            seen = set()
            recipients_clean = []
            for r in recipients:
                if r not in seen:
                    seen.add(r)
                    recipients_clean.append(r)

            logging.info(f"Enviando via SMTP para: {recipients_clean}")
            print(f"Enviando via SMTP para: {recipients_clean}")

            try:
                logging.info(f"Assunto a enviar: {assunto!r}")
                print(f"Assunto a enviar: {assunto!r}")
                headers = '\n'.join([f"{k}: {v}" for k, v in msg.items()])
                logging.debug(f"Email headers:\n{headers}")
                print(f"Email headers:\n{headers}")
            except Exception as dbg_e:
                logging.warning(f"Falha ao logar headers do email: {dbg_e}")

            smtp.send_message(msg, from_addr=remetente, to_addrs=recipients_clean)

        logging.info(f"SMTP: E-mail enviado para {destinatario} - {file_name} - Assunto: {assunto}")
        print(f"SMTP: E-mail enviado para {destinatario} - {file_name} - Assunto: {assunto}")
        return True

    except smtplib.SMTPAuthenticationError as e:
        logging.error(f"SMTP Authentication failed: {e}")
        print(f"SMTP Authentication failed: {e}")
        return False
    except Exception as e:
        logging.error(f"Erro SMTP ao enviar para {destinatario} com {caminho_arquivo}: {e}")
        logging.exception(traceback.format_exc())
        print(f"Erro SMTP: {e}")
        return False

def enviar_email_outlook(destinatario, assunto, corpo, caminho_arquivo, cc=None):
    if win32com is None:
        logging.error("Outlook COM não disponível (pywin32 não instalado).")
        print("Outlook COM não disponível.")
        return False

    temp_path = None
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)  # olMailItem
        mail.To = destinatario
        if cc and isinstance(cc, list) and len(cc) > 0:
            cc_str = '; '.join(cc)
            mail.CC = cc_str
            print(f"Adicionando destinatários em cópia no Outlook: {cc_str}")
            logging.info(f"Destinatários em cópia no Outlook: {cc_str}")
        mail.Subject = assunto
        mail.Body = corpo

        # Cria cópia temporária para anexar (evita lock do Outlook sobre o arquivo original)
        base_name = os.path.basename(caminho_arquivo)
        temp_path = os.path.join(tempfile.gettempdir(), f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{base_name}")
        shutil.copy2(caminho_arquivo, temp_path)
        mail.Attachments.Add(temp_path)

        mail.Send()
        logging.info(f"Outlook: E-mail enviado para {destinatario} - {os.path.basename(caminho_arquivo)}")
        print(f"Outlook: E-mail enviado para {destinatario} - {os.path.basename(caminho_arquivo)}")

        # remover temporário
        try:
            if temp_path and os.path.exists(temp_path):
                os.remove(temp_path)
        except Exception:
            logging.warning(f"Não foi possível remover temporário: {temp_path}")

        return True
    except Exception as e:
        logging.error(f"Erro Outlook ao enviar para {destinatario} com {caminho_arquivo}: {e}")
        logging.exception(traceback.format_exc())
        print(f"Erro Outlook: {e}")
        try:
            if temp_path and os.path.exists(temp_path):
                os.remove(temp_path)
        except Exception:
            pass
        return False

def processar_arquivos():
    try:
        verificar_diretorios()
        arquivos_processados = 0

        for info in arquivos_info:
            keyword = info["keyword"]
            destinatario = info["email"]
            cc_emails = info.get("cc", [])

            print(f"\nProcessando e-mails para keyword '{keyword}':")
            print(f"Destinatário principal: {destinatario}")
            print(f"Destinatários em cópia: {cc_emails}")

            all_files = [f for f in os.listdir(base_dir) if os.path.isfile(os.path.join(base_dir, f))]

            def matches_keyword(filename, keyword):
                fn = filename.lower()
                k = keyword.lower().strip()

                # correspondência direta do keyword completo
                if k in fn:
                    return True

                # tokens do keyword (ex.: ["zupper", "viagens"])
                tokens = [t for t in re.split(r"\W+", k) if t]
                if not tokens:
                    return False

                # Prioridade: primeira palavra (marca/nome) - busca por palavra inteira
                first = tokens[0]
                if first:
                    if re.search(r'\b' + re.escape(first) + r'\b', fn):
                        return True
                    # também tentar início do filename (sem precisar de word boundary)
                    if fn.startswith(first):
                        return True

                # Tentar casar pelas primeiras 2-3 palavras juntas (prefixo)
                prefix = ' '.join(tokens[:3])
                if prefix and prefix in fn:
                    return True

                return False

            arquivos_encontrados = [f for f in all_files if matches_keyword(f, keyword)]

            print(f"Encontrados {len(arquivos_encontrados)} arquivos para '{keyword}'")
            nao_casaram = [f for f in all_files if f not in arquivos_encontrados]
            if arquivos_encontrados:
                print(f"Arquivos selecionados: {arquivos_encontrados}")
            else:
                print(f"Nenhum arquivo correspondeu ao keyword '{keyword}'. Arquivos disponíveis (primeiros 20): {nao_casaram[:20]}")

            for arquivo in arquivos_encontrados:
                caminho_arquivo = os.path.join(base_dir, arquivo)
                try:
                    print(f"Processando: {caminho_arquivo}")

                    enviado = False
                    assunto = f"ASSUNTO DO EMAIL {keyword}"

                    # Tenta SMTP primeiro se credenciais estiverem presentes
                    if remetente and senha:
                        enviado = enviar_email_smtp(destinatario, assunto, mensagem_padrao, caminho_arquivo, cc=cc_emails)
                        logging.info(f"Resultado envio SMTP (arquivo={arquivo}): {enviado}")
                        print(f"Resultado envio SMTP (arquivo={arquivo}): {enviado}")

                    # Se SMTP falhar ou credenciais ausentes, tenta Outlook COM
                    if not enviado:
                        enviado = enviar_email_outlook(destinatario, assunto, mensagem_padrao, caminho_arquivo, cc=cc_emails)

                    if enviado:
                        destino = os.path.join(arquivo_morto_dir, arquivo)
                        try:
                            shutil.move(caminho_arquivo, destino)
                            logging.info(f"Arquivo movido para Arquivo morto: {arquivo}")
                            print(f"Arquivo {arquivo} movido para 'Arquivo morto'")
                        except Exception as mv_e:
                            logging.error(f"Erro ao mover arquivo {arquivo}: {mv_e}")
                            logging.exception(traceback.format_exc())
                            print(f"Erro ao mover arquivo: {mv_e}")
                        arquivos_processados += 1
                    else:
                        logging.warning(f"Não foi possível enviar o arquivo: {arquivo}")
                        print(f"Não foi possível enviar: {arquivo}")

                except Exception as e:
                    logging.error(f"Erro ao processar arquivo {arquivo}: {e}")
                    logging.exception(traceback.format_exc())
                    print(f"Erro geral processando {arquivo}: {e}")

        print(f"\nProcessamento concluído: {arquivos_processados} arquivos processados")
        logging.info(f"Processamento concluído: {arquivos_processados} arquivos processados")

    except Exception as e:
        logging.error(f"Erro na rotina principal: {e}")
        logging.exception(traceback.format_exc())
        print(f"Erro na rotina principal: {e}")
        raise

if __name__ == "__main__":
    processar_arquivos()