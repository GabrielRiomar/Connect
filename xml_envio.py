# -*- coding: utf-8 -*-
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime, timedelta
import shutil
import zipfile
import os
import pathlib

# Configurações do servidor SMTP
smtp_host = 'smtp.gmail.com'
smtp_port = 587
smtp_user = 'origem@gmail.com'
smtp_pass = 'senha'

# Diretório de destino na área de trabalho
dest_dir = os.path.join(pathlib.Path.home(), 'Desktop', 'XML ENVIADOS')
os.makedirs(dest_dir, exist_ok=True)

# Criação do objeto de mensagem
msg = MIMEMultipart()
msg['From'] = smtp_user
msg['To'] = 'destino@email.com'
msg['Subject'] = 'Assunto do Email'

# Adiciona o corpo do email
corpo_email = 'Texto do corpo do email'
msg.attach(MIMEText(corpo_email, 'plain'))

try:
    # Obter o mês atual e o mês anterior
    data_atual = datetime.now()
    mes_atual = data_atual.strftime('%Y%m')
    mes_anterior = (data_atual.replace(day=1) - timedelta(days=1)).strftime('%Y%m')

    # Gerar o caminho da pasta e do arquivo ZIP do mês anterior
    pasta_path = fr'C:\SmartPDV\dfes\xmls\03783085000110\mfe\envios\{mes_anterior}'
    zip_path = os.path.join(dest_dir, f'{mes_anterior}.zip')

    print('Compactando arquivo...')
    # Compactar a pasta em formato ZIP
    shutil.make_archive(os.path.splitext(zip_path)[0], 'zip', pasta_path)

    # Renomear o arquivo ZIP compactado para o nome correto
    os.rename(os.path.splitext(zip_path)[0] + '.zip', zip_path)

    print('Arquivo compactado com sucesso.')

    # Mover o arquivo ZIP compactado para a pasta "XML ENVIADOS" na área de trabalho
    dest_file_path = os.path.join(dest_dir, os.path.basename(zip_path))
    shutil.move(zip_path, dest_file_path)

    # Adicionar o anexo ao email
    with open(dest_file_path, 'rb') as anexo:
        part = MIMEBase('application', 'zip')
        part.set_payload(anexo.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="{mes_anterior}.zip"')
        msg.attach(part)

    # Envio do email
    with smtplib.SMTP(smtp_host, smtp_port) as server:
        server.starttls()
        server.login(smtp_user, smtp_pass)
        server.sendmail(smtp_user, 'destino1@gmail.com', msg.as_string())
        server.quit()

    print('Arquivo enviado com sucesso.')

except Exception as e:
    # Registra o erro
    print('Ocorreu um erro durante a execucao:', str(e).encode('unicode_escape'))

input('Pressione Enter para sair...')
