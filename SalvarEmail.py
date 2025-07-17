import os
import win32com.client

# Conectar ao Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
conta = outlook.Folders.Item("SEU_EMAIL_AQUI")  # Substitua pelo e-mail da conta
caixa_entrada = conta.Folders.Item("Caixa de Entrada")  # Pasta Inbox
pasta_reports = caixa_entrada.Folders.Item("Teste")  # Nome da Subpasta Ex:'Teste'

# Pasta destino dos anexos
pasta_destino = r"C:\Caminho\Onde\Salvar\Arquivos"
if not os.path.exists(pasta_destino):
    os.makedirs(pasta_destino)

# Filtrar por e-mails não lidos
mensagens = pasta_reports.Items
mensagens = mensagens.Restrict("[Unread] = true")
total_mensagens = mensagens.Count
print(f"Total de mensagens NÃO LIDAS na pasta: {total_mensagens}")


nomes_permitidos = ["ARQUIVO1", "ARQUIVO2"]# Nomes permitidos nos anexos

# Evitar sobrescrever arquivos
def gerar_nome_unico(caminho_base):
    if not os.path.exists(caminho_base):
        return caminho_base
    base, extensao = os.path.splitext(caminho_base)
    contador = 1
    novo_caminho = f"{base}.{contador}{extensao}"
    while os.path.exists(novo_caminho):
        contador += 1
        novo_caminho = f"{base}.{contador}{extensao}"
    return novo_caminho

# Salvar anexos PDF filtrados
for i, mensagem in enumerate(mensagens, start=1):
    if mensagem.Class == 43 and mensagem.Attachments.Count > 0:
        for anexo in mensagem.Attachments:
            nome_arquivo = anexo.FileName
            nome_minusculo = nome_arquivo.lower()
            if nome_minusculo.endswith(".pdf") and any(chave in nome_minusculo for chave in nomes_permitidos):
                caminho_arquivo = os.path.join(pasta_destino, nome_arquivo)
                caminho_arquivo = gerar_nome_unico(caminho_arquivo)
                anexo.SaveAsFile(caminho_arquivo)
                print(f"Mensagem {i} - Anexo PDF salvo: {os.path.basename(caminho_arquivo)}")
