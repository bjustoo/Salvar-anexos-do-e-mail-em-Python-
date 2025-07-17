# SalvareE-mails-em-Python

## 📩 𝐀𝐮𝐭𝐨𝐦𝐚𝐭𝐢𝐳𝐚𝐧𝐝𝐨 𝐩𝐫𝐨𝐜𝐞𝐬𝐬𝐨𝐬 𝐜𝐨𝐦 𝐏𝐲𝐭𝐡𝐨𝐧 + 𝐎𝐮𝐭𝐥𝐨𝐨𝐤

Este projeto automatiza o processo de extração e salvamento de anexos PDF de e-mails não lidos em uma subpasta específica do Outlook. Ideal para quem recebe relatórios recorrentes e deseja organizá-los automaticamente em uma pasta local.

## ✨ Funcionalidades

- Conexão com o Outlook via `win32com.client`
- Filtragem de e-mails **não lidos** em uma subpasta personalizada
- Salvamento automático de **anexos PDF** com nomes específicos
- Geração de nomes únicos para evitar sobrescrita de arquivos
- Criação automática da pasta de destino, se necessário

---

## ⚙️ Requisitos

- Sistema operacional **Windows**
- Microsoft Outlook instalado e configurado
- Python 3.x
- Bibliotecas Python:
  - `pywin32`
  - `os` (já incluída na biblioteca padrão)

Para instalar o `pywin32`, execute:

```bash
pip install pywin32
