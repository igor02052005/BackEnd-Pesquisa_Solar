import os
import re
import pandas as pd
import warnings
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# Configure your own credentials
client_id = "719743115049-5ljql6s0ai361ncstf31k32m5tj8dukg.apps.googleusercontent.com"
client_secret = "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"
TargetfolderId = '1oeKWVEDNPyPLVn5TZVE5IqQtcC2bT9xb'  # Updated folder ID
TEMP_FOLDER = r"C:\\Users\\igort\\OneDrive\\Desktop\\VSCODE\\TEMP"  # Temporary folder

# Define the scope for Google Drive API
SCOPES = ['https://www.googleapis.com/auth/drive']

def generate_tokens(client_id, client_secret):
    flow = InstalledAppFlow.from_client_config(
        client_config={
            "web": {
                "client_id": client_id,
                "client_secret": client_secret,
                "redirect_uris": ["urn:ietf:wg:oauth:2.0:oob"],
                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                "token_uri": "https://accounts.google.com/o/oauth2/token",
            }
        },
        scopes=SCOPES
    )
    flow.run_local_server(port=0)
    return flow.credentials.token, flow.credentials.refresh_token

def authenticate_with_token(token):
    creds = Credentials(
        token=token['token'],
        refresh_token=token['refresh_token'],
        token_uri=token['token_uri'],
        client_id=token['client_id'],
        client_secret=token['client_secret'],
        scopes=token['scopes']
    )
    if not creds.valid:
        if creds.expired and creds.refresh_token:
            creds.refresh(Request())
    return creds

def upload_file_to_drive(file_path, token, folder_id):
    creds = authenticate_with_token(token)
    service = build('drive', 'v3', credentials=creds)
    file_name = os.path.basename(file_path)

    # Check if the file already exists in the target folder
    query = f"'{folder_id}' in parents and name = '{file_name}' and trashed = false"
    results = service.files().list(q=query, fields="files(id, name)").execute()
    items = results.get('files', [])

    if items:
        # File exists, update it
        file_id = items[0]['id']
        media = MediaFileUpload(file_path, resumable=True)
        updated_file = service.files().update(fileId=file_id, media_body=media).execute()
        print(f'Updated File ID: {updated_file.get("id")}')
    else:
        # File does not exist, create a new one
        file_metadata = {
            'name': file_name,
            'parents': [folder_id]  # Specify the folder ID as the parent ID
        }
        media = MediaFileUpload(file_path, resumable=True)
        new_file = service.files().create(body=file_metadata,
                                          media_body=media,
                                          fields='id').execute()
        print(f'Created File ID: {new_file.get("id")}')

def process_soliscloud(file_path):
    df = pd.read_excel(file_path, skiprows=6, usecols=['Time', 'Inverter SN', 'Today Yield(kWh)'])
    df = df.rename(columns={'Time': 'DATA', 'Inverter SN': 'ID', 'Today Yield(kWh)': 'GERAÇÃO'})
    df = df[['DATA', 'GERAÇÃO', 'ID']]
    return df

def process_solarman(file_path):
    df = pd.read_excel(file_path, usecols=['Número de série', 'Tempo', 'Produção(kWh)'])
    df = df.rename(columns={'Número de série': 'ID', 'Tempo': 'DATA', 'Produção(kWh)': 'GERAÇÃO'})
    df = df[['DATA', 'GERAÇÃO', 'ID']]
    return df

def extract_numbers_after_last_underscore(string):
    # Usamos uma expressão regular para pegar os números após o último '_'
    match = re.search(r'_(\d+)$', string)
    if match:
        return match.group(1)  # Retorna o número encontrado
    else:
        return None  # Caso não haja um número após o último '_'

def process_isolarcloud(file_path):
    # Lê o arquivo Excel completo para obter os dados iniciais
    df = pd.read_excel(file_path, header=None)  # Não usa cabeçalho para garantir leitura da primeira linha
    
    # Obtém o valor da primeira linha da primeira coluna como ID
    id_inversor = df.iloc[0, 0]  # Primeiro valor da primeira coluna (linha 0, coluna 0)
    
    # Aplica a função de extração ao valor de 'id_inversor'
    id_inversor_extraido = extract_numbers_after_last_underscore(str(id_inversor))

    # Recarrega o arquivo Excel com a leitura das colunas necessárias (sem 'ID' ainda)
    df_filtered = pd.read_excel(
        file_path,
        skiprows=1,  # Ignora a primeira linha
        usecols=['Horário', 'Potência CC total(kW)']
    )

    # Renomeia as colunas
    df_filtered = df_filtered.rename(columns={
        'Horário': 'DATA',
        'Potência CC total(kW)': 'GERAÇÃO'
    })

    # Rearranja as colunas na ordem desejada
    df_filtered = df_filtered[['DATA', 'GERAÇÃO']]

    # Adiciona a coluna 'ID' ao DataFrame filtrado
    df_filtered['ID'] = id_inversor_extraido
    
    return df_filtered

def process_canadian(file_path):
    df = pd.read_excel(file_path, usecols=['SN', 'Tempo atualizado', 'Produção(kWh)'])
    df = df.rename(columns={'SN': 'ID', 'Tempo atualizado': 'DATA', 'Produção(kWh)': 'GERAÇÃO'})
    df = df[['DATA', 'GERAÇÃO', 'ID']]
    return df

def get_files_from_drive(folder_id, token):
    creds = authenticate_with_token(token)
    service = build('drive', 'v3', credentials=creds)

    query = f"'{folder_id}' in parents and mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' and trashed = false"
    results = service.files().list(q=query, fields="files(id, name)").execute()
    items = results.get('files', [])

    if not items:
        print(f"No files found in folder {folder_id}.")
        return []

    file_paths = []
    for item in items:
        file_id = item['id']
        file_name = item['name']
        temp_file_path = os.path.join(TEMP_FOLDER, file_name)

        # Ensure the temporary folder exists
        os.makedirs(TEMP_FOLDER, exist_ok=True)

        request = service.files().get_media(fileId=file_id)
        with open(temp_file_path, 'wb') as f:
            downloader = MediaIoBaseDownload(f, request)
            done = False
            while not done:
                status, done = downloader.next_chunk()
        file_paths.append(temp_file_path)
    return file_paths

def clean_temp_folder():
    if os.path.exists(TEMP_FOLDER):
        for file in os.listdir(TEMP_FOLDER):
            file_path = os.path.join(TEMP_FOLDER, file)
            if os.path.isfile(file_path):
                os.remove(file_path)


def process_and_combine_files(token, output_file):
    soliscloud_folder = '1RFr5IXSq1LDEoK7c1jKQ-6o5tma5zzgg'
    solarman_folder = '13xDDTxXKt1PPbZwuMtANKPxM0L4s1dpL'
    isolarcloud_folder = '1LLOHDSHTZTETMy8bcnxYzD3LdwDM1sS8'
    canadian_folder = '1Sa5_6JuUBtoVtnvWuGc-LtVqDA9n5umQ'  

    all_dataframes = []

    # Process SolisCloud
    soliscloud_files = get_files_from_drive(soliscloud_folder, token)
    for file in soliscloud_files:
        all_dataframes.append(process_soliscloud(file))

    # Process Solarman
    solarman_files = get_files_from_drive(solarman_folder, token)
    for file in solarman_files:
        all_dataframes.append(process_solarman(file))

    # Process iSolarCloud
    isolarcloud_files = get_files_from_drive(isolarcloud_folder, token)
    for file in isolarcloud_files:
        all_dataframes.append(process_isolarcloud(file))

    # Process Canadian
    canadian_files = get_files_from_drive(canadian_folder, token)
    for file in canadian_files:
        all_dataframes.append(process_canadian(file))

    # Combine all dataframes
    combined_df = pd.concat(all_dataframes, ignore_index=True)

    # Save to Excel
    create_excel_with_table(combined_df, output_file)

    # Clean temporary folder
    clean_temp_folder()

    print(f"Arquivo combinado salvo em: {output_file}")

def create_excel_with_table(df, output_file):
    wb = Workbook()
    ws = wb.active
    ws.title = "Dados"

    for row in dataframe_to_rows(df, index=False, header=True):
        ws.append(row)

    # Define a tabela, mas sem usar estilos adicionais para evitar avisos
    table_range = f"A1:{chr(65 + len(df.columns) - 1)}{len(df) + 1}"
    table = Table(displayName="TabelaDados", ref=table_range)
    table.tableStyleInfo = None  # Remover estilos para evitar o aviso "no default style"
    ws.add_table(table)

    # Salva o arquivo Excel
    wb.save(output_file)

if __name__ == "__main__":
    output_file = 'GERAÇÃO_COMBINADAS.xlsx'  # Output file updated

    access_token, refresh_token = generate_tokens(client_id, client_secret)
    tokenAuth = {
        "token": access_token,
        "refresh_token": refresh_token,
        "token_uri": "https://oauth2.googleapis.com/token",
        "client_id": client_id,
        "client_secret": client_secret,
        "scopes": SCOPES
    }

    process_and_combine_files(tokenAuth, output_file)
    upload_file_to_drive(output_file, tokenAuth, TargetfolderId)
