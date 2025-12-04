from __future__ import annotations
# """
# criar_treinamentos_classroom.py

# Cria automaticamente:
# - Cursos (turmas) no Google Classroom
# - Tópicos dentro dos cursos
# - Materiais com anexos DOCX enviados ao Google Drive

# SOLUÇÃO PARA O ERRO 403 (INSUFFICIENT SCOPES):
# O erro HttpError 403 na criação de tópico indica que o token de acesso (token.json)
# não possui a permissão "classroom.topics". Mesmo que o scope esteja na lista SCOPES,
# se ele foi adicionado após a primeira autenticação, o token antigo é inválido.
# A solução é EXCLUIR o arquivo 'token.json' para forçar uma nova autenticação
# e consentimento com TODOS os scopes.

# Pré-requisitos:
# - credentials.json salvo em C:\Users\Administrador\Documents\Estudos\python\treinamentos
# - APIs Google Classroom e Google Drive ativadas no projeto
# - pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib
# """

import os
import time
from pathlib import Path
from typing import Dict, Any, List, Tuple

from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from googleapiclient.errors import HttpError
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# ============================================================
# 1) CONFIGURAÇÕES GERAIS
# ============================================================

# Escopos: Classroom (criar cursos, materiais, tópicos) + Drive (upload arquivos)
SCOPES = [
    "https://www.googleapis.com/auth/classroom.courses",              # criar/editar cursos
    "https://www.googleapis.com/auth/classroom.courseworkmaterials", # criar materiais
    "https://www.googleapis.com/auth/classroom.topics",              # criar tópicos (CAUSA DO ERRO 403)
    "https://www.googleapis.com/auth/drive.file",                    # enviar arquivos pro Drive
]


# Pasta onde estão credentials.json e onde será criado o token.json
BASE_DIR = Path(r"C:\Users\Administrador\Documents\Estudos\python\treinamentos")

# Pasta onde estão os módulos DOCX
MODULOS_DIR = Path(
    r"C:\Users\Administrador\Documents\Empresas\Frascominas\RH\Treinamentos\MODULOS"
)

CREDENTIALS_FILE = BASE_DIR / "credentials.json"
TOKEN_FILE = BASE_DIR / "token.json"

# Opcional: id de uma pasta específica no Drive para guardar os DOCX.
# Se deixar como None, os arquivos irão para o "Meu Drive" raiz.
DRIVE_FOLDER_ID = None  # ex: "1AbCdEfG..."


# ============================================================
# 2) AUTENTICAÇÃO GOOGLE
# ============================================================

def get_credentials() -> Credentials:
    """
    Carrega credenciais do token.json; se não existir, usa credentials.json
    para gerar token via fluxo OAuth no navegador.

    Se o escopo for alterado, o 'token.json' deve ser deletado para forçar
    o consentimento com as novas permissões.
    """
    creds = None
    if TOKEN_FILE.exists():
        print(f"[{time.strftime('%H:%M:%S')}] Tentando carregar credenciais de {TOKEN_FILE.name}...")
        creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)
    # Se não há credenciais válidas, faz login
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            print(f"[{time.strftime('%H:%M:%S')}] Credenciais expiradas. Tentando renovar...")
            creds.refresh(Request())
        else:
            print(f"[{time.strftime('%H:%M:%S')}] Credenciais inválidas ou inexistentes.")
            print("Iniciando fluxo OAuth no navegador. Por favor, autorize.")
            try:
                flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_FILE), SCOPES)
                creds = flow.run_local_server(port=0)
            except FileNotFoundError:
                print(f"\n[ERRO FATAL] Arquivo de credenciais '{CREDENTIALS_FILE.name}' não encontrado.")
                print("Certifique-se de que o arquivo 'credentials.json' está na pasta correta.")
                exit(1)
            except Exception as e:
                print(f"\n[ERRO FATAL] Falha no fluxo de autenticação: {e}")
                exit(1)
        # Salva o token para a próxima vez
        with open(TOKEN_FILE, "w", encoding="utf-8") as token:
            token.write(creds.to_json())
        print(f"[{time.strftime('%H:%M:%S')}] Novo token salvo em {TOKEN_FILE.name}.")
    else:
         print(f"[{time.strftime('%H:%M:%S')}] Credenciais carregadas e válidas.")
    return creds


def get_services() -> Tuple[Any, Any]:
    """
    Retorna (classroom_service, drive_service).
    """
    creds = get_credentials()
    classroom_service = build("classroom", "v1", credentials=creds)
    drive_service = build("drive", "v3", credentials=creds)
    return classroom_service, drive_service


# ============================================================
# 3) ESTRUTURA DOS CURSOS / TÓPICOS / MATERIAIS (MANTIDA)
# ============================================================

# Cada curso terá:
# - name: nome da turma
# - section: texto adicional (ex: Frascominas)
# - topics: lista de tópicos, cada um com:
#   - name: nome do tópico
#   - materials: lista de materiais, cada um com:
#       - title: título no Classroom
#       - filename: nome do arquivo DOCX na pasta MODULOS

COURSES_STRUCTURE: List[Dict[str, Any]] = [
    {
        "name": "Treinamento Geral Obrigatório",
        "section": "Frascominas - Todos os colaboradores",
        "topics": [
            {
                "name": "Segurança do Trabalho",
                "materials": [
                    {
                        "title": "Módulo I - Treinamento e Segurança do Trabalho",
                        "filename": "MÓDULO I - Treinamento e segurança do trabalho.docx",
                    },
                    {
                        "title": "Treinamento Novo - Segurança",
                        "filename": "Treinamento novo segurança.docx",
                    },
                ],
            },
            {
                "name": "Atribuições e Responsabilidades",
                "materials": [
                    {
                        "title": "Módulo II - Atribuições e Responsabilidades",
                        "filename": "Módulo II - Atribuiões e responabilidades.docx",
                    }
                ],
            },
            {
                "name": "Primeiros Socorros",
                "materials": [
                    {
                        "title": "Módulo XIV - Primeiros Socorros",
                        "filename": "Módulo XIV - Primeiros Socorros.docx",
                    }
                ],
            },
            {
                "name": "Limpeza e Manutenção",
                "materials": [
                    {
                        "title": "Módulo XII - Limpeza e Manutenção Chão de Fábrica",
                        "filename": "Módulo XII - Limpeza e Manutenção chão de Fábrica.docx",
                    }
                ],
            },
            {
                "name": "Tipos de Máquinas e Funcionamento",
                "materials": [
                    {
                        "title": "Módulo XV - Tipos de Máquinas e Funcionamento",
                        "filename": "Módulo XV - Tipos de Máquinas e funcionamento.docx",
                    }
                ],
            },
        ],
    },
    {
        "name": "Produção PE - Treinamentos Operacionais",
        "section": "Frascominas - Produção Polietileno",
        "topics": [
            {
                "name": "Processo de Fabricação das Garrafas de Polietileno",
                "materials": [
                    {
                        "title": "Módulo III - Processo de Fabricação das Garrafas de Polietileno",
                        "filename": "Módulo III Processo de Fabricação das Garrafas de Polietileno.docx",
                    }
                ],
            },
            {
                "name": "Seladora",
                "materials": [
                    {
                        "title": "Módulo V - Seladora",
                        "filename": "Módulo V - Seladora.docx",
                    }
                ],
            },
            {
                "name": "Conhecimento e Manuseio do Cabeçote",
                "materials": [
                    {
                        "title": "Módulo VIII - Conhecimento e Manuseio Cabeçote",
                        "filename": "Módulo VIII - Conhecimento e Manuseio cabeçote.docx",
                    }
                ],
            },
            {
                "name": "Conhecimento e Manuseio Degolador",
                "materials": [
                    {
                        "title": "Módulo IX - Conhecimento e Manuseio Degolador",
                        "filename": "Módulo IX - Conhecimento e Manuseio Degolador.docx",
                    }
                ],
            },
            {
                "name": "Esteiras",
                "materials": [
                    {
                        "title": "Módulo XIII - Esteiras",
                        "filename": "Módulo XIII - Esteiras.docx",
                    }
                ],
            },
        ],
    },
    {
        "name": "Produção PET - Treinamentos Operacionais",
        "section": "Frascominas - Produção PET",
        "topics": [
            {
                "name": "Processo de Fabricação de PET",
                "materials": [
                    {
                        "title": "Módulo IV - Processo de Fabricação de PET",
                        "filename": "Módulo IV – Processo de fabricação de Pets.docx",
                    }
                ],
            },
            {
                "name": "Preparação e Configuração da Máquina de Sopro",
                "materials": [
                    {
                        "title": "Módulo XI - Preparação e Configuração da Máquina de Sopro",
                        "filename": "Módulo XI - Preparação e Config Maquina de Sopro.docx",
                    }
                ],
            },
            {
                "name": "Operação dos Sistemas de Resfriamento",
                "materials": [
                    {
                        "title": "Módulo X - Operação dos Sistemas de Resfriamento",
                        "filename": "Módulo X Operação sos Sistemas de Resfriamento.docx",
                    }
                ],
            },
            {
                "name": "Processo de Moagem",
                "materials": [
                    {
                        "title": "Módulo VII - Processo de Moagem",
                        "filename": "Módulo VII - Processo de Moagem.docx",
                    }
                ],
            },
        ],
    },
    {
        "name": "Treinamentos de Suporte - Almoxarifado, Paletização, Empilhadeira",
        "section": "Frascominas - Suporte Operacional",
        "topics": [
            {
                "name": "Almoxarifado",
                "materials": [
                    {
                        "title": "Módulo XVI - Almoxarifado",
                        "filename": "Modulo XVI - Almoxarifado.docx",
                    }
                ],
            },
            {
                "name": "Paletização",
                "materials": [
                    {
                        "title": "Módulo XVII - Paletização",
                        "filename": "Módulo XVII- Paletização.docx",
                    }
                ],
            },
            {
                "name": "Empilhadeira",
                "materials": [
                    {
                        "title": "Módulo XVIII - Empilhadeira",
                        "filename": "Módulo XVIII - Empilhadeira.docx",
                    }
                ],
            },
        ],
    },
]


# ============================================================
# 4) FUNÇÕES AUXILIARES: DRIVE E CLASSROOM
# ============================================================

def upload_to_drive(drive_service, local_path: Path) -> str:
    """
    Faz upload de um arquivo local ao Google Drive e retorna o fileId.
    """
    file_metadata = {"name": local_path.name}
    if DRIVE_FOLDER_ID:
        file_metadata["parents"] = [DRIVE_FOLDER_ID]

    media = MediaFileUpload(
        str(local_path),
        mimetype=(
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        ),
        resumable=True,
    )

    try:
        file = (
            drive_service.files()
            .create(body=file_metadata, media_body=media, fields="id")
            .execute()
        )
        file_id = file["id"]
        print(f"[{time.strftime('%H:%M:%S')}] [DRIVE] Upload concluído: {local_path.name} -> id={file_id}")
        return file_id
    except HttpError as err:
        print(f"[{time.strftime('%H:%M:%S')}] [DRIVE] ERRO no upload de {local_path.name}: {err}")
        raise


def create_course(classroom_service, name: str, section: str) -> str:
    """
    Cria um curso no Classroom e retorna o courseId.
    """
    course_body = {
        "name": name,
        "section": section,
        # O ownerId "me" usa o usuário autenticado
        "ownerId": "me",
        "courseState": "ACTIVE",
    }
    try:
        course = classroom_service.courses().create(body=course_body).execute()
        course_id = course["id"]
        print(f"[{time.strftime('%H:%M:%S')}] [CLASSROOM] Curso criado: {name} (id={course_id})")
        return course_id
    except HttpError as err:
        print(f"[{time.strftime('%H:%M:%S')}] [CLASSROOM] ERRO ao criar curso {name}: {err}")
        raise


def create_topic(classroom_service, course_id: str, topic_name: str) -> str:
    """
    Cria um tópico em um curso e retorna o topicId.
    """
    body = {"name": topic_name}
    try:
        topic = (
            classroom_service.courses()
            .topics()
            .create(courseId=course_id, body=body)
            .execute()
        )
        topic_id = topic["topicId"]
        print(f"[{time.strftime('%H:%M:%S')}] [CLASSROOM]   Tópico criado: {topic_name} (topicId={topic_id})")
        return topic_id
    except HttpError as err:
        print(f"[{time.strftime('%H:%M:%S')}] [CLASSROOM]   ERRO ao criar tópico {topic_name}: {err}")
        # Se for o erro 403, adiciona uma mensagem útil
        if err.resp.status == 403 and 'ACCESS_TOKEN_SCOPE_INSUFFICIENT' in str(err):
             print("\n!!! OCORREU O ERRO DE SCOPE 403 NA CRIAÇÃO DE TÓPICO !!!")
             print("SOLUÇÃO: Feche o programa e EXCLUA o arquivo 'token.json' na sua BASE_DIR.")
             print("Isso forçará a reautenticação com as permissões corretas.")
        raise


def create_material(
    classroom_service,
    course_id: str,
    topic_id: str,
    title: str,
    drive_file_id: str,
    description: str = "",
) -> None:
    """
    Cria um material (courseWorkMaterial) no Classroom associado a um arquivo do Drive.
    """
    body = {
        "title": title,
        "description": description,
        "materials": [
            {
                "driveFile": {
                    "driveFile": {
                        "id": drive_file_id,
                    }
                }
            }
        ],
        "topicId": topic_id,
        "state": "PUBLISHED" # Garante que o material é visível imediatamente
    }
    try:
        material = (
            classroom_service.courses()
            .courseWorkMaterials()
            .create(courseId=course_id, body=body)
            .execute()
        )
        material_id = material["id"]
        print(f"[{time.strftime('%H:%M:%S')}] [CLASSROOM]     Material criado: {title} (id={material_id})")
    except HttpError as err:
        print(f"[{time.strftime('%H:%M:%S')}] [CLASSROOM]     ERRO ao criar material {title}: {err}")
        raise


# ============================================================
# 5) PIPELINE PRINCIPAL
# ============================================================

def main():
    print("============================================")
    print("INÍCIO DO PROCESSAMENTO DE CRIAÇÃO DE CURSOS")
    print("============================================")
    print(f"Pasta Base: {BASE_DIR}")
    print(f"Pasta de Módulos (DOCX): {MODULOS_DIR}")
    
    try:
        # 1) Autentica e obtém serviços
        classroom_service, drive_service = get_services()

        # 2) Percorre a estrutura de cursos
        for course_def in COURSES_STRUCTURE:
            course_name = course_def["name"]
            section = course_def.get("section", "")

            print(f"\n--- Iniciando Curso: {course_name} ---")

            # 2.1) Cria curso
            course_id = create_course(classroom_service, course_name, section)

            # 2.2) Cria tópicos
            for topic_def in course_def.get("topics", []):
                topic_name = topic_def["name"]
                
                # Cria tópico
                topic_id = create_topic(classroom_service, course_id, topic_name)

                # 2.3) Cria materiais (upload para Drive + material no Classroom)
                for mat in topic_def.get("materials", []):
                    title = mat["title"]
                    # Normaliza o nome do arquivo para garantir que espaços/caracteres especiais
                    # do CODES_STRUCTURE coincidam com o nome real do arquivo.
                    filename = mat["filename"].replace(" ", " ") 

                    local_path = MODULOS_DIR / filename
                    
                    print(f"[{time.strftime('%H:%M:%S')}] Processando material: {title} (Arquivo: {filename})")
                    
                    if not local_path.exists():
                        print(
                            f"[{time.strftime('%H:%M:%S')}] [AVISO] Arquivo não encontrado: {local_path}. "
                            f"Pulei este material."
                        )
                        continue

                    # Upload para o Drive
                    file_id = upload_to_drive(drive_service, local_path)

                    # Descrição padrão
                    description = (
                        "Material obrigatório do treinamento. "
                        "Leia o conteúdo e, se houver formulário de avaliação, responda ao final."
                    )

                    # Cria material no Classroom
                    create_material(
                        classroom_service,
                        course_id,
                        topic_id,
                        title,
                        file_id,
                        description=description,
                    )

        print("\n==============================================")
        print("===== PROCESSO CONCLUÍDO COM SUCESSO ======")
        print("==============================================")
        print(
            "Verifique no Google Classroom os cursos criados, tópicos e materiais com os DOCX anexados."
        )

    except Exception as e:
        print(f"\n[ERRO FATAL NO PIPELINE PRINCIPAL]: {e}")
        print("O processo foi interrompido.")


if __name__ == "__main__":
    # Remove as linhas do 'if __name__' para que o script funcione no contexto do Immersive.
    # Vou manter aqui, mas é só um lembrete do contexto de execução.
    main()