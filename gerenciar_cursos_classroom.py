from __future__ import annotations
# """
# gerenciar_cursos_classroom.py

# Funções para listar, arquivar e excluir cursos do Google Classroom.
# - Para 'cancelar' um curso, você precisa primeiro ARQUIVÁ-LO.
# - A exclusão (DELETE) é permanente e só funciona para cursos ARQUIVADOS.
# """

import os
import time
from pathlib import Path
from typing import Dict, Any, List, Tuple

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# ============================================================
# 1) CONFIGURAÇÕES GERAIS (As mesmas do script anterior)
# ============================================================

# Escopos necessários para listar, arquivar e deletar
SCOPES = [
    "https://www.googleapis.com/auth/classroom.courses",              # Permite listar, criar, arquivar e deletar
]


# Pasta onde estão credentials.json e token.json
BASE_DIR = Path(r"C:\Users\Administrador\Documents\Estudos\python\treinamentos")
CREDENTIALS_FILE = BASE_DIR / "credentials.json"
TOKEN_FILE = BASE_DIR / "token.json"


# ============================================================
# 2) AUTENTICAÇÃO GOOGLE (Reutilizando as funções anteriores)
# ============================================================

def get_credentials() -> Credentials:
    """
    Carrega credenciais do token.json; se não existir, usa credentials.json
    para gerar token via fluxo OAuth no navegador.
    """
    creds = None
    if TOKEN_FILE.exists():
        print(f"[{time.strftime('%H:%M:%S')}] Tentando carregar credenciais...")
        # NOTA: O SCOPE foi alterado para incluir apenas 'classroom.courses'
        # o que é suficiente para arquivar/excluir.
        creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES) 

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            print(f"[{time.strftime('%H:%M:%S')}] Credenciais inválidas. Iniciando fluxo OAuth.")
            try:
                flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_FILE), SCOPES)
                creds = flow.run_local_server(port=0)
            except FileNotFoundError:
                print(f"\n[ERRO FATAL] Arquivo de credenciais '{CREDENTIALS_FILE.name}' não encontrado.")
                exit(1)
            
        with open(TOKEN_FILE, "w", encoding="utf-8") as token:
            token.write(creds.to_json())
        print(f"[{time.strftime('%H:%M:%S')}] Novo token salvo.")
    else:
         print(f"[{time.strftime('%H:%M:%S')}] Credenciais carregadas e válidas.")
    return creds


def get_classroom_service() -> Any:
    """
    Retorna o serviço do Google Classroom.
    """
    creds = get_credentials()
    classroom_service = build("classroom", "v1", credentials=creds)
    return classroom_service


# ============================================================
# 3) FUNÇÕES DE GERENCIAMENTO
# ============================================================

def list_courses(service: Any, course_state: str = "ACTIVE") -> List[Dict[str, str]]:
    """
    Lista todos os cursos em um determinado estado (ACTIVE, ARCHIVED, PROVISIONED, DECLINED).
    Retorna uma lista de dicionários com 'id' e 'name'.
    """
    print(f"\n[{time.strftime('%H:%M:%S')}] Buscando cursos no estado: {course_state}...")
    results = service.courses().list(courseStates=course_state).execute()
    courses = results.get("courses", [])
    
    course_list = []
    if not courses:
        print(f"[{time.strftime('%H:%M:%S')}] Nenhum curso encontrado no estado {course_state}.")
        return []
    
    print("-" * 50)
    print(f"Cursos Encontrados (Estado: {course_state}):")
    for course in courses:
        course_data = {
            "id": course["id"], 
            "name": course["name"],
            "section": course.get("section", "N/A")
        }
        course_list.append(course_data)
        print(f"ID: {course_data['id']} | Nome: {course_data['name']} ({course_data['section']})")
    print("-" * 50)
    return course_list


def archive_course(service: Any, course_id: str) -> bool:
    """
    Altera o estado do curso para ARCHIVED (arquivado).
    """
    try:
        course = service.courses().get(id=course_id).execute()
        
        if course.get("courseState") == "ARCHIVED":
            print(f"[{time.strftime('%H:%M:%S')}] [AVISO] Curso ID {course_id} já está ARQUIVADO.")
            return True

        course["courseState"] = "ARCHIVED"
        service.courses().update(id=course_id, body=course).execute()
        print(f"[{time.strftime('%H:%M:%S')}] [SUCESSO] Curso ID {course_id} ARQUIVADO com sucesso!")
        return True
    except HttpError as err:
        print(f"[{time.strftime('%H:%M:%S')}] [ERRO] Falha ao ARQUIVAR curso ID {course_id}: {err}")
        return False
    except Exception as e:
        print(f"[{time.strftime('%H:%M:%S')}] [ERRO] Ocorreu um erro inesperado: {e}")
        return False


def delete_archived_course(service: Any, course_id: str) -> bool:
    """
    Exclui permanentemente um curso ARQUIVADO. 
    Atenção: A exclusão é irreversível.
    """
    try:
        course = service.courses().get(id=course_id).execute()
        if course.get("courseState") != "ARCHIVED":
            print(f"[{time.strftime('%H:%M:%S')}] [AVISO] O curso ID {course_id} NÃO está ARQUIVADO. Exclua-o primeiro.")
            return False
            
        service.courses().delete(id=course_id).execute()
        print(f"[{time.strftime('%H:%M:%S')}] [SUCESSO] Curso ID {course_id} EXCLUÍDO permanentemente.")
        return True
    except HttpError as err:
        print(f"[{time.strftime('%H:%M:%S')}] [ERRO] Falha ao EXCLUIR curso ID {course_id}: {err}")
        return False
    except Exception as e:
        print(f"[{time.strftime('%H:%M:%S')}] [ERRO] Ocorreu um erro inesperado: {e}")
        return False


# ============================================================
# 4) EXEMPLO DE USO
# ============================================================

def run_manager():
    service = get_classroom_service()

    # PASSO 1: LISTAR CURSOS ATIVOS E ARQUIVADOS (para conferência)
    active_courses = list_courses(service, course_state="ACTIVE")
    archived_courses = list_courses(service, course_state="ARCHIVED")

    print("\n============================================")
    print("INSTRUÇÕES: ")
    print("1. Veja acima os IDs dos cursos que você quer remover.")
    print("2. Preencha a lista 'courses_to_remove' abaixo apenas com os IDs que deseja arquivar/excluir.")
    print("============================================")

    # IDs dos cursos que você QUER REMOVER (ativos ou arquivados)
    courses_to_remove = [
        "832440426942",
        "832439430803",
        "832439977461",
        "832438798170",
        "832437087456",
    ]

    # PASSO 2: ARQUIVAR (se ainda estiverem ativos) E EXCLUIR
    for course_id in courses_to_remove:
        print(f"\n>>> Processando curso ID {course_id}...")

        # 2.1) Arquiva se ainda não estiver arquivado
        archive_course(service, course_id)

        # 2.2) Exclui definitivamente (só funciona se estiver ARQUIVADO)
        delete_archived_course(service, course_id)

if __name__ == "__main__":
    run_manager()