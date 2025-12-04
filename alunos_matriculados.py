from __future__ import annotations
# -*- coding: utf-8 -*-

import os
import sys
from typing import Dict, Any, List

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# ------------------------------------------------------------
# CONFIG GOOGLE
# ------------------------------------------------------------
# Escopo mínimo para listar turmas e alunos (roster)
SCOPES = [
    "https://www.googleapis.com/auth/classroom.courses.readonly",
    "https://www.googleapis.com/auth/classroom.rosters.readonly", # Este é o essencial para a lista de alunos
]

CREDENTIALS_FILE = "credentials.json"
TOKEN_FILE = "token.json"


def get_credentials():
    """
    Carrega o token.json com os SCOPES corretos.
    Força re-autorização se o token não existir ou os escopos mudaram.
    """
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    if (
        not creds
        or not creds.valid
        or not set(SCOPES).issubset(set(creds.scopes or []))
    ):
        print("\n--- Necessário re-autorizar o acesso ---")
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDENTIALS_FILE):
                print("credentials.json não encontrado.")
                sys.exit(1)
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, "w", encoding="utf-8") as f:
            f.write(creds.to_json())
    return creds


# ------------------------------------------------------------
# CLASSROOM – ESCOLHER TURMA
# ------------------------------------------------------------
def escolher_turma(classroom_service) -> Dict[str, Any]:
    cursos: List[Dict[str, Any]] = []
    page_token = None

    while True:
        resp = classroom_service.courses().list(
            pageSize=100,
            courseStates=["ACTIVE"],
            pageToken=page_token
        ).execute()
        cursos.extend(resp.get("courses", []))
        page_token = resp.get("nextPageToken")
        if not page_token:
            break

    if not cursos:
        print("Nenhuma turma ativa encontrada.")
        sys.exit(0)

    print("\n=== TURMAS DISPONÍVEIS ===")
    for i, c in enumerate(cursos, start=1):
        print(f"{i:2d} - {c['name']} (ID {c['id']})")

    while True:
        op = input("\nEscolha o número da turma: ").strip()
        try:
            x = int(op)
            if 1 <= x <= len(cursos):
                return cursos[x - 1]
        except ValueError:
            pass
        print("Opção inválida.")

# ------------------------------------------------------------
# CLASSROOM – MAPA DE ALUNOS (email -> userId)
# ------------------------------------------------------------
def listar_e_baixar_alunos(classroom_service, course_id: str):
    """
    Lista todos os alunos da turma e exibe seus detalhes.
    """
    alunos: List[Dict[str, Any]] = []
    page_token = None

    try:
        while True:
            resp = classroom_service.courses().students().list(
                courseId=course_id,
                pageToken=page_token
            ).execute()

            alunos.extend(resp.get("students", []))
            page_token = resp.get("nextPageToken")
            if not page_token:
                break

        print(f"\n✅ Total de Alunos Carregados pela API: {len(alunos)}")
        
        if not alunos:
            print("AVISO: A API não retornou alunos. Verifique se o aluno está matriculado como 'Estudante'.")
            return

        print("\n--- DETALHES DOS ALUNOS ---")
        for i, st in enumerate(alunos, start=1):
            perfil = st.get("profile", {})
            email = perfil.get("emailAddress", "N/A")
            nome = perfil.get("name", {}).get("fullName", "N/A")
            user_id = st.get("userId", "N/A")
            
            print(f"{i:2d}. Nome: {nome} | E-mail: {email} | User ID: {user_id}")
        
    except HttpError as e:
        print(f"\nERRO: Falha ao listar alunos. O token de autorização pode estar insuficiente (SCOPES). Detalhes do erro: {e}")


# ------------------------------------------------------------
# MAIN
# ------------------------------------------------------------
def main():
    creds = get_credentials()
    print("\nEscopos do token carregado:")
    print(creds.scopes)

    classroom_service = build("classroom", "v1", credentials=creds)

    # Escolher turma
    curso = escolher_turma(classroom_service)
    course_id = curso["id"]
    print(f"\nTurma escolhida: {curso['name']} (ID {course_id})")

    listar_e_baixar_alunos(classroom_service, course_id)
    
    print("\nProcesso de listagem concluído.")


if __name__ == "__main__":
    main()