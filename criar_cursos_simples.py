from __future__ import annotations
"""
criar_cursos_simples.py

Script para CRIAR cursos (turmas) no Google Classroom,
sem criar tópicos nem materiais.

Você apenas informa os parâmetros no terminal:
- nome do curso
- seção (ex: Frascominas - Produção PET)
- sala (opcional)
- descrição (opcional)

Pré-requisitos:
- credentials.json em BASE_DIR
- APIs Classroom ativadas
- pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib

Campo	Para que serve	Precisa preencher?	Exemplo
name	Nome da turma	Sim	“Treinamento Geral Obrigatório”
section	Subtítulo / categoria	Não	“Frascominas – Produção PET”
room	Local físico/virtual	Não	“Online”
descriptionHeading	Descrição curta	Não	“Treinamento para novos colaboradores”
courseState	Estado da turma	Não	ACTIVE (padrão)

"""

import time
from pathlib import Path
from typing import Any

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# ============================================================
# CONFIGURAÇÕES
# ============================================================

SCOPES = [
    "https://www.googleapis.com/auth/classroom.courses",  # criar/editar/listar cursos
]

BASE_DIR = Path(r"C:\Users\Administrador\Documents\Estudos\python\treinamentos")
CREDENTIALS_FILE = BASE_DIR / "credentials.json"
TOKEN_FILE = BASE_DIR / "token_cursos.json"  # token separado só para este script (opcional)


# ============================================================
# AUTENTICAÇÃO
# ============================================================

def get_credentials() -> Credentials:
    creds = None
    if TOKEN_FILE.exists():
        print(f"[{time.strftime('%H:%M:%S')}] Tentando carregar token existente ({TOKEN_FILE.name})...")
        creds = Credentials.from_authorized_user_file(str(TOKEN_FILE), SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            print(f"[{time.strftime('%H:%M:%S')}] Token expirado. Renovando...")
            creds.refresh(Request())
        else:
            print(f"[{time.strftime('%H:%M:%S')}] Criando novo token via OAuth...")
            flow = InstalledAppFlow.from_client_secrets_file(str(CREDENTIALS_FILE), SCOPES)
            creds = flow.run_local_server(port=0)
            with open(TOKEN_FILE, "w", encoding="utf-8") as token:
                token.write(creds.to_json())
            print(f"[{time.strftime('%H:%M:%S')}] Novo token salvo em {TOKEN_FILE.name}.")
    else:
        print(f"[{time.strftime('%H:%M:%S')}] Token válido carregado.")

    return creds


def get_classroom_service() -> Any:
    creds = get_credentials()
    service = build("classroom", "v1", credentials=creds)
    return service


# ============================================================
# FUNÇÃO PARA CRIAR CURSO
# ============================================================

def criar_curso(
    service: Any,
    name: str,
    section: str = "",
    room: str = "",
    description: str = "",
    state: str = "ACTIVE",
) -> None:
    body = {
        "name": name,
        "ownerId": "me",          # usuário autenticado será o dono
        "courseState": state,
    }
    if section:
        body["section"] = section
    if room:
        body["room"] = room
    if description:
        body["descriptionHeading"] = description  # aparece no topo do curso

    try:
        course = service.courses().create(body=body).execute()
        cid = course["id"]
        code = course.get("enrollmentCode", "(sem código)")
        print(f"\n[SUCESSO] Curso criado:")
        print(f"  ID:           {cid}")
        print(f"  Nome:         {course['name']}")
        print(f"  Seção:        {course.get('section', '')}")
        print(f"  Sala:         {course.get('room', '')}")
        print(f"  Código de ingresso (para alunos): {code}")
        print("-" * 60)
    except HttpError as err:
        print(f"[ERRO] Falha ao criar curso '{name}': {err}")


# ============================================================
# MODO INTERATIVO
# ============================================================

def main():
    print("============================================")
    print("   CRIADOR SIMPLES DE CURSOS - CLASSROOM    ")
    print("============================================")

    service = get_classroom_service()

    while True:
        print("\nInforme os dados do novo curso.")
        print("Deixe o NOME em branco para encerrar.\n")

        name = input("Nome do curso (obrigatório): ").strip()
        if not name:
            print("Nenhum nome informado. Encerrando.")
            break

        section = input("Seção (ex: Frascominas - Produção PET) [opcional]: ").strip()
        room = input("Sala (ex: Sala 1, Online, Turno A) [opcional]: ").strip()
        description = input("Descrição curta [opcional]: ").strip()

        criar_curso(
            service,
            name=name,
            section=section,
            room=room,
            description=description,
        )

        continuar = input("\nCriar outro curso? (s/N): ").strip().lower()
        if continuar != "s":
            break

    print("\nFim do processo. Verifique os cursos em https://classroom.google.com")


if __name__ == "__main__":
    main()
