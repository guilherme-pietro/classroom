from __future__ import annotations
"""
limpar_tudo_classroom.py

Remove TODOS os:
- courseWorkMaterials (materiais)
- courseWork (tarefas/atividades)

dos cursos selecionados do Google Classroom.

NÃO exclui cursos nem alunos.
"""

import time
from pathlib import Path
from typing import Any, List, Dict

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# ============================================================
# 1) CONFIGURAÇÕES
# ============================================================

SCOPES = [
    "https://www.googleapis.com/auth/classroom.courses",              # listar cursos
    "https://www.googleapis.com/auth.classroom.coursework.me",       # (não é estritamente necessário, mas não atrapalha)
    "https://www.googleapis.com/auth/classroom.coursework.students", # deletar courseWork
    "https://www.googleapis.com/auth/classroom.courseworkmaterials", # listar/deletar materiais
]

BASE_DIR = Path(r"C:\Users\Administrador\Documents\Estudos\python\treinamentos")
CREDENTIALS_FILE = BASE_DIR / "credentials.json"
TOKEN_FILE = BASE_DIR / "token.json"


# ============================================================
# 2) AUTENTICAÇÃO
# ============================================================

def get_credentials() -> Credentials:
    creds = None
    if TOKEN_FILE.exists():
        print(f"[{time.strftime('%H:%M:%S')}] Tentando carregar token existente...")
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
# 3) FUNÇÕES DE APOIO
# ============================================================

def listar_cursos(service: Any) -> List[Dict[str, str]]:
    """
    Lista cursos em estados ACTIVE e ARCHIVED.
    Retorna lista de dicts com id, name, state.
    """
    cursos: List[Dict[str, str]] = []

    for estado in ["ACTIVE", "ARCHIVED"]:
        print(f"\n[{time.strftime('%H:%M:%S')}] Buscando cursos estado={estado}...")
        page_token = None
        while True:
            res = service.courses().list(
                courseStates=estado,
                pageToken=page_token
            ).execute()
            for c in res.get("courses", []):
                cursos.append(
                    {
                        "id": c["id"],
                        "name": c["name"],
                        "state": c.get("courseState", estado),
                        "section": c.get("section", ""),
                    }
                )
            page_token = res.get("nextPageToken")
            if not page_token:
                break

    print("\n===== CURSOS ENCONTRADOS =====")
    for c in cursos:
        print(f"ID: {c['id']} | Nome: {c['name']} | Estado: {c['state']} | Section: {c['section']}")
    print("================================")
    return cursos


def deletar_materiais(service: Any, course_id: str) -> int:
    """
    Deleta TODOS os courseWorkMaterials de um curso.
    Retorna quantidade apagada.
    """
    print(f"\n[{time.strftime('%H:%M:%S')}] Listando MATERIAIS do curso {course_id}...")
    page_token = None
    total = 0

    while True:
        try:
            res = service.courses().courseWorkMaterials().list(
                courseId=course_id,
                pageToken=page_token
            ).execute()
        except HttpError as err:
            print(f"[{time.strftime('%H:%M:%S')}] ERRO ao listar materiais do curso {course_id}: {err}")
            return total

        materials = res.get("courseWorkMaterial", [])
        if not materials and not page_token:
            print(f"[{time.strftime('%H:%M:%S')}] Nenhum material encontrado neste curso.")
            break

        for m in materials:
            mat_id = m["id"]
            mat_title = m.get("title", "(sem título)")
            print(f"[{time.strftime('%H:%M:%S')}]   Deletando material: {mat_title} (id={mat_id})...")
            try:
                service.courses().courseWorkMaterials().delete(
                    courseId=course_id,
                    id=mat_id
                ).execute()
                total += 1
            except HttpError as err:
                print(f"[{time.strftime('%H:%M:%S')}]   ERRO ao deletar {mat_title} (id={mat_id}): {err}")

        page_token = res.get("nextPageToken")
        if not page_token:
            break

    print(f"[{time.strftime('%H:%M:%S')}] Total de materiais deletados neste curso: {total}")
    return total


def deletar_coursework(service: Any, course_id: str) -> int:
    """
    Deleta TODOS os courseWork (tarefas, trabalhos) de um curso.
    Retorna quantidade apagada.
    """
    print(f"\n[{time.strftime('%H:%M:%S')}] Listando TAREFAS (courseWork) do curso {course_id}...")
    page_token = None
    total = 0

    while True:
        try:
            res = service.courses().courseWork().list(
                courseId=course_id,
                pageToken=page_token
            ).execute()
        except HttpError as err:
            print(f"[{time.strftime('%H:%M:%S')}] ERRO ao listar courseWork do curso {course_id}: {err}")
            return total

        works = res.get("courseWork", [])
        if not works and not page_token:
            print(f"[{time.strftime('%H:%M:%S')}] Nenhuma tarefa encontrada neste curso.")
            break

        for w in works:
            work_id = w["id"]
            work_title = w.get("title", "(sem título)")
            print(f"[{time.strftime('%H:%M:%S')}]   Deletando tarefa: {work_title} (id={work_id})...")
            try:
                service.courses().courseWork().delete(
                    courseId=course_id,
                    id=work_id
                ).execute()
                total += 1
            except HttpError as err:
                print(f"[{time.strftime('%H:%M:%S')}]   ERRO ao deletar {work_title} (id={work_id}): {err}")

        page_token = res.get("nextPageToken")
        if not page_token:
            break

    print(f"[{time.strftime('%H:%M:%S')}] Total de tarefas deletadas neste curso: {total}")
    return total


# ============================================================
# 4) PIPELINE PRINCIPAL
# ============================================================

def main():
    print("============================================")
    print(" APAGAR MATERIAIS E TAREFAS DOS CURSOS ")
    print("============================================")

    service = get_classroom_service()

    # 1) Lista todos os cursos (ativos + arquivados)
    cursos = listar_cursos(service)

    # 2) DEFINIR QUAIS CURSOS LIMPAR
    #
    # a) Para limpar TODOS os cursos listados:
    course_ids = [c["id"] for c in cursos]
    #
    # b) Se quiser limitar, comente a linha acima e use algo assim:
    # course_ids = [
    #     "ID_CURSO_1",
    #     "ID_CURSO_2",
    # ]

    print("\nCursos que terão TODOS os materiais e tarefas removidos:")
    for cid in course_ids:
        nome = next((c["name"] for c in cursos if c["id"] == cid), cid)
        print(f"- {cid} | {nome}")

    confirm = input("\nDIGITE 'APAGAR_TUDO' para confirmar a remoção de MATERIAIS + TAREFAS desses cursos: ").strip()
    if confirm != "APAGAR_TUDO":
        print("Operação cancelada pelo usuário.")
        return

    # 3) Limpa curso a curso
    for cid in course_ids:
        nome = next((c["name"] for c in cursos if c["id"] == cid), cid)
        print("\n======================================")
        print(f"Curso: {nome} (id={cid})")
        print("Removendo materiais...")
        deletar_materiais(service, cid)
        print("Removendo tarefas...")
        deletar_coursework(service, cid)

    print("\n==============================================")
    print("===== LIMPEZA DE MATERIAIS E TAREFAS CONCLUÍDA ========")
    print("==============================================")


if __name__ == "__main__":
    main()
