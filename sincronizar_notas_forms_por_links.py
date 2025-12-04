from __future__ import annotations
# -*- coding: utf-8 -*-

"""
sincronizar_notas_forms_por_links.py - COM SHEET_ID FIXO
"""

import os
import sys
import re
from typing import Dict, Any, List, Optional

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# ------------------------------------------------------------
# CONFIG GOOGLE
# ------------------------------------------------------------
SCOPES = [
    # Cursos, tarefas, submissões, alunos
    "https://www.googleapis.com/auth/classroom.courses.readonly",
    "https://www.googleapis.com/auth/classroom.coursework.students",
    "https://www.googleapis.com/auth/classroom.rosters.readonly",
    "https://www.googleapis.com/auth/classroom.student-submissions.students.readonly",
    
    # Forms (mantido)
    "https://www.googleapis.com/auth/forms.responses.readonly",
    "https://www.googleapis.com/auth/forms.body",  
    "https://www.googleapis.com/auth/drive.file", 
    
    # Sheets API
    "https://www.googleapis.com/auth/spreadsheets.readonly",
]

CREDENTIALS_FILE = "credentials.json"
TOKEN_FILE = "token.json"

# ====================================================================
# CONFIGURAÇÃO DE ID FIXO
# 
# Cole aqui o ID da Planilha que você acabou de usar.
# O script tentará usar este ID antes de pedir a entrada do usuário.
# ====================================================================
SHEET_ID_PADRAO = "1PaPwGxEVYPsLxetchqeD4WYxpn-vV8xlTMHKz6nE8uI"


def get_credentials():
    """
    Carrega o token.json com os SCOPES corretos.
    Se o token não existir ou os escopos mudaram, refaz a autorização.
    """
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    if (
        not creds
        or not creds.valid
        or not set(SCOPES).issubset(set(creds.scopes or []))
    ):
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
def mapear_alunos_por_email(classroom_service, course_id: str) -> Dict[str, str]:
    """
    Retorna dict: email_em_minúsculas -> userId
    INCLUI MAPA DE EMERGÊNCIA para contornar e-mails ocultos.
    """
    mapa: Dict[str, str] = {}
    page_token = None

    # ====================================================================
    # FALLBACK DE EMERGÊNCIA (CÓPIA DO SEU ÚLTIMO ESTADO DE TRABALHO)
    FALLBACK_EMAIL_MAP = {
        "martinezelaine838@gmail.com": "110514942462896649781"
    }
    # ====================================================================

    total_alunos_api = 0
    
    # 1. Tentar mapear alunos da API
    while True:
        resp = classroom_service.courses().students().list(
            courseId=course_id,
            pageToken=page_token
        ).execute()

        for st in resp.get("students", []):
            total_alunos_api += 1
            perfil = st.get("profile", {})
            email = perfil.get("emailAddress")
            user_id = st.get("userId")
            
            if email and user_id:
                mapa[email.strip().lower()] = user_id
            elif not email and user_id:
                # Aviso sobre o aluno Elaine Martinez
                print(f"ATENÇÃO: Aluno '{perfil.get('name', {}).get('fullName', 'Desconhecido')}' (ID: {user_id}) ignorado - EMAIL OCULTO pela API.")


        page_token = resp.get("nextPageToken")
        if not page_token:
            break

    # 2. Adicionar alunos do mapa de fallback
    for email, user_id in FALLBACK_EMAIL_MAP.items():
        if user_id not in mapa.values():
             mapa[email.strip().lower()] = user_id
             print(f"✅ Adicionado FALLBACK: {email} (ID: {user_id})")

    print(f"\nAlunos totais encontrados pela API: {total_alunos_api}")
    print(f"Alunos mapeados por e-mail (para sincronização): {len(mapa)}")
    return mapa


# ------------------------------------------------------------
# CLASSROOM – MAPA DE SUBMISSÕES (userId -> studentSubmission)
# ------------------------------------------------------------
def mapear_submissoes_por_user(
    classroom_service,
    course_id: str,
    coursework_id: str
) -> Dict[str, Dict[str, Any]]:
    """
    Retorna dict: userId -> objeto studentSubmission
    """
    mapa: Dict[str, Dict[str, Any]] = {}
    page_token = None

    while True:
        resp = classroom_service.courses().courseWork().studentSubmissions().list(
            courseId=course_id,
            courseWorkId=coursework_id,
            pageToken=page_token
        ).execute()

        for sub in resp.get("studentSubmissions", []):
            user_id = sub.get("userId")
            if user_id:
                mapa[user_id] = sub

        page_token = resp.get("nextPageToken")
        if not page_token:
            break

    print(f"  Submissões carregadas: {len(mapa)} (userId -> studentSubmission)")
    return mapa


# ------------------------------------------------------------
# SHEETS – MAPA DE NOTAS (email -> totalScore)
# ------------------------------------------------------------
def mapear_notas_do_forms_via_sheets(sheets_service, sheet_id: str) -> Dict[str, float]:
    """
    Lê respostas do Forms pela Planilha e retorna:
        email_em_minúsculas -> totalScore (float)
    """
    notas: Dict[str, float] = {}

    # Lendo apenas o cabeçalho para descobrir a posição das colunas
    cabecalho = sheets_service.spreadsheets().values().get(
        spreadsheetId=sheet_id,
        range='A1:Z1' # Assume que as colunas estão em A a Z na linha 1
    ).execute().get('values', [[]])[0]
    
    # Tenta localizar as colunas pelo nome. Caso contrário, assume colunas 2 e 3 (índices 1 e 2)
    try:
        email_col_idx = cabecalho.index('Nome de usuário')
        score_col_idx = cabecalho.index('Pontuação total')
        print(f"INFO: Colunas 'Nome de usuário' (índice {email_col_idx}) e 'Pontuação total' (índice {score_col_idx}) localizadas pelo nome.")
    except ValueError:
        print("\nAVISO: Não foi possível localizar as colunas 'Nome de usuário' e 'Pontuação total' pelo nome.")
        print("Assumindo que estão nas colunas B (índice 1) e C (índice 2) da Planilha.")
        email_col_idx = 1
        score_col_idx = 2
        # Verificação básica para evitar IndexError se a planilha for muito pequena
        if len(cabecalho) < 3:
             print("ERRO: A planilha é muito pequena. Colunas B e C não existem.")
             return {}
    
    # Lendo todas as linhas de dados (a partir da linha 2)
    range_data = sheets_service.spreadsheets().values().get(
        spreadsheetId=sheet_id,
        range='A2:Z' # Assume que os dados começam na linha 2
    ).execute().get('values', [])
    
    print(f"\n  Lendo {len(range_data)} respostas na Planilha vinculada...")

    for row in range_data:
        try:
            email = row[email_col_idx].strip()
            score_str = row[score_col_idx].split('/')[0].strip() # Pega só a nota (ex: "1.00 / 10" -> "1.00")
            total_score = float(score_str)
        except (IndexError, ValueError):
            # Ignora linhas mal formatadas ou sem dados nas colunas esperadas
            continue
        
        email_key = email.lower()
        antiga = notas.get(email_key)
        # Mantém a maior nota se houver mais de uma submissão
        if antiga is None or total_score > antiga:
            notas[email_key] = total_score

    print(f"  Respostas com nota: {len(notas)} (email -> totalScore)")
    return notas


# ------------------------------------------------------------
# CLASSROOM – BUSCAR MAX_POINTS DA TAREFA
# ------------------------------------------------------------
def obter_max_points_do_coursework(coursework: Dict[str, Any]) -> Optional[float]:
    mp = coursework.get("maxPoints")
    if mp is not None:
        print(f"  maxPoints da tarefa: {mp}")
    else:
        print("  Aviso: tarefa sem maxPoints definido.")
    return mp


# ------------------------------------------------------------
# CLASSROOM – APLICAR NOTAS
# ------------------------------------------------------------
def aplicar_notas(
    classroom_service,
    course_id: str,
    coursework_id: str,
    mapa_email_para_user: Dict[str, str],
    mapa_user_para_sub: Dict[str, Dict[str, Any]],
    notas_forms: Dict[str, float],
    max_points: Optional[float],
):
    sucesso = 0
    ignorados_sem_aluno = 0
    ignorados_sem_sub = 0

    for email_key, nota in notas_forms.items():
        user_id = mapa_email_para_user.get(email_key)
        if not user_id:
            print(f"    [IGNORADO] Sem aluno no Classroom com e-mail: {email_key}")
            ignorados_sem_aluno += 1
            continue

        sub = mapa_user_para_sub.get(user_id)
        if not sub:
            print(f"    [IGNORADO] Aluno {email_key} (userId {user_id}) não tem submissão para essa tarefa.")
            ignorados_sem_sub += 1
            continue

        sub_id = sub["id"]

        nota_aplicada = float(nota)
        if max_points is not None and nota_aplicada > float(max_points):
            nota_aplicada = float(max_points)

        body = {
            "draftGrade": nota_aplicada,
            "assignedGrade": nota_aplicada,
        }

        try:
            classroom_service.courses().courseWork().studentSubmissions().patch(
                courseId=course_id,
                courseWorkId=coursework_id,
                id=sub_id,
                body=body,
                updateMask="draftGrade,assignedGrade",
            ).execute()
            print(f"    [OK] {email_key} -> nota {nota_aplicada} (submission {sub_id})")
            sucesso += 1
        except HttpError as e:
            print(f"    [ERRO] Aplicando nota para {email_key} (submission {sub_id}): {e}")

    print("  Resumo da tarefa:")
    print(f"    Notas aplicadas com sucesso : {sucesso}")
    print(f"    Ignorados (sem aluno)       : {ignorados_sem_aluno}")
    print(f"    Ignorados (sem submissão)   : {ignorados_sem_sub}")


# ------------------------------------------------------------
# EXTRair formId A PARTIR DO LINK
# ------------------------------------------------------------
FORM_URL_RE = re.compile(r"/forms/d/(?:e/)?([a-zA-Z0-9_-]+)/")


def extrair_form_id_de_link_material(material: Dict[str, Any]) -> Optional[str]:
    """
    Se o material tiver um link para Google Forms, extrai o formId da URL.
    """
    link = material.get("link")
    if not link:
        return None

    url = link.get("url", "")
    if "docs.google.com/forms" not in url:
        return None

    m = FORM_URL_RE.search(url)
    if not m:
        return None

    return m.group(1)


# ------------------------------------------------------------
# CLASSROOM – LISTAR TODOS OS COURSEWORK
# ------------------------------------------------------------
def listar_courseworks(classroom_service, course_id: str) -> List[Dict[str, Any]]:
    works: List[Dict[str, Any]] = []
    page_token = None

    while True:
        resp = classroom_service.courses().courseWork().list(
            courseId=course_id,
            pageToken=page_token,
            pageSize=100
        ).execute()

        works.extend(resp.get("courseWork", []))
        page_token = resp.get("nextPageToken")
        if not page_token:
            break

    print(f"\nCourseWorks encontrados na turma: {len(works)}")
    return works


# ------------------------------------------------------------
# CLASSROOM – BUSCAR MAX_POINTS DA TAREFA
# ------------------------------------------------------------
def obter_max_points_do_coursework(coursework: Dict[str, Any]) -> Optional[float]:
    mp = coursework.get("maxPoints")
    if mp is not None:
        print(f"  maxPoints da tarefa: {mp}")
    else:
        print("  Aviso: tarefa sem maxPoints definido.")
    return mp


# ------------------------------------------------------------
# MAIN
# ------------------------------------------------------------
def main():
    creds = get_credentials()
    print("Escopos do token carregado:")
    print(creds.scopes)

    classroom_service = build("classroom", "v1", credentials=creds)
    sheets_service = build("sheets", "v4", credentials=creds) 

    # Escolher turma
    curso = escolher_turma(classroom_service)
    course_id = curso["id"]
    print(f"\nTurma escolhida: {curso['name']} (ID {course_id})")

    # Mapa alunos (email -> userId) – faz 1x por turma
    mapa_email_para_user = mapear_alunos_por_email(classroom_service, course_id)
    if not mapa_email_para_user:
        print("Nenhum aluno mapeado. O mapa de emergência está vazio ou incorreto.")
        sys.exit(0)

    # Listar todos os courseWorks
    works = listar_courseworks(classroom_service, course_id)
    if not works:
        print("Nenhuma atividade encontrada para essa turma.")
        sys.exit(0)

    # NOVO: USAR ID PADRÃO
    sheet_id = SHEET_ID_PADRAO
    if not sheet_id:
        # Se o ID Padrão for vazio, solicita ao usuário
        sheet_id = input("\n[PASSO NOVO] Por favor, cole o ID da Planilha do Google (Google Sheet) vinculada ao Forms: ").strip()
    
    if not sheet_id:
         print("ID da Planilha não fornecido. Não é possível sincronizar notas.")
         sys.exit(0)
    else:
         print(f"Planilha de sincronização configurada: {sheet_id}")


    # Para cada courseWork, procurar Forms vinculados via LINK e sincronizar
    for cw in works:
        cw_id = cw["id"]
        titulo = cw.get("title", "(sem título)")
        work_type = cw.get("workType", "")
        print(f"\n=== TAREFA: {titulo} (ID {cw_id}, workType={work_type}) ===")

        # Assumindo que a única Planilha de respostas é a que foi configurada.
        print(f"  Tentando sincronizar notas usando a Planilha {sheet_id} com a tarefa {cw_id}...")

        # CHAMADA CORRIGIDA: Usando Sheets API
        try:
            notas_forms = mapear_notas_do_forms_via_sheets(sheets_service, sheet_id)
        except HttpError as e:
            print(f"  ERRO: Falha ao acessar Planilha {sheet_id}. Verifique se o ID está correto ou se a Planilha está compartilhada com a sua conta. Erro: {e}")
            continue

        if not notas_forms:
            print("  Nenhuma nota encontrada na Planilha (vazia ou colunas incorretas). Pulando.")
            continue

        mapa_user_para_sub = mapear_submissoes_por_user(
            classroom_service, course_id, cw_id
        )
        if not mapa_user_para_sub:
            print("  Nenhuma submissão para essa tarefa no Classroom. Pulando.")
            continue

        max_points = obter_max_points_do_coursework(cw)
        aplicar_notas(
            classroom_service,
            course_id,
            cw_id,
            mapa_email_para_user,
            mapa_user_para_sub,
            notas_forms,
            max_points,
        )

    print("\nProcesso de sincronização concluído para TODAS as tarefas com Forms (por link) da turma.")


if __name__ == "__main__":
    main()