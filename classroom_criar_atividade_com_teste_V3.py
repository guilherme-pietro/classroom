from __future__ import annotations
# -*- coding: utf-8 -*-

import os
import sys
import re
from typing import List, Dict, Any, Optional

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# ------------------------------------------------------------
# CONFIG GOOGLE
# ------------------------------------------------------------
SCOPES = [
    "https://www.googleapis.com/auth/classroom.courses.readonly",
    "https://www.googleapis.com/auth/classroom.coursework.students",
    "https://www.googleapis.com/auth/classroom.topics.readonly",  # <-- NOVO
    "https://www.googleapis.com/auth/forms.body",
    "https://www.googleapis.com/auth/drive.file",
]


CREDENTIALS_FILE = "credentials.json"
TOKEN_FILE = "token.json"


def get_credentials():
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)

    if not creds or not creds.valid:
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
def escolher_turma(classroom_service):
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
# CLASSROOM – ESCOLHER TEMA (TÓPICO) DA TURMA
# ------------------------------------------------------------
def escolher_tema(classroom_service, curso) -> Optional[str]:
    """
    Lista os temas (tópicos) da turma e permite escolher um.
    Retorna o topicId escolhido ou None para 'sem tema'.
    """
    topics_resp = classroom_service.courses().topics().list(
        courseId=curso["id"]
    ).execute()

    temas = topics_resp.get("topic", [])

    if not temas:
        print("\nNenhum tema cadastrado para esta turma. A atividade será criada SEM tema.")
        return None

    print("\n=== TEMAS DISPONÍVEIS PARA A TURMA ===")
    print(" 0 - (Sem tema)")
    for i, t in enumerate(temas, start=1):
        print(f"{i:2d} - {t['name']} (ID {t['topicId']})")

    while True:
        op = input("\nEscolha o número do tema (0 para sem tema): ").strip()
        try:
            x = int(op)
            if x == 0:
                return None
            if 1 <= x <= len(temas):
                return temas[x - 1]["topicId"]
        except ValueError:
            pass
        print("Opção inválida.")


# ------------------------------------------------------------
# PARSER – FORMATO:
#
# P1: Pergunta ...
# A) alternativa
# B) alternativa
# C) alternativa
# D) alternativa
# E) alternativa
# G: B
#
# P2: ...
# ...
# ------------------------------------------------------------
def parse_bloco(bloco: str) -> List[Dict[str, Any]]:
    """
    Formato esperado:

    P1: Por que a segurança no trabalho é essencial no ambiente industrial?
    A) ...
    B) ...
    C) ...
    D) ...
    E) ...
    G: B

    P2: ...
    A) ...
    ...
    G: C
    """
    texto = bloco.strip("\n")
    if not texto:
        print("Nenhum texto informado.")
        sys.exit(1)

    linhas = texto.splitlines()

    blocos: List[List[str]] = []
    bloco_atual: List[str] = []

    for linha in linhas:
        s = linha.strip()
        if not s:
            if bloco_atual:
                bloco_atual.append("")
            continue

        if re.match(r"^P\d*:", s, re.IGNORECASE):
            if bloco_atual:
                blocos.append(bloco_atual)

            depois = s.split(":", 1)[1].strip() if ":" in s else ""
            if depois:
                bloco_atual = [depois]
            else:
                bloco_atual = []
        else:
            bloco_atual.append(linha)

    if bloco_atual:
        blocos.append(bloco_atual)

    if not blocos:
        print("ERRO: Nenhum bloco de pergunta encontrado (P1:, P2:, ...).")
        sys.exit(1)

    perguntas: List[Dict[str, Any]] = []

    for bloco_p in blocos:
        linhas_p = [l.strip() for l in bloco_p if l.strip()]
        if not linhas_p:
            continue

        enunciado = linhas_p[0]

        alternativas_textos: List[str] = []
        alternativas_letras: List[str] = []
        letra_correta: str | None = None

        for l in linhas_p[1:]:
            m_alt = re.match(r"([A-Z])\)\s*(.+)", l, re.IGNORECASE)
            if m_alt:
                letra = m_alt.group(1).lower()
                texto_alt = m_alt.group(2).strip()
                alternativas_letras.append(letra)
                alternativas_textos.append(texto_alt)
                continue

            m_g = re.match(r"G:\s*([A-Z])", l, re.IGNORECASE)
            if m_g:
                letra_correta = m_g.group(1).lower()
                continue

        if len(alternativas_textos) < 2:
            print(f"ERRO: pergunta sem alternativas suficientes -> {enunciado}")
            sys.exit(1)

        if not letra_correta:
            print(f"ERRO: pergunta sem gabarito (linha 'G: X') -> {enunciado}")
            sys.exit(1)

        if letra_correta not in alternativas_letras:
            print(
                f"ERRO: gabarito '{letra_correta.upper()}' não corresponde a "
                f"nenhuma alternativa em -> {enunciado}"
            )
            sys.exit(1)

        idx_correta = alternativas_letras.index(letra_correta)

        perguntas.append(
            {
                "title": enunciado,
                "alternativas": alternativas_textos,
                "correta_idx": idx_correta,
            }
        )

    if not perguntas:
        print("ERRO: não consegui montar nenhuma pergunta.")
        sys.exit(1)

    return perguntas


# ------------------------------------------------------------
# FORMS – CRIAR QUIZ (coletando e-mail para vincular ao aluno)
# ------------------------------------------------------------
def criar_quiz_forms(forms_service, titulo, questoes):
    # cria o formulário
    form = forms_service.forms().create(
        body={"info": {"title": titulo}}
    ).execute()

    form_id = form["formId"]

    # transformar em quiz e coletar e-mail do respondente
    requests = [
        {
            "updateSettings": {
                "settings": {
                    "quizSettings": {"isQuiz": True},
                    "emailCollectionType": "RESPONDER_INPUT",
                },
                "updateMask": "quizSettings.isQuiz,emailCollectionType",
            }
        }
    ]

    # adicionar questões
    for i, q in enumerate(questoes):
        alternativas = q["alternativas"]
        correta = alternativas[q["correta_idx"]]

        item = {
            "title": q["title"],
            "questionItem": {
                "question": {
                    "required": True,
                    "choiceQuestion": {
                        "type": "RADIO",
                        "options": [{"value": a} for a in alternativas],
                        "shuffle": False,
                    },
                    "grading": {
                        "pointValue": 1,
                        "correctAnswers": {"answers": [{"value": correta}]},
                    },
                }
            },
        }

        requests.append({
            "createItem": {
                "item": item,
                "location": {"index": i}
            }
        })

    forms_service.forms().batchUpdate(
        formId=form_id, body={"requests": requests}
    ).execute()

    form_final = forms_service.forms().get(formId=form_id).execute()
    responder_uri = form_final.get("responderUri")

    print("\nForm criado com sucesso.")
    print("Link para responder:", responder_uri)

    return form_id, responder_uri


# ------------------------------------------------------------
# CLASSROOM – CRIAR ATIVIDADE COM LINK DO FORMS
# ------------------------------------------------------------
def criar_atividade_classroom(classroom_service, curso, titulo, link, topic_id: Optional[str]):
    body = {
        "title": titulo,
        "description": "Avaliação criada automaticamente.",
        "materials": [{"link": {"url": link, "title": titulo}}],
        "workType": "ASSIGNMENT",
        "state": "PUBLISHED",
        "maxPoints": 10,
    }

    if topic_id:
        body["topicId"] = topic_id

    trabalho = (
        classroom_service.courses()
        .courseWork()
        .create(courseId=curso["id"], body=body)
        .execute()
    )
    print(f"Tarefa criada no Classroom (ID: {trabalho['id']}).")


# ------------------------------------------------------------
# MAIN
# ------------------------------------------------------------
def main():
    creds = get_credentials()
    print("Escopos do token carregado:")
    print(creds.scopes)

    classroom_service = build("classroom", "v1", credentials=creds)
    forms_service = build("forms", "v1", credentials=creds)

    turma = escolher_turma(classroom_service)
    topic_id = escolher_tema(classroom_service, turma)

    titulo = input("\nTítulo da avaliação: ").strip() or "Avaliação"

    print("\nCole TODAS as perguntas, alternativas e o gabarito no formato:")
    print("P1: Pergunta")
    print("A) ...")
    print("B) ...")
    print("C) ...")
    print("D) ...")
    print("E) ...")
    print("G: B   (gabarito da questão)")
    print("P2: ...")
    print("...")
    print("\nQuando terminar de colar, digite UMA LINHA SÓ com: fim")
    print("e pressione ENTER.\n")

    linhas_bloco = []
    while True:
        try:
            linha = input()
        except EOFError:
            break
        if linha.strip().lower() == "fim":
            break
        linhas_bloco.append(linha)

    bloco = "\n".join(linhas_bloco)

    questoes = parse_bloco(bloco)

    form_id, link = criar_quiz_forms(forms_service, titulo, questoes)
    criar_atividade_classroom(classroom_service, turma, titulo, link, topic_id)

    print("\nProcesso concluído.")
    print("Link da avaliação:", link)


if __name__ == "__main__":
    main()
