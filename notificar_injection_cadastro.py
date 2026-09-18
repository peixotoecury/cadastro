# -*- coding: utf-8 -*-
"""
notificar_injection_cadastro.py — LAWgico Cadastro

Roda a cada poucos minutos (Tarefa Agendada do Windows, ex: a cada 15 min,
dias uteis, horario comercial): checa se algum cadastro novo veio com
injection_nivel != 'LIMPO' (detectado no index.html ao salvar, campo
Observacao/Link) desde a ultima checagem, e manda 1 e-mail de alerta pra
controladoria via automacao do Outlook (.Send() de verdade).

Guarda o timestamp da ultima checagem em ULTIMO_CHECK_PATH (arquivo local,
fora do git) pra nunca notificar o mesmo cadastro duas vezes.

Se nada suspeito foi cadastrado desde a ultima checagem, nao manda e-mail
nenhum (evita spam de "nada aconteceu").

MODO_TESTE=True (padrao): manda pra TESTE_EMAIL em vez do destinatario real,
com aviso no assunto -- usar assim ate a usuaria validar por alguns dias.
Depois, mudar pra False (mesmo fluxo do notificar_conclusao.py, Compromissos).
"""
import json
import re
import logging
import sys
from datetime import datetime, timezone
from pathlib import Path

import requests
import win32com.client

from pdf_injection_scan import verificar_link

sys.stdout.reconfigure(encoding='utf-8')

SUPABASE_URL = "https://sydamnqagkdmczmgkvso.supabase.co"
SUPABASE_KEY = "sb_publishable_kA_IDYtEdATSygg7EajrdQ_cEi4d57N"
HEADERS = {"apikey": SUPABASE_KEY, "Authorization": f"Bearer {SUPABASE_KEY}", "Content-Type": "application/json"}

CADASTRO_URL = "https://peixotoecury.github.io/cadastro/"
DESTINATARIO_REAL = "controladoria@peixotoecury.com.br"

# ── MODO TESTE — deixar True ate a usuaria validar, depois mudar pra False ──
MODO_TESTE = False
TESTE_EMAIL = "claude.controladoria@peixotoecury.com.br"

ULTIMO_CHECK_PATH = Path(__file__).parent / "_ultimo_check_injection.json"

LOG = logging.getLogger("notificar_injection_cadastro")
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

COR_NIVEL = {"CRÍTICO": "#C0392B", "ALTO": "#E67E22", "MÉDIO": "#D4AC0D", "BAIXO": "#5B6B7A"}


def ler_ultimo_check():
    if not ULTIMO_CHECK_PATH.exists():
        # Primeira execucao: nao varre o historico inteiro, comeca a contar a
        # partir de agora (senao manda 1 e-mail gigante com tudo que ja foi
        # cadastrado desde sempre).
        return datetime.now(timezone.utc).isoformat()
    return json.loads(ULTIMO_CHECK_PATH.read_text(encoding="utf-8"))["ultimo_check"]


def salvar_ultimo_check(timestamp_iso):
    ULTIMO_CHECK_PATH.write_text(json.dumps({"ultimo_check": timestamp_iso}), encoding="utf-8")


def buscar_suspeitos_desde(ultimo_check):
    r = requests.get(
        f"{SUPABASE_URL}/rest/v1/cadastros_deloitte",
        headers=HEADERS, timeout=30,
        params={
            "created_at": f"gt.{ultimo_check}",
            "select": "*",
            "order": "created_at.asc",
        },
    )
    r.raise_for_status()
    return r.json()


ORDEM = ["LIMPO", "BAIXO", "MÉDIO", "ALTO", "CRÍTICO"]


def verificar_pdfs(itens):
    """Baixa e analisa o PDF de cada cadastro ainda nao verificado; grava o resultado no Supabase."""
    for c in itens:
        cats = c.get("injection_categorias") or ""
        if "[PDF verificado" in cats:
            c["_pdf"] = {"verificado": True}
            continue
        res = verificar_link(c.get("link_caso"))
        c["_pdf"] = res
        if not res.get("verificado"):
            continue
        nivel_txt = c.get("injection_nivel") or "LIMPO"
        nivel = max(nivel_txt, res["nivel"], key=lambda n: ORDEM.index(n) if n in ORDEM else 0)
        score = max(c.get("injection_score") or 0, res["score"])
        partes = [x for x in cats.split(" | ") if x] + res["categorias"]
        sufixo = "[PDF verificado]" if res["achados"] else "[PDF verificado, sem achado]"
        try:
            requests.patch(f"{SUPABASE_URL}/rest/v1/cadastros_deloitte", headers=HEADERS, timeout=30,
                           params={"id": f"eq.{c['id']}"},
                           json={"injection_nivel": nivel, "injection_score": score,
                                 "injection_categorias": (" | ".join(dict.fromkeys(partes)) + " " + sufixo).strip()})
        except Exception as e:
            LOG.warning(f"Nao gravou resultado do PDF de {c.get('id')}: {e}")
        c["injection_nivel"], c["injection_score"] = nivel, score


def montar_corpo(itens):
    def status(c):
        nivel = c.get("injection_nivel") or "SEM VERIFICAÇÃO"
        cor = COR_NIVEL.get(nivel, "#1E7A3E" if nivel == "LIMPO" else "#5B6B7A")
        bruto = c.get("injection_categorias") or ""
        anexo = re.search(r"\[PDF anexado[^\]]*\]", bruto)
        cats = re.sub(r"\[PDF[^\]]*\]", "", bruto).strip(" |")
        r = c.get("_pdf") or {}
        if r.get("verificado"):
            pdf = f"PDF lido ({r.get('paginas', '?')} pág.)" + (" — sem texto (escaneado), só metadados" if r.get("sem_texto") else "")
            for a in (r.get("achados") or [])[:3]:
                pdf += f"<br>▸ {a['nome']} [{a['origem']}]: {a['trecho'][:100]}"
        else:
            motivo = r.get('motivo', 'sem detalhe')
            if anexo and "sem texto" not in anexo.group(0):
                pdf = f"Link do caso não acessível pelo servidor ({motivo}) — verificação feita pelo anexo"
            else:
                pdf = f"⚠ PDF NÃO verificado: {motivo}"
        if anexo:
            pdf = "Anexo enviado no cadastro: " + anexo.group(0).strip("[]") + "<br>" + pdf
        return nivel, cor, cats, pdf

    def linha(c):
        nivel, cor, cats, pdf = status(c)
        obs = (c.get("observacao") or "")[:200]
        return (f"<tr><td style='padding:4px 8px'>{c.get('numero_processo') or '—'}</td>"
                f"<td style='padding:4px 8px'>{c.get('nome_reclamante') or '—'}</td>"
                f"<td style='padding:4px 8px'><b style='color:{cor}'>{nivel}</b> (score {c.get('injection_score')})</td>"
                f"<td style='padding:4px 8px;color:#5B6B7A;font-size:12px'>{cats or '—'}<br>{pdf}</td>"
                f"<td style='padding:4px 8px;color:#5B6B7A;font-size:12px'>{obs}</td></tr>")

    suspeitos = [c for c in itens if (c.get("injection_nivel") or "LIMPO") != "LIMPO"]
    corpo = f"Verificação de prompt injection dos <b>{len(itens)}</b> cadastro(s) novo(s) desde a última checagem"
    corpo += (f" — <b style='color:#C0392B'>{len(suspeitos)} com alerta</b>:<br><br>" if suspeitos
              else " — <b style='color:#1E7A3E'>nenhum alerta</b>:<br><br>")
    corpo += ("<table border='1' cellspacing='0' style='border-collapse:collapse;border-color:#DDE6EC;font-size:13px'>"
              "<tr style='background:#003B5C;color:#fff'><th>Processo</th><th>Reclamante</th><th>Resultado</th>"
              "<th>Categorias / PDF</th><th>Observação</th></tr>" + "".join(linha(c) for c in itens) + "</table>")
    corpo += (f"<br><a href='{CADASTRO_URL}'>Ver painel completo</a><br><br>"
              f"Atenciosamente,<br>Controladoria — Peixoto e Cury Advogados")
    return corpo, len(suspeitos)


def enviar_email(destinatario, assunto, corpo_html):
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)  # olMailItem
    mail.To = destinatario
    mail.Subject = assunto
    mail.HTMLBody = corpo_html
    mail.Send()


def main():
    ultimo_check = ler_ultimo_check()
    agora = datetime.now(timezone.utc).isoformat()

    suspeitos = buscar_suspeitos_desde(ultimo_check)
    LOG.info(f"Suspeitos desde {ultimo_check}: {len(suspeitos)}")

    if not suspeitos:
        salvar_ultimo_check(agora)
        LOG.info("Nada novo. Concluído.")
        return

    verificar_pdfs(suspeitos)
    corpo, n_alerta = montar_corpo(suspeitos)
    destinatario_real = DESTINATARIO_REAL
    assunto = (f"🔴 {n_alerta} alerta(s) de prompt injection em {len(suspeitos)} cadastro(s) — Cadastro Deloitte" if n_alerta
               else f"🟢 {len(suspeitos)} cadastro(s) verificado(s), sem prompt injection — Cadastro Deloitte")
    if MODO_TESTE:
        destinatario_real = TESTE_EMAIL
        assunto = f"[TESTE — seria p/ {DESTINATARIO_REAL}] {assunto}"

    LOG.info(f"Enviando pra {destinatario_real} "
             f"({'MODO TESTE, real=' + DESTINATARIO_REAL if MODO_TESTE else 'real'}) — {len(suspeitos)} item(ns)")
    enviar_email(destinatario_real, assunto, corpo)

    salvar_ultimo_check(agora)
    LOG.info("Concluído.")


if __name__ == "__main__":
    main()
