# -*- coding: utf-8 -*-
"""
pdf_injection_scan.py — baixa o PDF do link_caso (SharePoint/OneDrive, link
de compartilhamento + download=1) e procura prompt injection: texto visivel,
texto oculto (fonte < 3pt, cor branca, fora da pagina), metadados e as mesmas
regras (INJ_RULES) do index.html do Cadastro, portadas do JS.
Roda no servidor/PC da controladoria (sem CORS), chamado por notificar_injection_cadastro.py.
"""
import re
from pathlib import Path

import fitz  # pymupdf
import requests

HTML = Path(__file__).parent / "index.html"
PESO = {"CRÍTICO": 40, "ALTO": 20, "MÉDIO": 8, "BAIXO": 2}
MAX_BYTES = 80 * 1024 * 1024
MAX_PAGINAS = 60


def _carregar_regras():
    src = HTML.read_text(encoding="utf-8")
    ini = src.index("const INJ_RULES = [")
    fim = src.index("function injAplicaRegras")
    bloco = src[ini:fim]
    regras = []
    for m in re.finditer(r"\{\s*id:'([^']+)',\s*sev:'([^']+)',\s*name:'([^']+)',\s*patterns:\[(.*?)\](?=\s*(?:,\s*validator|\}))", bloco, re.S):
        rid, sev, nome, pats = m.groups()
        # regras com validator dependem de codigo JS: ignoradas (evita falso positivo)
        resto = bloco[m.end():m.end() + 60]
        if resto.lstrip().startswith(", validator") or resto.lstrip().startswith(",validator"):
            continue
        compiladas = []
        for lit in re.finditer(r"/((?:\\.|[^/\\\n])+)/([gimsuy]*)", pats):
            corpo, flags = lit.groups()
            f = re.I if "i" in flags else 0
            try:
                compiladas.append(re.compile(corpo, f | re.S if "s" in flags else f))
            except re.error:
                pass
        if compiladas:
            regras.append((rid, sev, nome, compiladas))
    return regras


REGRAS = _carregar_regras()
UNICODE_INVISIVEL = re.compile(r"[​-‏‪-‮⁠-⁤﻿]{3,}")


def baixar_pdf(link, timeout=120):
    """Retorna bytes do PDF ou (None, motivo)."""
    if not link or not link.lower().startswith("http"):
        return None, "link nao e http (arquivo local ou vazio)"
    url = link + ("&" if "?" in link else "?") + "download=1"
    try:
        r = requests.get(url, timeout=timeout, stream=True, allow_redirects=True)
        if r.status_code != 200:
            return None, f"HTTP {r.status_code}"
        buf = b""
        for chunk in r.iter_content(1024 * 1024):
            buf += chunk
            if len(buf) > MAX_BYTES:
                return None, "PDF acima de 80MB"
        if not buf.startswith(b"%PDF-"):
            return None, "link nao devolveu PDF (pede login ou nao e PDF)"
        return buf, None
    except Exception as e:
        return None, f"erro no download: {e}"


def _aplicar_regras(texto, origem):
    achados = []
    for rid, sev, nome, pats in REGRAS:
        for p in pats:
            m = p.search(texto)
            if m:
                achados.append({"nome": nome, "sev": sev, "origem": origem,
                                "trecho": texto[max(0, m.start() - 40):m.end() + 40].replace("\n", " ")})
                break
    if UNICODE_INVISIVEL.search(texto):
        achados.append({"nome": "Caracteres Unicode invisíveis em sequência", "sev": "ALTO", "origem": origem, "trecho": ""})
    return achados


def analisar_pdf(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    achados = []
    visivel, oculto = [], []
    for i, page in enumerate(doc):
        if i >= MAX_PAGINAS:
            break
        rect = page.rect
        d = page.get_text("dict")
        for b in d.get("blocks", []):
            for l in b.get("lines", []):
                for s in l.get("spans", []):
                    t = (s.get("text") or "").strip()
                    if len(t) < 3:
                        continue
                    x0, y0, x1, y1 = s["bbox"]
                    escondido = (s["size"] < 3 or (s["color"] == 0xFFFFFF) or
                                 x1 < rect.x0 or x0 > rect.x1 or y1 < rect.y0 or y0 > rect.y1)
                    (oculto if escondido else visivel).append(t)
    tv, to = " ".join(visivel), " ".join(oculto)
    achados += _aplicar_regras(tv, "texto visível")
    achados += _aplicar_regras(to, "texto OCULTO")
    # texto oculto em volume relevante ja e suspeito por si so (ignora resquicio pequeno)
    if len(to) > 80 and not any(a["origem"] == "texto OCULTO" for a in achados):
        achados.append({"nome": "Texto oculto (fonte minúscula/branco/fora da página)", "sev": "MÉDIO",
                        "origem": "texto OCULTO", "trecho": to[:160]})
    meta = " ".join(str(v) for v in (doc.metadata or {}).values() if v)
    achados += _aplicar_regras(meta, "metadados")
    n_pag = min(len(doc), MAX_PAGINAS)
    tem_texto = len(tv) > 200
    doc.close()

    score = min(sum(PESO.get(a["sev"], 0) for a in achados), 100)
    nivel = "CRÍTICO" if score >= 40 else "ALTO" if score >= 20 else "MÉDIO" if score >= 8 else "BAIXO" if score > 0 else "LIMPO"
    return {"score": score, "nivel": nivel, "achados": achados, "paginas": n_pag,
            "sem_texto": not tem_texto,  # PDF escaneado: so imagem, texto nao verificavel
            "categorias": sorted({a["nome"] for a in achados})}


def verificar_link(link):
    """Retorna dict com 'verificado' (bool), 'motivo' e o resultado da analise."""
    pdf, motivo = baixar_pdf(link)
    if pdf is None:
        return {"verificado": False, "motivo": motivo}
    try:
        res = analisar_pdf(pdf)
    except Exception as e:
        return {"verificado": False, "motivo": f"erro ao ler o PDF: {e}"}
    res["verificado"] = True
    return res


if __name__ == "__main__":
    import sys, json
    print(f"{len(REGRAS)} regras carregadas")
    print(json.dumps(verificar_link(sys.argv[1]), ensure_ascii=False, indent=1)[:1500])
