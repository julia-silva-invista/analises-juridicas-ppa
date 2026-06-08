# -*- coding: utf-8 -*-
import re
import time
import requests
from bs4 import BeautifulSoup
from datetime import datetime

ESAJ_TRIBUNAIS = {
    "TJSP (ESAJ)": "https://esaj.tjsp.jus.br",
    "TJBA (ESAJ)": "https://esaj.tjba.jus.br",
    "TJCE (ESAJ)": "https://esaj.tjce.jus.br",
    "TJMS (ESAJ)": "https://esaj.tjms.jus.br",
    "TJSC (ESAJ)": "https://esaj.tjsc.jus.br",
    "TJAL (ESAJ)": "https://www2.tjal.jus.br",
    "TJAM (ESAJ)": "https://consultasaj.tjam.jus.br",
}

_UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"


def _limpar(txt):
    return re.sub(r"\s+", " ", (txt or "").strip())


def _extrair_conversation_id(html):
    m = re.search(r'name=["\']conversationId["\']\s+value=["\']([^"\']*)', html or "", re.I)
    if m:
        return m.group(1)
    m = re.search(r'conversationId=([A-Za-z0-9._-]+)', html or "")
    return m.group(1) if m else ""


def _formatar_numero_cnj(numero):
    n = re.sub(r"\D", "", numero or "")
    if len(n) == 20:
        return f"{n[:7]}-{n[7:9]}.{n[9:13]}.{n[13]}.{n[14:16]}.{n[16:]}"
    return numero.strip()


def _nova_sessao(base_url):
    s = requests.Session()
    s.headers.update({
        "User-Agent": _UA,
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "pt-BR,pt;q=0.9,en-US;q=0.8,en;q=0.7",
        "Connection": "keep-alive",
        "Upgrade-Insecure-Requests": "1",
    })
    resp = s.get(f"{base_url}/cpopg/open.do", timeout=30, allow_redirects=True)
    return s, _extrair_conversation_id(resp.text)


def _montar_params_doc(doc, conversation_id):
    return {
        "cbPesquisa": "DOCPARTE",
        "dadosConsulta.valorConsulta": doc,
        "cdForo": "-1",
        "conversationId": conversation_id or "",
        "dadosConsulta.localPesquisa.cdLocal": "-1",
        "paginaConsulta": "1",
    }


def _tem_captcha(html):
    texto = (html or "").lower()
    sinais = [
        "captcha",
        "g-recaptcha",
        "hcaptcha",
        "digite o texto",
        "imagem de controle",
    ]
    return any(s in texto for s in sinais)


def _executar_busca(session, base_url, params):
    headers = {
        "Referer": f"{base_url}/cpopg/open.do",
        "Origin": base_url,
    }
    resp = session.get(
        f"{base_url}/cpopg/search.do",
        params=params,
        headers=headers,
        timeout=30,
        allow_redirects=True,
    )
    if resp.status_code == 200 and not _tem_captcha(resp.text):
        return resp

    # Alguns SAJ aceitam melhor o envio como formulário, mantendo os mesmos cookies.
    post_resp = session.post(
        f"{base_url}/cpopg/search.do",
        data=params,
        headers=headers,
        timeout=30,
        allow_redirects=True,
    )
    return post_resp


def _parsear_lista(soup, base_url, session, cpf_cnpj, diag):
    """Extrai processos da página de lista de resultados."""
    processos = []

    # ESAJ lista resultados em <a> com classe linkProcesso ou dentro de tabela
    links = soup.select("a.linkProcesso") or soup.select("a[href*='show.do']")

    if links:
        diag.append(f"      [esaj/lista] {len(links)} link(s) encontrado(s)")
        for link in links[:50]:  # limite de segurança
            href = link.get("href", "")
            if not href.startswith("http"):
                href = base_url + href
            try:
                resp = session.get(href, timeout=30)
                if resp.status_code == 200:
                    proc = _parsear_detalhe(BeautifulSoup(resp.text, "lxml"), cpf_cnpj)
                    if proc:
                        processos.append(proc)
                time.sleep(0.3)
            except Exception:
                continue
        return processos

    # alternativa: tabela de resultados direta
    tabela = soup.find("table", id=re.compile(r"tabelaProcessos|resultado", re.I))
    if not tabela:
        tabela = soup.find("table", class_=re.compile(r"resultado|processo", re.I))
    if tabela:
        for tr in tabela.find_all("tr")[1:]:
            tds = tr.find_all("td")
            if not tds:
                continue
            numero_raw = _limpar(tds[0].get_text()) if tds else ""
            classe = _limpar(tds[1].get_text()) if len(tds) > 1 else ""
            data = _limpar(tds[2].get_text()) if len(tds) > 2 else ""
            polo_ativo = _limpar(tds[3].get_text()) if len(tds) > 3 else ""
            polo_passivo = _limpar(tds[4].get_text()) if len(tds) > 4 else ""
            if numero_raw:
                processos.append([
                    _formatar_numero_cnj(numero_raw),
                    cpf_cnpj, data, classe,
                    None, None,
                    polo_ativo, polo_passivo, "",
                ])
        diag.append(f"      [esaj/tabela] {len(processos)} linha(s)")

    return processos


def _parsear_detalhe(soup, cpf_cnpj):
    """Extrai dados de uma página de detalhe de processo individual."""
    numero_span = soup.find("span", id="numeroProcesso") or \
                  soup.find(id=re.compile(r"numeroProcesso|numProcesso", re.I))
    if not numero_span:
        return None

    numero = _formatar_numero_cnj(_limpar(numero_span.get_text()))

    # classe
    classe = ""
    cls_el = soup.find(id=re.compile(r"classeProcesso", re.I)) or \
             soup.find(class_=re.compile(r"classeProcesso", re.I))
    if cls_el:
        classe = _limpar(cls_el.get_text())

    # data de distribuição
    data = ""
    for label in soup.find_all(string=re.compile(r"distribui", re.I)):
        parent = label.parent
        if parent:
            txt = _limpar(parent.get_text())
            m = re.search(r"\d{2}/\d{2}/\d{4}", txt)
            if m:
                data = m.group()
                break

    # partes — polo ativo / passivo
    polo_ativo, polo_passivo = [], []
    for div in soup.find_all(class_=re.compile(r"nomeParteEsquerda|nomeParte", re.I)):
        txt = _limpar(div.get_text())
        if not txt:
            continue
        header = div.find_previous(string=re.compile(r"autor|requerente|reclamante|exequente|impetrante", re.I))
        if header:
            polo_ativo.append(txt)
        else:
            polo_passivo.append(txt)

    # última movimentação
    situacao = ""
    movs = soup.select(".descricaoMovimentacao")
    if movs:
        situacao = _limpar(movs[0].get_text())

    return [
        numero, cpf_cnpj, data, classe,
        None, None,
        " | ".join(polo_ativo), " | ".join(polo_passivo), situacao,
    ]


def buscar_esaj(tribunal_nome, base_url, cpf_cnpj):
    """Scrapa o portal ESAJ buscando por CPF/CNPJ. Retorna (linhas, diag)."""
    diag = []
    doc = re.sub(r"\D", "", cpf_cnpj or "")
    if not doc:
        return [], diag

    try:
        session, conversation_id = _nova_sessao(base_url)
        params = _montar_params_doc(doc, conversation_id)
        resp = _executar_busca(session, base_url, params)
        diag.append(f"      [esaj] HTTP {resp.status_code}")

        if resp.status_code != 200:
            return [], diag

        if _tem_captcha(resp.text):
            diag.append("      [esaj] CAPTCHA detectado — renovando sessão e tentando novamente")
            time.sleep(2)
            session, conversation_id = _nova_sessao(base_url)
            params = _montar_params_doc(doc, conversation_id)
            resp = _executar_busca(session, base_url, params)
            diag.append(f"      [esaj/retry] HTTP {resp.status_code}")
            if resp.status_code != 200:
                return [], diag
            if _tem_captcha(resp.text):
                diag.append("      [esaj] CAPTCHA persistente — o tribunal bloqueou a consulta automatizada neste momento")
                return [], diag

        soup = BeautifulSoup(resp.text, "lxml")

        # página de detalhe direto (1 único resultado)
        if soup.find("span", id="numeroProcesso") or \
           soup.find(id=re.compile(r"numeroProcesso", re.I)):
            proc = _parsear_detalhe(soup, doc)
            processos = [proc] if proc else []
            diag.append(f"      [esaj/detalhe] {len(processos)} processo(s)")
            return processos, diag

        # página de lista
        processos = _parsear_lista(soup, base_url, session, doc, diag)
        return processos, diag

    except Exception as e:
        diag.append(f"      [esaj] ERRO: {e}")
        return [], diag
