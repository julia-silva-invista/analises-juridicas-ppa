# -*- coding: utf-8 -*-
"""Bateria de testes de fidelidade do Checklist RJ (sem chamar o Gemini de verdade).

Roda com:
    python tests/test_checklist_rj.py
ou:
    pytest tests/test_checklist_rj.py -v
"""
from __future__ import annotations

import copy
import os
import re
import sys
import types as pytypes

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

# ── Stubs de dependências externas (não precisamos de rede/API key) ────────
try:
    import gradio  # noqa: F401
except ImportError:
    sys.modules["gradio"] = pytypes.ModuleType("gradio")

try:
    from google import genai  # noqa: F401
    from google.genai import types as genai_types  # noqa: F401
except ImportError:
    google_mod = pytypes.ModuleType("google")
    genai_mod = pytypes.ModuleType("google.genai")
    genai_types_mod = pytypes.ModuleType("google.genai.types")
    genai_types_mod.GenerateContentConfig = lambda **k: k
    genai_types_mod.Content = lambda **k: k
    genai_types_mod.Part = lambda **k: k
    genai_mod.types = genai_types_mod
    genai_mod.Client = object
    google_mod.genai = genai_mod
    sys.modules["google"] = google_mod
    sys.modules["google.genai"] = genai_mod
    sys.modules["google.genai.types"] = genai_types_mod

if "utils" not in sys.modules:
    utils_mod = pytypes.ModuleType("utils")
    utils_mod._retry = lambda fn, tentativas=3, espera_base=10: fn()
    sys.modules["utils"] = utils_mod

import checklist_rj as cr  # noqa: E402
from docx import Document  # noqa: E402


# ══════════════════════════════════════════════════════════════════════════
# Infraestrutura de teste
# ══════════════════════════════════════════════════════════════════════════

class FakeClient:
    """Simula client.models.generate_content(...).text sem chamar a API."""

    def __init__(self, resposta: str):
        self._resposta = resposta
        self.chamadas = 0
        self.models = self

    def generate_content(self, **kwargs):
        self.chamadas += 1
        resp = pytypes.SimpleNamespace()
        resp.text = self._resposta
        return resp


# Padrões que NUNCA deveriam aparecer no documento final — indicam vazamento
# de instrução de formato (do prompt ou do molde oficial) para o texto visível.
_PADROES_VAZAMENTO = [
    r"\(fls\.\)",
    r"\(referência\)",
    r"\(referencia\)",
    r"\(Mov\.\s*/\s*ID\s*/\s*fls\.\s*/\s*Evento\)",
    r"\[caso haja",
    r"indicar páginas",
    r"indicar paginas",
    r"\[Nome",
    r"\(referência não localizada\)\s*$",  # aceitável isolado, mas nunca em excesso — ver nota abaixo
]
# "(referência não localizada)" é uma saída LEGÍTIMA do prompt (não um vazamento
# de placeholder) — removido da lista de proibidos; mantido comentado para
# referência. Lista efetiva de vazamento:
_PADROES_VAZAMENTO = [
    r"\(fls\.\)",
    r"\(referência\)",
    r"\(referencia\)",
    r"\(Mov\.\s*/\s*ID\s*/\s*fls\.\s*/\s*Evento\)",
    r"\[caso haja",
    r"indicar páginas",
    r"indicar paginas",
    r"\[Nome",
]
_REGEX_VAZAMENTO = re.compile("|".join(_PADROES_VAZAMENTO), re.IGNORECASE)
_REGEX_REFERENCIA_REAL = re.compile(r"\((?:Mov\.|fls\.|Evento|ID)\s+\S+", re.IGNORECASE)


def _iter_celulas(doc: Document):
    for ti, table in enumerate(doc.tables):
        for ri, row in enumerate(table.rows):
            for ci, cell in enumerate(row.cells):
                yield ti, ri, ci, cell.text


def achar_vazamentos(docx_path: str):
    doc = Document(docx_path)
    achados = []
    for ti, ri, ci, texto in _iter_celulas(doc):
        if texto and _REGEX_VAZAMENTO.search(texto):
            achados.append((ti, ri, ci, texto.strip()[:120]))
    return achados


def contar_referencias(docx_path: str) -> int:
    doc = Document(docx_path)
    total = 0
    for _, _, _, texto in _iter_celulas(doc):
        if texto and _REGEX_REFERENCIA_REAL.search(texto):
            total += 1
    return total


# ══════════════════════════════════════════════════════════════════════════
# Fixtures sintéticas de `dados` (simulando extrações boas e ruins)
# ══════════════════════════════════════════════════════════════════════════

def _base_limpa(rotulo_ref: str) -> dict:
    ref = lambda n: f"({rotulo_ref} {n})"  # noqa: E731
    return {
        "rj_numero": "1234567-89.2024.8.11.0000",
        "vara": f"1ª Vara Empresarial {ref(1)}",
        "data_analise": "31/07/2026",
        "requerentes": f"Empresa Teste Ltda · 12.345.678/0001-90 {ref(10)}",
        "advogados_requerentes": f"Dr. Fulano de Tal {ref(10)}",
        "administrador_judicial": f"AJ Consultoria {ref(15)}",
        "data_pedido": f"01/01/2024 {ref(1)}",
        "data_deferimento": f"15/01/2024 {ref(20)}",
        "consolidacao_substancial": f"Deferido {ref(340)}",
        "periodo_blindagem": f"Ativo {ref(350)}",
        "previsao_encerramento_stay": f"15/01/2025 {ref(350)}",
        "stay_prorrogavel": f"Sim {ref(400)}",
        "recursos_relevantes": [
            {"recurso": f"Agravo de Instrumento nº 111 {ref(50)}", "status": f"Pendente {ref(50)}"},
        ],
        "imoveis_requerentes": [
            {"matricula": f"12.454 {ref(88)}", "cartorio": f"1º RI de Cuiabá {ref(88)}",
             "descricao": f"Fazenda Boa Vista {ref(88)}", "proprietario": f"Empresa Teste {ref(88)}"},
        ],
        "imoveis_essenciais": [
            {"matricula": f"99.001 {ref(90)}", "cartorio": f"2º RI {ref(90)}",
             "descricao": f"Sede da empresa {ref(90)}", "proprietario": f"Empresa Teste {ref(90)}"},
        ],
        "prj_classe_ii": {"desagio": f"30% {ref(200)}", "carencia": f"12 meses {ref(200)}",
                          "parcelas": f"60 {ref(200)}", "juros": f"2% a.a. {ref(200)}",
                          "correcao": f"IPCA {ref(200)}"},
        "prj_classe_iii": {"desagio": f"40% {ref(210)}", "carencia": f"24 meses {ref(210)}",
                           "parcelas": f"72 {ref(210)}", "juros": f"1% a.a. {ref(210)}",
                           "correcao": f"INPC {ref(210)}"},
        "qgc": {"classe_i": f"R$ 500.000,00 {ref(300)}", "classe_ii": f"R$ 2.000.000,00 {ref(300)}",
                "classe_iii": f"R$ 8.000.000,00 {ref(300)}", "classe_iv": f"R$ 100.000,00 {ref(300)}",
                "total": f"R$ 10.600.000,00 {ref(300)}"},
        "agc_situacao": f"Convocada {ref(320)}",
        "agc_1a": f"10/03/2025 {ref(320)}", "agc_2a": f"20/03/2025 {ref(320)}",
        "agc_continuacao": f"27/03/2025 {ref(320)}",
        "recuperandos": [
            {"nome": f"Empresa Teste Ltda {ref(10)}", "ecac": f"R$ 300.000,00 {ref(60)}",
             "divida_ativa": f"R$ 150.000,00 {ref(61)}"},
        ],
        "endividamento_fiscal_total": f"R$ 450.000,00 {ref(61)}",
        "documentos_salvos": {
            "peticao_inicial": {"status": "Salvo", "folhas": ref(1).strip("()")},
            "quadro_ativos": {"status": "Anexado", "folhas": ref(88).strip("()")},
            "pericia_previa": {"status": "Não existente", "folhas": ""},
            "laudo_imoveis": {"status": "Anexado", "folhas": ref(90).strip("()")},
            "ultimo_rma": {"status": "Anexado", "folhas": ref(95).strip("()")},
            "qgc_recuperando": {"status": "Anexado", "folhas": ref(300).strip("()")},
            "qgc_aj": {"status": "Anexado", "folhas": ref(305).strip("()")},
            "relatorio_divergencia": {"status": "Não existente", "folhas": ""},
            "prj_aditivos": {"status": "Anexado", "folhas": ref(200).strip("()")},
            "atas_agc": {"status": "Anexado", "folhas": ref(320).strip("()")},
        },
    }


def fixture_processo_longo() -> dict:
    dados = _base_limpa("Mov.")
    dados["recuperandos"] = [
        {"nome": f"Recuperando {i} Ltda (Mov. {100 + i})",
         "ecac": f"R$ {i * 10_000},00 (Mov. {200 + i})",
         "divida_ativa": f"R$ {i * 5_000},00 (Mov. {300 + i})"}
        for i in range(1, 11)
    ]
    dados["imoveis_requerentes"] = [
        {"matricula": f"{10_000 + i} (Mov. {400 + i})", "cartorio": f"{i}º RI (Mov. {400 + i})",
         "descricao": f"Imóvel rural {i} (Mov. {400 + i})", "proprietario": f"Recuperando {i} (Mov. {400 + i})"}
        for i in range(1, 21)
    ]
    dados["recursos_relevantes"] = [
        {"recurso": f"Recurso nº {i} (Mov. {500 + i})", "status": f"Pendente (Mov. {500 + i})"}
        for i in range(1, 8)
    ]
    return dados


def fixture_ocr_ruim() -> dict:
    dados = _base_limpa("fls.")
    dados["requerentes"] = "Empresa   Teste\nLtda .  12.345.678/0001-90   (fls. 10 )"
    dados["administrador_judicial"] = "AJ  Consultoria\n(fls . 15)"
    dados["consolidacao_substancial"] = "Deferido (fls 340)"  # sem ponto — ambíguo de propósito
    dados["previsao_encerramento_stay"] = "15/01/2025(fls.350)"  # sem espaço
    dados["qgc"]["total"] = "R$  10.600.000,00\n(fls. 300)"
    return dados


def fixture_vazio_total() -> dict:
    return {}


def fixture_tipos_malformados() -> dict:
    return {
        "rj_numero": None,
        "requerentes": None,
        "recursos_relevantes": None,
        "imoveis_requerentes": "não é uma lista",
        "imoveis_essenciais": [{"matricula": None, "cartorio": None}],
        "prj_classe_ii": None,
        "qgc": None,
        "recuperandos": [],
        "documentos_salvos": None,
    }


def fixture_checkbox_ambiguo() -> dict:
    dados = _base_limpa("Mov.")
    dados["consolidacao_substancial"] = "Deferido e Indeferido em parte (Mov. 12)"
    dados["periodo_blindagem"] = "Favorável"  # não é opção válida do campo — deve ficar tudo desmarcado
    return dados


def fixture_vazamento_proposital() -> dict:
    dados = _base_limpa("Mov.")
    dados["requerentes"] = "Empresa Teste Ltda (fls.)"
    dados["consolidacao_substancial"] = "Deferido (referência)"
    dados["qgc"]["total"] = "R$ 10.600.000,00 (Mov./ID/fls./Evento)"
    dados["documentos_salvos"]["peticao_inicial"] = {
        "status": "Salvo", "folhas": "[caso haja, indicar páginas]",
    }
    dados["recuperandos"][0]["nome"] = ""  # deve cair no fallback "Não consta"
    return dados


FIXTURES = {
    "clean_pje": lambda: _base_limpa("Mov."),
    "clean_esaj": lambda: _base_limpa("fls."),
    "clean_eproc": lambda: _base_limpa("Evento"),
    "clean_projudi": lambda: _base_limpa("ID"),
    "processo_longo_10_recuperandos": fixture_processo_longo,
    "ocr_ruim": fixture_ocr_ruim,
    "vazio_total": fixture_vazio_total,
    "tipos_malformados": fixture_tipos_malformados,
    "checkbox_ambiguo": fixture_checkbox_ambiguo,
    "vazamento_proposital": fixture_vazamento_proposital,
}

# fixtures que DEVEM sair 100% limpas (zero vazamento)
FIXTURES_ESPERAM_ZERO_VAZAMENTO = {
    "clean_pje", "clean_esaj", "clean_eproc", "clean_projudi",
    "processo_longo_10_recuperandos", "ocr_ruim", "vazio_total", "tipos_malformados",
    "checkbox_ambiguo",
}
# fixture que DEVE ter vazamento detectado (prova que o checador funciona)
FIXTURES_ESPERAM_VAZAMENTO = {"vazamento_proposital"}


# ══════════════════════════════════════════════════════════════════════════
# Testes de _cb (lógica de checkbox)
# ══════════════════════════════════════════════════════════════════════════

CASOS_CB = [
    (["Ativo", "Inativo"], "Ativo", "☑ Ativo"),
    (["Ativo", "Inativo"], "Inativo", "☑ Inativo"),
    (["Ativo", "Inativo"], "", None),  # nenhuma marcada
    (["Favorável", "Desfavorável", "Pendente"], "Desfavorável", "☑ Desfavorável"),
    (["Favorável", "Desfavorável", "Pendente"], "Favorável", "☑ Favorável"),
    (["Deferido", "Indeferido"], "Deferido (Mov. 12)", "☑ Deferido"),
    (["Deferido", "Indeferido"], "Indeferido (fls. 88)", "☑ Indeferido"),
    (["Sim", "Não"], "Não consta", None),
    (["Ativo", "Inativo"], "Deferido e Indeferido em parte (Mov. 12)", None),
    ([], "qualquer coisa", None),
]


def testar_cb():
    falhas = []
    for options, selected, esperado_marcado in CASOS_CB:
        resultado = cr._cb(options, selected)
        marcados = re.findall(r"☑ ([^☐☑]+?)(?:\s{2,}|$)", resultado)
        marcados = [m.strip() for m in marcados]
        if esperado_marcado is None:
            if marcados:
                falhas.append(f"_cb({options!r}, {selected!r}) marcou {marcados}, esperava nenhuma")
        else:
            esperado_texto = esperado_marcado.replace("☑ ", "").strip()
            if marcados != [esperado_texto]:
                falhas.append(f"_cb({options!r}, {selected!r}) => {marcados!r}, esperava [{esperado_texto!r}]")
    return falhas


# ══════════════════════════════════════════════════════════════════════════
# Testes de _extrair (parsing de resposta do "Gemini")
# ══════════════════════════════════════════════════════════════════════════

CASOS_EXTRAIR = [
    ("json_limpo", '{"rj_numero": "123"}', {"rj_numero": "123"}),
    ("com_cercas_markdown", '```json\n{"rj_numero": "123"}\n```', {"rj_numero": "123"}),
    ("com_cercas_sem_lang", '```\n{"rj_numero": "123"}\n```', {"rj_numero": "123"}),
    ("prosa_antes_e_depois", 'Aqui esta o JSON:\n{"rj_numero": "123"}\nFim.', "FALLBACK_VAZIO"),
    ("truncado", '{"rj_numero": "123", "vara":', "FALLBACK_VAZIO"),
    ("vazio", "", "FALLBACK_VAZIO"),
    ("nao_e_objeto", "[1, 2, 3]", "FALLBACK_VAZIO"),
]


def testar_extrair():
    falhas = []
    for nome, resposta, esperado in CASOS_EXTRAIR:
        client = FakeClient(resposta)
        try:
            resultado = cr._extrair("PROMPT ", "TEXTO", client, "modelo-fake")
        except Exception as e:  # noqa: BLE001
            falhas.append(f"_extrair[{nome}] levantou exceção: {e!r}")
            continue
        if esperado == "FALLBACK_VAZIO":
            if resultado != {}:
                falhas.append(f"_extrair[{nome}] esperava fallback {{}}, veio {resultado!r}")
        else:
            if resultado != esperado:
                falhas.append(f"_extrair[{nome}] esperava {esperado!r}, veio {resultado!r}")
    return falhas


# ══════════════════════════════════════════════════════════════════════════
# Execução principal
# ══════════════════════════════════════════════════════════════════════════

def rodar_bateria():
    resultados = []
    falhas_totais = []

    falhas_cb = testar_cb()
    resultados.append(("_cb (lógica de checkbox)", len(CASOS_CB), len(falhas_cb)))
    falhas_totais.extend(f"[_cb] {f}" for f in falhas_cb)

    falhas_extrair = testar_extrair()
    resultados.append(("_extrair (parsing da resposta)", len(CASOS_EXTRAIR), len(falhas_extrair)))
    falhas_totais.extend(f"[_extrair] {f}" for f in falhas_extrair)

    print("=" * 78)
    print("BATERIA DE FIDELIDADE — CHECKLIST RJ")
    print("=" * 78)

    for nome, fabrica in FIXTURES.items():
        dados = copy.deepcopy(fabrica())
        erro_build = None
        vazamentos = []
        n_refs = 0
        try:
            caminho = cr._build_checklist_rj(dados)
            vazamentos = achar_vazamentos(caminho)
            n_refs = contar_referencias(caminho)
        except Exception as e:  # noqa: BLE001
            erro_build = repr(e)

        status = "OK"
        if erro_build:
            status = "ERRO"
            falhas_totais.append(f"[{nome}] _build_checklist_rj levantou {erro_build}")
        elif nome in FIXTURES_ESPERAM_ZERO_VAZAMENTO and vazamentos:
            status = "FALHA (vazamento inesperado)"
            falhas_totais.append(f"[{nome}] vazamentos inesperados: {vazamentos}")
        elif nome in FIXTURES_ESPERAM_VAZAMENTO and not vazamentos:
            status = "FALHA (checador não detectou vazamento esperado)"
            falhas_totais.append(f"[{nome}] esperava vazamento detectado, não achou nenhum")

        print(f"{nome:38s} status={status:38s} refs={n_refs:3d} vazamentos={len(vazamentos)}")
        if vazamentos and nome in FIXTURES_ESPERAM_VAZAMENTO:
            for v in vazamentos:
                print(f"    (esperado) tabela={v[0]} linha={v[1]} col={v[2]}: {v[3]!r}")
        elif vazamentos:
            for v in vazamentos:
                print(f"    !! tabela={v[0]} linha={v[1]} col={v[2]}: {v[3]!r}")

    print("-" * 78)
    for nome, total, n_falhas in resultados:
        print(f"{nome:38s} {total - n_falhas}/{total} OK")

    print("=" * 78)
    if falhas_totais:
        print(f"RESULTADO: {len(falhas_totais)} problema(s) encontrado(s):")
        for f in falhas_totais:
            print(f"  - {f}")
    else:
        print("RESULTADO: nenhum problema encontrado.")
    print("=" * 78)
    return falhas_totais


if __name__ == "__main__":
    falhas = rodar_bateria()
    sys.exit(1 if falhas else 0)


# ── Wrappers pytest (opcional) ──────────────────────────────────────────────
def test_bateria_completa():
    falhas = rodar_bateria()
    assert not falhas, "\n".join(falhas)
