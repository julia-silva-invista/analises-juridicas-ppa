"""Testes das barreiras de fidelidade, OCR e rastreabilidade jurídica."""
from __future__ import annotations

import inspect
import sys
import tempfile
from pathlib import Path

import fitz

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import legal_prompts as lp
import matriculas
import processos
import rj
import utils


def test_regras_da_usuaria_foram_preservadas_no_prompt_comum():
    regras = lp.REGRAS_CONSOLIDACAO_PROCESSUAL
    regras_linha = " ".join(regras.split())
    assert "PÁGINA ABSOLUTA DO PDF" in regras
    assert "JAMAIS pode aparecer" in regras
    assert "pdf\" em minúsculas" in regras
    assert 'escreva exatamente "Não há"' in regras
    assert "número/título da cláusula" in regras
    assert "Mapeie TODOS os aditamentos principais" in regras
    assert "Se houver dúvida razoável sobre a relevância" in regras_linha
    assert "01/2002" in regras and "10/2000" in regras
    assert "ordem física das páginas" in regras


def test_contexto_vincula_pagina_local_a_absoluta_sem_promover_parte():
    contexto = lp.contexto_fonte_pdf("P4.PDF", 401, 800, [401, 410])
    assert "arquivo original: P4.PDF" in contexto
    assert "página local 1 deste trecho corresponde à página absoluta 401" in contexto
    assert "401, 410" in contexto
    assert "NÃO CITE ESTE CABEÇALHO" in contexto


def test_barreira_final_remove_referencia_interna_e_respeita_numero_de_pdfs():
    ruim = (
        "Ato deferido (Parte 3, pág. 25).\n"
        "Outro ato (fl. 44).\n"
        "Ato sem trilha (p. 19).\n"
        "Decisão (ID 9 | fl. 3 do PDF unico.pdf)."
    )
    unico = lp.normalizar_referencias_relatorio(ruim, multiplos_pdfs=False)
    assert "Parte 3" not in unico
    assert "pág. 25" not in unico
    assert "(p. 19)" not in unico
    assert "identificador processual não localizado | fl. 44" in unico
    assert "do pdf" not in unico.lower()

    multiplo = lp.normalizar_referencias_relatorio(
        "Decisão (ID 9 | fl. 3 do PDF P2.PDF).", multiplos_pdfs=True
    )
    assert "do pdf P2.pdf" in multiplo

    objeto = lp.normalizar_referencias_objeto(
        {"itens": [{"referencia": "Ato (Parte 2, pág. 8)"}]},
        multiplos_pdfs=False,
    )
    assert "Parte 2" not in objeto["itens"][0]["referencia"]


def test_deteccao_local_distingue_texto_pesquisavel_de_pagina_digitalizada():
    path = Path(tempfile.mktemp(prefix="fidelidade_ocr_", suffix=".pdf"))
    try:
        doc = fitz.open()
        p1 = doc.new_page()
        p1.insert_textbox(
            fitz.Rect(50, 50, 550, 750),
            "Texto pesquisável completo do contrato " * 50,
            fontsize=10,
        )
        p2 = doc.new_page()
        pix = fitz.Pixmap(fitz.csGRAY, 50, 50, bytes([220]) * 2500, False)
        p2.insert_image(p2.rect, pixmap=pix)
        doc.save(path)
        doc.close()
        assert utils._paginas_digitalizadas_pdf(str(path)) == [2]
        assert utils._paginas_digitalizadas_pdf(str(path), offset=400) == [402]
    finally:
        path.unlink(missing_ok=True)


def test_fluxos_escolhem_modelo_forte_para_paginas_digitalizadas():
    fonte_proc = inspect.getsource(processos._proc_analisar_impl)
    fonte_rj = inspect.getsource(rj._rj_analisar_impl)
    fonte_mat = inspect.getsource(matriculas._mat_analisar_pdf)
    for fonte in (fonte_proc, fonte_rj, fonte_mat):
        assert "GEMINI_MODEL_OCR" in fonte
        assert "paginas" in fonte.lower() and "digitalizadas" in fonte.lower()


def test_rj_nao_descarta_o_final_da_fonte_longa():
    fonte = inspect.getsource(rj._rj_analisar_impl)
    relacionados = inspect.getsource(rj._rj_processar_relacionados)
    assert "[:LIMITE_CONSOLIDACAO_RJ]" not in fonte
    assert "[:LIMITE_CONSOLIDACAO_RJ]" not in relacionados
    assert "_rj_preservar_fonte_longa" in fonte
    assert "_rj_preservar_fonte_longa" in relacionados

    bloco = "=" * 60 + "\nPARTE {i}/4\n" + ("x" * 120)
    texto = "\n\n".join(bloco.format(i=i) for i in range(1, 5))
    lotes = rj._rj_dividir_fonte_em_lotes(texto, limite=250)
    assert len(lotes) == 4
    assert "".join(lotes).count("x") == texto.count("x")


def test_fluxos_bloqueiam_relatorio_quando_um_chunk_falha():
    fonte_proc = inspect.getsource(processos._proc_analisar_impl)
    fonte_proc_rel = inspect.getsource(processos._proc_processar_relacionados)
    fonte_rj = inspect.getsource(rj._rj_analisar_impl)
    fonte_rj_rel = inspect.getsource(rj._rj_processar_relacionados)
    for fonte in (fonte_proc, fonte_proc_rel, fonte_rj, fonte_rj_rel):
        assert "nenhum relatório parcial" in fonte or "relatório parcial foi bloqueado" in fonte
    assert "paginas_omitidas" in inspect.getsource(processos._proc_extrair_chunk)
    assert "paginas_omitidas" in inspect.getsource(rj._rj_extrair_chunk)


def test_fluxos_redividem_recursivamente_apenas_quando_tokens_estouram():
    path = Path(tempfile.mktemp(prefix="token_split_", suffix=".pdf"))
    doc = fitz.open()
    for indice in range(4):
        page = doc.new_page()
        page.insert_text((72, 72), f"PAGINA {indice + 1}")
    doc.save(path)
    doc.close()

    try:
        for modulo, nome_adaptativo, nome_fileapi, nome_inline in (
            (processos, "_proc_extrair_chunk_adaptativo", "_proc_extrair_chunk_fileapi", "_proc_extrair_chunk"),
            (rj, "_rj_extrair_chunk_adaptativo", "_rj_extrair_chunk_fileapi", "_rj_extrair_chunk"),
        ):
            adaptativo = getattr(modulo, nome_adaptativo)
            original_fileapi = getattr(modulo, nome_fileapi)
            original_inline = getattr(modulo, nome_inline)
            caminhos_vistos = []

            def fake_fileapi(args):
                idx, chunk_path, offset, *_rest = args
                caminhos_vistos.append(Path(chunk_path))
                with fitz.open(chunk_path) as chunk_doc:
                    quantidade = len(chunk_doc)
                if quantidade > 1:
                    raise RuntimeError(
                        "400 INVALID_ARGUMENT: The input token count exceeds the maximum number of tokens allowed 1048576"
                    )
                return idx, f"pagina absoluta {offset + 1}", "fileapi fake"

            def inline_nao_deve_rodar(_args):
                raise AssertionError("excesso de tokens não deve tentar novamente o chunk inteiro inline")

            setattr(modulo, nome_fileapi, fake_fileapi)
            setattr(modulo, nome_inline, inline_nao_deve_rodar)
            try:
                _idx, texto, nota = adaptativo(
                    (0, str(path), 400, 1000, 1, object(), "P1.PDF", [])
                )
            finally:
                setattr(modulo, nome_fileapi, original_fileapi)
                setattr(modulo, nome_inline, original_inline)

            assert [f"pagina absoluta {p}" in texto for p in range(401, 405)] == [True] * 4
            assert [texto.index(f"pagina absoluta {p}") for p in range(401, 405)] == sorted(
                texto.index(f"pagina absoluta {p}") for p in range(401, 405)
            )
            assert "divisão adaptativa por tokens" in nota
            temporarios = {p for p in caminhos_vistos if p != path}
            assert temporarios and all(not p.exists() for p in temporarios)
    finally:
        path.unlink(missing_ok=True)


if __name__ == "__main__":
    test_regras_da_usuaria_foram_preservadas_no_prompt_comum()
    test_contexto_vincula_pagina_local_a_absoluta_sem_promover_parte()
    test_barreira_final_remove_referencia_interna_e_respeita_numero_de_pdfs()
    test_deteccao_local_distingue_texto_pesquisavel_de_pagina_digitalizada()
    test_fluxos_escolhem_modelo_forte_para_paginas_digitalizadas()
    test_rj_nao_descarta_o_final_da_fonte_longa()
    test_fluxos_bloqueiam_relatorio_quando_um_chunk_falha()
    test_fluxos_redividem_recursivamente_apenas_quando_tokens_estouram()
    print("8 testes de fidelidade jurídica: OK")
