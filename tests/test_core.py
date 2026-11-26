import re
import zipfile

import openpyxl
import pytest
from openpyxl.styles import Font, PatternFill

from excel_ajuster import ajustar_planilha


def _valores(ws):
    return [[c.value for c in linha] for linha in ws.iter_rows()]


def _cenarios(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Cenarios"
    ws.append(["ID", "Cenário", "Passo", "Resultado esperado", "Status"])
    # CT-01: 3 passos, ID e cenário mesclados verticalmente
    ws.append(["CT-01", "Login válido", "Abrir tela", "Tela aberta", "OK"])
    ws.append([None, None, "Informar usuário", "Campo preenchido", "OK"])
    ws.append([None, None, "Clicar entrar", "Home exibida", "OK"])
    ws.merge_cells("A2:A4")
    ws.merge_cells("B2:B4")
    ws["A2"].font = Font(bold=True)
    ws["A2"].fill = PatternFill("solid", fgColor="FFFF00")
    # CT-02: linha "alta" (2 linhas, todas as colunas mescladas verticalmente)
    ws.append(["CT-02", "Login inválido", "Senha errada", "Erro exibido", "NOK"])
    ws.append([None] * 5)
    for col in "ABCDE":
        ws.merge_cells(f"{col}5:{col}6")
    # Mesclagem horizontal: resultado esperado + status
    ws.append(["CT-03", "Logout", "Clicar sair", "Bloqueado", None])
    ws.merge_cells("D7:E7")
    ws.row_dimensions[8].height = 40
    ws.append(["CT-04", "Fim", "x", "y", "OK"])
    wb.save(path)


def test_desmescla_repete_e_remove_linha_redundante(tmp_path):
    src = tmp_path / "cenarios.xlsx"
    _cenarios(src)
    rel = ajustar_planilha(src)
    assert rel.ok, rel.texto()

    wb = openpyxl.load_workbook(tmp_path / "cenarios_ajustado.xlsx")
    ws = wb["Cenarios"]
    assert not ws.merged_cells.ranges
    assert _valores(ws) == [
        ["ID", "Cenário", "Passo", "Resultado esperado", "Status"],
        ["CT-01", "Login válido", "Abrir tela", "Tela aberta", "OK"],
        ["CT-01", "Login válido", "Informar usuário", "Campo preenchido", "OK"],
        ["CT-01", "Login válido", "Clicar entrar", "Home exibida", "OK"],
        ["CT-02", "Login inválido", "Senha errada", "Erro exibido", "NOK"],
        ["CT-03", "Logout", "Clicar sair", "Bloqueado", "Bloqueado"],
        ["CT-04", "Fim", "x", "y", "OK"],
    ]
    # linhas únicas
    linhas = [tuple(l) for l in _valores(ws)]
    assert len(linhas) == len(set(linhas))
    # estilo replicado
    assert ws["A3"].font.bold and ws["A4"].fill.fgColor.rgb.endswith("FFFF00")
    # altura da linha acompanhou a remoção (linha 8 -> 7)
    assert ws.row_dimensions[7].height == 40
    assert rel.abas[0].linhas_removidas == [6]
    # original intacto
    assert openpyxl.load_workbook(src)["Cenarios"].merged_cells.ranges


def test_manter_linhas_redundantes(tmp_path):
    src = tmp_path / "c.xlsx"
    _cenarios(src)
    rel = ajustar_planilha(src, remover_linhas_redundantes=False)
    assert rel.ok
    ws = openpyxl.load_workbook(rel.saida).active
    assert ws.max_row == 8
    assert _valores(ws)[5] == _valores(ws)[4]


def test_duplicatas_originais_nao_sao_removidas(tmp_path):
    src = tmp_path / "d.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["ID", "Passo"])
    ws.append(["CT-1", "a"])
    ws.append(["CT-1", "a"])  # duplicata já existente: dado do usuário
    ws.append(["X", None])
    ws.append([None, None])
    ws.merge_cells("A4:B5")
    wb.save(src)
    rel = ajustar_planilha(src)
    assert rel.ok
    ws = openpyxl.load_workbook(rel.saida).active
    assert _valores(ws) == [["ID", "Passo"], ["CT-1", "a"], ["CT-1", "a"], ["X", "X"]]
    assert rel.abas[0].duplicatas_originais == [(2, 3)]
    assert "ATENÇÃO" in rel.texto()


def _injetar_mesclagem(path, ref):
    """Mescla um intervalo direto no XML, mantendo os valores das células internas
    (simula arquivos gerados por outras ferramentas)."""
    with zipfile.ZipFile(path) as z:
        itens = {n: z.read(n) for n in z.namelist()}
    nome = "xl/worksheets/sheet1.xml"
    xml = itens[nome].decode()
    xml = re.sub(r"(</sheetData>)", rf'\1<mergeCells count="1"><mergeCell ref="{ref}"/></mergeCells>', xml)
    itens[nome] = xml.encode()
    with zipfile.ZipFile(path, "w") as z:
        for n, d in itens.items():
            z.writestr(n, d)


def test_valor_escondido_em_mesclagem_e_preservado(tmp_path):
    src = tmp_path / "h.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Visível", "Escondido"])
    wb.save(src)
    _injetar_mesclagem(src, "A1:B1")
    assert openpyxl.load_workbook(src).active.merged_cells.ranges
    rel = ajustar_planilha(src)
    assert rel.ok, rel.texto()
    ws = openpyxl.load_workbook(rel.saida).active
    assert _valores(ws) == [["Visível", "Escondido"]]
    assert rel.abas[0].valores_escondidos_preservados == ["B1"]


def test_com_formulas_nao_remove_linhas(tmp_path):
    src = tmp_path / "f.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["A", 1])
    ws.append([None, None])
    ws.append(["Total", "=SUM(B1:B2)"])
    ws.merge_cells("A1:A2")
    ws.merge_cells("B1:B2")
    wb.save(src)
    rel = ajustar_planilha(src)
    assert rel.ok
    ws = openpyxl.load_workbook(rel.saida).active
    assert ws.max_row == 3 and ws["B3"].value == "=SUM(B1:B2)"
    assert rel.abas[0].avisos


def test_multiplas_abas_e_formato_invalido(tmp_path):
    src = tmp_path / "m.xlsx"
    wb = openpyxl.Workbook()
    wb.active.append(["a", None])
    wb.active.merge_cells("A1:B1")
    ws2 = wb.create_sheet("Outra")
    ws2.append(["x"])
    wb.save(src)
    rel = ajustar_planilha(src)
    assert rel.ok
    out = openpyxl.load_workbook(rel.saida)
    assert out.sheetnames == ["Sheet", "Outra"]
    assert _valores(out["Sheet"]) == [["a", "a"]]

    bad = tmp_path / "x.xls"
    bad.write_bytes(b"")
    with pytest.raises(ValueError):
        ajustar_planilha(bad)
    with pytest.raises(ValueError):
        ajustar_planilha(src, src)
