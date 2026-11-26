"""Núcleo do Excel-Ajuster: desmescla células e repete os dados sem perder nada.

Fluxo por planilha (aba):
  1. Lê os valores "escondidos" dentro de intervalos mesclados (o openpyxl os
     descarta ao abrir o arquivo normalmente, então lemos em modo read-only).
  2. Desmescla cada intervalo e preenche todas as células com o valor e o
     estilo da célula superior esquerda. Se uma célula tinha um valor próprio
     escondido, ele é preservado (nada é sobrescrito).
  3. Remove apenas as linhas que ficaram duplicadas *por causa* da
     desmesclagem (linha sem nenhum conteúdo próprio, idêntica à anterior).
     Duplicatas que já existiam no arquivo original são apenas relatadas.
  4. Valida o resultado: nenhuma mesclagem restante e todo valor original
     presente no arquivo final.
"""

from __future__ import annotations

import copy
import os
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl
from openpyxl.cell.cell import MergedCell
from openpyxl.utils import get_column_letter


@dataclass
class RelatorioAba:
    nome: str
    mesclagens_removidas: int = 0
    celulas_preenchidas: int = 0
    valores_escondidos_preservados: list[str] = field(default_factory=list)
    linhas_removidas: list[int] = field(default_factory=list)
    duplicatas_originais: list[tuple[int, int]] = field(default_factory=list)
    avisos: list[str] = field(default_factory=list)


@dataclass
class Relatorio:
    entrada: str
    saida: str
    abas: list[RelatorioAba] = field(default_factory=list)
    erros_validacao: list[str] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not self.erros_validacao

    def texto(self) -> str:
        linhas = [f"Arquivo de entrada: {self.entrada}", f"Arquivo gerado:     {self.saida}", ""]
        for a in self.abas:
            linhas.append(f"[Aba '{a.nome}']")
            linhas.append(f"  Mesclagens removidas:  {a.mesclagens_removidas}")
            linhas.append(f"  Células preenchidas:   {a.celulas_preenchidas}")
            if a.valores_escondidos_preservados:
                linhas.append(
                    "  Valores escondidos em mesclagens preservados: "
                    + ", ".join(a.valores_escondidos_preservados)
                )
            if a.linhas_removidas:
                linhas.append(
                    "  Linhas redundantes removidas (duplicadas pela desmesclagem, sem dado próprio): "
                    + ", ".join(map(str, a.linhas_removidas))
                )
            for orig, dup in a.duplicatas_originais:
                linhas.append(
                    f"  ATENÇÃO: linha {dup} já era idêntica à linha {orig} no arquivo original "
                    "(mantida para não perder dados — revise manualmente)."
                )
            for aviso in a.avisos:
                linhas.append(f"  AVISO: {aviso}")
            linhas.append("")
        if self.ok:
            linhas.append("Validação: OK — nenhuma célula mesclada restante e nenhum dado perdido.")
        else:
            linhas.append("Validação: FALHOU")
            linhas.extend(f"  - {e}" for e in self.erros_validacao)
        return "\n".join(linhas)


def _vazio(v) -> bool:
    return v is None or (isinstance(v, str) and v.strip() == "")


def _valores_escondidos(caminho: Path, intervalos_por_aba: dict[str, list], keep_vba: bool) -> dict:
    """Retorna {aba: {(linha, coluna): valor}} para células que NÃO são a superior
    esquerda de uma mesclagem mas possuem valor gravado no arquivo."""
    resultado: dict[str, dict[tuple[int, int], object]] = {}
    alvo = {aba: ints for aba, ints in intervalos_por_aba.items() if ints}
    if not alvo:
        return resultado
    wb_ro = openpyxl.load_workbook(caminho, read_only=True, keep_vba=keep_vba)
    try:
        for aba, intervalos in alvo.items():
            ws = wb_ro[aba]
            if not hasattr(ws, "iter_rows"):
                continue
            min_r = min(i.min_row for i in intervalos)
            max_r = max(i.max_row for i in intervalos)
            encontrados = {}
            for linha in ws.iter_rows(min_row=min_r, max_row=max_r):
                for cel in linha:
                    if getattr(cel, "value", None) is None or not hasattr(cel, "column"):
                        continue
                    r, c = cel.row, cel.column
                    for i in intervalos:
                        if i.min_row <= r <= i.max_row and i.min_col <= c <= i.max_col:
                            if (r, c) != (i.min_row, i.min_col):
                                encontrados[(r, c)] = cel.value
                            break
            if encontrados:
                resultado[aba] = encontrados
    finally:
        wb_ro.close()
    return resultado


def _assinatura_linha(ws, r: int, max_col: int) -> tuple:
    return tuple(ws.cell(row=r, column=c).value for c in range(1, max_col + 1))


def _riscos_ao_remover_linhas(ws) -> list[str]:
    """Estruturas que o openpyxl não reajusta ao excluir linhas."""
    riscos = []
    for linha in ws.iter_rows():
        if any(isinstance(c.value, str) and c.value.startswith("=") for c in linha):
            riscos.append("fórmulas")
            break
    if ws.conditional_formatting:
        riscos.append("formatação condicional")
    if ws.data_validations.dataValidation:
        riscos.append("validação de dados")
    if ws.tables:
        riscos.append("tabelas")
    return riscos


def ajustar_planilha(
    entrada: str | os.PathLike,
    saida: str | os.PathLike | None = None,
    remover_linhas_redundantes: bool = True,
) -> Relatorio:
    entrada = Path(entrada)
    if not entrada.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {entrada}")
    ext = entrada.suffix.lower()
    if ext not in (".xlsx", ".xlsm", ".xltx", ".xltm"):
        raise ValueError(
            f"Formato '{ext}' não suportado. Salve o arquivo como .xlsx no Excel "
            "(Arquivo > Salvar como > Pasta de Trabalho do Excel) e tente novamente."
        )
    if saida is None:
        saida = entrada.with_name(f"{entrada.stem}_ajustado{entrada.suffix}")
    saida = Path(saida)
    if saida.resolve() == entrada.resolve():
        raise ValueError("O arquivo de saída não pode ser o mesmo da entrada (o original é sempre preservado).")

    keep_vba = ext in (".xlsm", ".xltm")
    wb = openpyxl.load_workbook(entrada, keep_vba=keep_vba)
    relatorio = Relatorio(entrada=str(entrada), saida=str(saida))

    intervalos_por_aba = {
        ws.title: list(ws.merged_cells.ranges) for ws in wb.worksheets
    }
    escondidos = _valores_escondidos(entrada, intervalos_por_aba, keep_vba)

    # Fotografia dos valores originais (para validação final): {aba: {(r, c): valor}}
    originais: dict[str, dict[tuple[int, int], object]] = {}
    esperados_por_aba: dict[str, dict[tuple[int, int], object]] = {}
    mapa_linhas: dict[str, dict[int, int]] = {}

    for ws in wb.worksheets:
        rel = RelatorioAba(nome=ws.title)
        relatorio.abas.append(rel)
        intervalos = intervalos_por_aba[ws.title]
        esc = escondidos.get(ws.title, {})

        originais[ws.title] = {
            (c.row, c.column): c.value
            for linha in ws.iter_rows()
            for c in linha
            if not isinstance(c, MergedCell) and c.value is not None
        }
        for (r, c), v in esc.items():
            originais[ws.title][(r, c)] = v

        # Células com conteúdo próprio no original (antes de preencher).
        proprias = {pos for pos, v in originais[ws.title].items() if not _vazio(v)}

        # Duplicatas que já existiam antes (apenas relatadas, nunca removidas).
        max_col = ws.max_column
        vistos: dict[tuple, int] = {}
        for r in range(1, ws.max_row + 1):
            assin = tuple(originais[ws.title].get((r, c)) for c in range(1, max_col + 1))
            if all(_vazio(v) for v in assin):
                continue
            if assin in vistos:
                rel.duplicatas_originais.append((vistos[assin], r))
            else:
                vistos[assin] = r

        # Desmesclar e preencher.
        esperado: dict[tuple[int, int], object] = dict(originais[ws.title])
        for intervalo in intervalos:
            min_r, min_c, max_r, max_c = (
                intervalo.min_row, intervalo.min_col, intervalo.max_row, intervalo.max_col
            )
            ws.unmerge_cells(str(intervalo))
            rel.mesclagens_removidas += 1
            origem = ws.cell(row=min_r, column=min_c)
            valor = origem.value
            for r in range(min_r, max_r + 1):
                for c in range(min_c, max_c + 1):
                    if (r, c) == (min_r, min_c):
                        continue
                    destino = ws.cell(row=r, column=c)
                    destino._style = copy.copy(origem._style)
                    if (r, c) in esc:
                        destino.value = esc[(r, c)]
                        rel.valores_escondidos_preservados.append(
                            f"{get_column_letter(c)}{r}"
                        )
                    else:
                        destino.value = valor
                        if valor is not None:
                            rel.celulas_preenchidas += 1
                    esperado[(r, c)] = destino.value

        # Remover linhas que ficaram duplicadas apenas por causa da desmesclagem.
        remover: list[int] = []
        if remover_linhas_redundantes and intervalos:
            max_col = ws.max_column
            for r in range(2, ws.max_row + 1):
                if any((r, c) in proprias for c in range(1, max_col + 1)):
                    continue  # a linha tem dado próprio: nunca remover
                atual = _assinatura_linha(ws, r, max_col)
                if all(_vazio(v) for v in atual):
                    continue  # linha vazia original: não mexer
                if atual == _assinatura_linha(ws, r - 1, max_col):
                    remover.append(r)
            if remover:
                riscos = _riscos_ao_remover_linhas(ws)
                if riscos:
                    rel.avisos.append(
                        f"{len(remover)} linha(s) duplicada(s) pela desmesclagem NÃO foram removidas "
                        f"porque a aba contém {', '.join(riscos)} (remover poderia quebrar referências). "
                        f"Linhas: {', '.join(map(str, remover))}."
                    )
                    remover = []

        mapa = {}
        deslocamento = 0
        remover_set = set(remover)
        for r in range(1, ws.max_row + 1):
            if r in remover_set:
                deslocamento += 1
            else:
                mapa[r] = r - deslocamento
        for r in sorted(remover, reverse=True):
            alturas = {k: copy.copy(v) for k, v in ws.row_dimensions.items() if k > r}
            ws.delete_rows(r, 1)
            for k in alturas:
                if k in ws.row_dimensions:
                    del ws.row_dimensions[k]
            for k, dim in alturas.items():
                dim.index = k - 1
                ws.row_dimensions[k - 1] = dim
        rel.linhas_removidas = remover
        mapa_linhas[ws.title] = mapa
        esperados_por_aba[ws.title] = esperado

    saida.parent.mkdir(parents=True, exist_ok=True)
    wb.save(saida)

    relatorio.erros_validacao = _validar(saida, originais, esperados_por_aba, mapa_linhas, keep_vba)
    return relatorio


def _validar(saida: Path, originais, esperados, mapa_linhas, keep_vba: bool) -> list[str]:
    erros: list[str] = []
    wb = openpyxl.load_workbook(saida, keep_vba=keep_vba)
    for aba, valores in originais.items():
        if aba not in wb.sheetnames:
            erros.append(f"Aba '{aba}' ausente no arquivo gerado.")
            continue
        ws = wb[aba]
        if ws.merged_cells.ranges:
            erros.append(f"Aba '{aba}': ainda há {len(ws.merged_cells.ranges)} mesclagem(ns).")
        mapa = mapa_linhas[aba]
        # 1) Todo valor original continua presente (na linha mapeada, ou numa linha
        #    removida cujo conteúdo é idêntico ao da linha mantida acima).
        for (r, c), v in valores.items():
            nr = mapa.get(r)
            if nr is None:
                # linha removida: o valor deve existir na mesma coluna da linha mantida anterior
                acima = max((k for k in mapa if k < r), default=None)
                nr = mapa.get(acima) if acima else None
            if nr is None or ws.cell(row=nr, column=c).value != v:
                erros.append(f"Aba '{aba}': valor original de {get_column_letter(c)}{r} ({v!r}) não encontrado.")
        # 2) Toda célula desmesclada recebeu o valor esperado.
        for (r, c), v in esperados[aba].items():
            nr = mapa.get(r)
            if nr is None:
                continue
            if ws.cell(row=nr, column=c).value != v:
                erros.append(
                    f"Aba '{aba}': {get_column_letter(c)}{nr} deveria conter {v!r}."
                )
    wb.close()
    return erros
