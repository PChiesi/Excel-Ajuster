"""Interface de linha de comando.

Uso:
    python -m excel_ajuster cenarios.xlsx [outro.xlsx ...] [-o saida.xlsx]
    python -m excel_ajuster pasta_com_planilhas/

Sem argumentos, pergunta o caminho do arquivo (útil ao dar duplo clique).
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .core import ajustar_planilha

EXTENSOES = {".xlsx", ".xlsm", ".xltx", ".xltm"}


def _expandir(caminhos: list[str]) -> list[Path]:
    arquivos: list[Path] = []
    for c in caminhos:
        p = Path(c.strip().strip('"'))
        if p.is_dir():
            arquivos.extend(
                sorted(
                    f for f in p.iterdir()
                    if f.suffix.lower() in EXTENSOES
                    and not f.stem.endswith("_ajustado")
                    and not f.name.startswith("~$")
                )
            )
        else:
            arquivos.append(p)
    return arquivos


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="excel_ajuster",
        description=(
            "Remove todas as células mescladas, repetindo o dado em cada célula "
            "desmesclada, sem perder nenhuma informação. O arquivo original nunca é alterado: "
            "é gerado um novo arquivo '<nome>_ajustado.xlsx'."
        ),
    )
    parser.add_argument("arquivos", nargs="*", help="Planilhas (.xlsx/.xlsm) ou pastas")
    parser.add_argument("-o", "--saida", help="Caminho do arquivo de saída (apenas com 1 arquivo)")
    parser.add_argument(
        "--manter-linhas-redundantes",
        action="store_true",
        help="Não remove linhas que ficaram idênticas à anterior por causa da desmesclagem",
    )
    args = parser.parse_args(argv)

    interativo = not args.arquivos
    if interativo:
        caminho = input("Arraste a planilha para cá (ou digite o caminho) e tecle Enter: ")
        args.arquivos = [caminho]

    arquivos = _expandir(args.arquivos)
    if not arquivos:
        print("Nenhuma planilha encontrada.", file=sys.stderr)
        return 2
    if args.saida and len(arquivos) > 1:
        parser.error("--saida só pode ser usado com um único arquivo")

    codigo = 0
    for arq in arquivos:
        print("=" * 70)
        try:
            rel = ajustar_planilha(
                arq,
                args.saida,
                remover_linhas_redundantes=not args.manter_linhas_redundantes,
            )
        except Exception as exc:  # noqa: BLE001 - mostrar erro amigável ao usuário
            print(f"ERRO ao processar {arq}: {exc}", file=sys.stderr)
            codigo = 1
            continue
        print(rel.texto())
        if not rel.ok:
            codigo = 1

    if interativo:
        input("\nPressione Enter para sair...")
    return codigo
