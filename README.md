# Excel-Ajuster

Ajusta planilhas de **cenários de teste** removendo **todas as células mescladas** e
**repetindo o dado** em cada célula desmesclada — sem perder nenhuma informação e
garantindo que cada linha de teste fique completa e única.

## Versão web (sem instalar nada)

Abra o arquivo [`index.html`](index.html) no navegador (duplo clique), envie a planilha e
clique em **Baixar planilha ajustada**. Tudo roda localmente no navegador: a planilha não é
enviada para nenhum servidor. Para disponibilizar para o time, ative o GitHub Pages
(Settings > Pages > Branch `main`, pasta `/`) e compartilhe o link.

A versão web trabalha direto no XML do `.xlsx`, então estilos, comentários, imagens,
gráficos, filtros e macros (`.xlsm`) ficam exatamente como estavam. Abas com fórmulas,
comentários, links, imagens, validação de dados ou formatação condicional não têm linhas
removidas (apenas aviso), para não quebrar referências.

## O que ele faz

| Antes (mesclado)                 | Depois (ajustado)                       |
|----------------------------------|-----------------------------------------|
| `CT-01` mesclado em 3 linhas     | `CT-01` repetido nas 3 linhas (1 por passo) |
| Cabeçalho mesclado em 2 colunas  | Texto repetido nas 2 colunas            |
| Linha "alta" (2 linhas mescladas)| Uma única linha (a cópia redundante é removida) |

Regras de segurança:

- **O original nunca é alterado**: é gerado `<nome>_ajustado.xlsx` ao lado.
- O valor **e a formatação** (fonte, cor, alinhamento, bordas, formato numérico) da
  célula mesclada são copiados para todas as células do intervalo.
- **Valores escondidos** dentro de mesclagens (comuns em arquivos gerados por outras
  ferramentas) são detectados e preservados — nunca sobrescritos.
- **Linhas únicas**: só são removidas as linhas que ficaram idênticas à anterior
  *exclusivamente por causa da desmesclagem* e que não têm nenhum dado próprio.
  Linhas que **já eram** duplicadas no original **não são apagadas** — aparecem como
  `ATENÇÃO` no relatório para revisão manual.
- Se a aba tiver fórmulas, formatação condicional, validação de dados ou tabelas,
  nenhuma linha é removida (para não quebrar referências) e um aviso é exibido.
- Ao final, o arquivo gerado é **reaberto e validado**: nenhuma mesclagem restante e
  todo valor original presente. Se algo falhar, o relatório mostra `Validação: FALHOU`.
- Todas as abas são processadas; macros de `.xlsm` são mantidas.

## Versão em Python (linha de comando)

### Instalação

Requer Python 3.9+.

```bash
pip install -r requirements.txt
```

### Uso

```bash
# um arquivo
python -m excel_ajuster cenarios.xlsx

# vários arquivos ou uma pasta inteira
python -m excel_ajuster plano1.xlsx plano2.xlsx
python -m excel_ajuster pasta_de_planilhas/

# escolher o nome de saída
python -m excel_ajuster cenarios.xlsx -o cenarios_limpo.xlsx

# não remover as linhas que ficaram repetidas pela desmesclagem
python -m excel_ajuster cenarios.xlsx --manter-linhas-redundantes
```

**No Windows**: arraste a(s) planilha(s) ou uma pasta sobre o arquivo `ajustar.bat`.

Também pode ser usado como biblioteca:

```python
from excel_ajuster import ajustar_planilha

relatorio = ajustar_planilha("cenarios.xlsx")
print(relatorio.texto())
```

## Formatos

`.xlsx`, `.xlsm`, `.xltx`, `.xltm`. Arquivos `.xls` antigos: abra no Excel e use
*Salvar como > Pasta de Trabalho do Excel (.xlsx)* antes.

## Testes

```bash
python -m pytest
```
