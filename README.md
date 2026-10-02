# Pipeline de Análise de Maturidade BIM

Sistema automatizado para avaliar a maturidade BIM de modelos Revit (.rvt)
com base em 14 critérios objetivos, gerando dois dashboards independentes:

- **Streamlit** — para uso pessoal e iteração rápida
- **Power BI** — para compartilhamento com o time e apresentações

---

## Sumário

1. [Estrutura do projeto](#1-estrutura-do-projeto)
2. [Pré-requisitos](#2-pré-requisitos)
3. [Como rodar](#3-como-rodar)
4. [O que cada script faz](#4-o-que-cada-script-faz)
5. [Os 14 checks automáticos](#5-os-14-checks-automáticos)
6. [Métricas técnicas extraídas](#6-métricas-técnicas-extraídas)
7. [Como interpretar as notas](#7-como-interpretar-as-notas)
8. [Manutenção](#8-manutenção)
9. [Solução de problemas](#9-solução-de-problemas)
10. [Changelog](#10-changelog)

---

## 1. Estrutura do projeto

```
X:\06_COMPARTILHADOS DE PROJETOS\ARQUITETURA\DIEGO\
│
├── PY_ORQUESTRACAO\                              ← Scripts Python 3 (fora do Revit)
│   ├── bim_paths.py                              ← Caminhos centralizados
│   ├── run_batch.py                              ← Dispara o pyRevit para cada .rvt
│   ├── consolidar.py                             ← JSONs → Excel (Streamlit)
│   ├── exportar_bi.py                            ← JSONs → CSVs (Power BI)
│   ├── dashboard.py                              ← Dashboard Streamlit
│   ├── run_all.bat                               ← Pipeline completo (duplo-clique)
│   ├── run_only_consolidar.bat                   ← Só reconsolida Excel + CSVs
│   ├── run_only_dashboard.bat                    ← Só abre o Streamlit
│   └── README.md                                 ← Este arquivo
│
├── PYREVIT SCRIPTS\BIM.extension\                ← Extensão pyRevit (IronPython 2.7)
│   └── BIM.tab\Analise.panel\Extrair.pushbutton\
│       ├── script.py                             ← Roda dentro do Revit 2020
│       └── bim_paths.py                          ← CÓPIA do orquestrador
│
└── TEMPLATE_LIMPEZA\ANALISE VIEW TEMPLATES EM USO\
    ├── MODELOS ATIVOS\                           ← Entrada: arquivos .rvt
    ├── _RESULTADOS\                              ← Saída principal
    │   ├── metricas_<nome>.json                  ← 1 JSON por modelo
    │   ├── dashboard_bim.xlsx                    ← Excel do Streamlit
    │   ├── batch.log                             ← Log do pyRevit CLI
    │   └── _debug_script.log                     ← Log detalhado do script.py
    ├── _BI_EXPORT\                               ← Saída para Power BI
    │   ├── 2026-09\                              ← Subpasta por período (histórico)
    │   │   ├── bi_modelos.csv
    │   │   ├── bi_checks.csv
    │   │   ├── bi_metricas.csv
    │   │   └── bi_resumo.csv
    │   └── 2026-10\                              ← Próximo mês (automático)
    └── 01-AUDIT-BIM.xlsx                         ← Pesos e descrições dos checks
```

---

## 2. Pré-requisitos

### 2.1 Softwares necessários

| Software | Versão | Como instalar |
|---|---|---|
| Python 3.10+ | 3.14 testado | Já instalado em `C:\Users\<user>\AppData\Local\Python\` |
| Revit 2020 | 20.0.0.377 | Já instalado |
| pyRevit CLI | 4.x+ | https://github.com/eirannejad/pyRevit/releases |
| Power BI Desktop | qualquer | Microsoft Store ou site oficial |

### 2.2 Pacotes Python

Instalar uma vez, no terminal (**não em arquivo .py**):

```powershell
python -m pip install --upgrade openpyxl streamlit pandas plotly
```

### 2.3 Configuração do pyRevit

1. Instalar o pyRevit (executável `.exe` do GitHub)
2. No Revit → aba **pyRevit** → **Settings** → `Custom Extension Directories`, adicionar:
   ```
   X:\06_COMPARTILHADOS DE PROJETOS\ARQUITETURA\DIEGO\PYREVIT SCRIPTS
   ```
3. Confirmar que o CLI está acessível:
   ```powershell
   & "C:\Program Files\pyRevit CLI\bin\pyrevit.exe" run --help
   ```

### 2.4 Caminho do pyRevit no `run_batch.py`

Confirmar que a constante `PYREVIT_CLI` no `run_batch.py` aponta para o `.exe` correto:

```python
PYREVIT_CLI = r"C:\Program Files\pyRevit CLI\bin\pyrevit.exe"
```

---

## 3. Como rodar

### 3.1 Fluxo completo (mais comum)

1. Coloque os `.rvt` a analisar em `MODELOS ATIVOS\`
2. **Duplo-clique em `run_all.bat`**
3. Aguarde **25–40 min** (o Revit abre/fecha para cada modelo)
4. O dashboard Streamlit abre automaticamente em `http://localhost:8501`
5. Para ver no Power BI: abrir o `.pbix` e clicar em **Atualizar**

**O que `run_all.bat` faz:**
```
[1/4] python run_batch.py     → Revit processa cada .rvt → metricas_*.json
[2/4] python consolidar.py    → JSONs → dashboard_bim.xlsx (Streamlit)
[3/4] python exportar_bi.py   → JSONs → CSVs em _BI_EXPORT/2026-09/
[4/4] streamlit run dashboard.py  → abre dashboard web
```

### 3.2 Rodar só o dashboard Streamlit

Quando o Excel já existe e você só quer visualizar:

**Duplo-clique em `run_only_dashboard.bat`**

### 3.3 Reconsolidar sem re-rodar o Revit

Quando mudar a régua de notas no `consolidar.py` ou os pesos no `01-AUDIT-BIM.xlsx`:

**Duplo-clique em `run_only_consolidar.bat`**

Regenera o `dashboard_bim.xlsx` e os CSVs em `_BI_EXPORT/` em segundos.

### 3.4 Atualizar o Power BI

Depois de rodar qualquer um dos scripts acima:

1. Abra o arquivo `.pbix` no Power BI Desktop
2. `Página Inicial → Atualizar` (ou `Ctrl+Alt+F5`)
3. O Power BI lê automaticamente os CSVs de `_BI_EXPORT/2026-09/`

---

## 4. O que cada script faz

### `run_batch.py` (Python 3)

- Lista todos os `.rvt` de `MODELOS ATIVOS\`
- Para cada um, chama o pyRevit CLI:
  ```
  pyrevit run script.py <modelo> --revit=2020 --purge --allowdialogs
  ```
- **Workaround para nomes acentuados:** se o arquivo tiver caracteres não-ASCII (ex: `NANÔBE`), copia para `C:\Temp\bim_ascii\` antes de processar e apaga depois
- Tempo: **~2–4 min por modelo**

### `script.py` (roda dentro do Revit, IronPython 2.7)

- Recebe a lista de modelos via variável global `__models__`
- Para cada um:
  1. Abre o modelo com `Application.OpenDocumentFile()`
  2. Coleta **26 métricas** (14 checks + 12 auxiliares)
  3. Grava `metricas_<nome>.json` em `_RESULTADOS\`
  4. Fecha o modelo sem salvar
- **Não usa** `revit.doc` (retorna None em batch); abre explicitamente

### `consolidar.py` (Python 3)

- Lê todos os `metricas_*.json`
- Lê os **pesos** da aba `ARC` do `01-AUDIT-BIM.xlsx`
- **Calibra a régua** por min-max relativo ao lote
- Gera `dashboard_bim.xlsx` com 3 abas:
  - `Detalhado` — cada check × modelo × nota
  - `Resumo` — ranking com status
  - `Não mensuráveis` — checks excluídos e motivo

### `exportar_bi.py` (Python 3)

- Mesmos dados do `consolidar.py`, mas em formato **star schema** para Power BI
- Gera **4 CSVs** em `_BI_EXPORT/<AAAA-MM>/`:
  - `bi_modelos.csv` — 16 colunas (dimensão de modelos com métricas técnicas)
  - `bi_checks.csv` — 14 linhas (dimensão de checks)
  - `bi_metricas.csv` — 154 linhas (fato: 14 checks × 11 modelos)
  - `bi_resumo.csv` — 11 linhas (fato agregado por modelo)
- **Encoding UTF-8 com BOM + separador `;` + decimal `,`** → compatível com Excel PT-BR e Power BI

### `dashboard.py` (Streamlit)

- Lê `dashboard_bim.xlsx`
- 4 visuais: KPIs, Ranking, Heatmap, Radar
- Atualiza automaticamente quando o Excel muda

---

## 5. Os 14 checks automáticos

Os **pesos** vêm da coluna E da aba `ARC` do `01-AUDIT-BIM.xlsx`.

| Row | Check | Peso | Tipo | O que mede |
|:---:|---|:---:|---|---|
| 5 | Nomenclatura do arquivo RVT | 1 | binário | Nome segue padrão `^(WD[_\d]\|\d+).*[-_](EX\|EP\|ARQ)[-_]?R\d+` |
| 6 | Elementos voando | 1 | menor melhor | Famílias fora do envelope vertical/horizontal do modelo |
| 8 | DWG referenciados | 2 | menor melhor | DWGs **importados** (não vinculados) |
| 9 | Modelagem no local | 2 | menor melhor | Famílias in-place |
| 10 | Níveis/vistas fora do padrão | 1 | menor melhor | Nomes sem prefixo padrão |
| 13 | Folhas fora do padrão | 2 | menor melhor | Numeração fora do padrão |
| 15 | Carimbos incompletos | 3 | menor melhor | Folhas com parâmetros de carimbo vazios |
| 16 | Famílias genéricas | 1 | menor melhor | Famílias na categoria "Modelo Genérico" |
| 18 | Tabelas fora do padrão | 3 | menor melhor | Tabelas de quantitativos sem nomenclatura |
| 23 | Total de tags | 1 | maior melhor | Identificadores aplicados (mais é melhor) |
| 25 | Avisos graves | 1 | menor melhor | Warnings do Revit |
| **75** | **Vínculos RVT** | **2** | menor melhor | `RevitLinkInstance` (links externos RVT) |
| **76** | **Vínculos CAD** | **1** | menor melhor | `CADLinkType` (CADs vinculados, não importados) |
| **77** | **Vistas sem template** | **2** | menor melhor | Vistas sem view template aplicado |

### Checks NÃO mensuráveis (avaliação humana)

Registrados na aba `Não mensuráveis` do Excel, não afetam a nota:

- Clash em ACC (row 7)
- Organização de folhas (row 11)
- Conteúdo das folhas (row 12)
- Preenchimento de Capa (row 14)
- Fidelidade às categorias (row 17)
- Identidade visual (row 19)
- Nomenclatura geral de famílias (row 20)
- Configurações de publicação (row 24)

### Checks excluídos por bug de API

- Ponto zero (row 21) — `ProjectLocation.Point` não mede o que queremos no Revit 2020
- Níveis fixados (row 22) — `Level.IsPinned` não exposto; fallback zerou o check

---

## 6. Métricas técnicas extraídas

Além dos 14 checks, o `script.py` extrai **7 métricas auxiliares** que vão para o Power BI (não afetam a nota):

| Métrica | Descrição | Onde aparece |
|---|---|---|
| `tamanho_arquivo_mb` | Tamanho do .rvt em disco | `bi_modelos.csv` |
| `qtd_familias_unicas` | Famílias distintas carregadas | `bi_modelos.csv` |
| `qtd_instancias_familia` | Total de FamilyInstances | `bi_modelos.csv` |
| `qtd_vinculos_rvt` | (duplicado como check 75) | `bi_modelos.csv` |
| `qtd_vinculos_cad` | (duplicado como check 76) | `bi_modelos.csv` |
| `qtd_worksets` | Worksets no projeto (-1 se falhar) | `bi_modelos.csv` |
| `qtd_vistas_sem_template` | (duplicado como check 77) | `bi_metricas.csv` |

---

## 7. Como interpretar as notas

### Régua de calibração (relativa ao lote)

O sistema usa **min-max relativo**, não absoluto:

| Situação | Nota |
|---|---|
| Melhor modelo do lote em um check | **10** |
| Pior modelo do lote em um check | **0** |
| Modelos intermediários | Distribuídos 0–10 |

**Importante:** nota 10 significa "o melhor **deste** lote", não "perfeito".

### Status final

| Nota final | Status | Cor |
|---|---|---|
| ≥ 8,0 | **Otimo** | Verde |
| 6,0–7,9 | **Bom** | Azul |
| 4,0–5,9 | **Regular** | Amarelo |
| < 4,0 | **Critico** | Vermelho |

### Cálculo

```
Nota Final = Σ(nota_check × peso) / Σ(pesos)
```

Checks com falha de coleta (`valor = -1`) são **ignorados**.

---

## 8. Manutenção

### 8.1 Ajustar a régua de notas

Edite `consolidar.py`, função `metrica_para_nota`. Depois:

```powershell
run_only_consolidar.bat
```

Regenera o Excel e os CSVs em segundos (sem abrir o Revit).

### 8.2 Ajustar os pesos dos checks

Edite a **coluna E** (PESO) da aba `ARC` do `01-AUDIT-BIM.xlsx`. Depois:

```powershell
run_only_consolidar.bat
```

### 8.3 Adicionar novos modelos

1. Coloque os `.rvt` em `MODELOS ATIVOS\`
2. Rode `run_all.bat`

O sistema processa **todos** os arquivos da pasta. Modelos já processados têm os JSONs sobrescritos.

### 8.4 Alterar o que é medido

1. Edite `script.py` no pushbutton
2. Limpe o cache do pyRevit:
   ```powershell
   Remove-Item "$env:APPDATA\pyRevit-Master\pyRevit\*.pyc" -Recurse -Force
   Remove-Item "$env:LOCALAPPDATA\Temp\pyRevit*" -Recurse -Force
   ```
3. Rode `run_all.bat`

### 8.5 Sincronizar `bim_paths.py` entre as duas pastas

O `bim_paths.py` existe em **dois lugares** (orquestrador e pushbutton). Sempre que editar o primeiro, copie para o segundo:

```powershell
Copy-Item "X:\...\PY_ORQUESTRACAO\bim_paths.py" `
          "X:\...\PYREVIT SCRIPTS\BIM.extension\BIM.tab\Analise.panel\Extrair.pushbutton\bim_paths.py" `
          -Force
```

### 8.6 Evolução temporal (histórico de análises)

O `exportar_bi.py` grava em subpastas por período (`_BI_EXPORT/2026-09/`, `_BI_EXPORT/2026-10/`, ...). Cada rodada em um mês novo **cria uma nova pasta** automaticamente.

No Power BI, você pode:
- Apontar para a **pasta raiz** `_BI_EXPORT/` com `Obter Dados → Pasta` (lê todos os meses)
- Ou apontar para a subpasta específica do mês atual

---

## 9. Solução de problemas

### "python não é reconhecido"

Feche e reabra o PowerShell. Se persistir, use o caminho completo:

```powershell
& "C:\Users\diego.bonadiman\AppData\Local\Python\pythoncore-3.14-64\python.exe" script.py
```

### "streamlit não é reconhecido"

Use o módulo em vez do executável:

```powershell
python -m streamlit run dashboard.py
```

### Revit abre mas o JSON não é gerado

Verifique os logs em ordem:

1. **`batch.log`** — mostra o comando chamado e o `exit code`
2. **`_debug_script.log`** — mostra passo-a-passo do `script.py` dentro do Revit

Procure por `[ERRO]` ou `Traceback`.

### Erros de encoding (BOM)

Se aparecer:
```
SyntaxError: invalid non-printable character U+FEFF
```

O arquivo foi salvo com BOM. Remova:

```powershell
python -c "p=r'caminho\do\arquivo.py'; s=open(p,'rb').read(); open(p,'wb').write(s.replace(b'\xef\xbb\xbf', b'', 1)); print('BOM removido')"
```

**Prevenção:** no VSCode, use sempre `UTF-8` (não `UTF-8 with BOM`). Confira no canto inferior direito.

### Modelo com nome acentuado (ex: `149_NANÔBE`)

Já tratado automaticamente: o `run_batch.py` copia para `C:\Temp\bim_ascii\` com nome sanitizado e apaga depois.

### Power BI não atualiza os dados

Verifique em ordem:

1. `Página Inicial → Transformar Dados → Configurações da Fonte de Dados` → confirme que aponta para `_BI_EXPORT\2026-09\` (não raiz)
2. `Página Inicial → Atualizar` (refresh)
3. Se colunas novas não aparecerem: `Transformar Dados → tabela → Atualizar Visualização`
4. Se persistir: exclua a tabela no Power Query e re-importe o CSV

### Valores com 100× o esperado (ex: `106748` em vez de `1067,48`)

O separador decimal está sendo lido como milhar. Corrija:

1. `Transformar Dados → bi_modelos`
2. Coluna `tamanho_mb` → clique no ícone do tipo → `Número Decimal`
3. Se persistir: `Página Inicial → Opções → Carregamento de Dados → Limpar Cache`

### `qtd_worksets = -1` em quase todos os modelos

Bug conhecido do IronPython 2.7 com `ICollection<WorksetId>.Count`. A métrica fica como informativa (não entra em nenhum check). Pode ser ignorada.

---

## 10. Changelog

### v3.0 (outubro/2026) — Versão atual

- **7 novas métricas** extraídas do Revit:
  - `tamanho_arquivo_mb`, `qtd_familias_unicas`, `qtd_instancias_familia`
  - `qtd_vinculos_rvt`, `qtd_vinculos_cad`, `qtd_worksets`, `qtd_vistas_sem_template`
- **3 novos checks** (rows 75, 76, 77): vínculos RVT, vínculos CAD, vistas sem template
- **`exportar_bi.py`** — geração de CSVs star schema para Power BI
- **`_BI_EXPORT/<período>/`** — versionamento automático por mês
- **Decimal com vírgula** nos CSVs (compatibilidade PT-BR)
- **Dashboard Power BI** com 5 páginas (Sumário, Executiva, Detalhamento, Análise Cruzada, Modelo Individual)
- **Drill-through** funcional por modelo

### v2.0 (setembro/2026)

- **Calibração relativa** (min-max) em vez de régua linear
- **`run_all.bat`** + 2 `.bat` auxiliares
- Correção de BOM em todos os scripts
- Fallback para nomes acentuados (`C:\Temp\bim_ascii\`)
- Fallback de abertura de modelo (`__models__` em vez de `revit.doc`)

### v1.0 (setembro/2026) — Versão inicial

- Pipeline básico: `run_batch.py`, `script.py`, `consolidar.py`, `dashboard.py`
- 11 checks automáticos (rows 5, 6, 8, 9, 10, 13, 15, 16, 18, 23, 25)
- Dashboard Streamlit com ranking, heatmap e radar

---

## Contatos

**Diego Bonadiman** — mantenedor do pipeline

Última atualização: outubro/2026
