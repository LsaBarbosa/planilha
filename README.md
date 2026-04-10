# Planilha - Macros VBA

Este repositório contém uma base funcional de macros VBA para gestão simples de estudos em Excel.

## Como usar

1. Abra o arquivo Excel habilitado para macros (`.xlsm`).
2. Importe os módulos `.bas` deste repositório no Editor VBA.
3. Execute a macro `Main` do módulo `modWorkbookBaseUI`.

## Fluxo principal

- `InitializeWorkbook`: cria/garante as planilhas `Config`, `Studies` e `Index`.
- `CreateStudy`: cria estudos com ID único.
- `RecalculateStudyScores`: calcula score dos estudos ativos.
- `SyncIndexSheet`: atualiza métricas consolidadas.
- `RunReleaseChecks`: executa validação básica de estrutura.

## Convenções

- Todos os módulos usam `Option Explicit`.
- Status válidos: `ACTIVE`, `INACTIVE`.
- Data/hora registrada no formato `yyyy-mm-dd hh:nn:ss`.
- IDs de estudo são únicos; inclusão duplicada gera erro.
