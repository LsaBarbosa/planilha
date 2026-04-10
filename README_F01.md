# F01 - Workbook Base & UI

Esta entrega inicia a **primeira feature da Sprint 1 (F01 Workbook Base & UI)**.

## Escopo coberto
- Criação/garantia das abas base na ordem correta:
  - Instructions
  - Study Index
  - Dimension Register
  - TEMPLATE
- Estrutura inicial de layout de cada aba
- Cabeçalho documental nas abas aplicáveis
- Rodapé documental nas abas aplicáveis
- Aba Instructions em inglês técnico
- Identidade visual inicial:
  - cabeçalhos em azul escuro com fonte branca
  - células de entrada em azul claro
  - células calculadas em cinza claro
  - gridlines desativadas
- Placeholder de logo corporativo

## Limites desta entrega
- Não implementa ainda:
  - Units global
  - validações numéricas
  - Add New Study
  - lookup de Dim. ID
  - cálculos
  - sincronização do Study Index
  - VeryHidden/proteção estrutural final

Esses itens pertencem às features seguintes do backlog.

## Arquivos
- `ThisWorkbook.cls`
- `modConfig.bas`
- `modSheetFormatting.bas`
- `modWorkbookBaseUI.bas`

## Como instalar
1. Abra o VBA Editor (`ALT + F11`).
2. Importe os módulos `.bas` e o `ThisWorkbook.cls`.
3. Em um workbook habilitado para macro (`.xlsm`), execute:
   - `modConfig.InstallFeatureF01`
4. Salve, feche e reabra o arquivo para validar o `Workbook_Open`.

## Observação sobre o logo
O arquivo da imagem corporativa não foi anexado. Por isso, a rotina insere um **placeholder visual**.
Quando o logo real for enviado, a rotina pode ser ajustada para carregar a imagem oficial.

## Próximo passo recomendado
Após validar a base visual e a navegação, o próximo desenvolvimento natural é a **F02 Configuração Global**.
