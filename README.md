# Tratamento BDA + Ajuste de Share

## Navegação

- `/` abre a página inicial com BDA LATAM e BDA GOL.
- `/latam` abre o programa atual de tratamento e ajuste de share.
- `/gol` abre o tratamento GOL, conferência IATA FINAL 255, relatórios e ajuste de share.
- Os filtros de CUPOM, TIPO, CONSOLIDADA 2 e DATA_EMISSAO são painéis suspensos; em telas menores, abrem dentro do próprio cartão.

## Novidades da V4

- Filtros redesenhados com cartões, busca, checkboxes e contador de selecionados.
- CUPOM, TIPO, CONSOLIDADA 2 e DATA_EMISSAO continuam controlando exatamente o universo do share.
- Explicação em linguagem simples para cada pergunta do ajuste de share.
- Ao exportar o ZIP, é gerado `RELATORIO_INICIAL_SHARE_E_FORMATACAO.pdf` com:
  - share JJ / G3 / AD;
  - total BRL;
  - período, linhas e bilhetes;
  - filtros usados;
  - auditoria da formatação;
  - exclusões realizadas;
  - tratamento ART/MAX.
- Se a meta de LATAM for igual ao share existente, nenhum bilhete é retirado.
- Se a meta for menor que o share atual, o processamento é bloqueado com uma mensagem explicativa.
- A seleção para ajuste é feita no universo filtrado, mas a exclusão física continua sendo por BILHETE COMPLETO.

## Instalação

```bash
python -m venv .venv
```

Windows:

```bash
.venv\Scripts\activate
pip install -r requirements.txt
python app.py
```

Abra no navegador:

```text
http://127.0.0.1:5000
```

## ZIP gerado

- `SHARE_ATUAL_E_AUDITORIA.pdf` e `SHARE_ATUAL_E_AUDITORIA.txt` na raiz do ZIP, com os indicadores anteriores ao ajuste de share e a auditoria da formatação.
- `01_ORIGINAL/`
- `02_FORMATACAO/`
  - `BASE_FORMATADA.csv`
  - `LINHAS_EXCLUIDAS_FORMATACAO.csv`
- `03_WEST/`
  - `WEST (...) COM CONSOLIDADA.csv`
  - `WEST (...).csv`
- `04_AJUSTE_SHARE/`
  - `LINHAS_EXCLUIDAS_AJUSTE_SHARE.csv`
  - `RELATORIO_SHARE.xlsx`
  - `RELATORIO_INICIAL_SHARE_E_FORMATACAO.pdf`
- `05_FINAL/`
  - `BASE FINAL 50 COLUNAS.csv`
  - arquivo nomeado pelo período, por exemplo `01 A 08.csv`
- `RESUMO_PROCESSAMENTO.txt`

## BDA GOL

1. Envie a base GOL original em CSV ou XLSX. A tela mostra os IATAs CORP sugeridos pelo tratador V4 para editar, excluir ou adicionar.
2. Confira a auditoria da formatação. Para vendas GOL em ART/MAX, o sistema identifica na coluna `IATA_AGENCIA` os valores que terminam em `255`. Só as vendas GOL em ART/MAX sem esse final recebem `BASE_VENDA=CGB`, `CONSOLIDADA_1=43` e `CONSOLIDADA_2=ARARAUNA TURISMO`. As outras vendas mantêm sua base. O IATA permanece numérico.
3. Envie opcionalmente uma base de outro período. O sistema confere se ela já está formatada; se não estiver, aplica o tratamento GOL e informa quantas linhas foram formatadas e excluídas. O relatório completo só aparece quando essa base é enviada.
4. Escolha se deseja análise por ADVP e se a divisão vale para todos os relatórios ou somente CONS, CORP ou GOV; CONS é o padrão. Antes do ajuste, altere os filtros CUPOM, TIPO, BASE de cada tipo de relatório e DATA_EMISSAO. Os títulos mostram o intervalo real de emissão de cada relatório.
5. Se a base complementar não tiver DATA_EMISSAO, envie também no campo Origem das datas um CSV/XLSX que a contenha. O sistema só recupera a data quando todas as vendas correspondem por bilhete, cupom, companhia, IATA, valor e data de embarque.
6. O ZIP GOL contém GOL_RELATORIOS.pdf com os indicadores do tratamento e os relatórios originais e ajustados, além das bases CSV e do arquivo com dinâmicas.
7. Informe metas GOL para cada relatório no período completo, na semana e/ou por ADVP; marque os dias da semana que devem ficar fora do ajuste. Somente bilhetes LATAM e AZUL inteiramente pertencentes à semana original podem ser excluídos, preservando a LATAM em média pelo menos 3 pontos percentuais à frente da GOL e limitando a queda da AZUL a 2 pontos percentuais. Se ao alterar o ADVP curto estipulado o share geral não for alcançado, o ADVP longo também é ajustado. A tela mostra qualquer meta ou margem não atingida.
8. Baixe o ZIP. O pacote traz o arquivo principal ajustado (`GOL PRINCIPAL.csv` / `GOL FINAL.csv`) com as vendas do período atual, e a base completa ajustada (`GOL BASE COMPLETA.csv`) contendo a união dos demais períodos + período atual, ambos no modelo operacional de 20 colunas GOL (separador `;`, datas `dd/mm/aaaa`, vírgula decimal e IATA numérico). O ZIP traz também os originais enviados, a base formatada, as auditorias e a planilha com dinâmicas e relatórios.
