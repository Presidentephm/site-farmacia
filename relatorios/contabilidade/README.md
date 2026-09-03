# Relatório de folha para a contabilidade

## Arquivo

`RELATORIO_HOLERITE_AGOSTO_2026_pgto_05-09-2026.xlsx`

- **Competência:** agosto/2026 (vendas de 01/08/2026 a 31/08/2026)
- **Pagamento:** 05/09/2026
- **Fonte das comissões/incentivos:** InovaFarma — Relatório detalhado de comissão de
  vendedor, extraído em 03/09/2026

## Abas

| Aba | Conteúdo |
|---|---|
| CAPA | Identificação, instruções de uso e lista de pendências |
| HOLERITE AGO.26 | Relatório principal: rubricas × funcionários, com proventos, descontos e líquido |
| LISTA CONTABILIDADE | Os mesmos lançamentos em formato de lista (funcionário, rubrica, tipo, valor) |
| BASE INOVAFARMA AGO.26 | Apuração por vendedor: venda bruta, descontos, venda líquida, comissão e incentivos |
| JULHO.26 REVISADO | Folha de julho/2026 reorganizada e conferida (valores pagos preservados) |
| PARÂMETROS | Calendário do mês, fator do DSR e valores fixos |

## Convenção de cores

- Célula amarela com texto azul: valor digitado — preencher/conferir
- Célula laranja: valor repetido de julho/2026 — conferir se ainda vale
- Texto verde: valor vinculado à aba BASE INOVAFARMA AGO.26
- Célula verde com texto preto: resultado de fórmula — não alterar

## Filtros

As abas HOLERITE AGO.26, JULHO.26 REVISADO, LISTA CONTABILIDADE e
BASE INOVAFARMA AGO.26 têm filtro na linha de cabeçalho e painéis
congelados (o cabeçalho e a coluna de rubrica ficam fixos ao rolar).

## Definições da empresa (agosto/2026)

- AGNOR: mantida a comissão fixa de R$ 2.000,00 — o complemento sobre o
  apurado é calculado automaticamente a partir da aba PARÂMETROS
- JOEL: férias em agosto/2026 — falta informar o período
- CAMILA: contratada recentemente; agosto é o primeiro mês completo
- DEAN: não recebe comissão sobre vendas — rubrica zerada na folha

## Demonstrativo individual

A aba GANHOS DO COLABORADOR tem uma lista suspensa com os nomes: ao
escolher um colaborador, a tabela mostra os ganhos e descontos dele
(puxados da aba HOLERITE AGO.26) e o resumo de vendas do mês no
InovaFarma. A aba já está configurada para imprimir em uma página.

## Férias

- JOEL: acordo interno de saída em 04/08 e retorno em 04/09; o recibo de
  férias saiu como 01 a 31/08/2026 e foi pago no início de
  agosto — não se repetem neste holerite; salário e desconto de
  vale-transporte zerados no mês

## Lojas

| Loja | Arquivo |
|---|---|
| Arraial | `RELATORIO_HOLERITE_AGOSTO_2026_pgto_05-09-2026.xlsx` |
| Centro | `RELATORIO_HOLERITE_AGOSTO_2026_CENTRO_pgto_05-09-2026.xlsx` |
| Trancoso | `RELATORIO_HOLERITE_AGOSTO_2026_TRANCOSO_pgto_05-09-2026.xlsx` |

As três seguem a mesma estrutura de abas. Diferenças por loja:

- **Trancoso**: a comissão paga é o dobro da apurada no InovaFarma. O
  multiplicador fica na aba PARÂMETROS e a linha de comissão já sai
  multiplicada; os incentivos ficam pelo valor apurado
- **Trancoso**: UILLIAN é o gerente e não recebe comissão
- **Centro**: ARIANE é a gerente e não recebe comissão

## Definições de agosto/2026 por loja

**Centro**
- ARIANE e PEDRO não recebem comissão
- GENECIR: férias de 01 a 30/08 — salário zerado; falta confirmar o dia 31/08
- SERGIO (diária R$ 150,00) e ANA CELIA (diária R$ 100,00) são folguistas:
  ficam na aba FOLGUISTAS, para o contas a pagar, fora do holerite
- GILSON, UILLIAN e FABIANA não recebem comissão pela loja
- SARA: o saldo dela no Centro (R$ 2,65) foi somado à folha do Arraial

**Trancoso**
- Comissão paga em dobro; incentivos simples, sem dobrar
- Julho foi lançado por média; a diferença gerada pela comissão dobrada
  será reduzida nas premiações

## Apontamento do ponto (aba PONTO AGO.26)

Cada arquivo tem uma aba com as ocorrências de ponto do mês. Os totais
entram sozinhos nas linhas HORAS EXTRAS e ADICIONAL NOTURNO do holerite
(via SUMIFS), então essas duas rubricas não devem ser digitadas à mão.

Critérios, com os parâmetros na aba PARÂMETROS:

- salário-hora = salário base ÷ 220
- hora extra = salário-hora × 1,5
- adicional noturno = salário-hora × 20%
- feriado trabalhado sem folga = 1 salário-dia a mais (salário ÷ 30)

## Competência setembro/2026 (pagamento 05/10/2026)

Arquivos `RELATORIO_HOLERITE_SETEMBRO_2026_<LOJA>_pgto_05-10-2026.xlsx`,
gerados por `gerar_relatorio_setembro_2026.py`.

São a folha já montada com o que se sabe hoje: salários, rubricas fixas,
parâmetros do mês e as ocorrências conhecidas. Comissões e incentivos
ficam zerados até o relatório do InovaFarma de setembro ser extraído no
início de outubro e lançado na aba BASE INOVAFARMA SET.26.

Calendário do DSR em setembro: 4 domingos mais o feriado de 07/09
(segunda-feira) ÷ 25 dias úteis = fator 0,200000.

Ocorrências já registradas:

- JOEL (Arraial): férias até 03/09, volta em 04/09 — salário proporcional
  a 27 dias (R$ 1.458,90)
- TAMILES (Trancoso): atestado de 02 a 08/09, abonado, sem desconto
- GENECIR (Centro): voltou das férias, salário integral

## Resumo consolidado

`RESUMO_3_LOJAS_AGOSTO_2026_pgto_05-09-2026.xlsx` reúne as três lojas
numa página só, por funcionário, com totais por loja e total geral. É o
arquivo para mandar à contabilidade junto com os três detalhados.
Gerado por `gerar_resumo_3_lojas.py`.

## Férias registradas

| Quem | Loja | Período | Situação |
|---|---|---|---|
| JOEL | Arraial | 01 a 31/08/2026 | recibo pago no início de agosto |
| GENECIR | Centro | 01 a 30/08/2026 | recibo pago em agosto |
| RENALDO | Centro | 01 a 30/09/2026 | entra na folha de setembro |
| VALDICK | Trancoso | 01 a 30/09/2026 | entra na folha de setembro |

Quem passa o mês inteiro de férias fica com salário zerado no holerite e
sem desconto de vale-transporte; férias e 1/3 são pagos em recibo próprio
e ficam registrados na linha FÉRIAS PAGAS À PARTE, fora do líquido.

## Média de comissões para férias

`MEDIA_COMISSOES_12_MESES_FERIAS.xlsx` (gerado por
`gerar_media_12_meses.py`) tem uma aba por colaborador — VALDICK,
RENALDO e MANOEL — com os 12 meses a preencher, a soma e a média mensal
calculadas por fórmula, mais o passo a passo de como extrair cada mês na
tela Comissão de Vendedores do InovaFarma.

Janelas de 12 meses: VALDICK e RENALDO de 01/09/2025 a 31/08/2026
(férias em 01 a 30/09/2026); MANOEL de 01/05/2025 a 30/04/2026 (férias
gozadas em maio/2026).
