# Saúde Gustavo — Alimentação pós-infarto

Repositório dedicado ao plano alimentar e cuidados de alimentação pós-infarto
agudo do miocárdio (25/08/2026). Separado do histórico médico geral para
facilitar a consulta do dia a dia.

## Estrutura

```
planilha/
  plano-alimentar-pos-infarto.xlsx   Planilha principal (ver abas abaixo)
  dados_alimentos.json               Fonte de verdade da lista de alimentos
  gerar_planilha.py                  Script que gera/regenera o .xlsx
referencia/
  dados-consulta-e-alimentacao.md    Dados clínicos e perguntas para a consulta
  historico-tratamento.md            Histórico cronológico + interpretação por diretrizes
  documentos/receita-alta-03-09-2026.jpg   Imagem da receita de alta
.claude/skills/adicionar-alimento/
  SKILL.md                           Skill para acrescentar alimentos à planilha
```

## A planilha

`planilha/plano-alimentar-pos-infarto.xlsx` tem 5 abas visíveis e 1 aba oculta
de apoio:

1. **Cuidados na Alimentação** — o que priorizar e reduzir, atenção especial a
   triglicerídeos/LDL (com as diferentes metas possíveis, ver fontes abaixo),
   a lista de medicações da alta e os avisos de interação medicamento x
   alimento (Ticagrelor x toranja, Enalapril x potássio/sal light, etc.), e um
   aviso pendente sobre função renal (ainda a confirmar com o médico).
2. **Plano Geral - Restaurante** — modelo de prato, o que pedir/evitar fora de
   casa, perguntas úteis ao garçom e exemplos de pratos que funcionam.
3. **Alimentos e Quantidades** — lista de alimentos comuns por categoria, com
   porção, frequência e classificação (Liberado / Moderar / Evitar).
4. **Montar Refeição** — monta café da manhã, almoço, lanche e jantar
   escolhendo alimentos em menus suspensos (já filtrados por refeição — não
   aparece arroz/feijão no café da manhã, por exemplo). Mostra sozinha a
   categoria, a porção sugerida e a classificação de cada item escolhido,
   sugere um alimento que combina bem com o item anterior, e sinaliza
   "⚠ Conflito" quando dois itens problemáticos (ou um "Evitar") caem na
   mesma refeição. Ver limitações conhecidas no rodapé da própria aba.
5. **Fontes e Referências** — de onde veio cada afirmação médica usada nas
   abas acima (bulas oficiais/Anvisa, estudos revisados por pares, NIH,
   diretrizes da SBC/AHA/ACC-ESC/KDIGO/ESUR/ESVS/ADA-SBD), com uma coluna
   "Área" (Alimentação / Clínico geral) e link clicável para cada fonte. A
   mesma lista, em formato mais fácil de copiar, está em
   `referencia/historico-tratamento.md`, seção "Referências médicas".
6. **Ref_Alimentos** (oculta) — tabelas auxiliares que alimentam as fórmulas
   e os menus suspensos da aba 4 (listas por refeição, sugestões de
   combinação). Não precisa mexer nela na mão.

A lista de alimentos é gerada a partir de `planilha/dados_alimentos.json` pelo
script `planilha/gerar_planilha.py` — isso garante que a formatação (cores,
categorias, larguras, fórmulas) fique sempre consistente ao adicionar novos
itens. Cada alimento no JSON tem um campo `refeicoes` que diz em quais das 4
refeições ele pode aparecer.

**Sobre as fórmulas da aba "Montar Refeição":** o LibreOffice headless deste
ambiente não conseguiu recalcular o arquivo automaticamente ao criar esta aba
(travou repetidamente, mesmo num arquivo mínimo de teste — parece ser uma
limitação do ambiente, não do arquivo). A lógica das fórmulas foi validada por
simulação em Python antes de publicar. Ao abrir o arquivo pela primeira vez no
Excel, Google Sheets ou LibreOffice Desktop, as fórmulas calculam normalmente
(comportamento padrão desses programas) — se notar algo estranho na primeira
abertura, avise para eu corrigir.

**Sobre os menus suspensos da coluna "Alimento":** eles usam "nomes definidos"
(`Lista_CafeDaManha`, `Lista_Almoco`, `Lista_Lanche`, `Lista_Jantar`, definidos
em `gerar_planilha.py`) em vez de apontar direto para a aba `Ref_Alimentos` —
o Excel não permite uma lista suspensa referenciar diretamente um intervalo em
outra aba, só um nome definido. Se algum dia o menu parecer vazio ou quebrado
de novo, é o primeiro lugar para checar.

## Como acrescentar um alimento novo

Basta pedir, em uma conversa com o Claude neste repositório:

> adicione sucrilhos

A skill `adicionar-alimento` (em `.claude/skills/adicionar-alimento/SKILL.md`)
classifica o alimento (categoria, porção, frequência, Liberado/Moderar/Evitar),
atualiza `dados_alimentos.json` e regenera a planilha automaticamente.

Toda sugestão de porção/classificação é uma orientação inicial baseada em
conhecimento nutricional geral — vale confirmar com nutricionista ou
cardiologista, especialmente para itens marcados como "Moderar" ou "Evitar".

## Pendências conhecidas (ver `referencia/`)

`referencia/historico-tratamento.md` foi atualizado (04/09/2026) com uma reavaliação clínica
mais completa, interpretando cada achado à luz de diretrizes médicas (KDIGO, ESUR, ESC,
ACC/AHA, ADA/SBD, ESVS/AHA-ASA) — ver as seções "Interpretação à luz de diretrizes médicas" e
"Informações que ainda faltam para uma avaliação completa" nesse arquivo. Nenhuma dessas
interpretações substitui a confirmação médica. Pendências que seguem em aberto:

- **Função renal:** possível proteinúria (exame certo a pedir: ACR/RAC) e uma elevação de
  creatinina que, por critério objetivo (KDIGO), já preenche Lesão Renal Aguda estágio 1,
  compatível com nefropatia por contraste (padrão temporal ESUR) — a planilha não aplica
  restrição renal enquanto isso não for confirmado/descartado pelo médico.
- Lista de medicações da alta já foi incorporada (seção "Medicações da alta hospitalar" de
  `referencia/historico-tratamento.md` e aba "Cuidados na Alimentação") — falta ainda
  confirmar com o médico a conduta exata em caso de esquecimento de dose de cada medicamento.
- Confirmar se o encaminhamento para reabilitação cardíaca (recomendação Classe 1) já foi
  feito na alta.
