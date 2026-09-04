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
  historico-tratamento.md            Histórico cronológico do tratamento
.claude/skills/adicionar-alimento/
  SKILL.md                           Skill para acrescentar alimentos à planilha
```

## A planilha

`planilha/plano-alimentar-pos-infarto.xlsx` tem 4 abas visíveis e 1 aba oculta
de apoio:

1. **Cuidados na Alimentação** — o que priorizar e reduzir, atenção especial a
   triglicerídeos/LDL, a lista de medicações da alta e os avisos de interação
   medicamento x alimento (Ticagrelor x toranja, Enalapril x potássio/sal
   light, etc.), e um aviso pendente sobre função renal (ainda a confirmar
   com o médico).
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
5. **Ref_Alimentos** (oculta) — tabelas auxiliares que alimentam as fórmulas
   da aba 4 (listas por refeição, sugestões de combinação). Não precisa
   mexer nela na mão.

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

- Função renal: possível proteinúria e quedas de TFG ainda não confirmadas
  pelo médico — a planilha não aplica restrição renal enquanto isso não for
  esclarecido.
- Lista de medicações da alta já foi incorporada (seção 1a de
  `referencia/dados-consulta-e-alimentacao.md` e aba "Cuidados na
  Alimentação") — falta ainda confirmar com o médico a conduta exata em caso
  de esquecimento de dose de cada medicamento.
