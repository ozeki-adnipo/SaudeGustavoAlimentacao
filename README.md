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

`planilha/plano-alimentar-pos-infarto.xlsx` tem 3 abas:

1. **Cuidados na Alimentação** — o que priorizar e reduzir, atenção especial a
   triglicerídeos/LDL, e avisos pendentes (função renal e medicação, ainda a
   confirmar com o médico).
2. **Plano Geral - Restaurante** — modelo de prato, o que pedir/evitar fora de
   casa, perguntas úteis ao garçom e exemplos de pratos que funcionam.
3. **Alimentos e Quantidades** — lista de alimentos comuns por categoria, com
   porção, frequência e classificação (Liberado / Moderar / Evitar).

A lista de alimentos é gerada a partir de `planilha/dados_alimentos.json` pelo
script `planilha/gerar_planilha.py` — isso garante que a formatação (cores,
categorias, larguras) fique sempre consistente ao adicionar novos itens.

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
- Lista de medicações da alta ainda não preenchida — pode haver interações
  entre medicamentos e alimentos (ex.: anticoagulante e vitamina K) a revisar
  assim que a lista estiver disponível.
