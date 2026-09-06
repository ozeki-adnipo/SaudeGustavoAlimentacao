---
name: adicionar-alimento
description: Acrescenta um novo alimento à planilha de alimentação pós-infarto (planilha/plano-alimentar-pos-infarto.xlsx). Use quando o usuário disser algo como "adicione X", "adiciona X na planilha", "inclui X na lista de alimentos" ou "atualiza X na planilha", onde X é um alimento ou bebida. Classifica o alimento, sugere porção/frequência/classificação e regenera a planilha.
---

# Adicionar alimento à planilha

Esta skill existe para que o usuário só precise digitar algo como **"adicione
sucrilhos"** e a planilha `planilha/plano-alimentar-pos-infarto.xlsx` seja
atualizada sozinha, na aba **"Alimentos e Quantidades"** — e passe a
aparecer também nos menus suspensos da aba **"Montar Refeição"** (nas
refeições marcadas no campo `refeicoes`) — mantendo a mesma formatação de
sempre.

## Contexto do paciente (usar para decidir a classificação)

- Pós-infarto agudo do miocárdio recente, dislipidemia mista: **LDL alto**
  (138 mg/dL; meta pós-IAM entre <40 e <55 mg/dL dependendo da diretriz —
  ver aba "Fontes e Referências"), **HDL baixo**, **triglicerídeos muito
  altos** (343 mg/dL) e colesterol total alto.
- Função renal com pendência ainda não confirmada pelo médico — por enquanto
  **não aplicar** restrição de sódio/potássio/proteína por causa disso, só o
  cuidado cardiovascular.
- Prioridade de cuidado, do mais importante para o menos: **açúcar e álcool**
  (por causa dos triglicerídeos) > **gordura saturada/trans** > **sódio** >
  carboidrato refinado em geral.
- Detalhes completos em `referencia/dados-consulta-e-alimentacao.md` e
  `referencia/historico-tratamento.md`, se precisar consultar algo específico.

### Medicações em uso (receita de alta, 03/09/2026) — SEMPRE checar interação

Todas de uso contínuo: **Aspirina Prevent 100mg** e **Ticagrelor 90mg** (dupla
antiagregação, 1 ano), **Rosucor 20mg** (rosuvastatina, 40mg/dia) + **Ezetimiba
10mg** (colesterol), **Maleato de Enalapril 5mg** (IECA — cuidado com pressão
baixa). Antes de classificar QUALQUER alimento novo, verifique se ele se encaixa
em algum destes pontos e, se sim, rebaixe a classificação e explique na
observação:

- **Toranja/grapefruit (fruta, suco ou qualquer preparo com toranja)** →
  `Evitar` sempre. Interage com o Ticagrelor (estudo publicado + a própria
  bula do Brilinta reconhecem o efeito, mesmo a bula dizendo que "não é
  esperado ser clinicamente relevante para a maioria" — o paciente toma 2
  antiagregantes, então a orientação conservadora é evitar). NÃO tem
  interação relevante conhecida com a Rosucor/rosuvastatina especificamente
  (via metabólica diferente) — o motivo de evitar é só o Ticagrelor.
- **Substitutos de sal / "sal light" (ricos em cloreto de potássio) e
  suplementos de potássio** → `Evitar` ou `Moderar` com nota de alerta.
  O Enalapril já tende a elevar o potássio; combinado com esses produtos o
  risco de hipercalemia aumenta, ainda mais com a função renal pendente.
  Alimentos naturalmente ricos em potássio (banana, batata, coco) continuam
  liberados em porção normal — o alerta é para excesso/concentrado, não para
  a fruta in natura.
- **Álcool, em qualquer forma** → `Evitar`. Além de piorar os triglicerídeos,
  soma-se ao risco de sangramento da dupla Aspirina + Ticagrelor.
- **Suplementos como "arroz de levedura vermelha" (red yeast rice) ou outros
  que se apresentem como "estatina natural"** → `Evitar`. Podem somar
  toxicidade com o Rosucor.

Se o alimento pedido não tiver relação com nenhum desses pontos, siga apenas o
raciocínio nutricional geral do passo 3 abaixo.

### Ao suspeitar de uma interação nova (não listada acima)

Se o alimento pedido puder interagir com alguma das 5 medicações e isso não
estiver coberto pelos 4 pontos acima (ex.: um alimento rico em vitamina K, se
algum dia entrar um anticoagulante; um suplemento ou fitoterápico específico),
**pesquise em fontes confiáveis antes de classificar** (WebSearch, se
disponível): priorize bula oficial (Anvisa/FDA/DailyMed), estudos revisados
por pares (PubMed/PMC), órgãos de saúde do governo (NIH) ou diretrizes de
sociedades médicas (SBC, AHA, ACC/ESC) — não confie só em conhecimento geral
para uma interação medicamentosa nova. Adicione a fonte encontrada à lista
`FONTES` no topo de `planilha/gerar_planilha.py` (ela alimenta a aba "Fontes e
Referências") e cite-a resumidamente na observação do alimento.

## Passo a passo

1. **Identifique o alimento** pedido pelo usuário (ex.: "sucrilhos",
   "requeijão", "vinho tinto").

2. **Classifique a categoria**, usando as categorias já existentes sempre que
   o alimento se encaixar em uma delas (senão, crie uma nova categoria):
   `Frutas`, `Grãos e Cereais`, `Proteínas`, `Laticínios`, `Bebidas`,
   `Doces e Sobremesas`, `Gorduras e Temperos`, `Vegetais e Verduras`.
   Se criar uma categoria nova, edite também `CATEGORY_ORDER` e, se fizer
   sentido, `CATEGORIA_COMBINA` no topo de `planilha/gerar_planilha.py`
   (senão a coluna "Sugestão" da aba "Montar Refeição" fica em branco para
   ela).

2a. **Defina em quais refeições o alimento se encaixa** (campo `refeicoes`,
    obrigatório — sem ele o alimento não aparece em nenhum menu suspenso da
    aba "Montar Refeição"): uma lista com um ou mais de `"Café da manhã"`,
    `"Almoço"`, `"Lanche"`, `"Jantar"` (grafia exata). Ex.: arroz/feijão só
    Almoço+Jantar; aveia só Café da manhã; água em todas.

2b. **Se o usuário disser a marca** (ex.: "esse é Polenghi", ou mandar foto de
    um rótulo com marca visível), preencha o campo opcional `marca` no JSON.
    Sem marca informada, deixe o campo de fora (ou `""`) — não invente uma
    marca. Se o usuário mandar uma foto de rótulo com a informação
    nutricional, use os valores reais do rótulo (sódio, gordura saturada,
    açúcar etc.) na observação em vez de estimar genericamente.

3. **Defina porção, frequência e classificação** (`Liberado` / `Moderar` /
   `Evitar`) com base em conhecimento nutricional geral e no contexto do
   paciente acima. Exemplos de raciocínio:
   - Alto teor de açúcar adicionado (cereal matinal açucarado, refrigerante
     comum, doces) → `Moderar` ou `Evitar`, com porção pequena e frequência
     baixa (ou "evitar").
   - Rico em gordura saturada/frituras (embutidos, frituras, queijos
     amarelos gordurosos) → `Moderar` ou `Evitar`.
   - Alimento in natura, rico em fibra/proteína magra/ômega-3, sem açúcar ou
     sódio excessivo → `Liberado`, com porção usual de uma refeição.
   - Na dúvida entre duas classificações, escolha a mais conservadora e
     explique o motivo na observação.

4. **Escreva uma observação** curta e objetiva. Sempre que a classificação for
   `Moderar` ou `Evitar`, ou houver incerteza nutricional relevante, inclua ao
   final da observação: *"Sugestão gerada por IA — confirmar com
   nutricionista/cardiologista."*

5. **Atualize `planilha/dados_alimentos.json`**:
   - Se o alimento já existir na lista (mesmo nome, ignorando maiúsculas
     /minúsculas, e mesma marca — ou ambos sem marca), **edite o item
     existente** em vez de duplicar. Se for a mesma comida mas o usuário deu
     um nome mais específico (ex.: "pão integral" → "pão de forma integral"),
     também é edição, não duplicata.
   - Senão, **acrescente um novo objeto** à lista, no formato (o campo
     `refeicoes` é obrigatório, ver passo 2a; `marca` é opcional, ver 2b):
     ```json
     {"categoria": "...", "alimento": "...", "marca": "...", "porcao": "...", "frequencia": "...", "classificacao": "Liberado|Moderar|Evitar", "observacao": "...", "refeicoes": ["Café da manhã", "Lanche"]}
     ```
   - **Nunca escreva nada no campo "Gostoso"** da aba 3 — é a avaliação
     pessoal do usuário (menu suspenso: Ruim/Normal/Bom/Muito bom), preenchida
     por ele diretamente na planilha, não algo que a IA decide. Ela não vive
     no JSON; `gerar_planilha.py` a preserva automaticamente entre execuções
     lendo o `.xlsx` anterior (função `load_gostoso_previo`).

6. **Regenere a planilha** rodando, a partir da raiz do repositório:
   ```bash
   python3 planilha/gerar_planilha.py
   ```
   Isso reconstrói o arquivo `.xlsx` inteiro (abas 1 e 2 continuam iguais; a
   aba 3 e a aba oculta `Ref_Alimentos` são remontadas a partir do JSON, e a
   aba 4 "Montar Refeição" passa a enxergar o alimento novo nos menus
   suspensos da(s) refeição(ões) marcada(s)), preservando toda a
   formatação/cores — inclusive os valores que o usuário já tiver preenchido
   na coluna "Gostoso" (o script lê o `.xlsx` anterior antes de sobrescrever).
   Se `openpyxl` não estiver instalado, rode `pip install openpyxl` antes.

   **Tente recalcular com LibreOffice** (`scripts/recalc.py` da skill xlsx)
   depois de gerar o arquivo, para confirmar que as fórmulas da aba "Montar
   Refeição" não geram erro. Se o recálculo travar/der timeout no ambiente
   (aconteceu na criação desta aba — parece ser uma limitação do sandbox, não
   do arquivo), não insista mais que uma tentativa extra: avise o usuário que
   não foi possível confirmar automaticamente, mas que a lógica das fórmulas
   foi validada por simulação em Python e que elas calculam normalmente ao
   abrir o arquivo num programa de verdade (Excel, Google Sheets, LibreOffice
   Desktop).

7. **Confirme para o usuário** o que foi feito: alimento, categoria, porção,
   frequência, classificação e o motivo da classificação — e pergunte se ele
   quer ajustar algum valor (ex.: porção diferente da sugerida).

## Observações importantes

- Nunca edite o `.xlsx` diretamente célula por célula — sempre passe pelo
  `dados_alimentos.json` + `gerar_planilha.py`, para não quebrar a
  formatação nem duplicar linhas.
- Se o usuário pedir para remover um alimento, remova o objeto correspondente
  do JSON e rode o script de novo.
- Se o usuário pedir várias alterações de uma vez, você pode editar o JSON
  várias vezes e rodar `gerar_planilha.py` só uma vez no final.
- **Sincronizar com o Google Sheets é um passo à parte, só quando pedido.**
  Por padrão esta skill só regenera o `.xlsx` local e o envia pelo chat — não
  mexe no Google Sheets automaticamente (o processo de subir um arquivo novo
  no Drive é manual e custoso, ver motivo em `planilha/SHEETS_SYNC.md`). Se o
  usuário pedir explicitamente para atualizar/sincronizar o Sheets também
  (ex.: "atualiza o sheets", "sincroniza com a nuvem"), siga o procedimento
  completo descrito em `planilha/SHEETS_SYNC.md` (baixar o Sheets atual antes
  de regenerar, para não perder o "Gostoso" preenchido lá; depois subir um
  Sheets novo e atualizar o link registrado nesse arquivo).
