---
name: adicionar-alimento
description: Acrescenta um novo alimento à planilha de alimentação pós-infarto (planilha/plano-alimentar-pos-infarto.xlsx). Use quando o usuário disser algo como "adicione X", "adiciona X na planilha", "inclui X na lista de alimentos" ou "atualiza X na planilha", onde X é um alimento ou bebida. Classifica o alimento, sugere porção/frequência/classificação e regenera a planilha.
---

# Adicionar alimento à planilha

Esta skill existe para que o usuário só precise digitar algo como **"adicione
sucrilhos"** e a planilha `planilha/plano-alimentar-pos-infarto.xlsx` seja
atualizada sozinha, na aba **"Alimentos e Quantidades"**, mantendo a mesma
formatação de sempre.

## Contexto do paciente (usar para decidir a classificação)

- Pós-infarto agudo do miocárdio recente, dislipidemia mista: **LDL alto**
  (meta pós-IAM <50 mg/dL), **HDL baixo**, **triglicerídeos muito altos**
  (343 mg/dL) e colesterol total alto.
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
  `Evitar` sempre. Interage com o Ticagrelor (aumenta o nível do remédio no
  sangue e o risco de sangramento).
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

## Passo a passo

1. **Identifique o alimento** pedido pelo usuário (ex.: "sucrilhos",
   "requeijão", "vinho tinto").

2. **Classifique a categoria**, usando as categorias já existentes sempre que
   o alimento se encaixar em uma delas (senão, crie uma nova categoria):
   `Frutas`, `Grãos e Cereais`, `Proteínas`, `Laticínios`, `Bebidas`,
   `Doces e Sobremesas`, `Gorduras e Temperos`, `Vegetais e Verduras`.

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
     /minúsculas), **edite o item existente** em vez de duplicar.
   - Senão, **acrescente um novo objeto** à lista, no formato:
     ```json
     {"categoria": "...", "alimento": "...", "porcao": "...", "frequencia": "...", "classificacao": "Liberado|Moderar|Evitar", "observacao": "..."}
     ```

6. **Regenere a planilha** rodando, a partir da raiz do repositório:
   ```bash
   python3 planilha/gerar_planilha.py
   ```
   Isso reconstrói o arquivo `.xlsx` inteiro (abas 1 e 2 continuam iguais; a
   aba 3 é remontada a partir do JSON), preservando toda a formatação/cores.
   Se `openpyxl` não estiver instalado, rode `pip install openpyxl` antes.

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
