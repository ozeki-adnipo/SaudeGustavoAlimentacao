# Sincronização com o Google Sheets

## Por que isso existe

O usuário pediu um lugar "na nuvem" para editar a planilha junto comigo, sem
precisar que eu reenvie o arquivo pelo chat toda vez. A opção escolhida foi o
Google Sheets.

## Método atual: importação manual pelo usuário (padrão desde 13/09/2026)

Depois de tentar por várias sessões o upload automático via
`mcp__Google_Drive__create_file` (ver seção "Método antigo" abaixo, mantida
só como histórico), ficou comprovado que **não é confiável**: o conteúdo do
`.xlsx` precisa ser colado como texto base64 (~40-45 mil caracteres) direto
num parâmetro de ferramenta — não há como referenciar um caminho de arquivo
local — e reproduzir um texto desse tamanho sem erro se mostrou pouco
confiável mesmo com verificação cuidadosa (chegou a produzir o **mesmo erro
de forma idêntica em duas tentativas separadas**, ou seja, não é só acaso).

**O fluxo agora é:**

1. Eu regenero e envio o `.xlsx` pelo chat normalmente (skill
   `adicionar-alimento`), como sempre.
2. Quando o usuário quiser que a nuvem também seja atualizada, ele mesmo
   importa esse arquivo no Sheets existente:
   - Abrir o link do Sheets (ver "Link atual" abaixo)
   - **Arquivo > Importar > Fazer upload** > selecionar o `.xlsx` recebido
   - Escolher **"Substituir planilha"** (não "Criar nova planilha" nem
     "Inserir novas abas" — essas trocam o link ou duplicam abas)
   - **Importar dados**

   Essa opção tem uma vantagem sobre o upload automático: **mantém sempre o
   mesmo link**, em vez de gerar um novo a cada sincronização.

3. **Risco a avisar sempre que o usuário for importar:** a importação
   *substitui a planilha inteira* pelo conteúdo do `.xlsx` local. Qualquer
   valor que o usuário tenha preenchido na coluna **"Gostoso"** (ou
   redimensionamento de coluna/linha) diretamente no Sheets **depois da
   última vez que eu baixei o Sheets para regenerar localmente** seria
   perdido nessa substituição — porque o arquivo local só conhece o que
   estava lá na última sincronização. Antes de instruir a importação,
   pergunte se o usuário mexeu em "Gostoso" no Sheets recentemente; se sim,
   **confirme primeiro que o fileId em "Link atual" ainda resolve**
   (`mcp__Google_Drive__get_file_metadata` — em 13/09/2026 o fileId então
   registrado **não resolvia mais**, ver seção "Link atual"; não repita o
   erro de assumir que o campo está certo sem checar) e só então baixe o
   Sheets atual (`mcp__Google_Drive__download_file_content` +
   `GOSTOSO_SHEETS_EXPORT=... python3 planilha/gerar_planilha.py`, mesmo
   mecanismo do método antigo, só que agora alimentando o `.xlsx` que será
   enviado pelo chat, não um upload direto) para fundir esses valores antes
   de gerar o arquivo que o usuário vai importar. Se o fileId não resolver,
   pare e peça o link atual ao usuário antes de prosseguir — não presuma
   qual "Gostoso" está valendo.

4. Depois de o usuário confirmar que importou, **atualizar a data em "Link
   atual" abaixo** (o link normalmente não muda, só a data/descrição do que
   foi incluído) e commitar junto com as mudanças de alimento.

## Link atual

- **Título:** Plano Alimentar Pós-Infarto — Gustavo Ozeki
- **Link:** https://docs.google.com/spreadsheets/d/12gqsq1bewjfHwMBZUGunugg2RDUSpkzUfvq9Y1AvAZg/edit
- **fileId:** `12gqsq1bewjfHwMBZUGunugg2RDUSpkzUfvq9Y1AvAZg`
- **Última importação bem-sucedida: 13/09/2026.** Nesse mesmo dia o link
  chegou a dar "not found" (usuário tinha apagado sem querer, depois
  restaurou — ver histórico completo no fim desta seção), e a primeira
  tentativa de importar não chegou a se efetivar por causa disso. Depois
  de restaurado, o usuário conseguiu importar pelo navegador do celular
  (precisou ativar "Versão para computador" no Chrome, já que o app do
  Sheets não tem a opção Arquivo > Importar). **Conferido via
  `read_file_content`: o Sheets agora tem** sashimi de salmão, sushi de
  salmão, creme de tahine, kare com lombo de porco, cream cheese light e
  os 12 vegetais novos — tudo presente. "Gostoso" = "Manga: Bom"
  continua preservado.
- <details><summary>Histórico do susto de 13/09 (apagar/restaurar)</summary>

  Mais cedo em 13/09 este fileId deu "Requested entity was not found" —
  o usuário confirmou que provavelmente tinha apagado a planilha sem
  querer. Pediu para eu checar se tinha sido restaurada, e sim:
  `get_file_metadata` voltou a funcionar normalmente no mesmo fileId, com
  o mesmo título e `mimeType` de Google Sheet nativo de sempre — mas com
  o conteúdo ainda no estado de 07/09 (sem os 4 alimentos/mudanças mais
  recentes), confirmando que a restauração trouxe de volta uma versão
  anterior à tentativa de importação daquele dia. A importação bem
  sucedida (parágrafo acima) só aconteceu depois disso.
  </details>

### Arquivo solto encontrado e removido (13/09/2026)
Durante a checagem do link atual apareceu um arquivo separado
`plano-alimentar-pos-infarto.xlsx` (fileId `1zKG4qAjrdZxUNaPilLeHkIywWKXgeN2V`)
— um `.xlsx` cru solto no Drive (não um Google Sheet nativo), criado às
05:23 de 13/09/2026, com conteúdo batendo com a versão do commit `a7348f9`
(sashimi/sushi/tahine/kare, sem cream cheese/vegetais). Provavelmente um
upload avulso feito sem querer durante a mesma confusão da exclusão. Não era
o link oficial — o usuário confirmou e o arquivo foi movido para a lixeira
do Drive (`trash_file`) no mesmo dia, para não confundir com o Sheets de
verdade.

### Links anteriores (não recebem mais atualizações)
- `1-Mz3ue4QsSVytQIuL9MJUFnLHGdwp1VmS2wcleT5aDc` — criado em 06/09/2026 (era
  da época do upload automático, substituído pelo link acima em 07/09/2026).
  Continua retornando "not found" em `get_file_metadata` (checado em
  13/09/2026) — diferente do link atual, este não voltou; presume-se
  excluído de vez.

## Quando sincronizar

Não é automático a cada alimento adicionado. Para pedidos comuns de
"adicione X" sem menção ao Sheets, o fluxo normal continua sendo só
regenerar o `.xlsx` local e enviar pelo chat (skill `adicionar-alimento`),
sem tocar no Drive. Só quando o usuário pedir explicitamente para sincronizar
(ou disser que já importou, como confirmação) é que este documento entra em
jogo — para lembrar do aviso do "Gostoso" e atualizar a data acima.

---

## Método antigo (upload automático via API — não usar mais)

Mantido aqui só como registro histórico de por que foi abandonado, caso
algum dia o conector do Google Drive ganhe uma forma confiável de escrever
conteúdo grande (ex.: um parâmetro que aceite caminho de arquivo em vez de
texto inline, ou um endpoint de upload em chunks).

**Nota de auditoria (13/09/2026):** o relato abaixo é uma observação direta
feita durante a própria sessão em que o método foi abandonado (tool calls de
diagnóstico mostraram o mesmo tamanho de arquivo errado — 6318 bytes em vez
de 10500 — em duas tentativas separadas de subir o mesmo trecho). Essas
tentativas de diagnóstico foram feitas em arquivos temporários no Drive
(apagados depois) e não geram commit no git, então **não há como auditar
esse relato a partir do histórico do repositório** — ele depende de confiar
no registro de quem estava na sessão. Fica registrado aqui para transparência,
não como fato verificável de forma independente.

O conector do Google Drive disponível neste ambiente (`mcp__Google_Drive__*`)
não tem uma forma de atualizar o conteúdo de um arquivo já existente no
Drive — só `create_file` (sempre cria um arquivo/link novo) e `update_file`
(só título e pasta, nunca conteúdo). Isso já forçava cada sincronização a
gerar um link novo. Mas o problema que efetivamente inviabilizou o método foi
outro: para chamar `create_file` era preciso colar o `.xlsx` inteiro
codificado em base64 (~40-45 mil caracteres) dentro do parâmetro
`base64Content` — sem poder referenciar um arquivo local. Tentativas de
reproduzir um texto desse tamanho (mesmo dividido em partes de 14000
caracteres, mesmo com verificação por `cmp` byte a byte) falharam de forma
reproduzível: um mesmo trecho saiu corrompido do mesmo jeito (tamanho
decodificado errado) em duas tentativas independentes, indicando um limite
estrutural de geração de texto verbatim muito longo, não só erro aleatório de
transcrição. Depois de várias sessões tentando (incluindo escrita
incremental com âncoras via `Edit`, uploads de diagnóstico em pedaços
menores para isolar o trecho com erro, etc.), a conclusão foi abandonar esse
caminho em favor da importação manual descrita acima — mais simples, mais
confiável, e com a vantagem extra de manter o mesmo link.
