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

## Link atual — ⚠️ PRECISA DE CONFIRMAÇÃO DO USUÁRIO (ver nota abaixo)

- **Título:** Plano Alimentar Pós-Infarto — Gustavo Ozeki
- **Link:** https://docs.google.com/spreadsheets/d/12gqsq1bewjfHwMBZUGunugg2RDUSpkzUfvq9Y1AvAZg/edit
- **fileId:** `12gqsq1bewjfHwMBZUGunugg2RDUSpkzUfvq9Y1AvAZg`
- **Status (checado em 13/09/2026 via `get_file_metadata`): este fileId não
  resolve mais** ("Requested entity was not found"). Uma busca por
  `mimeType = 'application/vnd.google-apps.spreadsheet'` em toda a conta
  conectada (`ozeki1@gmail.com`) também não encontra nenhum Google Sheet com
  esse título. Ou seja, **não há evidência de que este link ainda funcione**
  — não sabemos se foi excluído, se a importação de 13/09 não foi feita
  nesse link, ou se foi feita em outra conta Google.
- **O que existe de fato no Drive, verificado em 13/09/2026:** um arquivo
  `plano-alimentar-pos-infarto.xlsx` (fileId `1zKG4qAjrdZxUNaPilLeHkIywWKXgeN2V`,
  https://drive.google.com/file/d/1zKG4qAjrdZxUNaPilLeHkIywWKXgeN2V/view),
  criado às 05:23 de 13/09/2026. **Mas isso não é um Google Sheet nativo** —
  o `mimeType` é `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`
  (um `.xlsx` cru, como um anexo comum) e o título é o nome do arquivo, não
  "Plano Alimentar Pós-Infarto — Gustavo Ozeki". O tamanho em bytes bate
  exatamente com o commit `a7348f9` (a versão com sashimi/sushi/tahine/kare,
  **sem** cream cheese e sem os vegetais novos — ou seja, é o arquivo que foi
  mandado pelo chat *antes* da última rodada de alterações desta sessão, não
  a versão mais atual). Isso sugere que a ação do usuário em 13/09 foi um
  **upload avulso do arquivo pelo Drive**, não uma importação
  "Substituir planilha" dentro de um Sheets já existente como este documento
  supunha — o que explica por que o resultado é um `.xlsx` solto e não um
  Sheets nativo.
- **Ação necessária:** perguntar ao usuário qual é, hoje, o link real que ele
  usa para ver a planilha na nuvem, e substituir os três campos acima
  (Título/Link/fileId) por esse valor confirmado, ou registrar aqui que não
  existe mais um Sheets ativo até ele criar um novo.

### Links anteriores (não recebem mais atualizações — status não confirmado)
- `1-Mz3ue4QsSVytQIuL9MJUFnLHGdwp1VmS2wcleT5aDc` — criado em 06/09/2026 (era
  da época do upload automático, substituído pelo link acima em 07/09/2026).
  Também retornou "not found" em `get_file_metadata` no mesmo teste de
  13/09/2026 — não dá para confirmar se ainda existe (trashed ou excluído
  de vez) só com essa ferramenta.

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
