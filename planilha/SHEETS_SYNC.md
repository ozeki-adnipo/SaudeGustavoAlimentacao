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
   baixe o Sheets atual primeiro (`mcp__Google_Drive__download_file_content`
   + `GOSTOSO_SHEETS_EXPORT=... python3 planilha/gerar_planilha.py`, mesmo
   mecanismo do método antigo, só que agora alimentando o `.xlsx` que será
   enviado pelo chat, não um upload direto) para fundir esses valores antes
   de gerar o arquivo que o usuário vai importar.

4. Depois de o usuário confirmar que importou, **atualizar a data em "Link
   atual" abaixo** (o link normalmente não muda, só a data/descrição do que
   foi incluído) e commitar junto com as mudanças de alimento.

## Link atual

- **Título:** Plano Alimentar Pós-Infarto — Gustavo Ozeki
- **Link:** https://docs.google.com/spreadsheets/d/12gqsq1bewjfHwMBZUGunugg2RDUSpkzUfvq9Y1AvAZg/edit
- **fileId:** `12gqsq1bewjfHwMBZUGunugg2RDUSpkzUfvq9Y1AvAZg`
- **Última importação confirmada pelo usuário:** 13/09/2026 (inclui sashimi de
  salmão, sushi de salmão, creme de tahine, kare com lombo de porco — e, a
  partir desta atualização, cream cheese light Philadelphia e a lista
  ampliada de vegetais e verduras)

### Links anteriores (não recebem mais atualizações)
- `1-Mz3ue4QsSVytQIuL9MJUFnLHGdwp1VmS2wcleT5aDc` — criado em 06/09/2026 (era
  da época do upload automático, substituído pelo link acima em 07/09/2026)

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
