# Sincronização com o Google Sheets

## Por que isso existe

O usuário pediu um lugar "na nuvem" para editar a planilha junto comigo, sem
precisar que eu reenvie o arquivo pelo chat toda vez. A opção escolhida foi o
Google Sheets. Duas limitações do conector do Google Drive disponível neste
ambiente (`mcp__Google_Drive__*`) moldam como isso funciona:

1. **Não existe uma forma de atualizar o conteúdo de um arquivo já existente
   no Drive.** As únicas ferramentas de escrita são `create_file` (sempre cria
   um arquivo/link novo) e `update_file` (só título e pasta, nunca conteúdo).
   Não há API de células do Sheets neste conector. Ou seja: **cada
   sincronização gera um link novo** — isso foi confirmado com o usuário e
   aceito por ele.
2. **A coluna "Gostoso" é preenchida pelo usuário à mão, diretamente no
   Sheets.** Se eu simplesmente regenerar a planilha a partir do
   `dados_alimentos.json` e subir um Sheets novo, esses valores manuais somem
   — porque a fonte de verdade da regeneração é o JSON + o `.xlsx` local, que
   não sabem o que foi editado no Sheets. Por isso o procedimento abaixo
   sempre **baixa o Sheets atual antes de regenerar**, para fundir esses
   valores (ver `load_gostoso_previo()` em `gerar_planilha.py`).

## Link atual

- **Título:** Plano Alimentar Pós-Infarto — Gustavo Ozeki
- **fileId:** `1-Mz3ue4QsSVytQIuL9MJUFnLHGdwp1VmS2wcleT5aDc`
- **Link:** https://docs.google.com/spreadsheets/d/1-Mz3ue4QsSVytQIuL9MJUFnLHGdwp1VmS2wcleT5aDc/edit
- **Criado em:** 06/09/2026 (a partir do `.xlsx` gerado nesta mesma sessão)

Sempre que uma sincronização criar um Sheets novo, **atualize os três campos
acima** (título permanece igual; fileId e link mudam) para o próximo run desta
skill/procedimento saber de onde baixar o "Gostoso" antes de regenerar.

## Quando sincronizar

**Não é automático a cada alimento adicionado** — o passo de subir um arquivo
novo no Drive exige codificar o `.xlsx` inteiro em base64 e colar esse texto
num parâmetro de ferramenta (não há como referenciar um caminho de arquivo
local diretamente), o que é um processo manual e sensível a erro de
transcrição (ver seção "Procedimento" abaixo — inclui um passo de verificação
por causa disso). Só sincronize com o Sheets quando o usuário pedir
explicitamente (ex.: "atualiza o sheets também", "sincroniza com a planilha na
nuvem", "sobe pro Sheets de novo").

Para pedidos comuns de "adicione X" sem menção ao Sheets, o fluxo normal
continua sendo só regenerar o `.xlsx` local e enviar pelo chat (skill
`adicionar-alimento`), sem tocar no Drive.

## Procedimento (quando o usuário pedir sincronização)

1. **Baixar o Sheets atual como .xlsx:**
   `mcp__Google_Drive__download_file_content` com o `fileId` guardado acima e
   `exportMimeType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"`.
   O resultado traz `{fileContent: <base64>}` — se for grande, a ferramenta
   salva automaticamente num arquivo de resultado em
   `.../tool-results/...txt`; **não** tente copiar esse base64 manualmente.
   Em vez disso, use Python/Bash para ler esse JSON e decodificar direto para
   `planilha/_sheets_download.xlsx`:
   ```bash
   python3 -c "
   import json, base64
   with open('<caminho do tool-results .txt>') as f:
       d = json.load(f)
   with open('planilha/_sheets_download.xlsx', 'wb') as out:
       out.write(base64.b64decode(d['fileContent']))
   "
   ```
   Esse arquivo é só um insumo temporário (está no `.gitignore`) — nunca é
   commitado.

2. **Regenerar com o "Gostoso" do Sheets preservado:**
   ```bash
   GOSTOSO_SHEETS_EXPORT=planilha/_sheets_download.xlsx python3 planilha/gerar_planilha.py
   ```
   Isso funde os valores de "Gostoso" do Sheets baixado com os do `.xlsx`
   local (o Sheets vence em caso de conflito — ver `load_gostoso_previo()`).

3. **Codificar o novo `.xlsx` para upload, em partes pequenas e verificadas.**
   Não existe forma de passar um caminho de arquivo para `create_file` — o
   conteúdo tem que ser colado como texto no parâmetro `base64Content`. Como
   o texto costuma passar de 40 mil caracteres, colar tudo de uma vez arrisca
   erro de transcrição (já aconteceu: um caractere trocado numa sessão
   anterior, pego só por causa do passo de verificação abaixo — nunca pule
   essa verificação). Processo seguro, testado nesta sessão:
   ```bash
   cd planilha
   base64 -w0 plano-alimentar-pos-infarto.xlsx > /tmp/plano_b64.txt
   split -b 14000 -d /tmp/plano_b64.txt /tmp/plano_b64_part_
   ```
   Depois, para cada parte (`/tmp/plano_b64_part_00`, `_01`, `_02`, ...):
   `Read` o arquivo inteiro (cada parte é uma única linha ≤14000 caracteres,
   dentro do limite de truncamento do `Read`), copie o conteúdo (removendo só
   o prefixo `1\t` que o `Read` adiciona) e concatene tudo em ordem, sem
   separador, num arquivo novo via `Write` (ex.: `/tmp/plano_b64_final.txt`).
   **Depois, sempre confira com `cmp`:**
   ```bash
   cmp /tmp/plano_b64.txt /tmp/plano_b64_final.txt && echo IDENTICAL
   ```
   Se `cmp` apontar uma diferença, corrija só o trecho indicado (via `Edit`,
   com contexto suficiente para ser único) e rode `cmp` de novo — repita até
   `IDENTICAL`. Só depois disso use o conteúdo de `/tmp/plano_b64_final.txt`
   como `base64Content` na chamada de `create_file` abaixo.

4. **Subir como Sheets novo:**
   `mcp__Google_Drive__create_file` com:
   - `title`: "Plano Alimentar Pós-Infarto — Gustavo Ozeki" (mesmo título de
     sempre — não numerar/datar, o link é que muda)
   - `base64Content`: o conteúdo verificado no passo 3
   - `contentMimeType`: `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`
   - (não usar `disableConversionToGoogleType` — deixa converter para Sheets nativo)

5. **Atualizar este arquivo** (`planilha/SHEETS_SYNC.md`) com o novo `fileId`
   e link retornados, e commitar junto com as mudanças de alimento.

6. **Reportar ao usuário:** o link novo, e lembrar que o link antigo não
   recebe mais atualizações (mas continua existindo — pode ser arquivado
   depois, se o usuário preferir, via `trash_file`).

7. **Limpar o arquivo temporário:** apagar `planilha/_sheets_download.xlsx`
   (ou deixar — está no `.gitignore`, não afeta o repositório).
