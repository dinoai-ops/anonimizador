# Anonimizador Jurídico LGPD — v4 (app local)

Ferramenta local, gratuita e de código aberto em Python para anonimização de
dados pessoais e sensíveis em documentos jurídicos brasileiros (.docx, .txt e
.pdf), em conformidade com a LGPD.

**Não usa APIs externas.** Tudo é processado no seu computador — nenhum dado sai
da máquina. A anonimização é **reversível**: cada arquivo gera um "mapa de
reversão" guardado na pasta `cofre/`, que permite restaurar o documento original.

> ⚠️ **Aviso importante:** esta é uma ferramenta de **apoio** à anonimização.
> A detecção é automática e **não garante 100% de cobertura** — a revisão humana
> da lista de entidades (etapa "1 · Analisar") e do documento final é
> **indispensável** antes de qualquer compartilhamento. Veja a seção
> [Licença e aviso legal](#10-licença-e-aviso-legal).

---

## 1. O que instalar (uma única vez)

Um leigo consegue: são só dois passos.

### Passo 1 — Instalar o Python

1. Acesse **https://www.python.org/downloads/** e baixe o Python (versão 3.10 ou superior).
2. Execute o instalador e, **antes de clicar em "Install Now"**, marque a caixinha
   **"Add python.exe to PATH"** (fica no rodapé da janela). Esse detalhe é essencial.
3. Conclua a instalação normalmente.

### Passo 2 — Instalar as duas bibliotecas

1. Abra o **Prompt de Comando** (tecla Windows → digite `cmd` → Enter).
2. Copie e cole a linha abaixo e tecle Enter:

```
pip install python-docx pymupdf
```

3. Aguarde terminar (aparece "Successfully installed..."). Pode fechar a janela.

| Biblioteca | Para quê |
|---|---|
| `python-docx` | ler e gravar arquivos **.docx** (Word) |
| `pymupdf` | ler e gravar arquivos **.pdf** |

Arquivos **.txt** não exigem nada além do próprio Python.

**Para conferir se deu certo:** abra o Prompt de Comando e digite `python --version`
(deve responder algo como `Python 3.14.0`) e depois `pip show python-docx pymupdf`
(deve listar as duas bibliotecas).

---

## 2. Como usar o app (sem terminal)

Dê **duplo clique em `Anonimizador.pyw`**. A janela do app abre. Então:

1. **Selecionar arquivo…** — escolha o documento original (.docx, .txt ou .pdf).
2. **1 · Analisar** — o app lista todas as entidades detectadas: categoria,
   dado original, número de ocorrências e a **substituição** proposta
   (ex.: `[NOME_01]`).
3. **Revisar** (opcional):
   - **duplo clique** na coluna *Substituição* para trocar o texto de qualquer
     entidade (ex.: trocar `[NOME_01]` por `SÓCIA FUNDADORA`);
   - **um clique** na coluna *Usar* para desmarcar um falso positivo — a
     entidade desmarcada permanece intacta no documento.
4. **2 · Anonimizar e salvar** — gera o arquivo `_anonimizado` em `saida/`,
   o mapa de reversão em `cofre/` e o relatório LGPD em `relatorios/`.
   Para PDFs, o seletor **“Saída do PDF”** na barra de ações define o formato
   do arquivo anonimizado: **PDF**, **TXT** ou **DOCX** (a escolha fica salva
   para os próximos usos). Como o layout do PDF já é simplificado na saída,
   TXT/DOCX costumam ser mais práticos para envio a IAs ou edição posterior.
5. **Restaurar…** — processo inverso (regressão): selecione o arquivo
   anonimizado e o app devolve o original, usando o mapa do cofre. Funciona
   **tanto com os tokens padrão quanto com substituições customizadas**.
   Também aceita um **documento diferente do original** que contenha os mesmos
   tokens (ex.: parecer ou minuta revisada devolvidos por uma IA): se o mapa
   não for localizado pelo nome do arquivo, o app pede para escolher qual mapa
   do cofre aplicar, e informa quantos trechos foram efetivamente revertidos.

### Botões auxiliares

- **Modelos de token…** — define o texto default por categoria, válido para as
  próximas análises (salvo em `app_prefs.json`). Exemplos de modelo:
  `[{cat}_{n:02d}]` (padrão), `PARTE {n}`, `<<NOME {n}>>`. Todo modelo precisa
  conter `{n}` (número sequencial), senão duas pessoas receberiam o mesmo texto
  e a reversão ficaria impossível.
- **Config. LGPD…** — preenche finalidade, destinatário, base legal e categoria
  dos titulares. **Obrigatório antes da primeira anonimização** (esses dados
  entram no relatório e no log de compliance, art. 37 da LGPD).

### Proteções automáticas

- **Mesmo texto para dados diferentes** (ex.: substituir vários nomes por
  `XXX`) é permitido, mas o app pede **confirmação explícita**: esses trechos
  ficam **fora do mapa de reversão** e NÃO serão restaurados pela função
  *Restaurar* — permanecem anonimizados (regressão parcial). O restante do
  documento continua totalmente reversível.
- O app **avisa** se um texto customizado já existir no documento original
  (a restauração poderia substituir texto legítimo).

---

## 3. Estrutura de pastas

```
anonimizador-local/
├── Anonimizador.pyw   ← DUPLO CLIQUE AQUI para abrir o app
├── app.py             ← interface do app local
├── anonimizador.py    ← motor de anonimização (v4)
├── config.json        ← metadados LGPD (editável pelo app)
├── app_prefs.json     ← modelos de token do usuário (criado pelo app)
├── entrada/           ← originais (usado pelo modo linha de comando)
├── saida/             ← arquivos anonimizados e restaurados
├── cofre/             ← mapas de reversão  [CONFIDENCIAL]
├── relatorios/        ← relatórios LGPD + log_compliance.jsonl (append-only)
└── exemplos/          ← peças FICTÍCIAS para testar a ferramenta
```

> **Quer testar sem arriscar nada?** A pasta [`exemplos/`](exemplos) traz uma
> petição (PDF), um contrato social (DOCX) e uma procuração (TXT) com dados
> 100% fictícios. A petição contém, de propósito, um nome que **escapa** da
> detecção automática — leia o [`exemplos/LEIA-ME.txt`](exemplos/LEIA-ME.txt)
> e entenda por que a revisão humana é indispensável.

> **O `cofre/` é confidencial.** Cada mapa contém os dados originais e permite
> reverter a anonimização. Nunca envie o cofre junto com o arquivo anonimizado
> (por e-mail, IA, nuvem etc.). Sem o mapa, o arquivo anonimizado não é reversível.

---

## 4. O que é anonimizado

Tokens numerados por categoria; a mesma pessoa/empresa recebe **sempre o mesmo
token** em todo o documento (corpo, tabelas, cabeçalhos e assinaturas).

| Categoria | Exemplos de dado detectado |
|---|---|
| `NOME` | pessoas (qualificação, assinaturas, gatilhos: Sr., Dr., autor, réu, "Eu,", CN=…) e organizações (LTDA, S.A., razão social antes de CNPJ) |
| `DOCUMENTO` | CPF, CNPJ, RG ("carteira de identidade nº…"), OAB, CNH, PIS, Título de Eleitor, NIRE, protocolos SEI |
| `PROCESSO` | número CNJ, processo administrativo |
| `CONTATO` | e-mail, telefone, sites (http/https/www) |
| `ENDERECO` | logradouro completo até o CEP, CEP, cidade/UF, IP, matrícula de imóvel |
| `DADOS_BANCARIOS` | agência, conta, banco, chave PIX |
| `DADO_VEICULO` | placas (formato antigo e Mercosul) |
| `DADO_SENSIVEL` | nomeação/designação por decreto ou portaria |
| `VALOR` | R$ com ou sem extenso ("R$ 30.700,00 (trinta mil e setecentos reais)"), quantidades de quotas |
| `DATA` | 10/06/2026, "10 de junho de 2026" — para preservar datas, desmarque-as na revisão do app |

Referências legais (art. 1.052, Lei nº 10.406/2002 etc.) são **preservadas**.

---

## 5. Levar para outra máquina

Copie a **pasta inteira** `anonimizador-local` (pen drive, rede etc.) — é o mais
simples e leva junto cofre, relatórios e preferências.

Instalação mínima (apenas o app, sem histórico): copie
`anonimizador.py`, `app.py`, `Anonimizador.pyw`, `config.json` e, se quiser
manter seus modelos de token, `app_prefs.json`. As pastas `entrada/`, `saida/`,
`cofre/` e `relatorios/` são criadas automaticamente.

Na máquina de destino, faça a **seção 1** (Python com "Add to PATH" +
`pip install python-docx pymupdf`) e dê duplo clique em `Anonimizador.pyw`.

> Importante: para conseguir **restaurar** arquivos anonimizados antigos na nova
> máquina, leve também a pasta `cofre/` correspondente.

---

## 6. Uso avançado (linha de comando, opcional)

O modo antigo continua funcionando, com os tokens padrão e em lote:

```powershell
# anonimiza tudo o que estiver em entrada/
python anonimizador.py

# PDFs geram saída anonimizada em .txt (ou docx | pdf; padrão: pdf)
python anonimizador.py --saida-pdf txt

# restaura um arquivo anonimizado usando o mapa do cofre
python anonimizador.py --restaurar "saida\documento_anonimizado.docx"
```

---

## 7. Conformidade LGPD (art. 37 e art. 6º, VIII)

- **Relatório por arquivo** (`relatorios/<nome>_relatorio.json`): hash SHA-256 do
  original (cadeia de custódia), data/hora, operador, finalidade, destinatário,
  base legal, categoria dos titulares e contagens por categoria.
- **Log consolidado** (`relatorios/log_compliance.jsonl`): uma linha por
  processamento, append-only — registro permanente das operações de tratamento.
- **Mapa de reversão** (`cofre/<nome>_mapa.json`): pseudonimização reversível;
  guarda `texto substituído → dado original` e o hash do arquivo original.

---

## 8. Limitações conhecidas

- **PDF de saída**: o conteúdo é preservado, mas o layout visual é simplificado
  (texto corrido, fonte Helvetica). Para manter a formatação, prefira anonimizar
  o **.docx** de origem quando existir. Alternativamente, use o seletor
  **“Saída do PDF”** (ou `--saida-pdf txt|docx` na CLI) para gerar a versão
  anonimizada em TXT ou DOCX, formatos mais adequados para revisão e para
  envio a ferramentas de IA.
- **PDFs escaneados (imagem)** não são suportados — exigem OCR prévio. O app
  detecta e avisa.
- **Falsos negativos**: nomes soltos no meio de texto, sem qualificação nem
  gatilho (Sr., autor, assinatura etc.), podem escapar. Ao detectar um nome uma
  vez, o app o substitui em todas as demais ocorrências — mas **revise sempre a
  lista da análise e o documento final antes de compartilhar**.
- **Restauração e caixa**: o texto restaurado reproduz a primeira forma
  encontrada da entidade (ex.: "PRIMAVERA SABORES LTDA." pode voltar em
  maiúsculas onde o original usava "Primavera Sabores Ltda."). Nenhum dado é
  perdido; apenas a caixa das letras pode variar nesses pontos.

---

## 9. Problemas comuns

| Sintoma | Causa e solução |
|---|---|
| Nada acontece ao clicar no `Anonimizador.pyw` | Python instalado sem "Add to PATH" ou sem associação de arquivos .pyw. Reinstale o Python marcando a caixinha (o instalador oferece "Repair"). Alternativa: abra o Prompt de Comando na pasta e rode `pythonw app.py`. |
| Windows bloqueia um arquivo ("pode não ser seguro", SmartScreen ou Controle de Aplicativo Inteligente) | Arquivos baixados da internet recebem uma marcação de segurança do Windows. Antes de extrair, clique com o botão direito no **ZIP baixado** → **Propriedades** → marque **Desbloquear** → OK, e só então extraia. Desde a v4.2 o pacote não contém executáveis nem scripts (.bat/.exe) — apenas arquivos Python, abertos pelo Python que você mesmo instalou —, o que evita o bloqueio na grande maioria dos casos. |
| `'pip' não é reconhecido...` | Mesmo problema de PATH acima. Alternativa: use `py -m pip install python-docx pymupdf`. |
| "python-docx nao instalado" ao anonimizar | Rode `pip install python-docx` no Prompt de Comando. |
| "PyMuPDF nao instalado" ao abrir PDF | Rode `pip install pymupdf`. |
| "PDF sem texto extraivel" | O PDF é escaneado (imagem). Faça OCR antes ou use o .docx de origem. |
| "Preencha os metadados LGPD" | Clique em **Config. LGPD…** e preencha os 4 campos. |
| "Mapa de reversao nao encontrado" ao restaurar | O arquivo `_mapa.json` correspondente não está na pasta `cofre/`. A restauração só é possível com o mapa gerado na anonimização. |

---

## 10. Licença e aviso legal

Este software é distribuído gratuitamente sob a **licença MIT** — veja o
arquivo [LICENSE](LICENSE). Em resumo: você pode usar, copiar, modificar e
redistribuir livremente, inclusive para fins comerciais, desde que preserve o
aviso de copyright.

**Isenção de garantias e de responsabilidade.** O software é fornecido "no
estado em que se encontra" (*as is*), **sem garantia de qualquer natureza**,
expressa ou implícita. A detecção de dados pessoais é feita por padrões
automáticos (expressões regulares) e **não alcança, em nenhuma hipótese, 100%
de cobertura**: nomes e dados atípicos podem escapar à detecção. O uso da
ferramenta **não substitui a revisão humana** nem constitui, por si só,
cumprimento da LGPD. A responsabilidade pela conferência do documento
anonimizado antes de qualquer compartilhamento — e pelo tratamento de dados
pessoais em geral — é **exclusivamente do usuário**. Os autores não respondem
por danos decorrentes do uso do software, nos termos da licença MIT.

**Boas práticas mínimas:**
1. Revise sempre a lista de entidades na etapa "1 · Analisar".
2. Releia o documento final antes de enviá-lo a terceiros ou a ferramentas de IA.
3. Guarde a pasta `cofre/` em local seguro e **nunca a compartilhe** junto com
   os arquivos anonimizados.
4. Este software não presta consultoria jurídica; dúvidas sobre a LGPD devem
   ser dirigidas a profissional habilitado.
