# garoa.github.io

Site estático do **Garoa Hacker Clube**, publicado pelo GitHub Pages em
<https://garoa.github.io/>. Não tem backend nem banco de dados: o GitHub serve
direto os arquivos HTML deste repositório.

## O que tem aqui

```
index.html          Página de pagamento da mensalidade e doações
ghc_logo.png        Logo do Garoa (copiado de garoa.net.br)
contas/             Página de prestação de contas
  index.html        A página em si (gráficos e tabelas)
  dados.js          Dados agregados que a página lê (gerado por script)
  assets/           Logo em SVG
  scripts/          Scripts que geram dados.js e o logo
  planilhas/        Planilhas originais (.xlsx/.csv) — NÃO vão para o git
.gitignore          Ignora *.csv, *.xlsx, .venv/, .DS_Store e __pycache__/
```

## Página de pagamento (`/`)

Endereço: <https://garoa.github.io/>

Formas de contribuir, nesta ordem:

1. **Pix** (preferencial, porque não tem taxas): chave `pagamentos@garoa.net.br`,
   com um botão para copiar.
2. **Assinatura PayPal**: mensalidade recorrente de R$ 187. O botão vem do SDK do
   PayPal (plano `P-9TP3259813133982JMDPFP3Y`).
3. **Doação PayPal**: valor livre (botão hospedado `TPPZB2LQJDV2S`).

Para mudar o valor da mensalidade, troque o texto em `index.html` **e** o valor do
plano no PayPal: a página só exibe o valor, quem cobra é o plano.

## Prestação de contas (`/contas/`)

Endereço: <https://garoa.github.io/contas/>

Mostra as receitas e despesas do ano por mês e por categoria, os saldos, quanto
falta arrecadar até dezembro, o fundo de reserva e a previsão de fechamento do ano.

A página lê os números de `contas/dados.js`, que **só tem valores somados por
categoria** (aluguel, luz, mensalidades etc.), sem nomes nem dados pessoais. As
planilhas originais ficam fora do git pelo `.gitignore`.

## Scripts

Os scripts ficam em `contas/scripts/`. Eles acham os arquivos a partir da
própria pasta, mas os comandos abaixo supõem que você está em `contas/`:

```sh
cd contas
```

### Preparar o ambiente (uma vez só)

Precisa de Python 3. O `extrair.py` usa a biblioteca `openpyxl`, então crie um
ambiente virtual e instale:

```sh
python3 -m venv .venv
.venv/bin/pip install -r scripts/requirements.txt
```

A pasta `.venv/` fica fora do git.

### `extrair.py`: gera o `dados.js` a partir da planilha

Lê `contas/planilhas/Garoa HC - Prestação de Contas 2026.xlsx` e escreve
`contas/dados.js` só com os totais por categoria.

1. Copie a planilha atualizada para `contas/planilhas/`, com o mesmo nome.
2. Se fechou um mês novo, ajuste o começo do `scripts/extrair.py`:
   - acrescente a aba do mês em `MESES` (ex.: `"Out"`);
   - tire esse mês de `ABAS_SEGUINTES`, que lista os meses ainda abertos.
3. Rode:
   ```sh
   .venv/bin/python scripts/extrair.py
   ```
   O script mostra os totais de cada mês e confere com a aba **Resumo** da
   planilha. Se algum mês não bater, ele para com um erro e não gera o arquivo.
4. Abra a página (veja [Testar localmente](#testar-localmente)), confira o
   `dados.js` gerado (ele vai ficar público) e faça o commit.

Também no começo do `extrair.py` ficam:

- `REGRAS_RECEITA` e `REGRAS_DESPESA`: qual palavra no lançamento leva a qual
  categoria. Vale a primeira regra que casar.
- As premissas da previsão até dezembro: aluguel, IPTU, contabilidade e as médias
  usadas para luz, água, internet e mensalidades.

Para um ano novo, troque `ANO`, o nome do arquivo em `PLANILHA` e as listas
`MESES` e `ABAS_SEGUINTES`.

### `extrair_logo.py`: gera o logo em SVG

Baixa o cartaz SVG da wiki (`Garoa_Cartaz_inline.svg`) e salva só o guarda-chuva
em `contas/assets/garoa-logo.svg`. Não precisa de dependências:

```sh
python3 scripts/extrair_logo.py
```

Só é necessário se o logo mudar. A página não carrega esse arquivo: o desenho
está colado dentro do `contas/index.html` (no `<symbol id="garoa-logo">`). Se
gerar um novo, copie o conteúdo para lá.

## Testar localmente

A página de pagamento abre direto no navegador. Para a de contas, é melhor rodar
um servidor local na raiz do repositório:

```sh
python3 -m http.server
```

e abrir <http://localhost:8000/> e <http://localhost:8000/contas/>.

## Publicar

Basta fazer commit e push no `master`: o GitHub Pages atualiza o site sozinho em
um ou dois minutos.
