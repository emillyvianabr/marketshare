# Market Share — painel Grupo Iter

Painel estático (HTML + JavaScript, sem servidor) com PAX e market share dos atrativos do Rio.
Toda a informação vem de **um arquivo**: `data/pax.csv`.

## Estrutura

```
index.html              página do painel
assets/                 estilos, lógica (app.js), exportação para PPT (export-pptx.js), fontes Stolzl, logo e bibliotecas
data/pax.csv            dados mensais (é o único arquivo que você edita todo mês)
data/eventos.json       fatores externos mostrados nos gráficos (obras, clima, eventos)
data/config.json        regra do BioParque, textos de nota e textos do PPT exportado
scripts/validate.py     confere o CSV antes de publicar
.github/workflows/      validação + publicação automática
```

## Atualização mensal (≈ 2 minutos)

1. No GitHub, abra `data/pax.csv` e clique no lápis (Edit).
2. No fim do arquivo, acrescente as linhas do mês novo, uma por atração, em **milhares** e com ponto decimal:

   ```
   2026-10-01,Bondinho,168.2
   2026-10-01,AquaRio,95.4
   2026-10-01,Trem Corcovado,90.1
   2026-10-01,Paineiras,104.8
   2026-10-01,YupStar,20.3
   ```
   - A data é sempre o **dia 1** do mês (AAAA-MM-01).
   - **Cristo e Mercado não entram no arquivo**: o painel soma Trem + Paineiras e o mercado sozinho.
   - **BioParque**: pode ficar de fora enquanto não houver dado (veja abaixo).
3. Clique em **Commit changes**. A automação confere o arquivo e republica o painel em 1 a 2 minutos.
   Se houver erro (atração com nome errado, data inválida, mês incompleto), a publicação é bloqueada e a aba **Actions** mostra o motivo.

Quem prefere continuar no Excel: `python scripts/xlsx_to_csv.py PAX_Concorrencia_Set26_organizado.xlsx` regera o CSV a partir da aba Base_PAX.

## Fatores externos nos gráficos (obras, clima, eventos)

Nos gráficos de 13 meses aparecem símbolos sobre os meses em que algum fator externo pode ter afetado o PAX.
Clicar no símbolo abre, embaixo do gráfico, um quadro com a explicação; clicar de novo (ou no ×) fecha.
Tudo vem de **`data/eventos.json`** — uma lista, um bloco por fator:

```json
{
  "mes": "2026-03",            // AAAA-MM
  "ate": "2026-04",            // opcional: período (o símbolo aparece em cada mês)
  "tipo": "obras",             // obras 🚧 · sol ☀️ · chuva 🌧️ · evento 🎉 · outro ℹ️
  "titulo": "Obras no acesso", // título do quadro
  "descricao": "O que aconteceu e como afetou o PAX.",
  "impacto": "negativo",       // negativo · positivo · neutro (cor do contorno)
  "atracoes": ["Bondinho"]     // opcional
}
```

Vários fatores no mesmo mês ficam empilhados. O `validate.py` confere o arquivo antes de publicar.
Os itens que começam com **EXEMPLO** são só demonstração: troque pelos reais ou apague.
Para abrir com duplo clique, rode `python scripts/build_local.py` (inclui os fatores no `data/pax.js`).

## Exportar para PowerPoint

O botão **Exportar PPT** (canto direito dos filtros) baixa um `.pptx` com **o que estiver selecionado na tela**
(ano, comparação, mês de fechamento). Nome do arquivo: `Report_Pax_Atrativos_Rio_de_Janeiro_Set26.pptx`.

O arquivo segue o relatório mensal, em 2 slides (16:9):

1. **PAX** — texto de abertura, linhas dos últimos 13 meses, barras YTD (com a abertura Trem/Paineiras ao lado do Cristo),
   variações ao lado de cada série e tabela de variação mês a mês.
2. **Market share** — texto de abertura, linhas de share, barras YTD em % e variações em p.p.

- Os gráficos são **nativos do PowerPoint** (clique com o botão direito → *Editar dados*). Rótulos de variação, abertura do Cristo e tabela são formas/tabela editáveis.
- Os textos de abertura são gerados a partir dos números (mesmo estilo do relatório) — revise antes de enviar.
- Fonte usada: *Ofelia Display* (a mesma do relatório). Se não estiver instalada, o PowerPoint substitui por outra.
- Textos fixos (título, fonte, nota do BioParque no rodapé) ficam em `data/config.json`, bloco `"export"`.
- A biblioteca que monta o PPT (`assets/vendor/pptxgen.bundle.js`, PptxGenJS 4.0.1) só é carregada quando o botão é clicado.
  Funciona também abrindo o `index.html` por duplo clique.

## Critério do BioParque (painel = relatório)

Sem dados do BioParque em 2026, **toda comparação que envolva 2026 calcula o mercado sem o BioParque nos dois períodos**
(mesma regra do relatório em PPT; não há filtro para mudar isso). Comparações entre outros anos mantêm o BioParque.
O ano vem de `bioparque_zero_desde` em `data/config.json`; o texto do aviso fica em `nota_exclusao`.

## BioParque em 2026

Com a troca de gestão do parque, os dados do BioParque não estão chegando. O painel o trata como **zerado** a partir do mês definido em `data/config.json`:

```json
"bioparque_zero_desde": "2026-01"
```

- `2026-01` = zera o ano inteiro (configuração atual).
- `2026-08` = mantém jan–jul com os dados reais e zera de agosto em diante.
- Quando os dados voltarem, apague a linha `bioparque_zero_desde` (ou coloque uma data futura, como `2099-01`) e lance os valores no CSV.

Enquanto a regra estiver ativa, o painel mostra um aviso no topo, asterisco (*) nos números afetados e a nota no rodapé.
O texto da nota fica em `nota_bioparque` no mesmo arquivo.

## Publicar pela primeira vez

1. Crie o repositório e envie estes arquivos para a branch `main`.
2. Em **Settings → Pages → Source**, escolha **GitHub Actions**.
3. O endereço do painel aparece na aba **Actions**, ao fim da primeira execução.

**Atenção à privacidade:** o GitHub Pages é público por padrão e os dados são de concorrentes.
Para restringir o acesso é preciso GitHub Enterprise Cloud (Pages privado) ou hospedar em outro serviço com senha (Cloudflare Pages com Access, Netlify, etc.). O `index.html` já traz `noindex`, mas isso só evita buscadores, não protege o conteúdo.

## Ver no computador

```
python -m http.server 8000     # depois abra http://localhost:8000
```
Para abrir o `index.html` com duplo clique (sem servidor), rode `python scripts/build_local.py` antes; ele gera `data/pax.js`. O site publicado ignora esse arquivo.

## Marca

Cores, tipografia (Stolzl) e logo seguem o brandbook Grupo Iter 2023. O logo é usado sempre a partir dos arquivos oficiais em `assets/`.
