# Site Farmácia Tropical — seletor de loja no WhatsApp

Este repositório guarda o **snapshot do site publicado** em
https://farmaciatropical.netlify.app (projeto Netlify `farmaciatropical`,
site id `ddb70b5c-c978-4af9-961b-2995c2f8d0f0`) com a alteração descrita
abaixo já aplicada.

## O que mudou

Antes, todo botão "Pedir pelo WhatsApp" levava direto para o WhatsApp da
Loja Centro. Agora ele abre um seletor com as três lojas — Centro,
Arraial d'Ajuda e Trancoso — e envia o cliente para o WhatsApp da loja
escolhida, mantendo a mensagem já montada com o nome e o preço do produto.

O seletor entra em:

- `a.vitrine-pedir` — botão de cada card de oferta;
- `a.vitrine-zap` — botão do topo do site;
- `a.vitrine-diferencial-link` — botão "Pedir entrega".

Os links da seção "Nossas lojas" continuam indo direto para o número
daquela unidade: ali o cliente já escolheu a loja.

## Faixa "Marcas que trabalhamos"

Uma esteira com os logos dos laboratórios rolando continuamente, como havia
no site anterior (`farmacia-tropical.farmaeasy.com.br`). Entra logo acima da
seção "Quem somos" e pausa quando o mouse passa por cima.

- `publicado/assets/marcas-marquee.js` — monta e posiciona a seção.
- `publicado/marcas/*.webp` — os dez logos, trazidos do site anterior.

A vitrine monta as seções fora de ordem enquanto carrega os dados do
Supabase, e as que chegam depois entrariam antes desta faixa. Por isso o
script não só insere a seção: ele a recoloca acima de "Quem somos" sempre
que a ordem muda.

Os logos vêm com fundo branco próprio, então o cartão de cada um é branco
nos dois temas — no tema escuro, um cartão da cor do fundo deixaria cada
logo como um retângulo branco solto.

## Como foi feito

O código-fonte do app (React + Vite + Supabase) não está neste
repositório, então a alteração foi aplicada sobre o site já compilado:

- `publicado/assets/seletor-loja.js` — o seletor, escrito em JavaScript
  puro. Ele intercepta o clique nos botões acima e monta o link
  `wa.me` da loja escolhida preservando o `?text=` original.
- `publicado/index.html` — carrega os scripts acima.

As lojas vêm da tabela `fs_lojas` do Supabase (as mesmas que o site já
usa), buscadas por REST com a chave pública que já está no bundle. Se a
consulta falhar, o script cai numa lista embutida com as três lojas, para
o botão nunca ficar sem destino. Trocar um número no painel da equipe
continua refletindo no seletor sem precisar mexer no código.

O visual usa as variáveis de tema do próprio site (`--card`, `--text`,
`--border`, `--success`), então acompanha o tema claro e o escuro.

## Botão flutuante do WhatsApp

Botão redondo fixo no canto da tela, como no site anterior. Ele não tem
lógica própria de destino: carrega a classe `fs-zap-flutuante`, que o
`seletor-loja.js` intercepta, então o clique abre a mesma escolha entre as
três lojas. O `href` aponta para a Loja Centro e só entra em ação se aquele
script não carregar.

- `publicado/assets/whatsapp-flutuante.js`

Dois cuidados no arquivo:

- O selo "Powered by Netlify" é um iframe fixo no canto inferior direito,
  com z-index máximo. Ele cobria o botão e engolia o clique — por isso o
  script mede o selo e sobe o botão acima dele, refazendo a conta quando a
  janela muda de tamanho e quando o selo é injetado (ele chega depois do
  carregamento).
- O botão não aparece nas rotas `#/login` e `#/admin`: o painel da equipe
  não é lugar para um botão de venda.

## Banner "Arraial 24 horas"

Arte no mesmo padrão do banner que já existia (vermelho com corte diagonal
e faixa amarela): avisa que a loja de Arraial d'Ajuda abre 24 horas e traz
os horários de entrega de madrugada, 01h30 e 03h00.

- `publicado/banners/arraial-24h.webp` — 2400×760, o dobro do formato do
  carrossel (1200×380), para ficar nítido em tela retina.

A imagem está hospedada junto com o site porque o Supabase Storage recusa
upload com a chave pública — subir por lá exige estar logado como equipe.
O registro em `fs_banners` aponta para a URL do site. Se um dia a imagem
for subida pelo painel da equipe, basta trocar a `imagem_url` desse
registro e apagar o arquivo daqui.

Com dois banners cadastrados, o carrossel finalmente alterna sozinho — ele
já fazia isso, mas ficava parado porque só havia um banner.

## Transição entre os banners

`publicado/assets/banner-transicao.js`. A vitrine troca o `src` da mesma
imagem, então a mudança era seca. Ao perceber a troca, o script sobrepõe uma
cópia da arte anterior e a faz sumir em meio segundo, produzindo o cruzamento
entre as duas.

- As artes são pré-carregadas na abertura da página. Sem isso, a primeira
  troca mostraria o fundo vazio do cartão enquanto a imagem nova baixasse.
- As setas e os pontinhos ganham `z-index: 2` para continuarem acima da cópia
  que some.
- Quem tem "reduzir animações" ligado no sistema segue com a troca seca.

A ordem dos banners é o campo `ordem` da tabela `fs_banners`: hoje o de
Arraial 24 horas está em 0 e aparece primeiro.

## Carrossel de banners no celular

`publicado/assets/banner-mobile.js` conserta duas coisas, só em aparelhos de
toque e telas até 640px. No computador nada disso roda.

**Parava de passar sozinho.** A vitrine pausa o rodízio no `mouseenter` e só
retoma no `mouseleave`. Num aparelho de toque o `mouseenter` dispara no
primeiro toque e o `mouseleave` nunca vem, então o carrossel congelava para
sempre. Um vigia percebe que o banner ficou parado além de 7,5s (o rodízio
troca a cada 6s) e toca no próximo pontinho. No computador a pausa ao passar
o mouse continua valendo — foi verificada depois da mudança.

**As artes apareciam cortadas nas laterais.** Até 640px o CSS do site deixa o
quadro do banner mais alto (`aspect-ratio: 1200/520`) do que as artes
(1200/380), e o recorte é `cover`, que amplia a imagem até preencher a altura.
O script passa a dar ao quadro a proporção da própria imagem carregada, então
nada é cortado — e vale para qualquer arte que venha a ser cadastrada, não só
as duas de hoje.

## Prévia ao compartilhar

O `og:url` do `index.html` apontava para `farmastudio-tropical.netlify.app`,
que não é o endereço do site. Corrigido para
`https://farmaciatropical.netlify.app/`, senão o preview do link no
WhatsApp e nas redes pode sair inconsistente.

## Publicar de novo

Dentro de `publicado/`:

```shell
npx -y @netlify/mcp@latest --site-id ddb70b5c-c978-4af9-961b-2995c2f8d0f0 --proxy-path "<proxy-path>"
```

## Atenção

Se o site for republicado a partir da ferramenta que gera o código-fonte,
esta alteração é sobrescrita. Para ela ficar permanente, o mesmo
comportamento precisa ser levado para o fonte do app.
