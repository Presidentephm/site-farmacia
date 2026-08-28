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

## Como foi feito

O código-fonte do app (React + Vite + Supabase) não está neste
repositório, então a alteração foi aplicada sobre o site já compilado:

- `publicado/assets/seletor-loja.js` — o seletor, escrito em JavaScript
  puro. Ele intercepta o clique nos botões acima e monta o link
  `wa.me` da loja escolhida preservando o `?text=` original.
- `publicado/index.html` — carrega o script acima.

As lojas vêm da tabela `fs_lojas` do Supabase (as mesmas que o site já
usa), buscadas por REST com a chave pública que já está no bundle. Se a
consulta falhar, o script cai numa lista embutida com as três lojas, para
o botão nunca ficar sem destino. Trocar um número no painel da equipe
continua refletindo no seletor sem precisar mexer no código.

O visual usa as variáveis de tema do próprio site (`--card`, `--text`,
`--border`, `--success`), então acompanha o tema claro e o escuro.

## Publicar de novo

Dentro de `publicado/`:

```shell
npx -y @netlify/mcp@latest --site-id ddb70b5c-c978-4af9-961b-2995c2f8d0f0 --proxy-path "<proxy-path>"
```

## Atenção

Se o site for republicado a partir da ferramenta que gera o código-fonte,
esta alteração é sobrescrita. Para ela ficar permanente, o mesmo
comportamento precisa ser levado para o fonte do app.
