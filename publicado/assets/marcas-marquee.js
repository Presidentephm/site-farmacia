/* Farmácia Tropical — faixa "Marcas que trabalhamos".
   Insere, antes da seção "Quem somos", uma esteira com os logos dos
   laboratórios rolando continuamente, como havia no site anterior. */
(function () {
  "use strict";

  var MARCAS = [
    { nome: "EMS", img: "/marcas/ems.webp" },
    { nome: "Eurofarma", img: "/marcas/eurofarma.webp" },
    { nome: "Medley", img: "/marcas/medley.webp" },
    { nome: "Neo Química", img: "/marcas/neoquimica.webp" },
    { nome: "Cimed", img: "/marcas/cimed.webp" },
    { nome: "Aché", img: "/marcas/ache.webp" },
    { nome: "Biolab", img: "/marcas/biolab.webp" },
    { nome: "Bayer", img: "/marcas/bayer.webp" },
    { nome: "Sanofi", img: "/marcas/sanofi.webp" },
    { nome: "Hypera Pharma", img: "/marcas/hypera.webp" }
  ];

  var SEGUNDOS = 45;

  function estilos() {
    if (document.getElementById("fs-marcas-estilo")) return;
    var css = document.createElement("style");
    css.id = "fs-marcas-estilo";
    css.textContent = [
      ".fs-marcas-pista{position:relative;overflow:hidden;padding:4px 0;",
      "-webkit-mask-image:linear-gradient(to right,transparent,#000 8%,#000 92%,transparent);",
      "mask-image:linear-gradient(to right,transparent,#000 8%,#000 92%,transparent)}",
      ".fs-marcas-fila{display:flex;gap:14px;width:max-content;animation:fs-marcas-rola linear infinite}",
      ".fs-marcas-pista:hover .fs-marcas-fila{animation-play-state:paused}",
      /* Os logos vêm com fundo branco próprio, então o cartão é branco nos
         dois temas: no escuro, um cartão da cor do fundo deixaria cada logo
         como um retângulo branco solto. */
      ".fs-marcas-item{flex:0 0 auto;display:flex;align-items:center;justify-content:center;",
      "width:150px;height:92px;padding:12px;background:#fff;",
      "border:1px solid var(--border,#E6E2D8);border-radius:var(--radius,12px)}",
      ".fs-marcas-item img{max-width:100%;max-height:100%;width:auto;height:auto;object-fit:contain;",
      "filter:grayscale(1);opacity:.72;transition:filter .2s,opacity .2s}",
      ".fs-marcas-item:hover img{filter:none;opacity:1}",
      "@keyframes fs-marcas-rola{from{transform:translateX(0)}to{transform:translateX(-50%)}}",
      "@media (prefers-reduced-motion:reduce){.fs-marcas-fila{animation:none}",
      ".fs-marcas-pista{overflow-x:auto;-webkit-mask-image:none;mask-image:none}}"
    ].join("");
    document.head.appendChild(css);
  }

  function montarSecao() {
    var secao = document.createElement("section");
    secao.className = "vitrine-secao";
    secao.id = "marcas";
    secao.setAttribute("aria-labelledby", "fs-marcas-titulo");

    var titulo = document.createElement("h2");
    titulo.className = "vitrine-titulo";
    titulo.id = "fs-marcas-titulo";
    titulo.textContent = "Marcas que trabalhamos";

    var pista = document.createElement("div");
    pista.className = "fs-marcas-pista";

    var fila = document.createElement("div");
    fila.className = "fs-marcas-fila";
    fila.style.animationDuration = SEGUNDOS + "s";

    /* A lista entra duas vezes: a animação anda metade da fila e reinicia,
       então a emenda passa despercebida. A segunda cópia é decorativa. */
    for (var volta = 0; volta < 2; volta++) {
      MARCAS.forEach(function (marca) {
        var item = document.createElement("div");
        item.className = "fs-marcas-item";
        if (volta === 1) item.setAttribute("aria-hidden", "true");

        var img = document.createElement("img");
        img.src = marca.img;
        img.alt = volta === 0 ? marca.nome : "";
        img.width = 120;
        img.height = 60;
        img.loading = "lazy";
        img.decoding = "async";

        item.appendChild(img);
        fila.appendChild(item);
      });
    }

    pista.appendChild(fila);
    secao.appendChild(titulo);
    secao.appendChild(pista);
    return secao;
  }

  function inserir() {
    var sobre = document.getElementById("sobre");
    if (!sobre || !sobre.parentNode) return;
    var secao = document.getElementById("marcas");
    if (secao && secao.nextElementSibling === sobre) return;
    /* A vitrine monta as seções fora de ordem enquanto carrega os dados, e
       as que chegam depois entram antes desta. Por isso a faixa não é só
       inserida: ela é recolocada logo acima de "Quem somos" sempre que a
       ordem mudar. */
    if (!secao) { estilos(); secao = montarSecao(); }
    sobre.parentNode.insertBefore(secao, sobre);
  }

  inserir();
  /* A vitrine é montada em React: a seção "Quem somos" pode aparecer
     depois deste script e voltar a ser recriada em novas renderizações. */
  new MutationObserver(inserir).observe(document.body, { childList: true, subtree: true });
})();
