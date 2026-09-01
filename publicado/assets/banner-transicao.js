/* Farmácia Tropical — transição suave entre os banners.

   A vitrine troca o `src` da mesma imagem, então a mudança era seca. Aqui,
   ao perceber a troca, uma cópia da imagem anterior é sobreposta e some em
   meio segundo, o que produz o cruzamento entre as duas artes.

   As artes são pré-carregadas: sem isso, a primeira troca mostraria o fundo
   vazio do cartão enquanto a imagem nova ainda estivesse baixando. */
(function () {
  "use strict";

  var SUPABASE_URL = "https://bmumsvrveaisyiguvdhe.supabase.co";
  var SUPABASE_KEY = "sb_publishable_yrzXUevAAQtsrxKTRahCBQ_QnoKgVTh";
  var DURACAO = 500;

  var semAnimacao = window.matchMedia("(prefers-reduced-motion: reduce)");

  function estilos() {
    if (document.getElementById("fs-banner-transicao-estilo")) return;
    var css = document.createElement("style");
    css.id = "fs-banner-transicao-estilo";
    css.textContent = [
      ".fs-banner-fade{position:absolute;inset:0;z-index:1;pointer-events:none;",
      "background-size:cover;background-position:center;opacity:1;",
      "transition:opacity " + DURACAO + "ms ease}",
      ".fs-banner-fade.saindo{opacity:0}",
      /* As setas e os pontinhos precisam continuar acima da cópia que some. */
      ".banner-seta,.banner-pontos{z-index:2}"
    ].join("");
    document.head.appendChild(css);
  }

  function precarregar() {
    var url = SUPABASE_URL + "/rest/v1/fs_banners?select=imagem_url&ativo=eq.true&order=ordem";
    fetch(url, { headers: { apikey: SUPABASE_KEY } })
      .then(function (r) { return r.ok ? r.json() : []; })
      .then(function (linhas) {
        (linhas || []).forEach(function (l) {
          if (l && l.imagem_url) { var i = new Image(); i.src = l.imagem_url; }
        });
      })
      .catch(function () { /* sem pré-carga a transição ainda funciona */ });
  }

  function cruzar(quadro, urlAnterior) {
    estilos();
    var capa = document.createElement("div");
    capa.className = "fs-banner-fade";
    capa.style.backgroundImage = 'url("' + urlAnterior + '")';
    quadro.appendChild(capa);
    /* Lê uma propriedade de layout para o navegador registrar o estado
       inicial: sem isso ele agruparia as duas opacidades e não animaria. */
    void capa.offsetWidth;
    capa.classList.add("saindo");
    setTimeout(function () { if (capa.parentNode) capa.remove(); }, DURACAO + 120);
  }

  function vigiar() {
    var quadro = document.querySelector(".banner-quadro");
    if (!quadro) return false;
    new MutationObserver(function (registros) {
      if (semAnimacao.matches) return;
      registros.forEach(function (r) {
        if (r.attributeName !== "src") return;
        var anterior = r.oldValue;
        var atual = r.target.getAttribute("src");
        if (!anterior || anterior === atual) return;
        cruzar(quadro, anterior);
      });
    }).observe(quadro, { subtree: true, attributes: true,
                         attributeFilter: ["src"], attributeOldValue: true });
    return true;
  }

  precarregar();
  /* A vitrine monta o banner depois de buscar os dados. */
  if (!vigiar()) {
    var esperando = new MutationObserver(function () {
      if (vigiar()) esperando.disconnect();
    });
    esperando.observe(document.body, { childList: true, subtree: true });
  }
})();
