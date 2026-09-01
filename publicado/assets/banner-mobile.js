/* Farmácia Tropical — dois ajustes no carrossel de banners, só no celular.

   1) Ele parava de passar sozinho. A vitrine pausa o rodízio no mouseenter e
      só retoma no mouseleave; num aparelho de toque o mouseenter dispara no
      primeiro toque e o mouseleave nunca vem, então o carrossel congelava.
      Aqui um vigia percebe que o banner ficou parado tempo demais e toca no
      próximo pontinho.

   2) As artes apareciam cortadas nas laterais. Até 640px o quadro do banner
      fica mais alto (1200/520) do que as artes (1200/380) e o recorte é
      "cover", que amplia a imagem até preencher a altura. Aqui o quadro passa
      a seguir a proporção da própria imagem, então nada é cortado — e serve
      para qualquer arte que venha a ser cadastrada.

   No computador nada disso roda: o rodízio e o recorte originais continuam. */
(function () {
  "use strict";

  var PAUSA_MAXIMA = 7500;   /* o rodízio da vitrine troca a cada 6s */
  var INTERVALO = 1000;

  var semHover = window.matchMedia("(hover: none)");
  var telaPequena = window.matchMedia("(max-width: 640px)");

  var ultimoIndice = -1;
  var paradoDesde = Date.now();

  function pontos() { return document.querySelectorAll(".banner-pontos button"); }

  function indiceAtivo(lista) {
    for (var i = 0; i < lista.length; i++) {
      if (lista[i].classList.contains("ativo")) return i;
    }
    return -1;
  }

  function destravarRodizio() {
    if (!semHover.matches) return;
    var lista = pontos();
    if (lista.length < 2) return;
    var atual = indiceAtivo(lista);
    if (atual < 0) return;
    if (atual !== ultimoIndice) {
      ultimoIndice = atual;
      paradoDesde = Date.now();
      return;
    }
    if (Date.now() - paradoDesde >= PAUSA_MAXIMA) {
      paradoDesde = Date.now();
      lista[(atual + 1) % lista.length].click();
    }
  }

  function ajustarProporcao() {
    var quadro = document.querySelector(".banner-quadro");
    if (!quadro) return;
    if (!telaPequena.matches) { quadro.style.aspectRatio = ""; return; }
    var img = quadro.querySelector(".banner-img");
    if (!img || !img.naturalWidth || !img.naturalHeight) return;
    quadro.style.aspectRatio = img.naturalWidth + " / " + img.naturalHeight;
  }

  setInterval(function () {
    destravarRodizio();
    ajustarProporcao();
  }, INTERVALO);

  window.addEventListener("resize", ajustarProporcao);
  document.addEventListener("load", ajustarProporcao, true);  /* imagem recém-carregada */
  ajustarProporcao();
})();
