/* Farmácia Tropical — botão flutuante do WhatsApp.
   Fica fixo no canto da tela, como no site anterior. O clique é tratado
   pelo seletor-loja.js, que abre a escolha entre as três lojas: por isso o
   botão carrega a classe "fs-zap-flutuante", que aquele script intercepta.
   O href aponta para a primeira loja e só é usado se o seletor não carregar. */
(function () {
  "use strict";

  var ID = "fs-zap-flutuante";
  var LOJA_PADRAO = "(73) 8111-4509";  /* Loja Centro — apenas como reserva */
  var SAUDACAO = "Olá! Vim pelo site da Farmácia Tropical e gostaria de atendimento.";

  /* O painel da equipe não é lugar para um botão de venda. */
  var ROTAS_OCULTAS = /^#\/?(login|admin)/i;

  function estilos() {
    if (document.getElementById(ID + "-estilo")) return;
    var css = document.createElement("style");
    css.id = ID + "-estilo";
    css.textContent = [
      "." + ID + "{position:fixed;right:18px;bottom:18px;z-index:9990;",
      "display:flex;align-items:center;justify-content:center;width:56px;height:56px;",
      "border-radius:50%;background:var(--success,#00A651);color:#fff;",
      "box-shadow:0 4px 14px rgba(0,0,0,.28);transition:transform .16s,box-shadow .16s;",
      "animation:fs-zap-entra .25s ease-out}",
      "." + ID + ":hover,." + ID + ":focus-visible{transform:scale(1.07);",
      "box-shadow:0 6px 20px rgba(0,0,0,.34);outline:none}",
      "." + ID + ":focus-visible{box-shadow:0 0 0 3px #fff,0 0 0 6px var(--success,#00A651)}",
      "." + ID + " svg{width:30px;height:30px;fill:currentColor}",
      "@keyframes fs-zap-entra{from{opacity:0;transform:scale(.7)}to{opacity:1;transform:none}}",
      "@media (prefers-reduced-motion:reduce){." + ID + "{animation:none;transition:none}}",
      /* Em telas estreitas o botão desce um pouco para não cobrir os cards. */
      "@media (max-width:480px){." + ID + "{right:14px;bottom:14px;width:52px;height:52px}",
      "." + ID + " svg{width:27px;height:27px}}",
      /* Enquanto o seletor de lojas está aberto, o botão sai da frente. */
      "body:has(.fs-loja-fundo) ." + ID + "{opacity:0;pointer-events:none}"
    ].join("");
    document.head.appendChild(css);
  }

  var ICONE = '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12.04 2C6.58 2 2.13 6.45 2.13 11.91c0 1.75.46 3.45 1.32 4.95L2 22l5.25-1.38a9.9 9.9 0 0 0 4.79 1.22h.01c5.46 0 9.91-4.45 9.91-9.91 0-2.65-1.03-5.14-2.9-7.01A9.82 9.82 0 0 0 12.04 2m5.8 14.16c-.25.69-1.44 1.32-1.99 1.37-.53.05-1.02.24-3.45-.72-2.9-1.15-4.74-4.1-4.88-4.29-.14-.19-1.16-1.55-1.16-2.95s.73-2.09 1-2.38c.26-.29.57-.36.76-.36h.55c.18 0 .41-.03.64.49.24.57.81 1.97.88 2.11.07.14.12.31.02.5-.09.19-.14.31-.28.48-.14.16-.29.37-.42.49-.14.14-.28.29-.12.57.16.29.71 1.17 1.53 1.9 1.05.94 1.94 1.23 2.22 1.37.28.14.44.12.6-.07.17-.19.69-.81.88-1.09.19-.29.37-.24.62-.14.26.09 1.65.78 1.93.92.28.14.47.21.54.33.07.11.07.66-.18 1.34"/></svg>';

  function criar() {
    var a = document.createElement("a");
    a.className = ID;
    a.id = ID;
    a.href = "https://wa.me/55" + LOJA_PADRAO.replace(/\D/g, "") +
             "?text=" + encodeURIComponent(SAUDACAO);
    a.target = "_blank";
    a.rel = "noreferrer";
    a.title = "Pedir pelo WhatsApp";
    a.setAttribute("aria-label", "Pedir pelo WhatsApp");
    a.innerHTML = ICONE;
    return a;
  }

  /* O selo "Powered by Netlify" é um iframe fixo no canto inferior direito,
     com z-index máximo: sem isto ele cobre o botão e o clique não passa. */
  function posicionar() {
    var el = document.getElementById(ID);
    if (!el) return;
    var selo = document.getElementById("nl-badge-frame");
    var folga = window.innerWidth <= 480 ? 14 : 18;
    var altura = selo ? selo.getBoundingClientRect().height : 0;
    el.style.bottom = (altura ? altura + 6 : folga) + "px";
  }

  function ajustar() {
    var existente = document.getElementById(ID);
    if (ROTAS_OCULTAS.test(location.hash)) {
      if (existente) existente.remove();
      return;
    }
    if (!existente) {
      estilos();
      document.body.appendChild(criar());
    }
    posicionar();
  }

  ajustar();
  window.addEventListener("hashchange", ajustar);
  window.addEventListener("resize", posicionar);
  /* O selo é injetado depois do carregamento. */
  new MutationObserver(posicionar).observe(document.body, { childList: true });
})()
