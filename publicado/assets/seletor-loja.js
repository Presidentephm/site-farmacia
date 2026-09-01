/* Farmácia Tropical — seletor de loja para os botões "Pedir pelo WhatsApp".
   Ao clicar em um botão de pedido, abre um modal com as três lojas
   (Centro, Arraial d'Ajuda e Trancoso) e envia o cliente para o WhatsApp
   da loja escolhida, preservando a mensagem original do link. */
(function () {
  "use strict";

  var SUPABASE_URL = "https://bmumsvrveaisyiguvdhe.supabase.co";
  var SUPABASE_KEY = "sb_publishable_yrzXUevAAQtsrxKTRahCBQ_QnoKgVTh";

  /* Usado enquanto as lojas do banco não chegam, ou se a consulta falhar. */
  var LOJAS_PADRAO = [
    { nome: "Loja Centro", whatsapp: "(73) 8111-4509", bairro: "Manoel Carneiro", cidade: "Porto Seguro", horario: "Todos os dias, 07:30 às 22:00" },
    { nome: "Loja Arraial d'Ajuda", whatsapp: "(73) 3575-2287", bairro: "Arraial d'Ajuda", cidade: "Porto Seguro", horario: "Aberta 24 horas" },
    { nome: "Loja Trancoso", whatsapp: "(73) 3668-1770", bairro: "Trancoso", cidade: "Porto Seguro", horario: "Seg a Sáb 07:00 às 22:00 · Dom 07:00 às 20:00" }
  ];

  var SELETOR_BOTOES = "a.vitrine-pedir, a.vitrine-zap, a.vitrine-diferencial-link, a.fs-zap-flutuante";

  var lojas = LOJAS_PADRAO;

  function carregarLojas() {
    var url = SUPABASE_URL + "/rest/v1/fs_lojas" +
      "?select=nome,whatsapp,bairro,cidade,horario,ordem&ativa=eq.true&order=ordem";
    fetch(url, { headers: { apikey: SUPABASE_KEY } })
      .then(function (r) { return r.ok ? r.json() : null; })
      .then(function (dados) {
        if (!dados || !dados.length) return;
        var comZap = dados.filter(function (l) { return l && l.whatsapp; });
        if (comZap.length) lojas = comZap;
      })
      .catch(function () { /* mantém as lojas padrão */ });
  }

  function estilos() {
    if (document.getElementById("fs-seletor-loja-estilo")) return;
    var css = document.createElement("style");
    css.id = "fs-seletor-loja-estilo";
    css.textContent = [
      ".fs-loja-fundo{position:fixed;inset:0;z-index:9999;display:flex;align-items:center;justify-content:center;",
      "padding:16px;background:rgba(20,15,10,.55);backdrop-filter:blur(2px);animation:fs-loja-surge .15s ease-out}",
      ".fs-loja-caixa{width:100%;max-width:420px;max-height:88vh;overflow-y:auto;background:var(--card,#fff);",
      "color:var(--text,#241F1A);border:1px solid var(--border,#E6E2D8);border-radius:var(--radius,12px);",
      "box-shadow:var(--shadow-lg,0 8px 30px rgba(0,0,0,.25));animation:fs-loja-sobe .18s ease-out}",
      ".fs-loja-topo{display:flex;align-items:flex-start;gap:12px;padding:16px 16px 8px}",
      ".fs-loja-topo h2{margin:0;font-size:1.05rem;font-weight:800;line-height:1.3}",
      ".fs-loja-topo p{margin:4px 0 0;font-size:.82rem;color:var(--text-2,#5F584F);line-height:1.4}",
      ".fs-loja-fechar{margin-left:auto;flex:0 0 auto;width:32px;height:32px;border:0;border-radius:50%;cursor:pointer;",
      "background:transparent;color:var(--text-2,#5F584F);font-size:22px;line-height:1}",
      ".fs-loja-fechar:hover{background:var(--bg,#F7F6F2)}",
      ".fs-loja-lista{display:flex;flex-direction:column;gap:8px;padding:8px 16px 16px}",
      ".fs-loja-item{display:flex;align-items:center;gap:10px;width:100%;text-align:left;cursor:pointer;",
      "padding:12px 14px;border:1px solid var(--border,#E6E2D8);border-radius:var(--radius,12px);",
      "background:var(--card,#fff);color:inherit;font:inherit;transition:border-color .12s,transform .12s}",
      ".fs-loja-item:hover,.fs-loja-item:focus-visible{border-color:var(--success,#00A651);transform:translateY(-1px);outline:none}",
      ".fs-loja-item-txt{min-width:0}",
      ".fs-loja-item-txt strong{display:block;font-size:.94rem;font-weight:700}",
      ".fs-loja-item-txt span{display:block;font-size:.78rem;color:var(--text-2,#5F584F);margin-top:2px}",
      ".fs-loja-zap{flex:0 0 auto;margin-left:auto;display:inline-flex;align-items:center;justify-content:center;",
      "width:34px;height:34px;border-radius:50%;background:var(--success,#00A651);color:#fff}",
      ".fs-loja-zap svg{width:18px;height:18px;fill:currentColor}",
      "@keyframes fs-loja-surge{from{opacity:0}to{opacity:1}}",
      "@keyframes fs-loja-sobe{from{opacity:0;transform:translateY(12px)}to{opacity:1;transform:none}}",
      "@media (max-width:480px){.fs-loja-fundo{align-items:flex-end;padding:0}",
      ".fs-loja-caixa{max-width:none;border-radius:16px 16px 0 0;max-height:82vh}}"
    ].join("");
    document.head.appendChild(css);
  }

  var ICONE_ZAP = '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12.04 2C6.58 2 2.13 6.45 2.13 11.91c0 1.75.46 3.45 1.32 4.95L2 22l5.25-1.38a9.9 9.9 0 0 0 4.79 1.22h.01c5.46 0 9.91-4.45 9.91-9.91 0-2.65-1.03-5.14-2.9-7.01A9.82 9.82 0 0 0 12.04 2m5.8 14.16c-.25.69-1.44 1.32-1.99 1.37-.53.05-1.02.24-3.45-.72-2.9-1.15-4.74-4.1-4.88-4.29-.14-.19-1.16-1.55-1.16-2.95s.73-2.09 1-2.38c.26-.29.57-.36.76-.36h.55c.18 0 .41-.03.64.49.24.57.81 1.97.88 2.11.07.14.12.31.02.5-.09.19-.14.31-.28.48-.14.16-.29.37-.42.49-.14.14-.28.29-.12.57.16.29.71 1.17 1.53 1.9 1.05.94 1.94 1.23 2.22 1.37.28.14.44.12.6-.07.17-.19.69-.81.88-1.09.19-.29.37-.24.62-.14.26.09 1.65.78 1.93.92.28.14.47.21.54.33.07.11.07.66-.18 1.34"/></svg>';

  function digitos(v) { return String(v || "").replace(/\D/g, ""); }

  function linkDaLoja(loja, hrefOriginal) {
    var consulta = "";
    var pos = hrefOriginal.indexOf("?");
    if (pos !== -1) consulta = hrefOriginal.slice(pos);
    return "https://wa.me/55" + digitos(loja.whatsapp) + consulta;
  }

  function fechar(fundo, focoAnterior) {
    if (!fundo || !fundo.parentNode) return;
    fundo.parentNode.removeChild(fundo);
    document.body.style.overflow = "";
    if (focoAnterior && focoAnterior.focus) focoAnterior.focus();
  }

  function abrirSeletor(hrefOriginal) {
    estilos();
    var focoAnterior = document.activeElement;

    var fundo = document.createElement("div");
    fundo.className = "fs-loja-fundo";

    var caixa = document.createElement("div");
    caixa.className = "fs-loja-caixa";
    caixa.setAttribute("role", "dialog");
    caixa.setAttribute("aria-modal", "true");
    caixa.setAttribute("aria-label", "Escolha a loja para pedir pelo WhatsApp");

    var topo = document.createElement("div");
    topo.className = "fs-loja-topo";
    var titulo = document.createElement("div");
    titulo.innerHTML = "<h2>Pedir pelo WhatsApp</h2><p>Escolha a loja que vai atender você.</p>";
    var btnFechar = document.createElement("button");
    btnFechar.className = "fs-loja-fechar";
    btnFechar.type = "button";
    btnFechar.setAttribute("aria-label", "Fechar");
    btnFechar.innerHTML = "&times;";
    btnFechar.addEventListener("click", function () { fechar(fundo, focoAnterior); });
    topo.appendChild(titulo);
    topo.appendChild(btnFechar);

    var lista = document.createElement("div");
    lista.className = "fs-loja-lista";

    lojas.forEach(function (loja) {
      if (!digitos(loja.whatsapp)) return;
      var item = document.createElement("button");
      item.className = "fs-loja-item";
      item.type = "button";

      var texto = document.createElement("div");
      texto.className = "fs-loja-item-txt";
      var nome = document.createElement("strong");
      nome.textContent = loja.nome || "Loja";
      texto.appendChild(nome);

      var local = [loja.bairro, loja.cidade].filter(Boolean).join(" · ");
      [local, loja.horario].forEach(function (linha) {
        if (!linha) return;
        var s = document.createElement("span");
        s.textContent = linha;
        texto.appendChild(s);
      });

      var zap = document.createElement("span");
      zap.className = "fs-loja-zap";
      zap.innerHTML = ICONE_ZAP;

      item.appendChild(texto);
      item.appendChild(zap);
      item.addEventListener("click", function () {
        var destino = linkDaLoja(loja, hrefOriginal);
        fechar(fundo, null);
        window.open(destino, "_blank", "noopener");
      });
      lista.appendChild(item);
    });

    caixa.appendChild(topo);
    caixa.appendChild(lista);
    fundo.appendChild(caixa);

    fundo.addEventListener("click", function (ev) {
      if (ev.target === fundo) fechar(fundo, focoAnterior);
    });
    document.addEventListener("keydown", function aoTeclar(ev) {
      if (ev.key === "Escape") {
        document.removeEventListener("keydown", aoTeclar);
        fechar(fundo, focoAnterior);
      } else if (!fundo.parentNode) {
        document.removeEventListener("keydown", aoTeclar);
      }
    });

    document.body.appendChild(fundo);
    document.body.style.overflow = "hidden";
    var primeiro = lista.querySelector(".fs-loja-item");
    if (primeiro) primeiro.focus();
  }

  document.addEventListener("click", function (ev) {
    if (ev.defaultPrevented || ev.button !== 0 || ev.metaKey || ev.ctrlKey || ev.shiftKey || ev.altKey) return;
    var alvo = ev.target.closest ? ev.target.closest(SELETOR_BOTOES) : null;
    if (!alvo) return;
    /* Na seção "Nossas lojas" o cliente já escolheu a unidade. */
    if (alvo.closest(".vitrine-loja")) return;
    if (lojas.length < 2) return;
    ev.preventDefault();
    abrirSeletor(alvo.getAttribute("href") || "");
  }, true);

  carregarLojas();
})();
