(function () {
  "use strict";

  var WHATS_NUMBER = "5591999990000"; // fictício — projeto de portfólio
  var MSG_PADRAO = "Olá! Vim pelo site e quero um orçamento de ar-condicionado.";

  function linkWhats(mensagem) {
    return "https://wa.me/" + WHATS_NUMBER + "?text=" + encodeURIComponent(mensagem);
  }

  document.querySelectorAll(
    "#cta-header, #cta-hero, #cta-final, #whats-float"
  ).forEach(function (el) {
    el.href = linkWhats(MSG_PADRAO);
  });

  // ---- menu mobile ----
  var toggle = document.getElementById("nav-toggle");
  var mnav = document.getElementById("mnav");
  if (toggle && mnav) {
    toggle.addEventListener("click", function () {
      var open = mnav.classList.toggle("open");
      toggle.setAttribute("aria-expanded", open ? "true" : "false");
    });
    mnav.querySelectorAll("a").forEach(function (a) {
      a.addEventListener("click", function () {
        mnav.classList.remove("open");
        toggle.setAttribute("aria-expanded", "false");
      });
    });
  }

  // ---- calculadora de BTUs ----
  var TIERS = [7000, 9000, 12000, 18000, 21000, 24000, 30000, 36000, 48000, 60000];

  function tierAcima(valor) {
    for (var i = 0; i < TIERS.length; i++) {
      if (TIERS[i] >= valor) return TIERS[i];
    }
    return TIERS[TIERS.length - 1];
  }

  var form = document.getElementById("calc-form");
  if (form) {
    form.addEventListener("submit", function (e) {
      e.preventDefault();

      var area = parseFloat(document.getElementById("area").value) || 0;
      var pessoas = parseInt(document.getElementById("pessoas").value, 10) || 1;
      var ambiente = document.getElementById("ambiente").value;
      var sol = document.getElementById("sol").checked;

      var base = area * 600;
      if (ambiente === "cozinha") base *= 1.3;
      if (ambiente === "loja") base *= 1.1;
      if (pessoas > 2) base += (pessoas - 2) * 600;
      if (sol) base *= 1.15;

      var recomendado = tierAcima(base);

      var resultBox = document.getElementById("calc-result");
      var resultValue = document.getElementById("calc-value");
      resultValue.textContent = recomendado.toLocaleString("pt-BR") + " BTUs";
      resultBox.hidden = false;

      var msg =
        "Olá! Usei a calculadora do site: ambiente de " + area +
        " m², " + pessoas + " pessoa(s)" +
        (sol ? ", com sol direto" : "") +
        ". Recomendação: " + recomendado.toLocaleString("pt-BR") +
        " BTUs. Quero um orçamento.";
      document.getElementById("cta-calc").href = linkWhats(msg);
    });
  }
})();
