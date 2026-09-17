(function () {
  "use strict";

  var mapa = L.map("mapa", { zoomControl: true }).setView([-1.7, -79.0], 12);

  L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
    attribution: "&copy; OpenStreetMap",
    maxZoom: 19,
  }).addTo(mapa);

  var capaMarcadores = L.layerGroup().addTo(mapa);
  var selector = document.getElementById("selector-camion");
  var fechaRutaEl = document.getElementById("fecha-ruta");
  var resumenEl = document.getElementById("resumen-camion");
  var estadoCargaEl = document.getElementById("estado-carga");
  var datos = null;

  function formatoMoneda(valor) {
    if (typeof valor !== "number") return "";
    return "$" + valor.toFixed(2);
  }

  function pintarCamion(idCamion) {
    capaMarcadores.clearLayers();

    var camion = datos.camiones.find(function (c) { return c.id === idCamion; });
    if (!camion) return;

    var puntos = [];

    camion.clientes.forEach(function (cliente, i) {
      var marcador = L.marker([cliente.lat, cliente.lng]).addTo(capaMarcadores);
      var html =
        '<div class="popup-cliente">' +
        "<b>" + (i + 1) + ". " + escaparHtml(cliente.nombre) + "</b><br>" +
        escaparHtml(cliente.direccion) + "<br>" +
        '<span class="canal">' + escaparHtml(cliente.canal) + "</span>" +
        (cliente.venta ? '<div class="venta">' + formatoMoneda(cliente.venta) + "</div>" : "") +
        "</div>";
      marcador.bindPopup(html);
      puntos.push([cliente.lat, cliente.lng]);
    });

    if (puntos.length) {
      mapa.fitBounds(puntos, { padding: [30, 30] });
    }

    var totalVenta = camion.clientes.reduce(function (acc, c) { return acc + (c.venta || 0); }, 0);
    resumenEl.textContent =
      camion.clientes.length + " clientes" +
      (camion.chofer ? " · " + camion.chofer : "") +
      (totalVenta ? " · " + formatoMoneda(totalVenta) : "");
    resumenEl.classList.add("visible");
    resumenEl.style.display = "block";
  }

  function escaparHtml(texto) {
    if (!texto) return "";
    var div = document.createElement("div");
    div.textContent = texto;
    return div.innerHTML;
  }

  function poblarSelector() {
    selector.innerHTML = "";
    datos.camiones.forEach(function (camion) {
      var opcion = document.createElement("option");
      opcion.value = camion.id;
      opcion.textContent = camion.id + " (" + camion.clientes.length + ")";
      selector.appendChild(opcion);
    });
  }

  selector.addEventListener("change", function () {
    pintarCamion(selector.value);
  });

  fetch("data/latest.json?_=" + Date.now())
    .then(function (r) {
      if (!r.ok) throw new Error("No se pudo cargar data/latest.json");
      return r.json();
    })
    .then(function (json) {
      datos = json;
      fechaRutaEl.textContent = "Ruta del " + datos.fecha;
      poblarSelector();
      if (datos.camiones.length) {
        pintarCamion(datos.camiones[0].id);
      }
      estadoCargaEl.classList.add("oculto");
    })
    .catch(function (err) {
      estadoCargaEl.textContent = "Error cargando la ruta del dia. Avisa al encargado.";
      console.error(err);
    });
})();
