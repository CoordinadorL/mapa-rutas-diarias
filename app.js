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
  var pantallaCodigoEl = document.getElementById("pantalla-codigo");
  var inputCodigoEl = document.getElementById("input-codigo");
  var botonCodigoEl = document.getElementById("boton-codigo");
  var errorCodigoEl = document.getElementById("error-codigo");
  var datos = null;

  // El mapa (index.html/app.js) es publico. Los datos reales de clientes
  // viven en un repo PRIVADO aparte (mapa-rutas-datos) para no exponerlos.
  // Este codigo de acceso (token de solo lectura de ese repo) lo pide una
  // sola vez el navegador y despues queda guardado en ese celular.
  var REPO_DATOS = "CoordinadorL/mapa-rutas-datos";
  var ARCHIVO_DATOS = "latest.json";
  var CLAVE_STORAGE = "mrd_codigo";

  function obtenerCodigoGuardado() {
    try {
      return localStorage.getItem(CLAVE_STORAGE) || "";
    } catch (e) {
      return "";
    }
  }

  function guardarCodigo(codigo) {
    try {
      localStorage.setItem(CLAVE_STORAGE, codigo);
    } catch (e) {
      // si el navegador bloquea localStorage, sigue funcionando igual
      // pero va a pedir el codigo de nuevo la proxima vez
    }
  }

  function borrarCodigoGuardado() {
    try {
      localStorage.removeItem(CLAVE_STORAGE);
    } catch (e) {}
  }

  function pedirCodigo(mostrarError) {
    errorCodigoEl.classList.toggle("oculto", !mostrarError);
    pantallaCodigoEl.classList.remove("oculto");
    estadoCargaEl.classList.add("oculto");
    inputCodigoEl.focus();
  }

  function cargarDatos(codigo) {
    estadoCargaEl.textContent = "Cargando ruta del dia...";
    estadoCargaEl.classList.remove("oculto");

    fetch("https://api.github.com/repos/" + REPO_DATOS + "/contents/" + ARCHIVO_DATOS, {
      headers: {
        Authorization: "Bearer " + codigo,
        Accept: "application/vnd.github.raw+json",
      },
      cache: "no-store",
    })
      .then(function (r) {
        if (r.status === 401 || r.status === 403 || r.status === 404) {
          throw new Error("codigo-invalido");
        }
        if (!r.ok) throw new Error("No se pudo cargar la ruta del dia");
        return r.json();
      })
      .then(function (json) {
        guardarCodigo(codigo);
        datos = json;
        fechaRutaEl.textContent = "Ruta del " + datos.fecha;
        poblarSelector();
        if (datos.camiones.length) {
          pintarCamion(datos.camiones[0].id);
        }
        pantallaCodigoEl.classList.add("oculto");
        estadoCargaEl.classList.add("oculto");
      })
      .catch(function (err) {
        if (err.message === "codigo-invalido") {
          borrarCodigoGuardado();
          pedirCodigo(true);
        } else {
          estadoCargaEl.textContent = "Error cargando la ruta del dia. Avisa al encargado.";
        }
        console.error(err);
      });
  }

  botonCodigoEl.addEventListener("click", function () {
    var codigo = inputCodigoEl.value.trim();
    if (!codigo) return;
    cargarDatos(codigo);
  });

  inputCodigoEl.addEventListener("keydown", function (e) {
    if (e.key === "Enter") botonCodigoEl.click();
  });

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

  var codigoGuardado = obtenerCodigoGuardado();

  if (codigoGuardado) {
    cargarDatos(codigoGuardado);
  } else {
    pedirCodigo(false);
  }
})();
