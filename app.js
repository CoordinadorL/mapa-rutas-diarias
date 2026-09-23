(function () {
  "use strict";

  // Deja el mapa usable sin señal despues de la primera carga del dia.
  if ("serviceWorker" in navigator) {
    navigator.serviceWorker.register("sw.js").catch(function (err) {
      console.error("No se pudo activar el modo sin conexion:", err);
    });
  }

  var mapa = L.map("mapa", { zoomControl: true }).setView([-1.7, -79.0], 12);

  L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", {
    attribution: "&copy; OpenStreetMap",
    maxZoom: 19,
  }).addTo(mapa);

  var capaMarcadores = L.layerGroup().addTo(mapa);
  var ZOOM_MOSTRAR_ETIQUETAS = 16;

  function actualizarEtiquetas() {
    mapaEl.classList.toggle("mostrar-etiquetas", mapa.getZoom() >= ZOOM_MOSTRAR_ETIQUETAS);
  }

  mapa.on("zoomend", actualizarEtiquetas);

  // Boton "Mi ubicacion": muestra un punto azul con la posicion del chofer
  // y lo mantiene actualizado mientras este activado (para ver que tan
  // cerca esta de las proximas paradas). Nunca se manda a ningun lado,
  // solo se dibuja en el mapa de este celular.
  var marcadorYo = null;
  var idVigilarUbicacion = null;

  var ControlUbicacion = L.Control.extend({
    options: { position: "topleft" },
    onAdd: function () {
      var contenedor = L.DomUtil.create("div", "leaflet-bar control-ubicacion");
      var boton = L.DomUtil.create("a", "", contenedor);
      boton.href = "#";
      boton.title = "Mostrar mi ubicacion";
      boton.innerHTML = "📍";
      L.DomEvent.on(boton, "click", function (e) {
        L.DomEvent.stop(e);
        alternarMiUbicacion();
      });
      return contenedor;
    },
  });

  mapa.addControl(new ControlUbicacion());

  function alternarMiUbicacion() {
    if (idVigilarUbicacion !== null) {
      navigator.geolocation.clearWatch(idVigilarUbicacion);
      idVigilarUbicacion = null;
      if (marcadorYo) {
        mapa.removeLayer(marcadorYo);
        marcadorYo = null;
      }
      return;
    }

    if (!("geolocation" in navigator)) {
      mostrarEstadoGuardado("Este celular no puede compartir su ubicacion", true);
      return;
    }

    idVigilarUbicacion = navigator.geolocation.watchPosition(
      function (posicion) {
        var pos = [posicion.coords.latitude, posicion.coords.longitude];
        if (!marcadorYo) {
          marcadorYo = L.marker(pos, {
            icon: L.divIcon({ className: "marcador-yo-wrap", html: '<span class="marcador-yo"></span>', iconSize: [16, 16], iconAnchor: [8, 8] }),
            zIndexOffset: 1000,
          }).addTo(mapa);
          mapa.setView(pos, Math.max(mapa.getZoom(), 15));
        } else {
          marcadorYo.setLatLng(pos);
        }
      },
      function () {
        mostrarEstadoGuardado("No se pudo obtener tu ubicacion", true);
        idVigilarUbicacion = null;
      },
      { enableHighAccuracy: true, maximumAge: 5000 }
    );
  }

  var barraSuperiorEl = document.getElementById("barra-superior");
  var infoSuperiorEl = document.getElementById("info-superior");
  var mapaEl = document.getElementById("mapa");
  var selector = document.getElementById("selector-camion");
  var fechaRutaEl = document.getElementById("fecha-ruta");
  var resumenEl = document.getElementById("resumen-camion");
  var leyendaEl = document.getElementById("leyenda-colores");
  var estadoCargaEl = document.getElementById("estado-carga");
  var estadoGuardadoEl = document.getElementById("estado-guardado");
  var badgeAdminEl = document.getElementById("badge-admin");
  var pantallaCodigoEl = document.getElementById("pantalla-codigo");
  var inputCodigoEl = document.getElementById("input-codigo");
  var botonCodigoEl = document.getElementById("boton-codigo");
  var errorCodigoEl = document.getElementById("error-codigo");
  var botonCompartirEl = document.getElementById("boton-compartir");
  var panelCompartirEl = document.getElementById("panel-compartir");
  var cerrarCompartirEl = document.getElementById("cerrar-compartir");
  var inputCodigoCompartirEl = document.getElementById("input-codigo-compartir");
  var guardarCodigoCompartirEl = document.getElementById("guardar-codigo-compartir");
  var generarWhatsappEl = document.getElementById("generar-whatsapp");
  var generarQrEl = document.getElementById("generar-qr");
  var resultadoQrEl = document.getElementById("resultado-qr");
  var qrImagenEl = document.getElementById("qr-imagen");
  var descargarQrEl = document.getElementById("descargar-qr");
  var botonBuscarEl = document.getElementById("boton-buscar");
  var panelBuscarEl = document.getElementById("panel-buscar");
  var cerrarBuscarEl = document.getElementById("cerrar-buscar");
  var inputBuscarEl = document.getElementById("input-buscar");
  var resultadosBuscarEl = document.getElementById("resultados-buscar");
  var datos = null;
  var overrides = {};
  var camionActual = null;
  var marcadoresPorCodigo = {};
  var entregados = {};

  // El mapa (index.html/app.js) es publico. Los datos reales de clientes
  // viven en un repo PRIVADO aparte (mapa-rutas-datos) para no exponerlos.
  //
  // Hay dos tipos de codigo (token de ese repo privado):
  // - "codigo" (choferes): CLAVE_CODIGO, solo lectura.
  // - "admin" (encargado de rutas): CLAVE_ADMIN, lectura y escritura. Permite
  //   arrastrar los puntos del mapa para corregir su ubicacion; el cambio se
  //   guarda en overrides.json en el repo privado y todos los choferes lo ven
  //   despues (sin que tengan que hacer nada).
  //
  // Para no tener que pegar el codigo a mano (es largo), se puede mandar un
  // link que ya lo trae incluido: .../index.html#codigo=XXXX o #admin=XXXX
  // El navegador lo guarda una sola vez y despues borra el link de la barra
  // de direcciones.
  var REPO_DATOS = "CoordinadorL/mapa-rutas-datos";
  var ARCHIVO_DATOS = "latest.json";
  var ARCHIVO_OVERRIDES = "overrides.json";
  var CLAVE_CODIGO = "mrd_codigo";
  var CLAVE_ADMIN = "mrd_admin";
  var CLAVE_CODIGO_COMPARTIR = "mrd_codigo_compartir";

  var PALETA_COLORES = [
    "#e6194B", "#4363d8", "#3cb44b", "#f58231", "#911eb4",
    "#000000", "#f032e6", "#9A6324", "#800000", "#bfef45",
  ];

  function leerStorage(clave) {
    try {
      return localStorage.getItem(clave) || "";
    } catch (e) {
      return "";
    }
  }

  function guardarStorage(clave, valor) {
    try {
      localStorage.setItem(clave, valor);
    } catch (e) {}
  }

  function borrarStorage(clave) {
    try {
      localStorage.removeItem(clave);
    } catch (e) {}
  }

  function tomarCodigosDelLink() {
    var hash = window.location.hash || "";
    if (!hash) return;

    var params = new URLSearchParams(hash.replace(/^#/, ""));
    var codigo = params.get("codigo");
    var admin = params.get("admin");

    if (codigo) guardarStorage(CLAVE_CODIGO, codigo);
    if (admin) guardarStorage(CLAVE_ADMIN, admin);

    if (codigo || admin) {
      history.replaceState(null, "", window.location.pathname + window.location.search);
    }
  }

  function pedirCodigo(mostrarError) {
    errorCodigoEl.classList.toggle("oculto", !mostrarError);
    pantallaCodigoEl.classList.remove("oculto");
    estadoCargaEl.classList.add("oculto");
    inputCodigoEl.focus();
  }

  function base64Utf8(texto) {
    return btoa(unescape(encodeURIComponent(texto)));
  }

  function utf8Base64(base64) {
    return decodeURIComponent(escape(atob(base64)));
  }

  function pedirArchivoRepo(nombreArchivo, token) {
    return fetch("https://api.github.com/repos/" + REPO_DATOS + "/contents/" + nombreArchivo, {
      headers: {
        Authorization: "Bearer " + token,
        Accept: "application/vnd.github+json",
      },
      cache: "no-store",
    });
  }

  function cargarDatos(token, esAdmin) {
    estadoCargaEl.textContent = "Cargando ruta del dia...";
    estadoCargaEl.classList.remove("oculto");

    pedirArchivoRepo(ARCHIVO_DATOS, token)
      .then(function (r) {
        if (r.status === 401 || r.status === 403 || r.status === 404) {
          throw new Error("codigo-invalido");
        }
        if (!r.ok) throw new Error("No se pudo cargar la ruta del dia");
        return r.json();
      })
      .then(function (contenido) {
        datos = JSON.parse(utf8Base64(contenido.content));

        // overrides.json es opcional: si todavia no existe (nunca se ha
        // corregido ningun punto) no es un error.
        return pedirArchivoRepo(ARCHIVO_OVERRIDES, token)
          .then(function (r) {
            if (r.ok) return r.json();
            return null;
          })
          .then(function (contenidoOverrides) {
            overrides = contenidoOverrides ? JSON.parse(utf8Base64(contenidoOverrides.content)) : {};
          })
          .catch(function () {
            overrides = {};
          });
      })
      .then(function () {
        guardarStorage(esAdmin ? CLAVE_ADMIN : CLAVE_CODIGO, token);
        badgeAdminEl.classList.toggle("oculto", !esAdmin);
        botonCompartirEl.classList.toggle("oculto", !esAdmin);
        fechaRutaEl.textContent = "Ruta del " + datos.fecha;
        poblarSelector();
        if (datos.camiones.length) {
          pintarCamion(datos.camiones[0].id, esAdmin, token);
        }
        pantallaCodigoEl.classList.add("oculto");
        estadoCargaEl.classList.add("oculto");
      })
      .catch(function (err) {
        if (err.message === "codigo-invalido") {
          borrarStorage(CLAVE_CODIGO);
          borrarStorage(CLAVE_ADMIN);
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
    cargarDatos(codigo, false);
  });

  inputCodigoEl.addEventListener("keydown", function (e) {
    if (e.key === "Enter") botonCodigoEl.click();
  });

  function ajustarAltoMapa() {
    var top = barraSuperiorEl.offsetHeight + infoSuperiorEl.offsetHeight;
    infoSuperiorEl.style.top = barraSuperiorEl.offsetHeight + "px";
    mapaEl.style.top = top + "px";
    mapa.invalidateSize();
  }

  window.addEventListener("resize", ajustarAltoMapa);

  function formatoMoneda(valor) {
    if (typeof valor !== "number") return "";
    return "$" + valor.toFixed(2);
  }

  function iconoColor(color, entregado) {
    return L.divIcon({
      className: "marcador-color" + (entregado ? " entregado" : ""),
      html:
        '<span style="background:' + color + '"></span>' +
        (entregado ? '<i class="check-entregado">&#10003;</i>' : ""),
      iconSize: [18, 18],
      iconAnchor: [9, 9],
      popupAnchor: [0, -10],
    });
  }

  // "Marcar entregado" queda solo en el celular de cada quien (no se manda a
  // ningun repo): es para que el chofer no pierda la cuenta de a quien ya le
  // dejo el pedido. Se reinicia solo cada dia, junto con la ruta nueva.
  function claveEntregados() {
    return "mrd_entregados_" + (datos ? datos.fecha : "");
  }

  function cargarEntregados() {
    try {
      return JSON.parse(localStorage.getItem(claveEntregados()) || "{}");
    } catch (e) {
      return {};
    }
  }

  function marcarEntregado(codigo, entregado) {
    if (entregado) {
      entregados[codigo] = true;
    } else {
      delete entregados[codigo];
    }
    try {
      localStorage.setItem(claveEntregados(), JSON.stringify(entregados));
    } catch (e) {}
  }

  function mostrarEstadoGuardado(texto, esError) {
    estadoGuardadoEl.textContent = texto;
    estadoGuardadoEl.classList.toggle("error", !!esError);
    estadoGuardadoEl.classList.remove("oculto");
    clearTimeout(mostrarEstadoGuardado._t);
    mostrarEstadoGuardado._t = setTimeout(function () {
      estadoGuardadoEl.classList.add("oculto");
    }, 3000);
  }

  function guardarOverride(codigoCliente, lat, lng, tokenAdmin) {
    mostrarEstadoGuardado("Guardando ubicacion...", false);

    pedirArchivoRepo(ARCHIVO_OVERRIDES, tokenAdmin)
      .then(function (r) {
        return r.ok ? r.json() : null;
      })
      .then(function (actual) {
        var contenidoActual = actual ? JSON.parse(utf8Base64(actual.content)) : {};
        contenidoActual[codigoCliente] = { lat: lat, lng: lng };
        overrides = contenidoActual;

        var body = {
          message: "Corregir ubicacion cliente " + codigoCliente,
          content: base64Utf8(JSON.stringify(contenidoActual)),
        };
        if (actual && actual.sha) body.sha = actual.sha;

        return fetch("https://api.github.com/repos/" + REPO_DATOS + "/contents/" + ARCHIVO_OVERRIDES, {
          method: "PUT",
          headers: {
            Authorization: "Bearer " + tokenAdmin,
            Accept: "application/vnd.github+json",
          },
          body: JSON.stringify(body),
        });
      })
      .then(function (r) {
        if (!r.ok) throw new Error("No se pudo guardar (" + r.status + ")");
        mostrarEstadoGuardado("Ubicacion guardada", false);
      })
      .catch(function (err) {
        console.error(err);
        mostrarEstadoGuardado("Error al guardar, intenta de nuevo", true);
      });
  }

  function posicionCliente(cliente) {
    var ajuste = overrides[cliente.codigo];
    return ajuste ? [ajuste.lat, ajuste.lng] : [cliente.lat, cliente.lng];
  }

  function actualizarResumen(camion) {
    var totalVenta = camion.clientes.reduce(function (acc, c) { return acc + (c.venta || 0); }, 0);
    var totalEntregados = camion.clientes.reduce(function (acc, c) {
      return acc + (entregados[c.codigo] ? 1 : 0);
    }, 0);

    resumenEl.textContent =
      camion.clientes.length + " clientes" +
      (totalEntregados ? " · " + totalEntregados + " entregados" : "") +
      (camion.chofer ? " · " + camion.chofer : "") +
      (totalVenta ? " · " + formatoMoneda(totalVenta) : "");
    resumenEl.classList.add("visible");
  }

  function actualizarEstiloEntregado(marcador, cliente, entregado) {
    marcador.setIcon(iconoColor(cliente.color || "#4363d8", entregado));
    marcador.setTooltipContent((entregado ? "✓ " : "") + cliente.nombre);
  }

  function pintarCamion(idCamion, esAdmin, tokenAdmin) {
    capaMarcadores.clearLayers();

    var camion = datos.camiones.find(function (c) { return c.id === idCamion; });
    if (!camion) return;

    camionActual = camion;
    marcadoresPorCodigo = {};
    entregados = cargarEntregados();

    var puntos = [];

    camion.clientes.forEach(function (cliente, i) {
      var pos = posicionCliente(cliente);
      var estaEntregado = !!entregados[cliente.codigo];
      var marcador = L.marker(pos, {
        icon: iconoColor(cliente.color || "#4363d8", estaEntregado),
        draggable: !!esAdmin,
      }).addTo(capaMarcadores);

      var html =
        '<div class="popup-cliente">' +
        "<b>" + (i + 1) + ". " + escaparHtml(cliente.nombre) + "</b><br>" +
        escaparHtml(cliente.direccion) + "<br>" +
        (cliente.vendedor ? '<span class="canal">Vendedor ' + escaparHtml(cliente.vendedor) + "</span>" : "") +
        (cliente.venta ? '<div class="venta">' + formatoMoneda(cliente.venta) + "</div>" : "") +
        '<div class="fila-navegar">' +
        '<a class="btn-navegar btn-maps" target="_blank" rel="noopener" href="https://www.google.com/maps/dir/?api=1&destination=' + pos[0] + "," + pos[1] + '">Maps</a>' +
        '<a class="btn-navegar btn-waze" target="_blank" rel="noopener" href="https://waze.com/ul?ll=' + pos[0] + "," + pos[1] + '&navigate=yes">Waze</a>' +
        "</div>" +
        '<button type="button" class="btn-entregado' + (estaEntregado ? " hecho" : "") + '">' +
        (estaEntregado ? "↺ Deshacer entrega" : "✓ Marcar entregado") +
        "</button>" +
        (esAdmin ? '<div class="ayuda-admin">Arrastra el punto para corregir la ubicacion</div>' : "") +
        "</div>";
      marcador.bindPopup(html);

      marcador.bindTooltip((estaEntregado ? "✓ " : "") + cliente.nombre, {
        permanent: true,
        direction: "top",
        offset: [0, -8],
        className: "etiqueta-cliente",
      });

      marcador.on("popupopen", function () {
        var el = marcador.getPopup().getElement();
        var boton = el && el.querySelector(".btn-entregado");
        if (!boton) return;
        boton.onclick = function () {
          var nuevoEstado = !entregados[cliente.codigo];
          marcarEntregado(cliente.codigo, nuevoEstado);
          actualizarEstiloEntregado(marcador, cliente, nuevoEstado);
          boton.textContent = nuevoEstado ? "↺ Deshacer entrega" : "✓ Marcar entregado";
          boton.classList.toggle("hecho", nuevoEstado);
          actualizarResumen(camion);
        };
      });

      if (esAdmin) {
        marcador.on("dragend", function () {
          var nueva = marcador.getLatLng();
          guardarOverride(cliente.codigo, nueva.lat, nueva.lng, tokenAdmin);
        });
      }

      marcadoresPorCodigo[cliente.codigo] = marcador;
      puntos.push(pos);
    });

    if (puntos.length) {
      mapa.fitBounds(puntos, { padding: [30, 30] });
    }
    actualizarEtiquetas();
    actualizarResumen(camion);
    resumenEl.style.display = "block";

    pintarLeyenda(camion);
    ajustarAltoMapa();
  }

  function pintarLeyenda(camion) {
    var vendedores = [];
    var vistos = {};
    camion.clientes.forEach(function (c) {
      var v = c.vendedor || "";
      if (v && !vistos[v]) {
        vistos[v] = true;
        vendedores.push({ vendedor: v, color: c.color });
      }
    });

    if (vendedores.length <= 1) {
      leyendaEl.innerHTML = "";
      leyendaEl.classList.add("oculto");
      return;
    }

    leyendaEl.innerHTML = vendedores.map(function (v) {
      return '<span class="chip-vendedor"><i style="background:' + v.color + '"></i>' + escaparHtml(v.vendedor) + "</span>";
    }).join("");
    leyendaEl.classList.remove("oculto");
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
    var admin = leerStorage(CLAVE_ADMIN);
    pintarCamion(selector.value, !!admin, admin);
  });

  function linkParaChofer(codigo) {
    return window.location.origin + window.location.pathname + "#codigo=" + codigo;
  }

  function actualizarBotonesCompartir() {
    var hayCodigo = !!inputCodigoCompartirEl.value.trim();
    generarWhatsappEl.disabled = !hayCodigo;
    generarQrEl.disabled = !hayCodigo;
  }

  function pintarResultadosBuscar(texto) {
    if (!camionActual) return;

    var filtro = texto.trim().toUpperCase();
    var clientes = camionActual.clientes.filter(function (c) {
      return !filtro || c.nombre.toUpperCase().indexOf(filtro) !== -1;
    });

    if (!clientes.length) {
      resultadosBuscarEl.innerHTML = '<p class="sin-resultados">No se encontro ningun cliente.</p>';
      return;
    }

    resultadosBuscarEl.innerHTML = clientes.map(function (c) {
      return '<button type="button" class="resultado-cliente" data-codigo="' + escaparHtml(String(c.codigo)) + '">' +
        escaparHtml(c.nombre) +
        '<span class="direccion-resultado">' + escaparHtml(c.direccion) + "</span>" +
        "</button>";
    }).join("");

    Array.prototype.forEach.call(resultadosBuscarEl.querySelectorAll(".resultado-cliente"), function (boton) {
      boton.addEventListener("click", function () {
        var codigo = boton.getAttribute("data-codigo");
        var marcador = marcadoresPorCodigo[codigo];
        if (!marcador) return;
        panelBuscarEl.classList.add("oculto");
        mapa.setView(marcador.getLatLng(), 18);
        setTimeout(function () { marcador.openPopup(); }, 300);
      });
    });
  }

  botonBuscarEl.addEventListener("click", function () {
    inputBuscarEl.value = "";
    pintarResultadosBuscar("");
    panelBuscarEl.classList.remove("oculto");
    inputBuscarEl.focus();
  });

  cerrarBuscarEl.addEventListener("click", function () {
    panelBuscarEl.classList.add("oculto");
  });

  inputBuscarEl.addEventListener("input", function () {
    pintarResultadosBuscar(inputBuscarEl.value);
  });

  botonCompartirEl.addEventListener("click", function () {
    inputCodigoCompartirEl.value = leerStorage(CLAVE_CODIGO_COMPARTIR);
    actualizarBotonesCompartir();
    resultadoQrEl.classList.add("oculto");
    qrImagenEl.innerHTML = "";
    panelCompartirEl.classList.remove("oculto");
  });

  cerrarCompartirEl.addEventListener("click", function () {
    panelCompartirEl.classList.add("oculto");
  });

  inputCodigoCompartirEl.addEventListener("input", actualizarBotonesCompartir);

  guardarCodigoCompartirEl.addEventListener("click", function () {
    guardarStorage(CLAVE_CODIGO_COMPARTIR, inputCodigoCompartirEl.value.trim());
    actualizarBotonesCompartir();
    mostrarEstadoGuardado("Codigo guardado en este navegador", false);
  });

  generarWhatsappEl.addEventListener("click", function () {
    var codigo = inputCodigoCompartirEl.value.trim();
    if (!codigo) return;
    var mensaje =
      "Mapa de la ruta de hoy:\n" + linkParaChofer(codigo) +
      "\n\nAbrelo una sola vez, despues va a quedar guardado en tu celular.";
    window.open("https://wa.me/?text=" + encodeURIComponent(mensaje), "_blank");
  });

  generarQrEl.addEventListener("click", function () {
    var codigo = inputCodigoCompartirEl.value.trim();
    if (!codigo) return;

    qrImagenEl.innerHTML = "";
    new QRCode(qrImagenEl, {
      text: linkParaChofer(codigo),
      width: 220,
      height: 220,
    });

    setTimeout(function () {
      var canvas = qrImagenEl.querySelector("canvas");
      if (canvas) {
        descargarQrEl.href = canvas.toDataURL("image/png");
        descargarQrEl.classList.remove("oculto");
      } else {
        descargarQrEl.classList.add("oculto");
      }
      resultadoQrEl.classList.remove("oculto");
    }, 50);
  });

  tomarCodigosDelLink();

  var tokenAdmin = leerStorage(CLAVE_ADMIN);
  var tokenCodigo = leerStorage(CLAVE_CODIGO);

  if (tokenAdmin) {
    cargarDatos(tokenAdmin, true);
  } else if (tokenCodigo) {
    cargarDatos(tokenCodigo, false);
  } else {
    pedirCodigo(false);
  }
})();
