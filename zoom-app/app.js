/* ============================================================
   Zoom Agricultura — Detalles Administración
   Lógica de la app: lectura de Excel, filtros, ajuste de precios,
   generación de PDF, exportación de Excel actualizado, historial
   compartido (Firebase Firestore) o local (localStorage).
   ============================================================ */

// ---------- Detección dinámica de columnas por nombre de encabezado ----------
// En vez de asumir una posición fija (B, C, D...), la app busca en las primeras
// filas de cada solapa cuál es la fila de encabezados y qué columna corresponde
// a cada dato, por el texto del título. Así funciona sin importar el orden de
// columnas del Excel, y no se rompe si el año que viene se reordenan o se
// agregan/sacan columnas (como pasó de "Registro Zoom 25-26" a "26-27").
const ALIAS_ENCABEZADOS = {
  fecha: ["fecha"],
  tipo: ["tipo de trabajo"],
  cliente: ["cliente"],
  establecimiento: ["establecimiento"],
  cantidad: ["superficie", "dia", "dias", "días"],
  precioUnidad: ["$/ha", "$/dia", "$/día"],
  distancia: ["distancia"],
  precioKm: ["$/km"],
  subtotalUnidad: ["subtotal", "trabajo"],
  subtotalMov: ["movilidad"],
  importe: ["importe"]
};

function normalizarTexto(s) {
  return String(s).trim().toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "");
}

// Busca la fila de encabezados (la que tiene una celda "Fecha") entre las
// primeras filas de la solapa, y arma el mapa campo -> índice de columna.
function detectarColumnas(ws) {
  if (!ws || !ws["!ref"]) return null;
  const range = XLSX.utils.decode_range(ws["!ref"]);
  let headerRow = -1;
  for (let r = range.s.r; r <= Math.min(range.s.r + 10, range.e.r); r++) {
    for (let c = range.s.c; c <= range.e.c; c++) {
      const v = getCell(ws, r, c);
      if (typeof v === "string" && normalizarTexto(v) === "fecha") { headerRow = r; break; }
    }
    if (headerRow >= 0) break;
  }
  if (headerRow === -1) return null;

  const cols = {};
  for (let c = range.s.c; c <= range.e.c; c++) {
    const v = getCell(ws, headerRow, c);
    if (typeof v !== "string") continue;
    const norm = normalizarTexto(v);
    for (const campo of Object.keys(ALIAS_ENCABEZADOS)) {
      if (cols[campo] !== undefined) continue;
      if (ALIAS_ENCABEZADOS[campo].includes(norm)) cols[campo] = c;
    }
  }
  // Campos imprescindibles para poder leer la solapa.
  if (cols.fecha === undefined || cols.cliente === undefined || cols.tipo === undefined) return null;

  return { headerRow, dataStartRow: headerRow + 1, cols };
}

const HOJAS = [
  { nombre: "Trabajos", unidadLabel: "ha", cantidadLabel: "Ha", origen: "Fact. 1 (Ha)" },
  { nombre: "Trabajos 2", unidadLabel: "dia", cantidadLabel: "Días", origen: "Fact. 2 (Día)" }
];

// ---------- Estado global ----------
let workbookActual = null;
let nombreArchivoActual = "";
let registros = [];        // todos los registros parseados
let resultadosActuales = []; // registros que matchean el último filtro (con precios editables)

// ---------- Utilidades ----------
function getCell(ws, r, c) {
  const addr = XLSX.utils.encode_cell({ r, c });
  return ws[addr] ? ws[addr].v : undefined;
}
function setCellNum(ws, r, c, value) {
  const addr = XLSX.utils.encode_cell({ r, c });
  if (!ws[addr]) ws[addr] = { t: "n", v: value };
  else { ws[addr].v = value; ws[addr].t = "n"; }
}
function esVacio(v) {
  return v === undefined || v === null || (typeof v === "string" && v.trim() === "");
}
function formatoMoneda(n) {
  return new Intl.NumberFormat("es-AR", { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(n || 0);
}
function formatoFecha(d) {
  if (!(d instanceof Date) || isNaN(d)) return "";
  return d.toLocaleDateString("es-AR", { timeZone: "UTC" });
}
function claveReferencia(cliente, tipo) {
  return `${cliente}||${tipo}`.toLowerCase();
}

// ---------- Parseo del Excel ----------
function parsearWorkbook(wb) {
  const out = [];
  const hojasNoDetectadas = [];
  HOJAS.forEach((hoja) => {
    const ws = wb.Sheets[hoja.nombre];
    if (!ws || !ws["!ref"]) return;
    const det = detectarColumnas(ws);
    if (!det) { hojasNoDetectadas.push(hoja.nombre); return; }
    const cols = det.cols;
    const range = XLSX.utils.decode_range(ws["!ref"]);

    for (let r = det.dataStartRow; r <= range.e.r; r++) {
      const fecha = getCell(ws, r, cols.fecha);
      const cliente = getCell(ws, r, cols.cliente);
      const tipo = getCell(ws, r, cols.tipo);
      if (esVacio(fecha) || esVacio(cliente) || esVacio(tipo)) continue;
      if (!(fecha instanceof Date)) continue;

      const cantidad = Number(getCell(ws, r, cols.cantidad)) || 0;
      const precioUnidad = Number(getCell(ws, r, cols.precioUnidad)) || 0;
      const distancia = cols.distancia !== undefined ? Number(getCell(ws, r, cols.distancia)) || 0 : 0;
      const precioKm = cols.precioKm !== undefined ? Number(getCell(ws, r, cols.precioKm)) || 0 : 0;

      out.push({
        sheet: hoja.nombre,
        rowIdx: r,
        cols, // se guarda para poder escribir de vuelta en la columna correcta al exportar
        origen: hoja.origen,
        unidadLabel: hoja.unidadLabel,
        cantidadLabel: hoja.cantidadLabel,
        fecha,
        cliente: String(cliente).trim(),
        establecimiento: String(getCell(ws, r, cols.establecimiento) || "").trim(),
        tipo: String(tipo).trim(),
        cantidad,
        precioUnidadBase: precioUnidad,
        distancia,
        precioKmBase: precioKm,
        // valores "actuales" editables (arrancan iguales a los del excel)
        precioUnidadActual: precioUnidad,
        precioKmActual: precioKm
      });
    }
  });
  if (hojasNoDetectadas.length) {
    console.warn("No se pudieron detectar los encabezados en:", hojasNoDetectadas.join(", "));
  }
  return out;
}

function poblarFiltros() {
  const clientes = [...new Set(registros.map((r) => r.cliente))].sort((a, b) => a.localeCompare(b, "es"));
  const establecimientos = [...new Set(registros.map((r) => r.establecimiento).filter(Boolean))].sort((a, b) => a.localeCompare(b, "es"));
  const tipos = [...new Set(registros.map((r) => r.tipo).filter(Boolean))].sort((a, b) => a.localeCompare(b, "es"));

  const selCliente = document.getElementById("filtroCliente");
  const selEst = document.getElementById("filtroEstablecimiento");
  const selTipo = document.getElementById("filtroTipo");
  selCliente.innerHTML = clientes.map((c) => `<option value="${escapeHtml(c)}">${escapeHtml(c)}</option>`).join("");
  selEst.innerHTML = establecimientos.map((e) => `<option value="${escapeHtml(e)}">${escapeHtml(e)}</option>`).join("");
  selTipo.innerHTML = tipos.map((t) => `<option value="${escapeHtml(t)}">${escapeHtml(t)}</option>`).join("");
}

function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (m) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[m]));
}

// ---------- Carga del Excel ----------
async function cargarExcelDesdeArrayBuffer(buffer, nombre) {
  const wb = XLSX.read(buffer, { type: "array", cellDates: true });
  workbookActual = wb;
  nombreArchivoActual = nombre;
  registros = parsearWorkbook(wb);
  document.getElementById("estadoExcel").textContent =
    `${nombre} — ${registros.length} trabajos encontrados (Fact. 1 + Fact. 2)`;
  poblarFiltros();
  resultadosActuales = [];
  renderResultados([]);
}

async function cargarExcelBase() {
  document.getElementById("estadoExcel").textContent = "Cargando archivo base…";
  try {
    const resp = await fetch("assets/registro-zoom-base.xlsx");
    if (!resp.ok) throw new Error("No se pudo obtener el archivo base.");
    const buffer = await resp.arrayBuffer();
    await cargarExcelDesdeArrayBuffer(buffer, "Registro Zoom (base por defecto)");
  } catch (err) {
    console.warn("No se pudo cargar el Excel base automáticamente:", err);
    document.getElementById("estadoExcel").textContent =
      "No se pudo cargar el Excel base automáticamente (esto pasa si abrís el archivo con doble clic en vez de una URL publicada). Subí el Excel manualmente con el botón de la derecha.";
  }
}

document.getElementById("inputExcel").addEventListener("change", async (e) => {
  const file = e.target.files[0];
  if (!file) return;
  const buffer = await file.arrayBuffer();
  await cargarExcelDesdeArrayBuffer(buffer, file.name);
});

document.getElementById("btnUsarBase").addEventListener("click", () => {
  document.getElementById("inputExcel").value = "";
  cargarExcelBase();
});

// ---------- Filtros / búsqueda ----------
document.getElementById("btnBuscar").addEventListener("click", async () => {
  const desde = document.getElementById("filtroDesde").value ? new Date(document.getElementById("filtroDesde").value + "T00:00:00Z") : null;
  const hasta = document.getElementById("filtroHasta").value ? new Date(document.getElementById("filtroHasta").value + "T23:59:59Z") : null;
  const clientesSel = Array.from(document.getElementById("filtroCliente").selectedOptions).map((o) => o.value);
  const estSel = Array.from(document.getElementById("filtroEstablecimiento").selectedOptions).map((o) => o.value);
  const tipoSel = Array.from(document.getElementById("filtroTipo").selectedOptions).map((o) => o.value);

  let filtrados = registros.filter((r) => {
    if (desde && r.fecha < desde) return false;
    if (hasta && r.fecha > hasta) return false;
    if (clientesSel.length && !clientesSel.includes(r.cliente)) return false;
    if (estSel.length && !estSel.includes(r.establecimiento)) return false;
    if (tipoSel.length && !tipoSel.includes(r.tipo)) return false;
    return true;
  });

  // Aplicar precios guardados (compartidos) si existen para cliente+tipo
  for (const r of filtrados) {
    const guardado = await ZoomStore.getPrecio(claveReferencia(r.cliente, r.tipo));
    if (guardado) {
      r.precioUnidadActual = guardado.precioUnidad;
      r.precioKmActual = guardado.precioKm;
    } else {
      r.precioUnidadActual = r.precioUnidadBase;
      r.precioKmActual = r.precioKmBase;
    }
  }

  filtrados.sort((a, b) => a.fecha - b.fecha || a.cliente.localeCompare(b.cliente, "es"));
  resultadosActuales = filtrados;
  renderResultados(filtrados);

  document.getElementById("resumenBusqueda").textContent =
    `${filtrados.length} trabajo(s) encontrado(s).`;
});

// ---------- Render de la tabla de resultados ----------
function calcularFila(r) {
  r.subtotalUnidadActual = r.cantidad * r.precioUnidadActual;
  r.subtotalMovActual = r.distancia * r.precioKmActual;
  r.importeActual = r.subtotalUnidadActual + r.subtotalMovActual;
  return r;
}

function renderResultados(filas) {
  const cuerpo = document.getElementById("cuerpoResultados");
  if (!filas.length) {
    cuerpo.innerHTML = `<tr><td colspan="12" class="vacio">No se encontraron trabajos con esos filtros.</td></tr>`;
    document.getElementById("totalGeneral").textContent = "";
    return;
  }

  cuerpo.innerHTML = "";
  filas.forEach((r, idx) => {
    calcularFila(r);
    const ajustado = r.precioUnidadActual !== r.precioUnidadBase || r.precioKmActual !== r.precioKmBase;
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${formatoFecha(r.fecha)}</td>
      <td>${escapeHtml(r.cliente)}</td>
      <td>${escapeHtml(r.tipo)}</td>
      <td>${escapeHtml(r.establecimiento)}</td>
      <td>${r.cantidad} ${escapeHtml(r.cantidadLabel)}</td>
      <td><input type="number" step="0.01" class="precio-input" data-idx="${idx}" data-campo="precioUnidad" value="${r.precioUnidadActual}" style="${ajustado ? "border-color:#e08a3e;background:#fff8ef;" : ""}"></td>
      <td class="subtotalUnidad">${formatoMoneda(r.subtotalUnidadActual)}</td>
      <td>${r.distancia}</td>
      <td><input type="number" step="0.01" class="precio-input" data-idx="${idx}" data-campo="precioKm" value="${r.precioKmActual}" style="${ajustado ? "border-color:#e08a3e;background:#fff8ef;" : ""}"></td>
      <td class="subtotalMov">${formatoMoneda(r.subtotalMovActual)}</td>
      <td class="importe">${formatoMoneda(r.importeActual)}</td>
      <td><span class="origen-tag">${r.origen}</span></td>
    `;
    cuerpo.appendChild(tr);
  });

  actualizarTotales();
}

function actualizarTotales() {
  const total = resultadosActuales.reduce((acc, r) => acc + (r.importeActual || 0), 0);
  document.getElementById("totalGeneral").textContent = `Total: $ ${formatoMoneda(total)}`;
}

document.getElementById("cuerpoResultados").addEventListener("input", (e) => {
  const target = e.target;
  if (!target.classList.contains("precio-input")) return;
  const idx = Number(target.dataset.idx);
  const campo = target.dataset.campo;
  const valor = Number(target.value) || 0;
  const registro = resultadosActuales[idx];
  if (campo === "precioUnidad") registro.precioUnidadActual = valor;
  if (campo === "precioKm") registro.precioKmActual = valor;
  calcularFila(registro);

  const tr = target.closest("tr");
  tr.querySelector(".subtotalUnidad").textContent = formatoMoneda(registro.subtotalUnidadActual);
  tr.querySelector(".subtotalMov").textContent = formatoMoneda(registro.subtotalMovActual);
  tr.querySelector(".importe").textContent = formatoMoneda(registro.importeActual);
  actualizarTotales();
});

// ---------- Guardar precios ajustados (compartido) ----------
document.getElementById("btnGuardarPrecios").addEventListener("click", async () => {
  if (!resultadosActuales.length) return;
  const claves = new Set();
  for (const r of resultadosActuales) {
    const clave = claveReferencia(r.cliente, r.tipo);
    if (claves.has(clave)) continue;
    claves.add(clave);
    await ZoomStore.setPrecio(clave, {
      cliente: r.cliente,
      tipo: r.tipo,
      precioUnidad: r.precioUnidadActual,
      precioKm: r.precioKmActual,
      actualizado: new Date().toISOString()
    });
  }
  alert(`Precios guardados para ${claves.size} combinación(es) de cliente + tipo de trabajo.`);
  renderResultados(resultadosActuales); // refresca marcas de "ajustado"
});

// ---------- Generar PDF ----------
document.getElementById("btnGenerarPDF").addEventListener("click", async () => {
  if (!resultadosActuales.length) {
    alert("Primero buscá trabajos con los filtros.");
    return;
  }
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ orientation: "landscape", unit: "pt", format: "a4" });

  const logoBase64 = await imagenUrlABase64("assets/logo-pdf.jpg");
  const pageWidth = doc.internal.pageSize.getWidth();

  // El logo es cuadrado (1080x1080) — se mantiene 1:1 para que no se deforme.
  doc.addImage(logoBase64, "JPEG", 40, 20, 50, 50);
  doc.setFontSize(16);
  doc.setTextColor(51, 71, 60);
  doc.text("Detalle de trabajos realizados", pageWidth - 40, 45, { align: "right" });

  const desde = document.getElementById("filtroDesde").value || "—";
  const hasta = document.getElementById("filtroHasta").value || "—";
  const clientesSel = Array.from(document.getElementById("filtroCliente").selectedOptions).map((o) => o.value);
  const estSel = Array.from(document.getElementById("filtroEstablecimiento").selectedOptions).map((o) => o.value);
  const tipoSel = Array.from(document.getElementById("filtroTipo").selectedOptions).map((o) => o.value);
  const generadoPor = document.getElementById("generadoPor").value;

  doc.setFontSize(9);
  doc.setTextColor(90, 90, 90);
  const infoLineas = [
    `Período: ${desde} a ${hasta}`,
    `Cliente(s): ${clientesSel.length ? clientesSel.join(", ") : "Todos"}`,
    `Establecimiento(s): ${estSel.length ? estSel.join(", ") : "Todos"}`,
    `Tipo(s) de trabajo: ${tipoSel.length ? tipoSel.join(", ") : "Todos"}`,
    `Generado por: ${generadoPor}   |   Fecha de generación: ${new Date().toLocaleDateString("es-AR")}`
  ];
  doc.text(infoLineas, 40, 85);

  const filas = resultadosActuales.map((r) => [
    formatoFecha(r.fecha),
    r.cliente,
    r.tipo,
    r.establecimiento,
    `${r.cantidad} ${r.cantidadLabel}`,
    `$ ${formatoMoneda(r.precioUnidadActual)}`,
    `$ ${formatoMoneda(r.subtotalUnidadActual)}`,
    `${r.distancia} km`,
    `$ ${formatoMoneda(r.precioKmActual)}`,
    `$ ${formatoMoneda(r.subtotalMovActual)}`,
    `$ ${formatoMoneda(r.importeActual)}`
  ]);

  const total = resultadosActuales.reduce((acc, r) => acc + r.importeActual, 0);

  // Última fecha real incluida en el detalle, por cliente (para "Últimos por cliente").
  const porCliente = {};
  resultadosActuales.forEach((r) => {
    const fechaStr = r.fecha.toISOString().slice(0, 10);
    if (!porCliente[r.cliente] || fechaStr > porCliente[r.cliente]) {
      porCliente[r.cliente] = fechaStr;
    }
  });

  doc.autoTable({
    startY: 135,
    head: [["Fecha", "Cliente", "Tipo de trabajo", "Establecimiento", "Cant.", "Precio unidad", "Subtotal", "Distancia", "Precio km", "Subtotal mov.", "Importe"]],
    body: filas,
    styles: { fontSize: 7.5, cellPadding: 4 },
    headStyles: { fillColor: [51, 71, 60], textColor: 255 },
    foot: [["", "", "", "", "", "", "", "", "", "TOTAL", `$ ${formatoMoneda(total)}`]],
    footStyles: { fillColor: [231, 227, 216], textColor: [37, 54, 48], fontStyle: "bold" }
  });

  const nombreArchivo = `Detalle_ZoomAgricultura_${new Date().toISOString().slice(0, 10)}.pdf`;
  doc.save(nombreArchivo);

  await ZoomStore.addHistorial({
    fecha: new Date().toISOString(),
    generadoPor,
    desde, hasta,
    clientes: clientesSel.length ? clientesSel : ["Todos"],
    establecimientos: estSel.length ? estSel : ["Todos"],
    tipos: tipoSel.length ? tipoSel : ["Todos"],
    cantidadTrabajos: resultadosActuales.length,
    total,
    archivo: nombreArchivo,
    porCliente
  });
  cargarHistorial();
});

function imagenUrlABase64(url) {
  return fetch(url)
    .then((r) => r.blob())
    .then((blob) => new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onloadend = () => resolve(reader.result);
      reader.onerror = reject;
      reader.readAsDataURL(blob);
    }));
}

// ---------- Exportar Excel actualizado ----------
document.getElementById("btnExportarExcel").addEventListener("click", () => {
  if (!workbookActual || !resultadosActuales.length) {
    alert("Primero buscá trabajos con los filtros.");
    return;
  }
  resultadosActuales.forEach((r) => {
    const ws = workbookActual.Sheets[r.sheet];
    const cols = r.cols || {};
    if (cols.precioUnidad !== undefined) setCellNum(ws, r.rowIdx, cols.precioUnidad, r.precioUnidadActual);
    if (cols.subtotalUnidad !== undefined) setCellNum(ws, r.rowIdx, cols.subtotalUnidad, r.subtotalUnidadActual);
    if (cols.precioKm !== undefined) setCellNum(ws, r.rowIdx, cols.precioKm, r.precioKmActual);
    if (cols.subtotalMov !== undefined) setCellNum(ws, r.rowIdx, cols.subtotalMov, r.subtotalMovActual);
    if (cols.importe !== undefined) setCellNum(ws, r.rowIdx, cols.importe, r.importeActual);
  });
  const nombreArchivo = `Registro_Zoom_actualizado_${new Date().toISOString().slice(0, 10)}.xlsx`;
  XLSX.writeFile(workbookActual, nombreArchivo);
});

// ---------- Historial ----------
async function cargarHistorial() {
  const historial = await ZoomStore.getHistorial();
  const cont = document.getElementById("listaHistorial");
  if (!historial.length) {
    cont.innerHTML = `<div class="vacio">Todavía no se generó ningún detalle.</div>`;
    return;
  }
  cont.innerHTML = historial
    .sort((a, b) => new Date(b.fecha) - new Date(a.fecha))
    .map((h) => `
      <div class="historial-item">
        <div><strong>${new Date(h.fecha).toLocaleString("es-AR")}</strong> — ${escapeHtml(h.generadoPor || "—")}</div>
        <div class="meta">
          Período ${escapeHtml(h.desde)} a ${escapeHtml(h.hasta)} ·
          ${h.cantidadTrabajos} trabajo(s) · Total $ ${formatoMoneda(h.total)}
        </div>
        <div class="meta">
          <span class="pill">Clientes: ${escapeHtml((h.clientes || []).join(", "))}</span>
          <span class="pill">Establecimientos: ${escapeHtml((h.establecimientos || []).join(", "))}</span>
          <span class="pill">Tipos: ${escapeHtml((h.tipos || ["Todos"]).join(", "))}</span>
        </div>
      </div>
    `).join("");
}

// ---------- Últimos detalles por cliente ----------
let ultimosPorClienteCache = null; // se recalcula cada vez que se entra a la solapa

async function calcularUltimosPorCliente() {
  const historial = await ZoomStore.getHistorial();
  const porCliente = {}; // cliente -> { fecha, generadoEl, generadoPor }
  historial.forEach((h) => {
    if (!h.porCliente) return; // detalles viejos generados antes de este agregado
    Object.entries(h.porCliente).forEach(([cliente, fechaStr]) => {
      const actual = porCliente[cliente];
      if (!actual || fechaStr > actual.fecha) {
        porCliente[cliente] = { fecha: fechaStr, generadoEl: h.fecha, generadoPor: h.generadoPor };
      }
    });
  });
  return Object.entries(porCliente)
    .map(([cliente, datos]) => ({ cliente, ...datos }))
    .sort((a, b) => a.cliente.localeCompare(b.cliente, "es"));
}

function renderUltimos(lista, filtroTexto) {
  const cuerpo = document.getElementById("cuerpoUltimos");
  const texto = (filtroTexto || "").trim().toLowerCase();
  const filtrada = texto ? lista.filter((x) => x.cliente.toLowerCase().includes(texto)) : lista;

  if (!filtrada.length) {
    cuerpo.innerHTML = `<tr><td colspan="4" class="vacio">${lista.length ? "Ningún cliente coincide con la búsqueda." : "Todavía no hay detalles generados."}</td></tr>`;
    return;
  }
  cuerpo.innerHTML = filtrada.map((x) => `
    <tr>
      <td>${escapeHtml(x.cliente)}</td>
      <td><strong>${formatoFecha(new Date(x.fecha + "T00:00:00Z"))}</strong></td>
      <td>${new Date(x.generadoEl).toLocaleDateString("es-AR")}</td>
      <td>${escapeHtml(x.generadoPor || "—")}</td>
    </tr>
  `).join("");
}

document.getElementById("buscarClienteUltimos").addEventListener("input", (e) => {
  if (ultimosPorClienteCache) renderUltimos(ultimosPorClienteCache, e.target.value);
});

// Aviso de "última fecha detallada" cuando se elige un solo cliente en Filtros.
document.getElementById("filtroCliente").addEventListener("change", async () => {
  const sel = Array.from(document.getElementById("filtroCliente").selectedOptions).map((o) => o.value);
  const aviso = document.getElementById("avisoUltimoCliente");
  if (sel.length !== 1) {
    aviso.classList.add("oculto");
    return;
  }
  const lista = await calcularUltimosPorCliente();
  const datos = lista.find((x) => x.cliente === sel[0]);
  if (!datos) {
    aviso.classList.add("oculto");
    return;
  }
  aviso.classList.remove("oculto");
  aviso.innerHTML = `
    <span>Último detalle a <strong>${escapeHtml(sel[0])}</strong>: hasta el <strong>${formatoFecha(new Date(datos.fecha + "T00:00:00Z"))}</strong>.</span>
    <button type="button" class="secundario" id="btnUsarUltimaFecha">Usar como "Desde"</button>
  `;
  document.getElementById("btnUsarUltimaFecha").addEventListener("click", () => {
    const siguiente = new Date(datos.fecha + "T00:00:00Z");
    siguiente.setUTCDate(siguiente.getUTCDate() + 1);
    document.getElementById("filtroDesde").value = siguiente.toISOString().slice(0, 10);
  });
});

// ---------- Navegación entre vistas ----------
function ocultarTodasLasVistas() {
  document.getElementById("vistaDetalle").classList.add("oculto");
  document.getElementById("vistaUltimos").classList.add("oculto");
  document.getElementById("vistaHistorial").classList.add("oculto");
  document.getElementById("tabDetalle").classList.remove("activo");
  document.getElementById("tabUltimos").classList.remove("activo");
  document.getElementById("tabHistorial").classList.remove("activo");
}

document.getElementById("tabDetalle").addEventListener("click", () => {
  ocultarTodasLasVistas();
  document.getElementById("vistaDetalle").classList.remove("oculto");
  document.getElementById("tabDetalle").classList.add("activo");
});
document.getElementById("tabUltimos").addEventListener("click", async () => {
  ocultarTodasLasVistas();
  document.getElementById("vistaUltimos").classList.remove("oculto");
  document.getElementById("tabUltimos").classList.add("activo");
  ultimosPorClienteCache = await calcularUltimosPorCliente();
  renderUltimos(ultimosPorClienteCache, document.getElementById("buscarClienteUltimos").value);
});
document.getElementById("tabHistorial").addEventListener("click", async () => {
  ocultarTodasLasVistas();
  document.getElementById("vistaHistorial").classList.remove("oculto");
  document.getElementById("tabHistorial").classList.add("activo");
  await cargarHistorial();
});

// ---------- Inicialización ----------
(async function init() {
  await ZoomStore.init();
  if (ZoomStore.modo === "local") {
    document.getElementById("bannerLocal").classList.remove("oculto");
  }
  await cargarExcelBase();
})();
