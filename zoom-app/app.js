/* ============================================================
   Zoom Agricultura — Detalles Administración
   Lógica de la app: lectura de Excel, filtros, ajuste de precios,
   generación de PDF, exportación de Excel actualizado, historial
   compartido (Firebase Firestore) o local (localStorage).
   ============================================================ */

// ---------- Mapa de columnas (igual en "Trabajos" y "Trabajos 2") ----------
// Col A vacía. B=Fecha C=Cod.Trabajo D=Tipo E=Cod.Cliente F=Cliente
// G=Establecimiento H=Cantidad(Ha/Día) I=Precio unidad J=Subtotal unidad
// K=Cod.Vehiculo L=Vehiculo M=Distancia N=Precio km O=Subtotal mov. P=Importe
const COLS = {
  fecha: 1, codTrabajo: 2, tipo: 3, codCliente: 4, cliente: 5,
  establecimiento: 6, cantidad: 7, precioUnidad: 8, subtotalUnidad: 9,
  codVehiculo: 10, vehiculo: 11, distancia: 12, precioKm: 13,
  subtotalMov: 14, importe: 15
};
const HEADER_ROW = 3; // fila 4 (0-indexed) = encabezados
const DATA_START_ROW = 4; // fila 5 (0-indexed) = primer dato

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
  HOJAS.forEach((hoja) => {
    const ws = wb.Sheets[hoja.nombre];
    if (!ws || !ws["!ref"]) return;
    const range = XLSX.utils.decode_range(ws["!ref"]);
    for (let r = DATA_START_ROW; r <= range.e.r; r++) {
      const fecha = getCell(ws, r, COLS.fecha);
      const cliente = getCell(ws, r, COLS.cliente);
      const tipo = getCell(ws, r, COLS.tipo);
      if (esVacio(fecha) || esVacio(cliente) || esVacio(tipo)) continue;
      if (!(fecha instanceof Date)) continue;

      const cantidad = Number(getCell(ws, r, COLS.cantidad)) || 0;
      const precioUnidad = Number(getCell(ws, r, COLS.precioUnidad)) || 0;
      const distancia = Number(getCell(ws, r, COLS.distancia)) || 0;
      const precioKm = Number(getCell(ws, r, COLS.precioKm)) || 0;

      out.push({
        sheet: hoja.nombre,
        rowIdx: r,
        origen: hoja.origen,
        unidadLabel: hoja.unidadLabel,
        cantidadLabel: hoja.cantidadLabel,
        fecha,
        cliente: String(cliente).trim(),
        establecimiento: String(getCell(ws, r, COLS.establecimiento) || "").trim(),
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

  doc.addImage(logoBase64, "JPEG", 40, 25, 120, 40);
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
    archivo: nombreArchivo
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
    setCellNum(ws, r.rowIdx, COLS.precioUnidad, r.precioUnidadActual);
    setCellNum(ws, r.rowIdx, COLS.subtotalUnidad, r.subtotalUnidadActual);
    setCellNum(ws, r.rowIdx, COLS.precioKm, r.precioKmActual);
    setCellNum(ws, r.rowIdx, COLS.subtotalMov, r.subtotalMovActual);
    setCellNum(ws, r.rowIdx, COLS.importe, r.importeActual);
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

// ---------- Navegación entre vistas ----------
document.getElementById("tabDetalle").addEventListener("click", () => {
  document.getElementById("vistaDetalle").classList.remove("oculto");
  document.getElementById("vistaHistorial").classList.add("oculto");
  document.getElementById("tabDetalle").classList.add("activo");
  document.getElementById("tabHistorial").classList.remove("activo");
});
document.getElementById("tabHistorial").addEventListener("click", async () => {
  document.getElementById("vistaDetalle").classList.add("oculto");
  document.getElementById("vistaHistorial").classList.remove("oculto");
  document.getElementById("tabHistorial").classList.add("activo");
  document.getElementById("tabDetalle").classList.remove("activo");
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
