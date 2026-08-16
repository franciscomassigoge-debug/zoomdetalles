/* ============================================================
   ZoomStore — capa de almacenamiento compartido.
   Si firebase-config.js tiene una configuración real, usa Firestore
   (compartido entre Francisco y Tomás). Si no, usa localStorage
   (solo en este dispositivo) y lo señala con un banner en pantalla.
   ============================================================ */

const ZoomStore = (function () {
  let modo = "local"; // "local" | "firebase"
  let db = null;

  const LS_PRECIOS = "zoom_precios_v1";
  const LS_HISTORIAL = "zoom_historial_v1";

  function configuracionValida(cfg) {
    return cfg && cfg.apiKey && !String(cfg.apiKey).startsWith("TU_");
  }

  async function init() {
    const cfg = window.ZOOM_FIREBASE_CONFIG;
    if (configuracionValida(cfg) && window.firebase) {
      try {
        firebase.initializeApp(cfg);
        db = firebase.firestore();
        modo = "firebase";
      } catch (err) {
        console.warn("No se pudo inicializar Firebase, se usa almacenamiento local.", err);
        modo = "local";
      }
    } else {
      modo = "local";
    }
  }

  // ---------- Precios de referencia (por cliente + tipo de trabajo) ----------
  async function getPrecio(clave) {
    if (modo === "firebase") {
      const doc = await db.collection("precios").doc(encodeURIComponent(clave)).get();
      return doc.exists ? doc.data() : null;
    }
    const todos = JSON.parse(localStorage.getItem(LS_PRECIOS) || "{}");
    return todos[clave] || null;
  }

  async function setPrecio(clave, datos) {
    if (modo === "firebase") {
      await db.collection("precios").doc(encodeURIComponent(clave)).set(datos);
      return;
    }
    const todos = JSON.parse(localStorage.getItem(LS_PRECIOS) || "{}");
    todos[clave] = datos;
    localStorage.setItem(LS_PRECIOS, JSON.stringify(todos));
  }

  // ---------- Historial de PDFs generados ----------
  async function addHistorial(entry) {
    if (modo === "firebase") {
      await db.collection("historial").add(entry);
      return;
    }
    const lista = JSON.parse(localStorage.getItem(LS_HISTORIAL) || "[]");
    lista.push(entry);
    localStorage.setItem(LS_HISTORIAL, JSON.stringify(lista));
  }

  async function getHistorial() {
    if (modo === "firebase") {
      const snap = await db.collection("historial").orderBy("fecha", "desc").limit(200).get();
      return snap.docs.map((d) => d.data());
    }
    return JSON.parse(localStorage.getItem(LS_HISTORIAL) || "[]");
  }

  return {
    init,
    getPrecio,
    setPrecio,
    addHistorial,
    getHistorial,
    get modo() { return modo; }
  };
})();
