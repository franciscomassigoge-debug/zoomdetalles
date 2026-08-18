/*
 * CONFIGURACIÓN DE FIREBASE (para que el historial y los precios guardados
 * sean compartidos entre Francisco y Tomás desde cualquier dispositivo).
 *
 * Pasos para activarlo (gratis, 5-10 minutos):
 * 1. Entrar a https://console.firebase.google.com con una cuenta Google.
 * 2. "Agregar proyecto" → nombre, por ejemplo "zoom-agricultura" → crear.
 * 3. En el menú lateral: "Compilación" → "Firestore Database" → "Crear base
 *    de datos" → modo "Producción" → elegir una región (ej. southamerica-east1).
 * 4. En "Reglas" de Firestore, para que ambos usuarios (sin login) puedan
 *    leer y escribir, usar temporalmente:
 *      rules_version = '2';
 *      service cloud.firestore {
 *        match /databases/{database}/documents {
 *          match /{document=**} {
 *            allow read, write: if true;
 *          }
 *        }
 *      }
 *    (Esto es simple y funciona para uso interno de dos personas. No es
 *    seguro para una app pública con datos sensibles de terceros.)
 * 5. En el panel del proyecto: ícono de engranaje → "Configuración del
 *    proyecto" → bajar hasta "Tus apps" → ícono "</>" (Web) → registrar
 *    la app (nombre: "Zoom Detalles") → copiar el objeto firebaseConfig
 *    que te muestra.
 * 6. Pegar ese objeto acá abajo, reemplazando el de ejemplo.
 * 7. Guardar este archivo y volver a publicar la app.
 *
 * Mientras este archivo tenga los valores de ejemplo ("TU_API_KEY", etc.),
 * la app funciona en "modo local": el historial y los precios ajustados se
 * guardan solo en este dispositivo (localStorage), no se comparten.
 */

window.ZOOM_FIREBASE_CONFIG = {
  apiKey: "AIzaSyC3b8bv6Dt01M8Qwne27sBO8NIp8f3VM-E",
  authDomain: "zoom-agricultura.firebaseapp.com",
  projectId: "zoom-agricultura",
  storageBucket: "zoom-agricultura.firebasestorage.app",
  messagingSenderId: "670474028495",
  appId: "1:670474028495:web:9762ccd96a219658f53eba"
};
