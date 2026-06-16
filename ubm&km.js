/**
 * ============================================================================
 * UNIFICACIÓN DE MÓDULOS: KILOMETRAJE + VIAJES CAMPO + DETALLE NUEVA SECCIÓN
 * ============================================================================
 */

function procesarKilometrosYViajesCore(minDate, maxDate, hacerMerge = false) {
  const ssMaestro = SpreadsheetApp.getActiveSpreadsheet();
  const ID_SHEET_KILOMETROS = '1Wr-_P4mDvldif_cAx08sp7yT8uTUrajI2HQAJF6tnGM';

  let mapaKms = {};
  let viajesAgrupados = {};
  let arrayNuevaSeccion = []; 

  // --- 1. LECTURA DE CACHÉ VIEJO (Merge de api_km) ---
  if (hacerMerge) {
    try {
      const hojaKm = ssMaestro.getSheetByName('api_km');
      if (hojaKm && hojaKm.getLastRow() > 0) {
        let kmStr = "";
        hojaKm.getDataRange().getValues().forEach(row => {
          row.forEach(cell => { if (cell) kmStr += String(cell).replace(/^'/, ""); });
        });
        
        if (kmStr) {
          let cacheVieja = JSON.parse(kmStr);
          for (let chofer in cacheVieja) {
            mapaKms[chofer] = cacheVieja[chofer].filter(registro => {
              let partes = registro.fechaCorta.split('/');
              let fRegistro = new Date("20" + partes[2], partes[1] - 1, partes[0]);
              return fRegistro < minDate || fRegistro > maxDate; 
            });
          }
        }
      }
    } catch(e) { console.warn("No se pudo leer caché KM para merge. Generando desde cero."); }
  }

  // --- 2. LECTURA DE LA HOJA DE KILÓMETROS ---
  let dataKm;
  try {
    const ssKm = SpreadsheetApp.openById(ID_SHEET_KILOMETROS);
    const sheetKm = ssKm.getSheetByName('KM') || ssKm.getSheets()[0]; 
    dataKm = sheetKm.getDataRange().getValues();
  } catch(e) { 
    console.error("Error leyendo planilla de KM:", e); 
    return; 
  }

  // Helpers internos
  const parseNumberSafe = (val) => {
      if (typeof val === 'number') return val;
      let limpio = String(val || '').replace(/,/g, '.').replace(/[^0-9.-]/g, '');
      return parseFloat(limpio) || 0;
  };
  
  // Tu parseo de fecha exacto
  const parseSafeDateLocal = (rawDate) => {
     if (rawDate instanceof Date) return rawDate;
     let strDate = String(rawDate).split('-')[0].trim();
     let parts = strDate.split(/[\/\-]/);
     if (parts.length >= 3) {
        if(parts[0].length <= 2 && parts[2].length >= 2) { 
          let anio = parts[2].length === 2 ? "20" + parts[2] : parts[2];
          return new Date(anio, parseInt(parts[1], 10) - 1, parts[0]);
        } else if (parts[0].length === 4) { 
          return new Date(parts[0], parseInt(parts[1], 10) - 1, parts[2]);
        }
     }
     // Si tienes un parseSafeDate global, caerá aquí orgánicamente
     return new Date(strDate); 
  };

  const normalizarNombre = (n) => String(n).trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, ' ');
  const mesesAbrev = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"];

  let contadorExtraidos = 0;

  // --- 3. BUCLE MAESTRO ---
  for (let i = 1; i < dataKm.length; i++) {
    let row = dataKm[i];

    // Índices mapeados idénticamente a tu CSV/Planilla
    let dominioRaw  = row[0];   // Col A
    let fechaRaw    = row[1];   // Col B
    let nombreRaw   = row[2];   // Col C
    let livianoRaw  = row[3];   // Col D
    let euroRaw     = row[4];   // Col E
    let campoRaw    = row[5];   // Col F
    let infiniaDRaw = row[7];   // Col H
    let kmBackupRaw = row[8];   // Col I [8] (Respaldo)
    let kmBaseRaw   = row[10];  // Col K [10] (Kilometros)
    let hojaRutaRaw = row[19];  // Col T

    let nOriginal = String(nombreRaw).trim().toLowerCase(); 
    if (!nOriginal) continue;
    
    let nombreNorm = normalizarNombre(nOriginal); 
    let dObj = parseSafeDateLocal(fechaRaw);
    
    if (!dObj || isNaN(dObj.getTime())) continue;

    // ----- A) EXTRACCIÓN DE KILÓMETROS (Tu lógica exacta) -----
    if (dObj >= minDate && dObj <= maxDate) {
      let km = parseFloat(kmBaseRaw);
      
      // Si está vacía o no es un número, busca en el respaldo (Col I)
      if (isNaN(km) || km === 0) km = parseFloat(kmBackupRaw) || 0;        
      
      if (km > 0) {
        let dd = String(dObj.getDate()).padStart(2, '0');
        let mm = String(dObj.getMonth() + 1).padStart(2, '0');
        let yy = String(dObj.getFullYear()).slice(-2);
        let datePart = `${dd}/${mm}/${yy}`;
        
        if (!mapaKms[nOriginal]) mapaKms[nOriginal] = [];
        mapaKms[nOriginal].push({ fechaCorta: datePart, km: km }); 
      }
    }

    // ----- B) EXTRACCIÓN DE VIAJES CAMPO Y NUEVA SECCIÓN -----
    let livianoNum  = parseNumberSafe(livianoRaw);
    let euroNum     = parseNumberSafe(euroRaw);
    let campoNum    = parseNumberSafe(campoRaw);
    let infiniaDNum = parseNumberSafe(infiniaDRaw);
    let hojaStr     = String(hojaRutaRaw || "").trim();

    if (campoNum > 0 || livianoNum > 0 || euroNum > 0 || infiniaDNum > 0 || hojaStr !== "") {
      
      let mesAnio = `${mesesAbrev[dObj.getMonth()]}-${String(dObj.getFullYear()).slice(-2)}`;

      if (!viajesAgrupados[nombreNorm]) viajesAgrupados[nombreNorm] = {};
      if (!viajesAgrupados[nombreNorm][mesAnio]) {
          viajesAgrupados[nombreNorm][mesAnio] = { km: 0, liviano: 0, euro: 0, infiniaD: 0 };
      }
      viajesAgrupados[nombreNorm][mesAnio].km += campoNum;
      viajesAgrupados[nombreNorm][mesAnio].liviano += livianoNum;
      viajesAgrupados[nombreNorm][mesAnio].euro += euroNum;
      viajesAgrupados[nombreNorm][mesAnio].infiniaD += infiniaDNum;

      arrayNuevaSeccion.push({
          fecha: dObj.toISOString().split('T')[0], 
          dominio: String(dominioRaw || "").trim(),
          chofer: String(nombreRaw || "").trim(),
          liviano: livianoNum,
          euro: euroNum,
          campo: campoNum,
          infiniaD: infiniaDNum,
          hoja_ruta: hojaStr
      });
      contadorExtraidos++;
    }
  }

  // --- 4. ESCRITURA EN CACHÉ ---
  
  // Guardar api_km clásico
  let hojaKm = ssMaestro.getSheetByName('api_km');
  if (!hojaKm) { hojaKm = ssMaestro.insertSheet('api_km'); hojaKm.hideSheet(); }
  hojaKm.clearContents(); 
  
  let kmChunks = [];
  let kmStr = JSON.stringify(mapaKms);
  for (let i = 0; i < kmStr.length; i += 40000) { kmChunks.push(["'" + kmStr.substring(i, i + 40000)]); }
  if (kmChunks.length > 0) hojaKm.getRange(1, 1, kmChunks.length, 1).setValues(kmChunks);

  // Guardar nueva data agrupada (Reemplaza a las funciones sueltas)
  let cacheSheet = ssMaestro.getSheetByName("API_CACHE_BASICO");
  if (cacheSheet && typeof escribirChunksEnFila === 'function') {
    escribirChunksEnFila(cacheSheet, 7, JSON.stringify(viajesAgrupados));
    escribirChunksEnFila(cacheSheet, 12, JSON.stringify(arrayNuevaSeccion));
  }
  
  ssMaestro.toast(`Proceso OK. ${contadorExtraidos} viajes detallados cacheados.`, "✅ Éxito");
}

// ============================================================================
// TRIGGERS Y DISPARADORES
// ============================================================================

function generarJSONKilometros_Frecuente() {
  const hoy = new Date(); hoy.setHours(23, 59, 59, 999);
  const hace5Dias = new Date(hoy); hace5Dias.setDate(hoy.getDate() - 5); hace5Dias.setHours(0, 0, 0, 0);
  procesarKilometrosYViajesCore(hace5Dias, hoy, true); 
}

function generarJSONKilometros_Diario() {
  const hoy = new Date(); hoy.setHours(23, 59, 59, 999);
  const hace31Dias = new Date(hoy); hace31Dias.setDate(hoy.getDate() - 30); hace31Dias.setHours(0, 0, 0, 0);
  procesarKilometrosYViajesCore(hace31Dias, hoy, true);
}

function generarJSONKilometros_Completo() {
  const hoy = new Date(); hoy.setHours(23, 59, 59, 999);
  const hace1Ano = new Date(hoy); hace1Ano.setFullYear(hoy.getFullYear() - 1); hace1Ano.setHours(0, 0, 0, 0);
  procesarKilometrosYViajesCore(hace1Ano, hoy, false);
}

// Disparador manual arreglado (Solo llama al completo, el resto es redundante)
function ubmkm() {
  generarJSONKilometros_Completo();
}
