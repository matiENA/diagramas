/**
 * ============================================================================
 * EXTRACCIÓN MASIVA: KILOMETRAJE + VIAJES CAMPO + HOJAS DE RUTA
 * (Para uso con Triggers Temporales o Botón Manual)
 * ============================================================================
 */

function procesarKilometrosYViajesCore(minDate, maxDate, hacerMerge = false) {
  const ssMaestro = SpreadsheetApp.getActiveSpreadsheet();
  const ID_SHEET_KILOMETROS = '1Wr-_P4mDvldif_cAx08sp7yT8uTUrajI2HQAJF6tnGM';

  let mapaKms = {};
  let viajesAgrupados = {};
  let arrayNuevaSeccion = []; 
  
  let cacheSheet = ssMaestro.getSheetByName("API_CACHE_BASICO");
  
  // --- 1. LECTURA Y PRESERVACIÓN DE CACHÉ EXISTENTE (Merge) ---
  if (hacerMerge && cacheSheet) {
    // A) Preservar api_km clásico
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
    } catch(e) { console.warn("No se pudo procesar merge de api_km."); }

    // B) PRESERVAR EL HISTORIAL DE LA FILA 12 (Detalle de Viajes Nueva Sección)
    try {
      if (cacheSheet.getMaxRows() < 12) cacheSheet.insertRowsAfter(cacheSheet.getMaxRows(), 12 - cacheSheet.getMaxRows());
      let lastCol = cacheSheet.getLastColumn() || 1;
      let dataFila12 = cacheSheet.getRange(12, 1, 1, lastCol).getValues()[0];
      let jsonNuevaSeccionRaw = dataFila12.filter(String).map(c => String(c || "").replace(/^'/, "")).join("");
      
      if (jsonNuevaSeccionRaw) {
        let registrosAnteriores = JSON.parse(jsonNuevaSeccionRaw);
        if (Array.isArray(registrosAnteriores)) {
          registrosAnteriores.forEach(reg => {
            let fReg = new Date(reg.fecha + "T12:00:00");
            // Si el registro es viejo (fuera de ventana temporal), lo preservamos
            if (fReg < minDate || fReg > maxDate) {
              arrayNuevaSeccion.push(reg);
            }
          });
        }
      }
    } catch(e) { console.warn("No se pudo realizar merge de la fila 12 de viajes detallados."); }
  }

  // --- 2. LECTURA DE LA HOJA DE KILÓMETROS MAESTRA ---
  let dataKm;
  try {
    const ssKm = SpreadsheetApp.openById(ID_SHEET_KILOMETROS);
    const sheetKm = ssKm.getSheetByName('KM') || ssKm.getSheets()[0]; 
    dataKm = sheetKm.getDataRange().getValues();
  } catch(e) { 
    console.error("Error leyendo planilla de KM:", e); 
    return; 
  }

  const parseNumberSafe = (val) => {
      if (typeof val === 'number') return val;
      let limpio = String(val || '').replace(/,/g, '.').replace(/[^0-9.-]/g, '');
      return parseFloat(limpio) || 0;
  };
  
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
     return new Date(strDate); 
  };

  const normalizarNombre = (n) => String(n).trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, ' ');
  const mesesAbrev = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"];

  let contadorExtraidos = 0;

  // --- 3. BUCLE MAESTRO ---
  for (let i = 1; i < dataKm.length; i++) {
    let row = dataKm[i];

    let dominioRaw  = row[0];   // Col A
    let fechaRaw    = row[1];   // Col B
    let nombreRaw   = row[2];   // Col C
    let livianoRaw  = row[3];   // Col D
    let euroRaw     = row[4];   // Col E
    let campoRaw    = row[5];   // Col F
    let infiniaDRaw = row[7];   // Col H
    let kmBackupRaw = row[8];   // Col I (Respaldo viejo)
    let kmBaseRaw   = row[16];  // Col Q (Kilometros Totales)
    let hojaRutaRaw = row[19];  // Col T (Hoja de Ruta)

    let nOriginal = String(nombreRaw).trim().toLowerCase(); 
    if (!nOriginal) continue;
    
    let nombreNorm = normalizarNombre(nOriginal); 
    let dObj = parseSafeDateLocal(fechaRaw);
    
    if (!dObj || isNaN(dObj.getTime())) continue;

    // Solo extraemos si la fecha entra en la ventana requerida
    if (dObj >= minDate && dObj <= maxDate) {
      
      // A) KILÓMETROS GLOBALES
      let km = parseFloat(kmBaseRaw);
      if (isNaN(km) || km === 0) km = parseFloat(kmBackupRaw) || 0;        
      
      if (km > 0) {
        let dd = String(dObj.getDate()).padStart(2, '0');
        let mm = String(dObj.getMonth() + 1).padStart(2, '0');
        let yy = String(dObj.getFullYear()).slice(-2);
        let datePart = `${dd}/${mm}/${yy}`;
        
        if (!mapaKms[nOriginal]) mapaKms[nOriginal] = [];
        mapaKms[nOriginal].push({ fechaCorta: datePart, km: km }); 
      }

      // B) VIAJES CAMPO Y HOJAS DE RUTA DETALLADOS
      let livianoNum  = parseNumberSafe(livianoRaw);
      let euroNum     = parseNumberSafe(euroRaw);
      let campoNum    = parseNumberSafe(campoRaw);
      let infiniaDNum = parseNumberSafe(infiniaDRaw);
      let hojaStr     = String(hojaRutaRaw || "").trim();

      if (campoNum > 0 || livianoNum > 0 || euroNum > 0 || infiniaDNum > 0 || hojaStr !== "") {
          
          // Sumatoria mensual (Fila 7)
          let mesAnio = `${mesesAbrev[dObj.getMonth()]}-${String(dObj.getFullYear()).slice(-2)}`;
          if (!viajesAgrupados[nombreNorm]) viajesAgrupados[nombreNorm] = {};
          if (!viajesAgrupados[nombreNorm][mesAnio]) {
              viajesAgrupados[nombreNorm][mesAnio] = { km: 0, liviano: 0, euro: 0, infiniaD: 0 };
          }
          viajesAgrupados[nombreNorm][mesAnio].km += campoNum;
          viajesAgrupados[nombreNorm][mesAnio].liviano += livianoNum;
          viajesAgrupados[nombreNorm][mesAnio].euro += euroNum;
          viajesAgrupados[nombreNorm][mesAnio].infiniaD += infiniaDNum;

          // Detalle diario (Fila 12)
          arrayNuevaSeccion.push({
              fecha: dObj.toISOString().split('T')[0], 
              dominio: String(dominioRaw || "").trim(),
              chofer: String(nombreRaw || "").trim(),
              liviano: livianoNum,
              euro: euroNum,
              campo: campoNum,
              infiniaD: infiniaDNum,
              hoja_ruta: hojaStr !== "" ? hojaStr.split(',').map(s => s.trim()) : []
          });
          contadorExtraidos++;
      }
    }
  }

  // --- 4. ESCRITURA ATÓMICA EN CACHÉ ---
  
  // Guardar api_km clásico
  let hojaKm = ssMaestro.getSheetByName('api_km');
  if (!hojaKm) { hojaKm = ssMaestro.insertSheet('api_km'); hojaKm.hideSheet(); }
  hojaKm.clearContents(); 
  
  let kmChunks = [];
  let kmStr = JSON.stringify(mapaKms);
  for (let i = 0; i < kmStr.length; i += 40000) { kmChunks.push(["'" + kmStr.substring(i, i + 40000)]); }
  if (kmChunks.length > 0) hojaKm.getRange(1, 1, kmChunks.length, 1).setValues(kmChunks);

  // Guardar en API_CACHE_BASICO
  if (cacheSheet && typeof escribirChunksEnFila === 'function') {
    escribirChunksEnFila(cacheSheet, 7, JSON.stringify(viajesAgrupados));
    escribirChunksEnFila(cacheSheet, 12, JSON.stringify(arrayNuevaSeccion)); 
  }
  
  ssMaestro.toast(`Proceso OK. ${contadorExtraidos} viajes detallados cacheados.`, "✅ Éxito");
}

function generarJSONKilometros_Frecuente() {
  const hoy = new Date(); hoy.setHours(23, 59, 59, 999);
  const hace60Dias = new Date(hoy); hace60Dias.setDate(hoy.getDate() - 60); hace60Dias.setHours(0, 0, 0, 0);
  procesarKilometrosYViajesCore(hace60Dias, hoy, true); 
}

function generarJSONKilometros_Completo() {
  const hoy = new Date(); hoy.setHours(23, 59, 59, 999);
  const hace1Ano = new Date(hoy); hace1Ano.setFullYear(hoy.getFullYear() - 1); hace1Ano.setHours(0, 0, 0, 0);
  procesarKilometrosYViajesCore(hace1Ano, hoy, false);
}

function ubmkm() {
  generarJSONKilometros_Completo();
}

/**
 * ============================================================================
 * EXTRACCIÓN QUIRÚRGICA: onEdit PARA HOJA DE RUTA
 * (Captura modificaciones en la Columna T y actualiza EXCLUSIVAMENTE la Fila 12)
 * ============================================================================
 */

function alEditarKilometros(e) {
  if (!e || !e.range) return;

  const fila = e.range.getRow();
  const columna = e.range.getColumn();
  const sheet = e.source.getActiveSheet();
  const nombreHoja = sheet.getName();

  // 👉 Columna T (20): N° Hoja de Ruta. Ignoramos la fila 1 (encabezados)
  if (fila >= 2 && columna === 20 && nombreHoja.toUpperCase() === 'KM') {
    try {
      // 1. SOLUCIÓN AL CONTEXTO: Definimos explícitamente la planilla Maestra
      const ID_PLANILLA_MAESTRA = '1eQ9Y5diL5fwxYTxvseNgZJFbX-lSUQ13axbp3cLiqPc'; 
      
      let ssMaestro;
      try { ssMaestro = SpreadsheetApp.openById(ID_PLANILLA_MAESTRA); } 
      catch (err) { ssMaestro = SpreadsheetApp.getActiveSpreadsheet(); }

      const cacheSheet = ssMaestro.getSheetByName("API_CACHE_BASICO");
      if (!cacheSheet) { console.error("No se encontró API_CACHE_BASICO. Abortando."); return; }

      // 2. Extraemos los datos de la fila editada en el Sheet de KMs
      const rangoFila = sheet.getRange(fila, 1, 1, 20).getValues()[0];
      const dominioRaw  = rangoFila[0];   // Col A
      const fechaRaw    = rangoFila[1];   // Col B
      const nombreRaw   = rangoFila[2];   // Col C
      const livianoNum  = parseFloat(String(rangoFila[3] || '').replace(/,/g, '.').replace(/[^0-9.-]/g, '')) || 0;
      const euroNum     = parseFloat(String(rangoFila[4] || '').replace(/,/g, '.').replace(/[^0-9.-]/g, '')) || 0;
      const campoNum    = parseFloat(String(rangoFila[5] || '').replace(/,/g, '.').replace(/[^0-9.-]/g, '')) || 0;
      const infiniaDNum = parseFloat(String(rangoFila[7] || '').replace(/,/g, '.').replace(/[^0-9.-]/g, '')) || 0;
      
      // SOLUCIÓN A E.VALUE: Leemos la celda real
      const nuevoValorHR = String(rangoFila[19] || "").trim();

      if (!nombreRaw || !fechaRaw) return;

      // Normalización de claves idéntica al FrontEnd
      const normalizar = (n) => String(n).trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/\s+/g, ' ');
      const choferNorm = normalizar(nombreRaw);

      // Parseo seguro de fecha ISO para sincronizar con la Fila 12
      let fechaIso = "";
      if (fechaRaw instanceof Date) {
        let tempDate = new Date(fechaRaw.getTime() - (fechaRaw.getTimezoneOffset() * 60000));
        fechaIso = tempDate.toISOString().split('T')[0];
      } else {
        let parts = String(fechaRaw).split('-')[0].trim().split(/[\/\-]/);
        if (parts.length >= 3) {
          let aa = parts[2].length === 2 ? "20" + parts[2] : parts[2];
          fechaIso = `${aa}-${String(parts[1]).padStart(2,'0')}-${String(parts[0]).padStart(2,'0')}`;
        }
      }
      if (!fechaIso) return;

      // 3. Leemos la Fila 12 actual de API_CACHE_BASICO
      if (cacheSheet.getMaxRows() < 12) cacheSheet.insertRowsAfter(cacheSheet.getMaxRows(), 12 - cacheSheet.getMaxRows());
      let lastCol = cacheSheet.getLastColumn() || 1;
      let dataFila12 = cacheSheet.getRange(12, 1, 1, lastCol).getValues()[0];
      let jsonRaw = dataFila12.filter(String).map(c => String(c || "").replace(/^'/, "")).join("");
      
      let arrayViajes = [];
      if (jsonRaw) {
        try { arrayViajes = JSON.parse(jsonRaw); } catch(err) { arrayViajes = []; }
      }
      if (!Array.isArray(arrayViajes)) arrayViajes = [];

      // 4. Buscamos de forma quirúrgica si ya existe el viaje detallado en el caché
      let idx = arrayViajes.findIndex(viaje => normalizar(viaje.chofer) === choferNorm && viaje.fecha === fechaIso);
      let hojasParseadas = nuevoValorHR !== "" ? nuevoValorHR.split(',').map(s => s.trim()).filter(Boolean) : [];

      if (idx > -1) {
        // Si ya existía, actualizamos ÚNICAMENTE su listado de Hojas de Ruta
        arrayViajes[idx].hoja_ruta = hojasParseadas;
      } else {
        // Si no existía, inyectamos el nuevo objeto completo de forma nativa
        arrayViajes.push({
          fecha: fechaIso,
          dominio: String(dominioRaw || "").trim(),
          chofer: String(nombreRaw || "").trim(),
          liviano: livianoNum,
          euro: euroNum,
          campo: campoNum,
          infiniaD: infiniaDNum,
          hoja_ruta: hojasParseadas
        });
      }

      // 5. Volcado ultra veloz en chunks horizontales sobre la Fila 12 (Sin tocar fila 7 ni api_km)
      if (typeof escribirChunksEnFila === 'function') {
        escribirChunksEnFila(cacheSheet, 12, JSON.stringify(arrayViajes));
        sheet.getParent().toast(`Hoja Ruta [${hojasParseadas.join(',')}] sincronizada en servidor`, "⚡ Caché Actualizado");
      } else {
        console.error("La función escribirChunksEnFila no está disponible en este contexto.");
      }

    } catch (error) {
      console.error("Error crítico en inyección onEdit Fila 12:", error);
    }
  }
}
