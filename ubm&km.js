function procesarKilometrosCore(minDate, maxDate, hacerMerge = false) {
  const ssMaestro = SpreadsheetApp.getActiveSpreadsheet();
  let mapaKms = {};

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

  try {
    const ssKm = SpreadsheetApp.openById(ID_SHEET_KILOMETROS);
    const sheetKm = ssKm.getSheetByName('KM') || ssKm.getSheets()[0]; 
    const dataKm = sheetKm.getDataRange().getValues();
    
    for (let i = 1; i < dataKm.length; i++) {
      let n = String(dataKm[i][2]).trim().toLowerCase(); 
      if (!n) continue;
      
      let rawDate = dataKm[i][1]; 
      let dObj = (rawDate instanceof Date) ? rawDate : parseSafeDate(String(rawDate).split('-')[0].trim());
      
      if (dObj && dObj >= minDate && dObj <= maxDate) {
        let km = parseFloat(dataKm[i][10]);   

        // Si está vacía o no es un número, busca en el respaldo
      if (isNaN(km) || km === 0) km = parseFloat(dataKm[i][8]) || 0;        
        if (km > 0) {
          let dd = String(dObj.getDate()).padStart(2, '0');
          let mm = String(dObj.getMonth() + 1).padStart(2, '0');
          let yy = String(dObj.getFullYear()).slice(-2);
          let datePart = `${dd}/${mm}/${yy}`;
          
          if (!mapaKms[n]) mapaKms[n] = [];
          mapaKms[n].push({ fechaCorta: datePart, km: km }); 
        }
      }
    }
  } catch(e) { console.error("Error leyendo planilla de KM:", e); return; }

  let kmChunks = [];
  let kmStr = JSON.stringify(mapaKms);
  let chunkSize = 40000; 
  for (let i = 0; i < kmStr.length; i += chunkSize) {
    kmChunks.push(["'" + kmStr.substring(i, i + chunkSize)]);
  }

  let hojaKm = ssMaestro.getSheetByName('api_km');
  if (!hojaKm) { hojaKm = ssMaestro.insertSheet('api_km'); hojaKm.hideSheet(); }
  hojaKm.clearContents(); 
  if (kmChunks.length > 0) hojaKm.getRange(1, 1, kmChunks.length, 1).setValues(kmChunks);
  
  ssMaestro.toast(`KMs actualizados (${hacerMerge ? 'Merge Parcial' : 'Completo'}).`, "KM OK");
}

function generarJSONKilometros_Frecuente() {
  const hoy = new Date();
  hoy.setHours(23, 59, 59, 999);
  const hace5Dias = new Date(hoy);
  hace5Dias.setDate(hoy.getDate() - 5);
  hace5Dias.setHours(0, 0, 0, 0);
  procesarKilometrosCore(hace5Dias, hoy, true); 
}

function generarJSONKilometros_Diario() {
  const hoy = new Date();
  hoy.setHours(23, 59, 59, 999);
  const hace31Dias = new Date(hoy);
  hace31Dias.setDate(hoy.getDate() - 30);
  hace31Dias.setHours(0, 0, 0, 0);
  procesarKilometrosCore(hace31Dias, hoy, true);
}

function generarJSONKilometros_Completo() {
  const hoy = new Date();
  hoy.setHours(23, 59, 59, 999);
  const hace1Ano = new Date(hoy);
  hace1Ano.setFullYear(hoy.getFullYear() - 1);
  hace1Ano.setHours(0, 0, 0, 0);
  procesarKilometrosCore(hace1Ano, hoy, false);
}
 
  function ubmkm() {
    procesarKilometrosCore(minDate, maxDate, hacerMerge = false) 
    generarJSONKilometros_Frecuente();
    generarJSONKilometros_Diario();
    generarJSONKilometros_Completo();
  }