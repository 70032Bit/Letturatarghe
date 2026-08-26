/*
  Google Apps Script per Verifica Targhe da Chiave
  - Riceve GET: ?codice=...&stato=...&locazione=...&callback=...
  - Cerca il codice (targa) nel foglio 'Anagrafica' (col D, righe 10+)
  - Scrive la locazione in col B (Car Status)
  - Colora la riga in base allo stato
  - Incrementa contatori in F3, F4, F5
  - Scrive un LOG nel foglio 'Log' (Timestamp, Codice, Targa, Stato, Locazione)
  - Restituisce JSON/JSONP
*/

var CONFIG = {
  foglioAnagrafica: "Anagrafica",  // foglio con dati auto
  foglioLog: "Log",                // foglio dove salvare le scansioni
  rigaInizio: 10,                  // prima riga dati
  colCodice: 3,                    // 0-based: colonna D (codice/targa)
  colTarga: 3,                     // 0-based: colonna D (targa)
  colCarStatus: 1,                 // 0-based: colonna B (locazione)
  contaPulitaParcheggio: "F3",
  contaSporcaParcheggio: "F4",
  contaSporcaLavaggio: "F5"
};

function colorePerSituazione(stato, locazione) {
  if (stato === "Pulita" && locazione === "Parcheggio") return "#c6efce";
  if (stato === "Sporca" && locazione === "Parcheggio") return "#ef9a9a";
  if (stato === "Sporca" && locazione === "Lavaggio") return "#ffe598";
  return null;
}

function contaPerSituazione(stato, locazione) {
  if (stato === "Pulita" && locazione === "Parcheggio") return CONFIG.contaPulitaParcheggio;
  if (stato === "Sporca" && locazione === "Parcheggio") return CONFIG.contaSporcaParcheggio;
  if (stato === "Sporca" && locazione === "Lavaggio") return CONFIG.contaSporcaLavaggio;
  return null;
}

function doGet(e) {
  var callback = e.parameter.callback;
  var codice = String(e.parameter.codice || "").trim();
  var stato = String(e.parameter.stato || "").trim();
  var locazione = String(e.parameter.locazione || "").trim();

  var risposta = { targa: "", trovata: false, errore: "" };

  try {
    if (!codice) {
      throw new Error("Codice mancante");
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var anagrafica = ss.getSheetByName(CONFIG.foglioAnagrafica);
    if (!anagrafica) {
      throw new Error("Foglio '" + CONFIG.foglioAnagrafica + "' non trovato");
    }

    var dati = anagrafica.getDataRange().getValues();
    var targa = "";
    var rigaFoglio = -1;
    for (var i = CONFIG.rigaInizio - 1; i < dati.length; i++) {
      if (String(dati[i][CONFIG.colCodice]).trim() === codice) {
        targa = String(dati[i][CONFIG.colTarga]);
        rigaFoglio = i + 1;
        break;
      }
    }

    if (!targa) {
      risposta.targa = codice;
      risposta.trovata = false;
    } else {
      risposta.targa = targa;
      risposta.trovata = true;

      var colore = colorePerSituazione(stato, locazione);
      if (colore && rigaFoglio > 0) {
        anagrafica.getRange(rigaFoglio, 1, 1, anagrafica.getLastColumn()).setBackground(colore);
        anagrafica.getRange(rigaFoglio, CONFIG.colCarStatus + 1).setValue(locazione);
      }

      var cellaContatore = contaPerSituazione(stato, locazione);
      if (cellaContatore) {
        var valore = Number(anagrafica.getRange(cellaContatore).getValue()) || 0;
        anagrafica.getRange(cellaContatore).setValue(valore + 1);
      }

      var log = ss.getSheetByName(CONFIG.foglioLog);
      if (!log) {
        // Se il foglio Log non esiste, lo crea
        log = ss.insertSheet(CONFIG.foglioLog);
        log.appendRow(["Timestamp", "Codice", "Targa", "Stato", "Locazione"]);
      }
      log.appendRow([new Date(), codice, targa, stato, locazione]);
    }
  } catch (err) {
    risposta.errore = err.toString();
  }

  var json = JSON.stringify(risposta);
  if (callback) {
    return ContentService.createTextOutput(callback + "(" + json + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}
