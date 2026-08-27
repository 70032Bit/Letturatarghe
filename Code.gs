/*
  Google Apps Script per Verifica Targhe da Chiave
  - Riceve GET: ?codice=...&stato=...&locazione=...&callback=...
  - Riceve GET: ?azione=azzera&callback=... per azzerare le righe evidenziate
  - Cerca il codice (targa) nel foglio 'Database' (col C, righe 10+)
  - Scrive la locazione in col A (Car Status)
  - Colora la riga in base allo stato
  - Incrementa contatori in G3, G4, G5
  - Scrive un LOG nel foglio 'Log' (Timestamp, Codice, Targa, Stato, Locazione)
  - Restituisce JSON/JSONP
*/

var CONFIG = {
  foglioDatabase: "Database",      // foglio con dati auto
  foglioLog: "Log",                // foglio dove salvare le scansioni
  rigaInizio: 10,                  // prima riga dati
  colCodice: 2,                    // 0-based: colonna C (codice/targa)
  colTarga: 2,                     // 0-based: colonna C (targa)
  colCarStatus: 0,                 // 0-based: colonna A (locazione)
  contaPulitaParcheggio: "G3",
  contaSporcaParcheggio: "G4",
  contaSporcaLavaggio: "G5"
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

function normalizza(text) {
  return String(text).toUpperCase().replace(/[^A-Z0-9]/g, '');
}

function azzeraDatabase(database) {
  var lastRow = database.getLastRow();
  if (lastRow >= CONFIG.rigaInizio) {
    for (var r = CONFIG.rigaInizio; r <= lastRow; r++) {
      var sfondi = database.getRange(r, 1, 1, database.getLastColumn()).getBackgrounds()[0];
      var evidenziata = false;
      for (var c = 0; c < sfondi.length; c++) {
        if (sfondi[c] && sfondi[c] !== "#ffffff" && sfondi[c] !== "") {
          evidenziata = true;
          break;
        }
      }
      if (evidenziata) {
        database.getRange(r, 1, 1, database.getLastColumn()).clear();
      }
    }
  }
  // Azzera i contatori G3, G4, G5
  database.getRange("G3:G5").clearContent();
  return { azzerato: true, messaggio: "Righe evidenziate e contatori azzerati" };
}

function doGet(e) {
  var callback = e.parameter.callback;
  var azione = String(e.parameter.azione || "").trim();

  var risposta = {};

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var database = ss.getSheetByName(CONFIG.foglioDatabase);
    if (!database) {
      throw new Error("Foglio '" + CONFIG.foglioDatabase + "' non trovato");
    }

    if (azione === "azzera") {
      risposta = azzeraDatabase(database);
    } else {
      var codice = String(e.parameter.codice || "").trim();
      var stato = String(e.parameter.stato || "").trim();
      var locazione = String(e.parameter.locazione || "").trim();

      risposta = { targa: "", trovata: false, errore: "" };

      if (!codice) {
        throw new Error("Codice mancante");
      }

    var dati = database.getDataRange().getValues();
    var targa = "";
    var rigaFoglio = -1;
    var codicePulito = normalizza(codice);

    for (var i = CONFIG.rigaInizio - 1; i < dati.length; i++) {
      if (normalizza(dati[i][CONFIG.colCodice]) === codicePulito) {
        rigaFoglio = i + 1;
        targa = String(dati[i][CONFIG.colTarga]).trim() || codice;
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
        database.getRange(rigaFoglio, 1, 1, database.getLastColumn()).setBackground(colore);
        database.getRange(rigaFoglio, CONFIG.colCarStatus + 1).setValue(locazione);
      }

      var cellaContatore = contaPerSituazione(stato, locazione);
      if (cellaContatore) {
        var valore = Number(database.getRange(cellaContatore).getValue()) || 0;
        database.getRange(cellaContatore).setValue(valore + 1);
      }

      var log = ss.getSheetByName(CONFIG.foglioLog);
      if (!log) {
        // Se il foglio Log non esiste, lo crea
        log = ss.insertSheet(CONFIG.foglioLog);
        log.appendRow(["Timestamp", "Codice", "Targa", "Stato", "Locazione"]);
      }
      log.appendRow([new Date(), codice, targa, stato, locazione]);
    }
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
