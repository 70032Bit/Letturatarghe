/*
  Google Apps Script per Verifica Targhe da Chiave
  - Riceve GET: ?codice=...&stato=...&locazione=...&callback=...
  - Riceve GET: ?azione=azzera&callback=... per azzerare A10:G + contatori
  - Menu personalizzato 'Pulizia' in alto
  - Eliminazione fisica delle colonne H:X (dalla colonna 8 alla 24)
*/

var CONFIG = {
  foglioDatabase: "Database",      // foglio con dati auto
  foglioLog: "Log",                // foglio dove salvare le scansioni
  rigaInizio: 10,                  // prima riga dati
  colTarga: 1,                     // 0-based: colonna B (Plate)
  contaPulita: "F3",               // Contatore PARCHEGGIO PULITE
  contaSporche: "F4",              // Contatore PARCHEGGIO SPORCHE
  contaLavaggio: "F5"              // Contatore LAVAGGIO SPORCHE
};

function colorePerSituazione(stato, locazione) {
  if (stato === "Pulita" && locazione === "Parcheggio") return "#c6efce";
  if (stato === "Sporca" && locazione === "Parcheggio") return "#ef9a9a";
  if (stato === "Sporca" && locazione === "Lavaggio") return "#ffe598";
  return null;
}

function contaPerSituazione(stato, locazione) {
  if (stato === "Pulita" && locazione === "Parcheggio") return CONFIG.contaPulita;
  if (stato === "Sporca" && locazione === "Parcheggio") return CONFIG.contaSporche;
  if (stato === "Sporca" && locazione === "Lavaggio") return CONFIG.contaLavaggio;
  return null;
}

function colorePerStato(statoCodice) {
  if (statoCodice === "pulite") return "#c6efce";
  if (statoCodice === "sporche") return "#ef9a9a";
  if (statoCodice === "lavaggio") return "#ffe598";
  return null;
}

function contatorePerStato(statoCodice) {
  if (statoCodice === "pulite") return CONFIG.contaPulita;
  if (statoCodice === "sporche") return CONFIG.contaSporche;
  if (statoCodice === "lavaggio") return CONFIG.contaLavaggio;
  return null;
}

function cercaTargaParziale(database, testoRicerca) {
  var dati = database.getDataRange().getValues();
  testoRicerca = normalizza(testoRicerca);
  var risultati = [];
  if (testoRicerca === "") return risultati;

  for (var i = CONFIG.rigaInizio - 1; i < dati.length; i++) {
    if (dati[i] && dati[i][CONFIG.colTarga]) {
      var targaCorrente = normalizza(dati[i][CONFIG.colTarga]);
      if (targaCorrente.indexOf(testoRicerca) !== -1) {
        risultati.push({
          riga: i + 1,
          branch: dati[i][0] || "N/D",
          targa: dati[i][CONFIG.colTarga],
          group: dati[i][2] || "N/D",
          brand: dati[i][3] || "N/D",
          model: dati[i][4] || "N/D",
          colour: dati[i][5] || "N/D"
        });
      }
    }
  }
  return risultati;
}

function applicaStatoVeicolo(database, targaEsatta, statoCodice) {
  var dati = database.getDataRange().getValues();
  var targaCercata = normalizza(targaEsatta);
  var rigaFoglio = -1;

  for (var i = CONFIG.rigaInizio - 1; i < dati.length; i++) {
    if (dati[i] && dati[i][CONFIG.colTarga] && normalizza(dati[i][CONFIG.colTarga]) === targaCercata) {
      rigaFoglio = i + 1;
      break;
    }
  }

  if (rigaFoglio === -1) {
    throw new Error("Targa " + targaEsatta + " non trovata");
  }

  var colore = colorePerStato(statoCodice);
  if (colore) {
    database.getRange(rigaFoglio, 1, 1, database.getLastColumn()).setBackground(colore);
  }

  var cellaContatore = contatorePerStato(statoCodice);
  if (cellaContatore) {
    var valore = Number(database.getRange(cellaContatore).getValue()) || 0;
    database.getRange(cellaContatore).setValue(valore + 1);
  }

  return { success: true, message: "Targa " + targaEsatta + " registrata come " + statoCodice };
}

function normalizza(text) {
  return String(text).toUpperCase().replace(/[^A-Z0-9]/g, '');
}

// -------------------------------------------------------------
// CREAZIONE MENU IN ALTO ALL'APERTURA DEL FOGLIO
// -------------------------------------------------------------
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Pulizia")
    .addItem("Pulisci Tabella (A10:G) e Contatori", "menuPulisciDatabase")
    .addItem("Elimina Colonne H:X", "menuPulisciHX")
    .addToUi();
}

// -------------------------------------------------------------
// SEZIONE PULIZIA A10:G E CONTATORI
// -------------------------------------------------------------
function pulisciDatabase(chiamataDaMenu) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var database = ss.getSheetByName(CONFIG.foglioDatabase);

  if (!database) {
    throw new Error("Foglio '" + CONFIG.foglioDatabase + "' non trovato");
  }

  if (chiamataDaMenu === true) {
    var ui = SpreadsheetApp.getUi();
    var risposta = ui.alert("Conferma", "Sei sicuro di voler pulire TUTTA la tabella (A10:G) e i contatori G3:G5?", ui.ButtonSet.YES_NO);
    if (risposta !== ui.Button.YES) {
      return { azzerato: false, messaggio: "Operazione annullata" };
    }
  }

  var lastRow = database.getLastRow();
  
  if (lastRow >= CONFIG.rigaInizio) {
    var rangeTotale = database.getRange(CONFIG.rigaInizio, 1, lastRow - CONFIG.rigaInizio + 1, 7);
    rangeTotale.clearContent();       // Cancella testo e valori
    rangeTotale.setBackground(null);  // Rimuove i colori di sfondo (mantiene i bordi)
  }

  database.getRange("G3:G5").clearContent();

  if (chiamataDaMenu === true) {
    SpreadsheetApp.getUi().alert("Pulizia A10:G e contatori completata con successo!");
  }

  return { azzerato: true, messaggio: "Intervallo A10:G e contatori puliti" };
}

function menuPulisciDatabase() {
  pulisciDatabase(true);
}

// -------------------------------------------------------------
// SEZIONE ELIMINAZIONE COLONNE H:X
// -------------------------------------------------------------
function pulisciHX(chiamataDaMenu) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var database = ss.getSheetByName(CONFIG.foglioDatabase);

  if (!database) {
    throw new Error("Foglio '" + CONFIG.foglioDatabase + "' non trovato");
  }

  if (chiamataDaMenu === true) {
    var ui = SpreadsheetApp.getUi();
    var risposta = ui.alert("Conferma", "Sei sicuro di voler ELIMINARE definitivamente le colonne H:X?", ui.ButtonSet.YES_NO);
    if (risposta !== ui.Button.YES) {
      return;
    }
  }

  // Elimina le 17 colonne a partire dalla colonna 8 (H) fino alla colonna 24 (X)
  database.deleteColumns(8, 17);

  if (chiamataDaMenu === true) {
    SpreadsheetApp.getUi().alert("Eliminazione colonne H:X completata!");
  }
}

function menuPulisciHX() {
  pulisciHX(true);
}

// -------------------------------------------------------------
// WEB APP DO-GET
// -------------------------------------------------------------
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
      risposta = pulisciDatabase(false);
    } else if (azione === "cerca") {
      var targa = String(e.parameter.targa || "").trim();
      if (!targa) throw new Error("Targa mancante");
      risposta = { risultati: cercaTargaParziale(database, targa) };
    } else if (azione === "salva") {
      var targa = String(e.parameter.targa || "").trim();
      var statoCodice = String(e.parameter.stato || "").trim();
      if (!targa || !statoCodice) throw new Error("Targa o stato mancante");
      risposta = applicaStatoVeicolo(database, targa, statoCodice);
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
        if (normalizza(dati[i][CONFIG.colTarga]) === codicePulito) {
          rigaFoglio = i + 1;
          targa = String(dati[i][CONFIG.colTarga]).trim() || codice;
          break;
        }
      }

      if (!targa) {
        risposta.targa = codice;
        risposta.trovata = false;
      } else {
        risposta.trovata = true;
        risposta.targa = targa;

        var colore = colorePerSituazione(stato, locazione);
        if (colore && rigaFoglio > 0) {
          database.getRange(rigaFoglio, 1, 1, database.getLastColumn()).setBackground(colore);
        }

        var cellaContatore = contaPerSituazione(stato, locazione);
        if (cellaContatore) {
          var valore = Number(database.getRange(cellaContatore).getValue()) || 0;
          database.getRange(cellaContatore).setValue(valore + 1);
        }

        var log = ss.getSheetByName(CONFIG.foglioLog);
        if (!log) {
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