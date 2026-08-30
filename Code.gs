// -------------------------------------------------------------
// CONFIGURAZIONE PARAMETRI GLOBALI
// -------------------------------------------------------------
var CONFIG = {
  foglioDatabase: "Database", // Nome esatto del tuo foglio
  foglioLog: "Log",           // Foglio log scansioni
  rigaInizio: 10,             // Riga da cui iniziano i dati delle auto
  colonneTotali: 6,           // Quante colonne colorare (A:F = 6 colonne)
  colTarga: 1,                // 0-based: colonna B (Plate)
  cellaPulite: "F3",          // Cella contatore PARCHEGGIO PULITE
  cellaSporche: "F4",         // Cella contatore PARCHEGGIO SPORCHE
  cellaLavaggio: "F5"         // Cella contatore LAVAGGIO SPORCHE
};

// Colori di evidenziazione per l'intera riga (Codici HEX)
var IMPOSTAZIONI_STATO = {
  pulite:   { testo: "PARCHEGGIO PULITE",   colore: "#b7e1cd" }, // Verde
  sporche:  { testo: "PARCHEGGIO SPORCHE",  colore: "#f4c7c3" }, // Rosso
  lavaggio: { testo: "LAVAGGIO SPORCHE",    colore: "#fce8b2" }  // Giallo
};

// -------------------------------------------------------------
// NORMALIZZAZIONE
// -------------------------------------------------------------
function normalizza(text) {
  return String(text).toUpperCase().replace(/[^A-Z0-9]/g, '');
}

// -------------------------------------------------------------
// CREAZIONE MENU UNIFICATO IN ALTO
// -------------------------------------------------------------
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("⚙️ Gestione Fleet")
    .addItem("🚗 Pannello di Controllo", "apriSidebar")
    .addSeparator() 
    .addItem("🔄 Sincronizza Contatori", "menuSincronizzaContatori")
    .addItem("🧹 Ripristina Tabella", "menuPulisciDatabase")
    .addItem("🗑️ Elimina Colonne", "menuPulisciHX")
    .addToUi();
}

function apriSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Interfaccia')
      .setTitle('Fleet Check Serale')
      .setWidth(350)     
      .setHeight(580);   // Leggermente più alto per fare spazio al nuovo bottone
  SpreadsheetApp.getUi().showModelessDialog(html, 'Fleet Check Serale');
}

// -------------------------------------------------------------
// PULIZIA
// -------------------------------------------------------------
function pulisciDatabase(chiamataDaMenu) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var database = ss.getSheetByName(CONFIG.foglioDatabase);

  if (!database) {
    throw new Error("Foglio '" + CONFIG.foglioDatabase + "' non trovato");
  }

  if (chiamataDaMenu === true) {
    var ui = SpreadsheetApp.getUi();
    var risposta = ui.alert("Conferma", "Sei sicuro di voler pulire TUTTA la tabella (A10:F) e i contatori F3:F5?", ui.ButtonSet.YES_NO);
    if (risposta !== ui.Button.YES) {
      return { azzerato: false, messaggio: "Operazione annullata" };
    }
  }

  var lastRow = database.getLastRow();
  
  if (lastRow >= CONFIG.rigaInizio) {
    var rangeTotale = database.getRange(CONFIG.rigaInizio, 1, lastRow - CONFIG.rigaInizio + 1, CONFIG.colonneTotali);
    rangeTotale.clearContent();       // Cancella testo e valori
    rangeTotale.setBackground(null);  // Rimuove i colori di sfondo
  }

  database.getRange("F3:F5").clearContent();

  if (chiamataDaMenu === true) {
    SpreadsheetApp.getUi().alert("Pulizia A10:F e contatori completata con successo!");
  }

  return { azzerato: true, messaggio: "Intervallo A10:F e contatori puliti" };
}

function menuPulisciDatabase() {
  pulisciDatabase(true);
}

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

  database.deleteColumns(8, 17);

  if (chiamataDaMenu === true) {
    SpreadsheetApp.getUi().alert("Eliminazione colonne H:X completata!");
  }
}

function menuPulisciHX() {
  pulisciHX(true);
}

// -------------------------------------------------------------
// RICERCA E REGISTRAZIONE
// -------------------------------------------------------------
function cercaTargaParziale(testoRicerca) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var database = ss.getSheetByName(CONFIG.foglioDatabase);
  if (!database) return { errore: "Foglio '" + CONFIG.foglioDatabase + "' non trovato" };

  return cercaTargaParzialeInDatabase(database, testoRicerca);
}

function cercaTargaParzialeInDatabase(database, testoRicerca) {
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
          targa: dati[i][CONFIG.colTarga],
          brand: dati[i][3] || "",
          modello: dati[i][4] || "N/D"
        });
      }
    }
  }
  return risultati;
}

function registraStatoVeicolo(targaEsatta, codiceStato) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var database = ss.getSheetByName(CONFIG.foglioDatabase);
  if (!database) return { errore: "Foglio '" + CONFIG.foglioDatabase + "' non trovato" };
  
  return registraStatoVeicoloInDatabase(database, targaEsatta, codiceStato);
}

function registraStatoVeicoloInDatabase(database, targaEsatta, codiceStato) {
  var dati = database.getDataRange().getValues();
  var targaCercata = normalizza(targaEsatta);
  var rigaTrovata = -1;
  
  for (var i = CONFIG.rigaInizio - 1; i < dati.length; i++) {
    if (dati[i] && dati[i][CONFIG.colTarga] && normalizza(dati[i][CONFIG.colTarga]) === targaCercata) {
      rigaTrovata = i + 1;
      break;
    }
  }
  
  if (rigaTrovata === -1) {
    return { errore: "Impossibile trovare la targa " + targaEsatta + " nel foglio." };
  }
  
  var configStato = IMPOSTAZIONI_STATO[codiceStato];
  if (!configStato) return { errore: "Stato non valido." };
  
  var rangeRiga = database.getRange(rigaTrovata, 1, 1, CONFIG.colonneTotali);
  rangeRiga.setBackground(configStato.colore);
  
  aggiornaContatoriFlotta(database);
  
  return { successo: true, targa: targaEsatta, stato: configStato.testo, riga: rigaTrovata };
}

function menuSincronizzaContatori() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var database = ss.getSheetByName(CONFIG.foglioDatabase);
  if (!database) return { errore: "Foglio '" + CONFIG.foglioDatabase + "' non trovato" };
  
  aggiornaContatoriFlotta(database);
  return { successo: true };
}

function aggiornaContatoriFlotta(foglio) {
  var ultimaRiga = foglio.getLastRow();
  if (ultimaRiga < CONFIG.rigaInizio) return;
  
  var sfondiTarghe = foglio.getRange(CONFIG.rigaInizio, 2, ultimaRiga - CONFIG.rigaInizio + 1, 1).getBackgrounds();
  
  var conteggio = { pulite: 0, sporche: 0, lavaggio: 0 };
  
  var colVerde = IMPOSTAZIONI_STATO.pulite.colore.toLowerCase();
  var colRosso = IMPOSTAZIONI_STATO.sporche.colore.toLowerCase();
  var colGiallo = IMPOSTAZIONI_STATO.lavaggio.colore.toLowerCase();
  
  for (var i = 0; i < sfondiTarghe.length; i++) {
    var coloreCorrente = sfondiTarghe[i][0].toLowerCase();
    if (coloreCorrente === colVerde) {
      conteggio.pulite++;
    } else if (coloreCorrente === colRosso) {
      conteggio.sporche++;
    } else if (coloreCorrente === colGiallo) {
      conteggio.lavaggio++;
    }
  }
  
  foglio.getRange(CONFIG.cellaPulite).setValue(conteggio.pulite);
  foglio.getRange(CONFIG.cellaSporche).setValue(conteggio.sporche);
  foglio.getRange(CONFIG.cellaLavaggio).setValue(conteggio.lavaggio);
}

function codiceStatoDaSituazione(stato, locazione) {
  if (stato === "Pulita" && locazione === "Parcheggio") return "pulite";
  if (stato === "Sporca" && locazione === "Parcheggio") return "sporche";
  if (stato === "Sporca" && locazione === "Lavaggio") return "lavaggio";
  return null;
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
      risposta = { risultati: cercaTargaParzialeInDatabase(database, targa) };
    } else if (azione === "salva") {
      var targa = String(e.parameter.targa || "").trim();
      var statoCodice = String(e.parameter.stato || "").trim();
      if (!targa || !statoCodice) throw new Error("Targa o stato mancante");
      risposta = registraStatoVeicoloInDatabase(database, targa, statoCodice);
    } else if (azione === "sincronizza") {
      aggiornaContatoriFlotta(database);
      risposta = { successo: true, messaggio: "Contatori sincronizzati" };
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

        var codiceStato = codiceStatoDaSituazione(stato, locazione);
        if (codiceStato) {
          var esito = registraStatoVeicoloInDatabase(database, targa, codiceStato);
          if (esito.errore) throw new Error(esito.errore);
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
