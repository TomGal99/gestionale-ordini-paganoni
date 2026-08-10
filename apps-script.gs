/**
 * Google Apps Script — Web App per gestire gli ordini dalla dashboard
 * (index.html), senza dover aprire il Google Sheet a mano.
 *
 * NOTA TECNICA: le richieste arrivano via GET (parametri nell'URL), non POST.
 * Motivo: Apps Script risponde con un redirect 302, e per specifica i browser
 * (fetch) trasformano automaticamente una POST in GET perdendo il corpo della
 * richiesta quando seguono quel redirect — bug noto, non evitabile lato client.
 * GET invece non ha questo problema, quindi tutto passa da lì.
 *
 * Parametri GET supportati:
 *   ?action=evadi|elimina&data=...&cliente=...&prodotto=...&testo=...
 * Se "action" manca, di default fa "evadi" (retrocompatibile).
 *
 * DEPLOY (fallo dall'account Google che possiede il foglio: giuliana.paganoni@gmail.com):
 * 1. Apri il Google Sheet "Richieste Instagram - Le Gioie di Giuliana Paganoni".
 * 2. Estensioni → Apps Script.
 * 3. Cancella il contenuto di Code.gs, incolla questo file.
 * 4. Deploy → Gestisci implementazioni → icona matita sulla implementazione
 *    esistente → Versione: Nuova versione → Esegui il deployment.
 *    (Se non hai ancora nessuna implementazione: Deploy → Nuova implementazione →
 *    Applicazione web → Esegui come: Me, Chi ha accesso: Chiunque.)
 * 5. L'URL resta lo stesso di prima (finisce in /exec), non serve aggiornarlo
 *    su index.html se hai fatto "Nuova versione" su un deployment esistente.
 */

const SHEET_NAME = "Richieste"; // nome del foglio (tab) dentro lo spreadsheet
const COL_EVASO = "Evaso";

// Questa Web App è per forza "Chi ha accesso: Chiunque" (serve al browser per chiamarla
// senza login). Una chiave condivisa nel codice pubblico NON è vera sicurezza (chi legge
// il sorgente di index.html la trova), ma alza la soglia contro scraping/abusi casuali —
// difesa a basso costo, non protezione reale. Se la cambi qui, cambiala anche in index.html
// (costante APPS_SCRIPT_KEY) o le richieste dalla dashboard smettono di funzionare.
const SHARED_KEY = "paganoni-gioielli-2026";

function doGet(e) {
  const p = (e && e.parameter) || {};

  if (p.action === "list") {
    return handleList(p);
  }

  // Nessun parametro action/data → probabilmente una visita di controllo nel browser.
  if (!p.data && !p.cliente) {
    return jsonOut({ ok: true, info: "Web App attiva. Usa i parametri ?action=evadi|elimina&data=...&cliente=...&prodotto=...&testo=... oppure ?action=list&key=..." });
  }

  return handleRequest(p);
}

// Restituisce tutti gli ordini in JSON — sostituisce la lettura del CSV pubblico del
// foglio, così il foglio può tornare privato (Condividi → solo il tuo account) e i dati
// dei clienti (nomi, indirizzi, IBAN nel testo) non restano leggibili da chiunque trovi
// il link. Protetto dalla stessa SHARED_KEY delle altre azioni — stesso avvertimento:
// non è vera autenticazione, solo un filtro contro chi trova il link per caso.
function handleList(p) {
  if (p.key !== SHARED_KEY) {
    return jsonOut({ error: "Non autorizzato" }, 403);
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) {
    return jsonOut({ error: "Foglio '" + SHEET_NAME + "' non trovato" }, 500);
  }

  const values = sheet.getDataRange().getValues();
  if (values.length === 0) {
    return jsonOut({ orders: [] });
  }

  const headers = values[0].map(h => String(h).trim());
  const orders = values.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i] === undefined || row[i] === null ? "" : String(row[i]).trim(); });
    return obj;
  });

  return jsonOut({ orders });
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    return handleRequest(body);
  } catch (err) {
    return jsonOut({ error: String(err) }, 500);
  }
}

function handleRequest(p) {
  try {
    const { data, cliente, prodotto, testo, action, key } = p;

    if (key !== SHARED_KEY) {
      return jsonOut({ error: "Non autorizzato" }, 403);
    }

    if (!data || !cliente) {
      return jsonOut({ error: "Parametri mancanti (data, cliente)" }, 400);
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) {
      return jsonOut({ error: "Foglio '" + SHEET_NAME + "' non trovato" }, 500);
    }

    const values = sheet.getDataRange().getValues();
    const headers = values[0].map(h => String(h).trim());
    const idxData = headers.indexOf("Data");
    const idxCliente = headers.indexOf("Cliente");
    const idxProdotto = headers.indexOf("Prodotto");
    const idxTesto = headers.indexOf("TestoOriginale");
    const idxEvaso = headers.indexOf(COL_EVASO);

    let rowIndex = -1;
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const matchData = String(row[idxData]).trim() === String(data).trim();
      const matchCliente = String(row[idxCliente]).trim() === String(cliente).trim();
      const matchProdotto = idxProdotto === -1 || String(row[idxProdotto]).trim() === String(prodotto || "").trim();
      const matchTesto = idxTesto === -1 || String(row[idxTesto]).trim() === String(testo || "").trim();
      if (matchData && matchCliente && matchProdotto && matchTesto) {
        rowIndex = i + 1; // 1-based per getRange/deleteRow
        break;
      }
    }

    if (rowIndex === -1) {
      return jsonOut({ error: "Riga non trovata (i dati non combaciano esattamente col foglio)" }, 404);
    }

    if (action === "elimina") {
      sheet.deleteRow(rowIndex);
      return jsonOut({ ok: true, deleted: true });
    }

    if (idxEvaso === -1) {
      return jsonOut({ error: "Colonna 'Evaso' non trovata nell'intestazione" }, 500);
    }
    // "riapri" annulla un "Segna come spedito" fatto per sbaglio; default "evadi" lo imposta.
    const nuovoValore = action === "riapri" ? "No" : "Si";
    sheet.getRange(rowIndex, idxEvaso + 1).setValue(nuovoValore);
    return jsonOut({ ok: true, row: rowIndex, evaso: nuovoValore });
  } catch (err) {
    return jsonOut({ error: String(err) }, 500);
  }
}

function jsonOut(obj) {
  const output = ContentService.createTextOutput(JSON.stringify(obj));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}
