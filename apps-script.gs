/**
 * Google Apps Script — Web App per segnare un ordine come "Evaso/Spedito"
 * direttamente dalla dashboard (index.html), senza dover aprire il Google Sheet.
 *
 * DEPLOY (fallo dall'account Google che possiede il foglio: giuliana.paganoni@gmail.com):
 * 1. Apri il Google Sheet "Richieste Instagram - Le Gioie di Giuliana Paganoni".
 * 2. Estensioni → Apps Script.
 * 3. Cancella il contenuto di Code.gs, incolla questo file.
 * 4. In alto, cambia il nome del progetto (facoltativo) es. "Gestionale Ordini API".
 * 5. Deploy → Nuova implementazione → tipo "Applicazione web".
 *    - Esegui come: Me (giuliana.paganoni@gmail.com)
 *    - Chi ha accesso: Chiunque
 * 6. Autorizza i permessi richiesti (chiede conferma la prima volta).
 * 7. Copia l'URL della Web App (finisce in /exec) e incollalo in index.html
 *    al posto di APPS_SCRIPT_URL.
 * 8. Se in futuro modifichi questo script, devi ripetere "Nuova implementazione"
 *    (Gestisci implementazioni → Modifica → Nuova versione) altrimenti l'URL
 *    pubblico continua a servire la versione vecchia.
 */

const SHEET_NAME = "Richieste"; // nome del foglio (tab) dentro lo spreadsheet
const COL_EVASO = "Evaso";

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const { data, cliente, prodotto, testo, evaso } = body;

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

    if (idxEvaso === -1) {
      return jsonOut({ error: "Colonna 'Evaso' non trovata nell'intestazione" }, 500);
    }

    // Trova la riga che corrisponde esattamente ai campi passati (match più stringente possibile,
    // dato che non esiste un ID univoco per riga). Confronta Data+Cliente+Prodotto+TestoOriginale.
    let rowIndex = -1;
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const matchData = String(row[idxData]).trim() === String(data).trim();
      const matchCliente = String(row[idxCliente]).trim() === String(cliente).trim();
      const matchProdotto = idxProdotto === -1 || String(row[idxProdotto]).trim() === String(prodotto || "").trim();
      const matchTesto = idxTesto === -1 || String(row[idxTesto]).trim() === String(testo || "").trim();
      if (matchData && matchCliente && matchProdotto && matchTesto) {
        rowIndex = i + 1; // 1-based per getRange
        break;
      }
    }

    if (rowIndex === -1) {
      return jsonOut({ error: "Riga non trovata (i dati non combaciano esattamente col foglio)" }, 404);
    }

    const nuovoValore = evaso === false ? "No" : "Si";
    sheet.getRange(rowIndex, idxEvaso + 1).setValue(nuovoValore);

    return jsonOut({ ok: true, row: rowIndex, evaso: nuovoValore });
  } catch (err) {
    return jsonOut({ error: String(err) }, 500);
  }
}

function doGet(e) {
  return jsonOut({ ok: true, info: "Web App attiva. Usa POST per aggiornare una riga." });
}

function jsonOut(obj, status) {
  const output = ContentService.createTextOutput(JSON.stringify(obj));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}
