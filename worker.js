const ALLOWED_ORIGIN = "https://tomgal99.github.io";
const AZURE_ENDPOINT = "https://oper-ai-resource.services.ai.azure.com/anthropic/v1/messages";
const AZURE_MODEL = "claude-haiku-4-5";
const MAX_IMAGE_BASE64_CHARS = 8_000_000; // ~6MB immagine originale, margine sicuro sotto i limiti Azure/Anthropic

function corsHeaders() {
  return {
    "Access-Control-Allow-Origin": ALLOWED_ORIGIN,
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
  };
}

function jsonError(message, status, extra) {
  return new Response(JSON.stringify({ error: message, ...(extra || {}) }), {
    status,
    headers: { ...corsHeaders(), "Content-Type": "application/json" },
  });
}

export default {
  async fetch(request, env) {
    if (request.method === "OPTIONS") {
      return new Response(null, { headers: corsHeaders() });
    }

    if (request.method !== "POST") {
      return jsonError("Method not allowed", 405);
    }

    // Il CORS blocca solo le chiamate da browser: chiunque trovi l'URL può comunque
    // colpirlo con curl/script e consumare le chiamate Azure a pagamento. Controllare
    // l'header Origin non è una vera autenticazione (si può falsificare da script),
    // ma blocca gli abusi casuali/scanner automatici — difesa a basso costo, non zero.
    const origin = request.headers.get("Origin");
    if (origin !== ALLOWED_ORIGIN) {
      return jsonError("Richiesta non autorizzata", 403);
    }

    let body;
    try {
      body = await request.json();
    } catch (e) {
      return jsonError("Richiesta non valida (body non è JSON)", 400);
    }

    const { imageBase64, mediaType } = body;

    if (!imageBase64 || !mediaType) {
      return jsonError("Immagine mancante", 400);
    }

    if (imageBase64.length > MAX_IMAGE_BASE64_CHARS) {
      return jsonError(
        "Screenshot troppo grande. Fai uno screenshot più corto/ritagliato o riduci la qualità e riprova.",
        413
      );
    }

    const apiKey = (env.ANTHROPIC_API_KEY || "").trim();
    if (!apiKey) {
      console.error("ANTHROPIC_API_KEY mancante o vuota nei secret del Worker");
      return jsonError("Configurazione AI mancante sul server (contatta Tommaso).", 500);
    }

    const prompt = `Sei un assistente che legge screenshot di conversazioni Instagram (DM) di un negozio di gioielli artigianali fatti a mano ("Le Gioie di Giuliana Paganoni").

Analizza l'immagine ed estrai queste informazioni:

1. "cliente": il nome del cliente/acquirente, se visibile nell'header della chat o nel testo scritto dal cliente. Se non è visibile, scrivi "Sconosciuto".
2. "prodotto": TRASCRIVI ESATTAMENTE, così come scritti dal cliente o dal negozio nella chat, TUTTI gli articoli richiesti/ordinati, uno dopo l'altro separati da " · ". Includi quantità e prezzo se presenti (es. "70+85 catene · 240x4 ovetti · 240 anello jelly fruit · 430 anello topazio giallo · 10 spedizione"). NON scegliere un solo prodotto, NON riassumere, NON forzare la corrispondenza a un catalogo: riporta il dettaglio reale, anche se sono tanti articoli diversi. Se non riesci a distinguere articoli specifici, scrivi "Da specificare (vedi testo)".
3. "tipo": scrivi "Ordine confermato" se il cliente conferma l'acquisto (es. dice che paga, fa il bonifico, chiede l'indirizzo per la spedizione, dà i suoi dati), oppure "Richiesta info" se sta solo chiedendo informazioni/prezzi senza confermare.
4. "testo": il testo esatto e completo dello scambio rilevante (richiesta del cliente + eventuale riepilogo/conto del negozio), copiato così com'è scritto nello screenshot, mantenendo più righe se utile.

Rispondi SOLO con un oggetto JSON valido, senza markdown e senza testo aggiuntivo, con esattamente questa struttura:
{"cliente": "...", "prodotto": "...", "tipo": "...", "testo": "..."}`;

    let anthropicRes;
    try {
      anthropicRes = await fetch(AZURE_ENDPOINT, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          "x-api-key": apiKey,
          "api-key": apiKey,
          "anthropic-version": "2023-06-01",
        },
        body: JSON.stringify({
          model: AZURE_MODEL,
          max_tokens: 1500,
          messages: [
            {
              role: "user",
              content: [
                {
                  type: "image",
                  source: { type: "base64", media_type: mediaType, data: imageBase64 },
                },
                { type: "text", text: prompt },
              ],
            },
          ],
        }),
      });
    } catch (err) {
      console.error("Errore rete verso Azure/Anthropic:", err);
      return jsonError("Impossibile contattare il servizio AI. Riprova tra poco.", 502);
    }

    let data;
    try {
      data = await anthropicRes.json();
    } catch (e) {
      console.error("Risposta Azure non-JSON, status", anthropicRes.status);
      return jsonError("Risposta AI non valida dal server.", 502);
    }

    if (!anthropicRes.ok) {
      console.error("Errore Anthropic/Azure:", anthropicRes.status, JSON.stringify(data));
      return jsonError("Errore del servizio AI", 502, { details: data });
    }

    const textBlock = (data.content && data.content[0] && data.content[0].text) || "";
    const jsonMatch = textBlock.match(/\{[\s\S]*\}/);

    if (!jsonMatch) {
      console.error("Nessun JSON nella risposta AI:", textBlock);
      return jsonError("Risposta AI non interpretabile", 502, { raw: textBlock });
    }

    let extracted;
    try {
      extracted = JSON.parse(jsonMatch[0]);
    } catch (e) {
      console.error("JSON malformato dalla AI:", textBlock);
      return jsonError("JSON non valido dalla AI", 502, { raw: textBlock });
    }

    return new Response(JSON.stringify(extracted), {
      headers: { ...corsHeaders(), "Content-Type": "application/json" },
    });
  },
};
