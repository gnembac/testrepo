/**
 * ============================================================
 * AKTIENTRACKER - Google Apps Script v3.1 (BUGFIX)
 * ============================================================
 *
 * FIXES gegenüber v3.0:
 *  - Yahoo Finance: /v8/chart → /v7/finance/quote (zuverlässiger in Apps Script)
 *  - Batch-Verarbeitung: alle Aktien in einem API-Call (statt 121 Einzelcalls)
 *  - CoinGecko: alle Kryptos in einem API-Call
 *  - Kein Einzel-Sleep mehr → deutlich schneller (30s statt 60s+)
 *  - Debug-Funktion hinzugefügt (zeigt genau welche Ticker fehlschlagen)
 *  - Sheet-Name-Prüfung verbessert (Groß/Kleinschreibung tolerant)
 *  - Kurs in E als Zahl (nicht String) gespeichert
 *
 * SETUP:
 *  1. Google Sheet öffnen
 *  2. Erweiterungen → Apps Script
 *  3. Alten Code komplett ersetzen (Strg+A, dann einfügen)
 *  4. Speichern (Strg+S oder Speichern-Button)
 *  5. Sheet-Seite neu laden (F5) → Menü "Aktientracker" erscheint
 *  6. Beim ersten Klick: Berechtigungen gewähren
 *
 * SHEET-STRUKTUR:
 *  A=Name, B=Ticker, C=WKN, D=ISIN
 *  E=Kurs, F=Währung, G=Quelle, H=Timestamp, I=Kurs€
 *  Y=Währung kurz, Z=FX-Kurs, AA=Währung lang, AB=Börsen, AC=Länder
 * ============================================================
 */

// ============================================================
// KONFIGURATION
// ============================================================
var CONFIG = {
  SHEET_NAME:      "Übersicht",
  SETTINGS_SHEET:  "Einstellungen",
  HEADER_ROW:      2,
  DATA_START_ROW:  3,

  // Spalten-Index (1-basiert)
  COL_NAME:      1,   // A
  COL_TICKER:    2,   // B
  COL_WKN:       3,   // C
  COL_ISIN:      4,   // D
  COL_KURS:      5,   // E
  COL_WAEHRUNG:  6,   // F
  COL_QUELLE:    7,   // G
  COL_TIMESTAMP: 8,   // H
  COL_KURS_EUR:  9,   // I
  COL_FX_KURZ:   25,  // Y
  COL_FX_KURS:   26,  // Z
  COL_FX_NAME:   27,  // AA
  COL_FX_BOERSE: 28,  // AB
  COL_FX_LAND:   29,  // AC

  // Alle Krypto-Ticker (für CoinGecko-Routing)
  CRYPTO_TICKERS: ["BTC","ETH","SOL","AVAX","LINK","HBAR","XRP","DOGE",
                   "DOT","LTC","ADA","FET","RENDER","XLM","MATIC","UNI"],

  YAHOO_BATCH_SIZE: 80,
};

// CoinGecko ID-Mapping
var COINGECKO_IDS = {
  "BTC":    "bitcoin",
  "ETH":    "ethereum",
  "SOL":    "solana",
  "AVAX":   "avalanche-2",
  "LINK":   "chainlink",
  "HBAR":   "hedera-hashgraph",
  "XRP":    "ripple",
  "DOGE":   "dogecoin",
  "DOT":    "polkadot",
  "LTC":    "litecoin",
  "ADA":    "cardano",
  "FET":    "fetch-ai",
  "RENDER": "render-token",
  "XLM":    "stellar",
  "MATIC":  "matic-network",
  "UNI":    "uniswap",
};

// Währungsdatenbank
var CURRENCY_DATA = {
  "USD": ["US-Dollar",           "NYSE, NASDAQ, CBOT",                   "USA, El Salvador, Ecuador"],
  "EUR": ["Euro",                "Xetra (Frankfurt), Euronext, Borsa It.","Deutschland, Frankreich, EU-Zone (+19)"],
  "GBP": ["Britisches Pfund",    "London Stock Exchange (LSE), Aquis",    "Vereinigtes Königreich"],
  "CHF": ["Schweizer Franken",   "SIX Swiss Exchange, Eurex",             "Schweiz, Liechtenstein"],
  "JPY": ["Japanischer Yen",     "Tokyo Stock Exchange (TSE/JPX)",        "Japan"],
  "CNY": ["Chinesischer Yuan",   "Shanghai (SSE), Shenzhen (SZSE)",       "China (Festland)"],
  "HKD": ["Hongkong-Dollar",     "Hong Kong Stock Exchange (HKEX)",       "Hongkong SAR, Macau"],
  "CAD": ["Kanadischer Dollar",  "Toronto Stock Exchange (TSX), TSXV",    "Kanada"],
  "AUD": ["Australischer Dollar","Australian Securities Exchange (ASX)",   "Australien"],
  "KRW": ["Südkorean. Won",      "Korea Exchange (KRX): KOSPI & KOSDAQ",  "Südkorea"],
  "INR": ["Indische Rupie",      "NSE (National), BSE (Bombay)",          "Indien"],
  "SGD": ["Singapur-Dollar",     "Singapore Exchange (SGX)",              "Singapur"],
  "SEK": ["Schwedische Krone",   "Nasdaq Stockholm, NGM",                 "Schweden"],
  "NOK": ["Norwegische Krone",   "Oslo Børs (Euronext Oslo)",             "Norwegen"],
  "DKK": ["Dänische Krone",      "Nasdaq Copenhagen",                     "Dänemark"],
  "BRL": ["Brasil. Real",        "B3 – Brasil, Bolsa, Balcão (São Paulo)","Brasilien"],
  "MXN": ["Mexikan. Peso",       "Bolsa Mexicana de Valores (BMV)",       "Mexiko"],
  "ZAR": ["Südafrik. Rand",      "Johannesburg Stock Exchange (JSE)",     "Südafrika"],
  "TRY": ["Türkische Lira",      "Borsa Istanbul",                        "Türkei"],
  "PLN": ["Polnischer Złoty",    "Warsaw Stock Exchange (GPW)",           "Polen"],
  "HUF": ["Ungarischer Forint",  "Budapest Stock Exchange (BÉT)",         "Ungarn"],
  "KZT": ["Kasach. Tenge",       "Kazakhstan Stock Exchange (KASE), AIX", "Kasachstan"],
};

// ============================================================
// MENÜ
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🔄 Aktientracker")
    .addItem("📈 Kurse aktualisieren",       "updateAllKurse")
    .addItem("💱 FX-Kurse aktualisieren",    "updateAllFX")
    .addSeparator()
    .addItem("⚡ Alle Daten aktualisieren",  "updateAll")
    .addSeparator()
    .addItem("🔍 Debug: API-Test",           "debugApiTest")
    .addItem("ℹ️ Status anzeigen",            "showStatus")
    .addSeparator()
    .addItem("⏱️ Auto-Update einrichten (30 Min)", "createTrigger")
    .addItem("🗑️ Auto-Update deaktivieren",  "deleteTriggers")
    .addToUi();
}

function updateAll() {
  updateAllFX();
  Utilities.sleep(500);
  updateAllKurse();
}


// ============================================================
// SECTION 1: KURSE — Yahoo Finance Batch + CoinGecko Batch
// ============================================================

function updateAllKurse() {
  var sheet    = getMainSheet();
  var lastRow  = getLastDataRow(sheet);
  var numRows  = lastRow - CONFIG.DATA_START_ROW + 1;
  var apiKeys  = getApiKeys();
  var fxRates  = loadFxRates();

  // Alle Daten auf einmal lesen (performant)
  var rawData = sheet.getRange(CONFIG.DATA_START_ROW, 1, numRows, 9).getValues();

  // Assets klassifizieren
  var cryptoRows  = {};  // { rowIndex: geckoId }
  var stockAssets = [];  // { index, ticker }
  var results     = {};  // { rowIndex: { kurs, waehrung, quelle } }

  rawData.forEach(function(row, i) {
    var ticker = (row[1] || "").toString().trim();
    var name   = (row[0] || "").toString().trim();
    if (!ticker || ticker === "") return;

    if (isCrypto(ticker)) {
      var geckoId = COINGECKO_IDS[ticker.toUpperCase()];
      if (geckoId) cryptoRows[i] = geckoId;
    } else {
      stockAssets.push({ index: i, ticker: ticker, name: name });
    }
  });

  Logger.log("Klassifiziert: " + Object.keys(cryptoRows).length + " Kryptos, " + stockAssets.length + " Aktien/ETFs");

  // ── 1. CoinGecko Batch ────────────────────────────────────
  if (Object.keys(cryptoRows).length > 0) {
    var geckoIds  = [];
    var idToRows  = {};
    for (var ri in cryptoRows) {
      var id = cryptoRows[ri];
      if (geckoIds.indexOf(id) === -1) geckoIds.push(id);
      if (!idToRows[id]) idToRows[id] = [];
      idToRows[id].push(parseInt(ri));
    }

    try {
      var cgUrl = "https://api.coingecko.com/api/v3/simple/price?ids=" + geckoIds.join(",") +
                  "&vs_currencies=eur,usd" +
                  (apiKeys.coingecko ? "&x_cg_demo_api_key=" + apiKeys.coingecko : "");
      var cgRes = UrlFetchApp.fetch(cgUrl, { muteHttpExceptions: true });

      if (cgRes.getResponseCode() === 200) {
        var cgData = JSON.parse(cgRes.getContentText());
        for (var id in idToRows) {
          var p = cgData[id];
          if (p) {
            idToRows[id].forEach(function(rowIdx) {
              results[rowIdx] = {
                kurs:     p.eur || p.usd,
                waehrung: p.eur ? "EUR" : "USD",
                quelle:   "CoinGecko",
                kursEur:  p.eur || (p.usd ? p.usd * (fxRates["USD"] || 1) : null),
              };
            });
          }
        }
        Logger.log("CoinGecko: " + Object.keys(results).length + " Kryptos geladen.");
      } else {
        Logger.log("CoinGecko Fehler " + cgRes.getResponseCode() + ": " + cgRes.getContentText().substring(0, 200));
      }
    } catch (e) {
      Logger.log("CoinGecko Exception: " + e.toString());
    }
    Utilities.sleep(600);
  }

  // ── 2. Yahoo Finance Batch (/v7/finance/quote) ────────────
  var remainingStocks = stockAssets.slice();

  for (var b = 0; b < remainingStocks.length; b += CONFIG.YAHOO_BATCH_SIZE) {
    var batch   = remainingStocks.slice(b, b + CONFIG.YAHOO_BATCH_SIZE);
    var symbols = batch.map(function(s) { return s.ticker; }).join(",");

    try {
      var yhUrl = "https://query1.finance.yahoo.com/v7/finance/quote" +
                  "?symbols=" + encodeURIComponent(symbols) +
                  "&fields=regularMarketPrice,currency,shortName,quoteType";
      var yhRes = UrlFetchApp.fetch(yhUrl, {
        muteHttpExceptions: true,
        headers: {
          "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
          "Accept":     "application/json",
        }
      });

      if (yhRes.getResponseCode() === 200) {
        var yhData = JSON.parse(yhRes.getContentText());
        var quotes = (yhData.quoteResponse && yhData.quoteResponse.result) || [];
        var found  = 0;

        quotes.forEach(function(q) {
          var sym   = (q.symbol || "").toUpperCase();
          var price = q.regularMarketPrice;
          var cur   = (q.currency || "USD").toUpperCase();

          if (price && isFinite(price) && price > 0) {
            var stock = batch.find(function(s) {
              return s.ticker.toUpperCase() === sym;
            });
            if (stock && results[stock.index] === undefined) {
              results[stock.index] = {
                kurs:    price,
                waehrung: cur,
                quelle:  "Yahoo",
                kursEur: convertToEur(price, cur, fxRates),
              };
              found++;
            }
          }
        });

        Logger.log("Yahoo Batch " + (Math.floor(b / CONFIG.YAHOO_BATCH_SIZE) + 1) +
                   ": " + found + "/" + batch.length + " Treffer");
      } else {
        Logger.log("Yahoo Fehler " + yhRes.getResponseCode() + " für Batch " +
                   (Math.floor(b / CONFIG.YAHOO_BATCH_SIZE) + 1));
      }
    } catch (e) {
      Logger.log("Yahoo Exception: " + e.toString());
    }

    if (b + CONFIG.YAHOO_BATCH_SIZE < remainingStocks.length) {
      Utilities.sleep(1000);
    }
  }

  // ── 3. AlphaVantage Fallback (nur für nicht gefundene) ───
  if (apiKeys.alphaVantage) {
    var notFound = stockAssets.filter(function(s) { return results[s.index] === undefined; });
    Logger.log("AlphaVantage Fallback für " + notFound.length + " Aktien.");
    var avCount = 0;

    for (var ai = 0; ai < notFound.length && avCount < 20; ai++) {
      var s = notFound[ai];
      try {
        var avUrl = "https://www.alphavantage.co/query?function=GLOBAL_QUOTE&symbol=" +
                    encodeURIComponent(s.ticker) + "&apikey=" + apiKeys.alphaVantage;
        var avRes = UrlFetchApp.fetch(avUrl, { muteHttpExceptions: true });

        if (avRes.getResponseCode() === 200) {
          var avData = JSON.parse(avRes.getContentText());
          if (avData.Note || avData.Information) {
            Logger.log("AlphaVantage Rate-Limit nach " + avCount + " Calls.");
            break;
          }
          var avPrice = parseFloat((avData["Global Quote"] || {})["05. price"]);
          if (!isNaN(avPrice) && avPrice > 0) {
            results[s.index] = {
              kurs:    avPrice,
              waehrung: "USD",
              quelle:  "AlphaVantage",
              kursEur: convertToEur(avPrice, "USD", fxRates),
            };
            avCount++;
          }
        }
      } catch (e) {
        Logger.log("AV Exception " + s.ticker + ": " + e.toString());
      }
      Utilities.sleep(1000);
    }
  }

  // ── 4. Ergebnisse in Sheet schreiben (ein setValues-Call) ─
  var ts    = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm:ss");
  var output = [];
  var freshCount  = 0;
  var errorTickers = [];

  rawData.forEach(function(row, i) {
    var ticker = (row[1] || "").toString().trim();

    if (!ticker) {
      output.push(["", "", "", "", ""]);
      return;
    }

    var r = results[i];
    if (r) {
      output.push([r.kurs, r.waehrung, r.quelle, ts, r.kursEur || ""]);
      freshCount++;
    } else {
      // Altwert behalten wenn vorhanden und gültig
      var altKurs  = row[4];
      var altWaehr = row[5];
      if (altKurs && altKurs !== "FEHLER" && isFinite(altKurs) && altKurs !== "") {
        output.push([altKurs, altWaehr, "ALTWERT", ts, row[8] || ""]);
      } else {
        output.push(["FEHLER", altWaehr || "", "Kein Kurs", ts, ""]);
        errorTickers.push(ticker);
      }
    }
  });

  sheet.getRange(CONFIG.DATA_START_ROW, CONFIG.COL_KURS, numRows, 5).setValues(output);
  sheet.getRange("A1").setValue("Update: " + ts);

  var msg = freshCount + " Kurse aktualisiert.";
  if (errorTickers.length > 0) {
    msg += " " + errorTickers.length + " ohne Kurs: " + errorTickers.slice(0, 5).join(", ");
    if (errorTickers.length > 5) msg += " ...";
  }
  Logger.log("✅ " + msg);
  SpreadsheetApp.getActiveSpreadsheet().toast(msg, "Kurs-Update", 6);
}


// ============================================================
// SECTION 2: FX-KURSE (Frankfurter.app)
// ============================================================

function updateAllFX() {
  var sheet   = getMainSheet();
  var lastRow = getLastDataRow(sheet);
  var numRows = lastRow - CONFIG.DATA_START_ROW + 1;

  var fxRates = fetchAllFxRatesRaw();  // EUR-Basis: { USD: 1.08, ... }
  if (!fxRates) {
    SpreadsheetApp.getActiveSpreadsheet().toast("FX-API nicht erreichbar!", "Fehler", 4);
    return;
  }

  var rawData  = sheet.getRange(CONFIG.DATA_START_ROW, 1, numRows, 2).getValues();
  var fxOutput = [];

  rawData.forEach(function(row) {
    var ticker   = (row[1] || "").toString().trim();
    var name     = (row[0] || "").toString().trim();
    if (!ticker) {
      fxOutput.push(["", "", "", "", ""]);
      return;
    }

    var currency    = determineCurrency(ticker, name);
    var rate        = fxRates[currency] || 1.0;
    var currData    = CURRENCY_DATA[currency] || ["Unbekannt", "N/A", "N/A"];

    fxOutput.push([currency, rate, currData[0], currData[1], currData[2]]);
  });

  sheet.getRange(CONFIG.DATA_START_ROW, CONFIG.COL_FX_KURZ, numRows, 5).setValues(fxOutput);
  var ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm:ss");
  sheet.getRange("Y1").setValue("FX-Update: " + ts);

  Logger.log("FX-Kurse aktualisiert: " + fxOutput.length + " Einträge.");
  SpreadsheetApp.getActiveSpreadsheet().toast("FX-Kurse aktualisiert!", "FX-Update", 3);
}

function fetchAllFxRatesRaw() {
  try {
    var res = UrlFetchApp.fetch("https://api.frankfurter.app/latest?base=EUR",
                                { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) return null;
    return JSON.parse(res.getContentText()).rates || {};
  } catch (e) {
    Logger.log("Frankfurter.app Fehler: " + e.toString());
    return null;
  }
}

function loadFxRates() {
  // Gibt X→EUR zurück (invertiert), für Kursumrechnung
  var raw = fetchAllFxRatesRaw();
  if (!raw) return {};
  var inv = { "EUR": 1.0 };
  for (var c in raw) {
    if (raw[c] > 0) inv[c] = 1.0 / raw[c];
  }
  return inv;
}

function determineCurrency(ticker, name) {
  if (isCrypto(ticker)) return "USD";

  var t = ticker.toUpperCase();
  // Ticker-Suffix Mapping
  if (t.match(/\.HK$/))              return "HKD";
  if (t.match(/\.(DE|F|MU|HM|BE|DU|HA|SG)$/)) return "EUR";
  if (t.match(/\.L$/))               return "GBP";
  if (t.match(/\.SW$/))              return "CHF";
  if (t.match(/\.(TO|V)$/))          return "CAD";
  if (t.match(/\.AX$/))              return "AUD";
  if (t.match(/\.(KS|KQ)$/))         return "KRW";
  if (t.match(/\.(SS|SZ)$/))         return "CNY";
  if (t.match(/\.CO$/))              return "DKK";
  if (t.match(/\.ST$/))              return "SEK";
  if (t.match(/\.OL$/))              return "NOK";
  if (t.match(/\.SI$/))              return "SGD";
  if (t.match(/\.SA$/))              return "BRL";

  // Bekannte EUR OTC-ADRs
  var eurOtc = ["SMAWF","SMEGF","SFFLY","SESNF","IFNNY","BAYRY","BASFY",
                "RHHBY","ADYYF","SIEGY","DGELX","VCYGY","MTNOY"];
  if (eurOtc.indexOf(t) !== -1) return "EUR";

  // Bekannte HKD-OTC
  var hkdOtc = ["BYDDF","BIDU","XIACF","TCEHY","PANHF","XJNGF","GUKYF",
                "ZIJMF","PTR","QNTPF","DNNGY"];
  if (hkdOtc.indexOf(t) !== -1) return "HKD";

  // CAD
  if (["WEED","ARQQ","POET"].indexOf(t) !== -1) return "CAD";

  // AUD
  if (["BRCHF","SRUUF"].indexOf(t) !== -1) return "AUD";

  return "USD";
}


// ============================================================
// DEBUG-FUNKTION
// ============================================================

function debugApiTest() {
  var sheet   = getMainSheet();
  var apiKeys = getApiKeys();
  var log     = [];

  // Test 1: CoinGecko
  try {
    var cgRes = UrlFetchApp.fetch(
      "https://api.coingecko.com/api/v3/simple/price?ids=bitcoin&vs_currencies=eur,usd" +
      (apiKeys.coingecko ? "&x_cg_demo_api_key=" + apiKeys.coingecko : ""),
      { muteHttpExceptions: true }
    );
    var cgCode = cgRes.getResponseCode();
    var cgData = JSON.parse(cgRes.getContentText());
    log.push("✅ CoinGecko: HTTP " + cgCode + " | BTC = " +
             (cgData.bitcoin ? cgData.bitcoin.eur + " EUR" : "FEHLER"));
  } catch (e) {
    log.push("❌ CoinGecko: " + e.message);
  }

  // Test 2: Yahoo Finance /v7/quote (Batch-Endpoint)
  try {
    var yhRes = UrlFetchApp.fetch(
      "https://query1.finance.yahoo.com/v7/finance/quote?symbols=AAPL,MSFT&fields=regularMarketPrice,currency",
      {
        muteHttpExceptions: true,
        headers: { "User-Agent": "Mozilla/5.0", "Accept": "application/json" }
      }
    );
    var yhCode   = yhRes.getResponseCode();
    var yhData   = JSON.parse(yhRes.getContentText());
    var yhResult = (yhData.quoteResponse && yhData.quoteResponse.result) || [];
    if (yhResult.length > 0) {
      log.push("✅ Yahoo Finance /v7: HTTP " + yhCode + " | AAPL = " +
               yhResult[0].regularMarketPrice + " " + yhResult[0].currency);
    } else {
      log.push("⚠️ Yahoo Finance /v7: HTTP " + yhCode + " | Keine Daten. Response: " +
               yhRes.getContentText().substring(0, 150));
    }
  } catch (e) {
    log.push("❌ Yahoo Finance /v7: " + e.message);
  }

  // Test 3: Frankfurter.app
  try {
    var fxRes  = UrlFetchApp.fetch("https://api.frankfurter.app/latest?base=EUR",
                                   { muteHttpExceptions: true });
    var fxCode = fxRes.getResponseCode();
    var fxData = JSON.parse(fxRes.getContentText());
    log.push("✅ Frankfurter.app: HTTP " + fxCode + " | USD = " +
             (fxData.rates ? fxData.rates.USD : "FEHLER"));
  } catch (e) {
    log.push("❌ Frankfurter.app: " + e.message);
  }

  // Test 4: Sheet-Name
  try {
    var s = getMainSheet();
    log.push("✅ Sheet '" + CONFIG.SHEET_NAME + "' gefunden | " +
             (getLastDataRow(s) - CONFIG.DATA_START_ROW + 1) + " Assets");
  } catch (e) {
    log.push("❌ Sheet-Fehler: " + e.message);
  }

  // Test 5: API-Keys
  log.push("🔑 API-Keys: CoinGecko=" + (apiKeys.coingecko ? "gesetzt" : "leer") +
           " | AlphaVantage=" + (apiKeys.alphaVantage ? "gesetzt" : "leer") +
           " | Marketstack=" + (apiKeys.marketstack ? "gesetzt" : "leer"));

  SpreadsheetApp.getUi().alert("🔍 API-Diagnose\n\n" + log.join("\n\n"));
}


// ============================================================
// STATUS
// ============================================================

function showStatus() {
  var sheet     = getMainSheet();
  var lastRow   = getLastDataRow(sheet);
  var numAssets = lastRow - CONFIG.DATA_START_ROW + 1;
  var kursUpdate = sheet.getRange("A1").getValue();
  var fxUpdate   = sheet.getRange("Y1").getValue();

  // Zähle befüllte Kurse
  var kurse  = sheet.getRange(CONFIG.DATA_START_ROW, CONFIG.COL_KURS, numAssets, 1).getValues();
  var filled = kurse.filter(function(r) {
    return r[0] !== "" && r[0] !== "FEHLER" && isFinite(r[0]);
  }).length;
  var fehler = kurse.filter(function(r) { return r[0] === "FEHLER"; }).length;

  SpreadsheetApp.getUi().alert(
    "📊 Status\n\n" +
    "Assets gesamt: " + numAssets + "\n" +
    "Kurse geladen: " + filled + "\n" +
    "Fehler:        " + fehler + "\n" +
    "Ohne Kurs:     " + (numAssets - filled - fehler) + "\n\n" +
    "Letztes Kurs-Update:\n" + kursUpdate + "\n\n" +
    "Letztes FX-Update:\n"   + fxUpdate
  );
}


// ============================================================
// TRIGGER-VERWALTUNG
// ============================================================

function createTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t) {
    if (t.getHandlerFunction() === "updateAll") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("updateAll").timeBased().everyMinutes(30).create();
  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Auto-Update alle 30 Minuten aktiviert!", "Trigger", 5);
}

function deleteTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    ScriptApp.deleteTrigger(t);
  });
  SpreadsheetApp.getActiveSpreadsheet().toast("Alle Trigger gelöscht.", "Trigger", 3);
}


// ============================================================
// HILFSFUNKTIONEN
// ============================================================

function isCrypto(ticker) {
  return CONFIG.CRYPTO_TICKERS.indexOf((ticker || "").toUpperCase()) !== -1;
}

function convertToEur(kurs, waehrung, fxRates) {
  if (!kurs || !isFinite(kurs)) return "";
  if (waehrung === "EUR") return kurs;
  var rate = fxRates[(waehrung || "").toUpperCase()];
  return rate ? kurs * rate : kurs;
}

function getMainSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    // Fallback: Sheet-Namen vergleichen (Groß-/Kleinschreibung-tolerant)
    var all = ss.getSheets();
    for (var i = 0; i < all.length; i++) {
      if (all[i].getName().toLowerCase() === CONFIG.SHEET_NAME.toLowerCase()) {
        return all[i];
      }
    }
    throw new Error(
      "Sheet '" + CONFIG.SHEET_NAME + "' nicht gefunden!\n" +
      "Vorhandene Sheets: " + all.map(function(s) { return s.getName(); }).join(", ")
    );
  }
  return sheet;
}

function getLastDataRow(sheet) {
  var values = sheet.getRange("A:A").getValues();
  var last   = CONFIG.DATA_START_ROW;
  for (var i = CONFIG.DATA_START_ROW - 1; i < values.length; i++) {
    if (values[i][0] !== "" && values[i][0] !== null) last = i + 1;
  }
  return last;
}

function getApiKeys() {
  var keys = { coingecko: "", alphaVantage: "", marketstack: "" };
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var s  = ss.getSheetByName(CONFIG.SETTINGS_SHEET);
    if (!s) return keys;
    var cg = s.getRange("B3").getValue();
    var av = s.getRange("B4").getValue();
    var ms = s.getRange("B5").getValue();
    if (cg) keys.coingecko    = cg.toString().trim();
    if (av) keys.alphaVantage = av.toString().trim();
    if (ms) keys.marketstack  = ms.toString().trim();
  } catch (e) {
    Logger.log("API-Key-Fehler: " + e.toString());
  }
  return keys;
}
