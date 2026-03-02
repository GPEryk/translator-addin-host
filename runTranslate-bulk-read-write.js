/**
 * runTranslate-bulk-read-write.js — SAMODZIELNY skrypt tłumaczenia bulk
 *
 * Nie zależy od funkcji z taskpane.js (IIFE). Implementuje WSZYSTKO lokalnie:
 * getApiKey, callOpenAI, refreshGlossaryAll, applyGlossaryTokens, log, setProgress, itp.
 *
 * Odczyt zaznaczenia w małych partiach (MAX_CELLS_PER_SYNC), żeby uniknąć 500.
 * Tłumaczenie w paczkach po RECORDS_PER_BATCH (10): po każdej paczce od razu zapis
 * do Excela (na bieżąco), bez jednego wielkiego zapisu na końcu.
 */

// ======================== STAŁE KONFIGURACYJNE ========================

var HEADER_ROW = 1;
var MAX_ROWS_PER_SYNC = 1000;
var MAX_CELLS_PER_SYNC = 250;
var RECORDS_PER_BATCH = 10;
var API_DELAY_MS = 400;
var RETRY_DELAY_MS = 800;
var API_429_RETRY_DELAY_MS = 3000;
var OPENAI_API_URL = "https://api.openai.com/v1/chat/completions";
var OPENAI_MODEL = "gpt-4o-mini";

// ======================== CACHE GLOSARIUSZA ========================

var glossaryCache = null;
var plGlossaryCache = null;

// ======================== HELPERY UI ========================

function getApiKey() {
  try {
    if (typeof Office !== "undefined" && Office.context && Office.context.roamingSettings) {
      return Office.context.roamingSettings.get("openai_api_key") || "";
    }
  } catch (e) {
    console.log("roamingSettings not available, using localStorage");
  }
  try {
    return localStorage.getItem("openai_api_key") || "";
  } catch (e) {
    console.error("Storage not available:", e);
    return "";
  }
}

function normalizeHeader(h) {
  return (h != null ? h : "").toString().replace(/\u00A0/g, " ").trim().toUpperCase();
}

function log(msg) {
  var el = document.getElementById("log");
  if (el) {
    el.textContent += msg + "\n";
    el.scrollTop = el.scrollHeight;
  }
}

function setProgress(pct) {
  var bar = document.getElementById("progress");
  var fill = document.getElementById("progressBar");
  if (bar && fill) {
    if (pct != null) {
      bar.classList.remove("hidden");
      fill.style.width = pct + "%";
    } else {
      bar.classList.add("hidden");
    }
  }
}

function customAlert(msg) {
  return new Promise(function (resolve) {
    var modal = document.getElementById("alertModal");
    var message = document.getElementById("alertMessage");
    var okBtn = document.getElementById("alertOk");
    if (!modal || !message || !okBtn) {
      alert(msg);
      resolve();
      return;
    }
    message.textContent = msg;
    modal.classList.remove("hidden");
    okBtn.onclick = function () {
      modal.classList.add("hidden");
      okBtn.onclick = null;
      resolve();
    };
  });
}

// ======================== LOGI SZCZEGÓŁOWE ========================

function _detailedLogNow() {
  var d = new Date();
  return d.toTimeString().slice(0, 12);
}

function detailedLog(step, message, data) {
  var line = "[" + _detailedLogNow() + "] [" + step + "] " + message;
  var fullLine = line;
  if (data !== undefined && data !== null) {
    var dataStr = typeof data === "object" ? JSON.stringify(data, null, 2) : String(data);
    fullLine += "\n  " + dataStr.replace(/\n/g, "\n  ");
  }
  fullLine += "\n";

  var el = document.getElementById("detailedLog");
  if (el) {
    el.textContent += fullLine;
    el.scrollTop = el.scrollHeight;
  }
  var mainLog = document.getElementById("log");
  if (mainLog) {
    mainLog.textContent += line + "\n";
    mainLog.scrollTop = mainLog.scrollHeight;
  }
}

function clearDetailedLog() {
  var el = document.getElementById("detailedLog");
  if (el) el.textContent = "";
}

function _appendToLog(el, text) {
  if (!el) return;
  el.textContent += text + "\n";
  el.scrollTop = el.scrollHeight;
}

// ======================== GLOSARIUSZ: ODCZYT Z ARKUSZY EXCEL ========================

async function refreshGlossaryAll() {
  await Excel.run(async function (ctx) {
    var sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();

    var enGloss = {};
    var plGloss = {};

    for (var si = 0; si < sheets.items.length; si++) {
      var sh = sheets.items[si];
      var name = sh.name.toUpperCase();

      if (name.startsWith("EN-") && name.length >= 5) {
        var tgtLang = normalizeHeader(name.slice(3));
        if (!tgtLang) continue;

        var ur;
        try {
          ur = sh.getUsedRange();
          ur.load("values,rowCount,columnCount");
          await ctx.sync();
        } catch (e) {
          console.log("Sheet " + name + " is empty or error");
          continue;
        }

        var vals = ur.values || [];
        if (vals.length < 2) continue;

        var hdr = vals[0].map(function (h) { return normalizeHeader(h); });
        var enCol = hdr.indexOf("EN");
        var tCol = hdr.indexOf(tgtLang);
        if (enCol < 0 || tCol < 0) continue;

        for (var r = 1; r < vals.length; r++) {
          var enVal = (vals[r][enCol] || "").toString().trim();
          var tVal = (vals[r][tCol] || "").toString().trim();
          if (enVal && tVal) {
            if (!enGloss[enVal]) enGloss[enVal] = {};
            enGloss[enVal][tgtLang] = tVal;
          }
        }
      }

      if (name.startsWith("PL-") && name.length >= 5) {
        var tgtLang2 = normalizeHeader(name.slice(3));
        if (!tgtLang2) continue;

        var ur2;
        try {
          ur2 = sh.getUsedRange();
          ur2.load("values,rowCount,columnCount");
          await ctx.sync();
        } catch (e) {
          console.log("Sheet " + name + " is empty or error");
          continue;
        }

        var vals2 = ur2.values || [];
        if (vals2.length < 2) continue;

        var hdr2 = vals2[0].map(function (h) { return normalizeHeader(h); });
        var plCol = hdr2.indexOf("PL");
        var tCol2 = hdr2.indexOf(tgtLang2);
        if (plCol < 0 || tCol2 < 0) continue;

        for (var r2 = 1; r2 < vals2.length; r2++) {
          var plVal = (vals2[r2][plCol] || "").toString().trim();
          var tVal2 = (vals2[r2][tCol2] || "").toString().trim();
          if (plVal && tVal2) {
            if (!plGloss[plVal]) plGloss[plVal] = {};
            plGloss[plVal][tgtLang2] = tVal2;
          }
        }
      }
    }

    glossaryCache = enGloss;
    plGlossaryCache = plGloss;
  });
  return glossaryCache;
}

// ======================== TOKENIZACJA GLOSARIUSZA ========================
// Obsługa fraz wielowyrazowych: replace całych fraz, najdłuższe najpierw,
// z word-boundary (\b) żeby nie łapać podciągów w środku słów.

function _escapeRegex(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function applyGlossaryTokens(lines, targetLang, glossary) {
  var tokenMap = {};
  var tokenCounter = 0;

  var entries = [];
  var keys = Object.keys(glossary || {});
  for (var ki = 0; ki < keys.length; ki++) {
    var srcPhrase = keys[ki];
    var translations = glossary[srcPhrase];
    if (translations && translations[targetLang]) {
      entries.push({ src: srcPhrase, tgt: translations[targetLang] });
    }
  }

  entries.sort(function (a, b) { return b.src.length - a.src.length; });

  var tokenizedLines = lines.map(function (line) {
    var result = line;
    for (var ei = 0; ei < entries.length; ei++) {
      var entry = entries[ei];
      var escaped = _escapeRegex(entry.src);
      var regex = new RegExp("\\b" + escaped + "\\b", "gi");
      result = result.replace(regex, function () {
        tokenCounter++;
        var token = "[[G" + tokenCounter + "]]";
        tokenMap[token] = entry.tgt;
        return token;
      });
    }
    return result;
  });

  return { tokenizedLines: tokenizedLines, tokenMap: tokenMap };
}

function restoreGlossaryTokens(text, tokenMap) {
  var result = text;
  var tokens = Object.keys(tokenMap);
  for (var ti = 0; ti < tokens.length; ti++) {
    result = result.split(tokens[ti]).join(tokenMap[tokens[ti]]);
  }
  return result;
}

// ======================== WYWOŁANIE OPENAI API ========================

var LANG_NAMES = {
  EN: "English", PL: "Polish", DE: "German", FR: "French", ES: "Spanish",
  IT: "Italian", RO: "Romanian", RU: "Russian", UA: "Ukrainian", CS: "Czech",
  SE: "Swedish", NL: "Dutch", DK: "Danish", PT: "Portuguese", HU: "Hungarian",
  BG: "Bulgarian", HR: "Croatian", SK: "Slovak", SL: "Slovenian", FI: "Finnish",
  NO: "Norwegian", EL: "Greek", TR: "Turkish", JA: "Japanese", ZH: "Chinese",
  KO: "Korean", AR: "Arabic", HE: "Hebrew", TH: "Thai", VI: "Vietnamese"
};

async function callOpenAI(targetLang, tokenizedLines, sourceLang) {
  sourceLang = sourceLang || "EN";

  var srcName = LANG_NAMES[sourceLang] || sourceLang;
  var tgtName = LANG_NAMES[targetLang] || targetLang;

  var systemPrompt =
    "You are a professional translator for industrial/gastronomic equipment parts. " +
    "Translate product names and technical terms from " + srcName + " to " + tgtName + ". " +
    "ALWAYS translate technical terms to the target language (e.g. 'microswitch' → 'microîntrerupător' in Romanian). " +
    "Preserve brand names (e.g. Robot Coupe, Bosch), model codes, part numbers, units, and casing. " +
    "Do not add extra words or paraphrase. Return one translation per line in the same order. " +
    "Output only translations, nothing else. " +
    "Do NOT translate or alter tokens like [[G1]], [[G2]]; keep them exactly as they are.";

  var userPrompt =
    "Translate from " + srcName + " to " + tgtName + ":\n" + tokenizedLines.join("\n");

  var requestBody = {
    model: OPENAI_MODEL,
    messages: [
      { role: "system", content: systemPrompt },
      { role: "user", content: userPrompt }
    ],
    max_tokens: 2000,
    temperature: 0.3
  };

  var apiKey = getApiKey();
  if (!apiKey) throw new Error("Brak klucza API! Dodaj klucz w sekcji Konfiguracja.");

  var response = await fetch(OPENAI_API_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Authorization": "Bearer " + apiKey
    },
    body: JSON.stringify(requestBody)
  });

  if (!response.ok) {
    var errorText = await response.text();
    throw new Error("OpenAI API error: " + response.status + " - " + errorText);
  }

  var data = await response.json();
  var content =
    (data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content) || "";
  return content.split(/\r?\n/);
}

// ======================== GŁÓWNA FUNKCJA TŁUMACZENIA ========================

async function runTranslate() {
  _appendToLog(document.getElementById("detailedLog"), ">>> runTranslate BULK wywołane <<<");
  _appendToLog(document.getElementById("log"), ">>> runTranslate BULK wywołane <<<");

  var apiKey = getApiKey();
  if (!apiKey) {
    await customAlert("⚠️ Najpierw zapisz swój klucz OpenAI API w sekcji Konfiguracja!");
    return;
  }

  var skipFilled = document.getElementById("skipFilled").checked;
  var sourceLang = document.getElementById("sourceLangSelect").value || "EN";

  if (!glossaryCache) {
    await refreshGlossaryAll();
  }

  clearDetailedLog();
  detailedLog("START", "Rozpoczęto tłumaczenie", { sourceLang: sourceLang, skipFilled: skipFilled });

  log("Czytam zaznaczenie... (źródło: " + sourceLang + ")");

  try {
    await Excel.run(async function (ctx) {
      var sheet = ctx.workbook.worksheets.getActiveWorksheet();
      var sel = ctx.workbook.getSelectedRange();
      detailedLog("EXCEL", "Pobieranie zaznaczenia (load + sync)...", {});
      sel.load(["rowCount", "columnCount", "rowIndex", "columnIndex"]);
      await ctx.sync();
      detailedLog("EXCEL", "Sync zakończony – mam wymiary zaznaczenia", {});

      var rowCount = sel.rowCount;
      var columnCount = sel.columnCount;
      var rowIndex = sel.rowIndex;
      var columnIndex = sel.columnIndex;

      var totalCells = rowCount * columnCount;
      detailedLog("EXCEL_READ", "Odczytano zaznaczenie z arkusza", {
        rowCount: rowCount,
        columnCount: columnCount,
        totalCells: totalCells,
        rowIndex: rowIndex,
        columnIndex: columnIndex
      });
      if (totalCells > 3000) {
        log("⚠️ Duże zaznaczenie (" + totalCells + " komórek). Przy błędach spróbuj mniejszego zakresu.");
      }

      // Nagłówki zaznaczonych kolumn (języki docelowe)
      detailedLog("EXCEL", "Ładowanie nagłówków zaznaczonych kolumn...", { columnIndex: columnIndex, columnCount: columnCount });
      var selHeaderRange = sheet.getRangeByIndexes(HEADER_ROW - 1, columnIndex, 1, columnCount);
      selHeaderRange.load("values");
      await ctx.sync();

      var selHeaderRaw = selHeaderRange.values && selHeaderRange.values[0] ? selHeaderRange.values[0] : [];
      var header = selHeaderRaw.map(function (h) { return normalizeHeader(h); });
      detailedLog("HEADER_SEL", "Nagłówki zaznaczonych kolumn", { selHeaderRaw: selHeaderRaw, header: header });

      // Szukaj kolumny źródłowej (EN/PL) — najpierw w zaznaczeniu, potem w całym wierszu 1
      var srcCol = -1;
      var srcColInSel = header.indexOf(sourceLang);
      if (srcColInSel >= 0) {
        srcCol = columnIndex + srcColInSel;
        detailedLog("HEADER", "Kolumna źródłowa znaleziona w zaznaczeniu", { sourceLang: sourceLang, srcCol: srcCol });
      } else {
        detailedLog("HEADER", "Kolumna źródłowa nie jest w zaznaczeniu — szukam w całym wierszu 1...", { sourceLang: sourceLang });
        var usedRange = sheet.getUsedRange();
        usedRange.load("columnCount");
        await ctx.sync();

        var fullHeaderRange = sheet.getRangeByIndexes(HEADER_ROW - 1, 0, 1, usedRange.columnCount);
        fullHeaderRange.load("values");
        await ctx.sync();

        var fullHeaderRaw = fullHeaderRange.values && fullHeaderRange.values[0] ? fullHeaderRange.values[0] : [];
        var fullHeader = fullHeaderRaw.map(function (h) { return normalizeHeader(h); });
        srcCol = fullHeader.indexOf(sourceLang);
        detailedLog("HEADER_FULL", "Przeszukano cały wiersz 1", { fullHeader: fullHeader, srcCol: srcCol });
      }

      if (srcCol < 0) {
        log("BŁĄD: brak kolumny " + sourceLang + " w wierszu nagłówków (1). Sprawdź czy arkusz ma nagłówek " + sourceLang + ".");
        detailedLog("ERROR", "Brak kolumny źródłowej w całym wierszu 1", { sourceLang: sourceLang, header: header });
        return;
      }

      var currentGlossary = sourceLang === "PL" ? plGlossaryCache : glossaryCache;

      // ----- BULK READ w partiach -----
      var selValues = [];
      var sourceColValues = [];
      var maxRowsPerChunk = Math.max(1, Math.min(MAX_ROWS_PER_SYNC, Math.floor(MAX_CELLS_PER_SYNC / columnCount)));
      var totalChunks = Math.ceil(rowCount / maxRowsPerChunk);
      detailedLog("CHUNK_PLAN", "Plan odczytu partiami", {
        maxRowsPerChunk: maxRowsPerChunk,
        totalChunks: totalChunks,
        maxCellsPerSync: MAX_CELLS_PER_SYNC
      });
      if (totalChunks > 1) {
        log("Zaznaczenie: " + (rowCount * columnCount) + " komórek → odczyt w partiach (po max " + MAX_CELLS_PER_SYNC + " komórek).");
      }
      var is500 = function (e) {
        return e && (e.message || e.code || "" + e) && /500|Internal|RichApi|błąd wewnętrzny/i.test(e.message || e.code || "" + e);
      };
      var rowOffset = 0, chunkNum = 0;
      while (rowOffset < rowCount) {
        var chunkRows = Math.min(maxRowsPerChunk, rowCount - rowOffset);
        var chunkDone = false;
        while (!chunkDone && chunkRows >= 1) {
          try {
            if (totalChunks > 1) {
              log("  Czytam partię " + (chunkNum + 1) + "...");
              detailedLog("CHUNK_READ", "Partia odczytu " + (chunkNum + 1) + "/" + totalChunks, {
                chunkRows: chunkRows,
                rowOffset: rowOffset,
                cellsInChunk: chunkRows * columnCount
              });
            }
            var chunkSel = sheet.getRangeByIndexes(rowIndex + rowOffset, columnIndex, chunkRows, columnCount);
            var chunkSrc = sheet.getRangeByIndexes(rowIndex + rowOffset, srcCol, chunkRows, 1);
            chunkSel.load("values");
            chunkSrc.load("values");
            await ctx.sync();
            var cv = chunkSel.values || [];
            var sv = chunkSrc.values || [];
            for (var ci = 0; ci < cv.length; ci++) selValues.push(cv[ci] ? cv[ci].slice() : []);
            for (var si = 0; si < sv.length; si++) sourceColValues.push(sv[si] ? sv[si].slice() : []);
            detailedLog("CHUNK_READ_OK", "Odebrano partię " + (chunkNum + 1) + ": " + cv.length + " wierszy", {
              rowsRead: cv.length,
              sourceColSample: (sv.slice(0, 3) || []).map(function (r) { return r && r[0]; })
            });
            rowOffset += chunkRows;
            chunkNum++;
            chunkDone = true;
          } catch (e) {
            if (chunkRows > 1 && is500(e)) {
              chunkRows = Math.max(1, Math.floor(chunkRows / 2));
              detailedLog("CHUNK_500", "Błąd 500 przy odczycie – zmniejszam partię i ponawiam", {
                newChunkRows: chunkRows,
                waitMs: RETRY_DELAY_MS,
                error: (e && e.message) || String(e)
              });
              log("  Błąd serwera (500) – czekam " + (RETRY_DELAY_MS / 1000) + " s, ponawiam z mniejszą partią (" + chunkRows + " wierszy)...");
              await new Promise(function (r) { setTimeout(r, RETRY_DELAY_MS); });
            } else throw e;
          }
        }
      }

      var items = [];
      for (var r = 0; r < rowCount; r++) {
        for (var c = 0; c < columnCount; c++) {
          var absRow = rowIndex + r;
          var absCol = columnIndex + c;
          var tgtLang = header[c];
          var src = (sourceColValues[r] && sourceColValues[r][0] != null ? sourceColValues[r][0] : "").toString().trim();
          var tgt = (selValues[r] && selValues[r][c] != null ? selValues[r][c] : "").toString().trim();

          items.push({ absRow: absRow, absCol: absCol, r: r, c: c, tgtLang: tgtLang, src: src, tgt: tgt });
        }
      }

      detailedLog("ITEMS", "Zbudowano listę par (wiersz, kolumna, język docelowy, źródło, cel)", {
        totalItems: items.length,
        sample: items.slice(0, 5).map(function (it) {
          return { absRow: it.absRow, absCol: it.absCol, tgtLang: it.tgtLang, src: (it.src || "").slice(0, 40), tgt: (it.tgt || "").slice(0, 40) };
        })
      });

      var groups = new Map();
      for (var ii = 0; ii < items.length; ii++) {
        var it = items[ii];
        if (!it.src) continue;
        if (skipFilled && it.tgt) continue;
        if (it.tgtLang === sourceLang) continue;

        if (!groups.has(it.tgtLang)) groups.set(it.tgtLang, []);
        groups.get(it.tgtLang).push({ absRow: it.absRow, absCol: it.absCol, r: it.r, c: it.c, tgtLang: it.tgtLang, src: it.src, tgt: it.tgt });
      }

      var total = 0;
      groups.forEach(function (arr) { total += arr.length; });

      var byLang = {};
      groups.forEach(function (arr, k) { byLang[k] = arr.length; });
      detailedLog("GROUPS", "Przygotowano dane do tłumaczenia (grupy po językach)", { total: total, byLang: byLang });

      if (total === 0) {
        log("Brak danych do tłumaczenia.");
        detailedLog("DONE", "Brak danych do tłumaczenia – zakończono");
        return;
      }

      var done = 0;
      var batchSize = RECORDS_PER_BATCH;

      var langEntries = Array.from(groups.entries());
      for (var li = 0; li < langEntries.length; li++) {
        var lang = langEntries[li][0];
        var arr = langEntries[li][1];
        if (!lang) continue;
        log(sourceLang + " → " + lang + ": " + arr.length + " rekordów (zapis co " + batchSize + ")");
        detailedLog("LANG_START", "Start języka docelowego: " + lang, { recordCount: arr.length, batchSize: batchSize });

        for (var i = 0; i < arr.length; i += batchSize) {
          var batch = arr.slice(i, i + batchSize);
          var lines = batch.map(function (x) { return x.src; });
          var batchNum = Math.floor(i / batchSize) + 1;

          detailedLog("BATCH_SOURCE", "Paczka " + batchNum + ": teksty z Excela (źródło)", { lines: lines });

          var tokenResult = applyGlossaryTokens(lines, lang, currentGlossary);
          var tokenizedLines = tokenResult.tokenizedLines;
          var tokenMap = tokenResult.tokenMap;

          detailedLog("BATCH_TOKENIZED", "Paczka " + batchNum + ": po glosariuszu (wysyłane do API)", {
            tokenizedLines: tokenizedLines,
            tokenMapKeys: Object.keys(tokenMap),
            tokenMap: tokenMap
          });

          detailedLog("API_SEND", "Wysyłam do OpenAI API (" + sourceLang + " → " + lang + "), paczka " + batchNum, {
            url: OPENAI_API_URL,
            lineCount: tokenizedLines.length,
            linesSent: tokenizedLines
          });

          var is429 = function (e) {
            return e && (e.message || "" + e) && /429|rate limit|Too Many Requests/i.test(e.message || "" + e);
          };
          var translatedLines;
          try {
            translatedLines = await callOpenAI(lang, tokenizedLines, sourceLang);
          } catch (apiErr) {
            if (is429(apiErr)) {
              detailedLog("API_429", "Limit 429 – czekam i ponawiam", { waitMs: API_429_RETRY_DELAY_MS });
              log("  Limit API (429) – czekam " + (API_429_RETRY_DELAY_MS / 1000) + " s, ponawiam...");
              await new Promise(function (r) { setTimeout(r, API_429_RETRY_DELAY_MS); });
              translatedLines = await callOpenAI(lang, tokenizedLines, sourceLang);
            } else {
              detailedLog("API_ERROR", "Błąd wywołania OpenAI", { message: (apiErr && apiErr.message) || String(apiErr) });
              throw apiErr;
            }
          }

          detailedLog("API_RAW", "Odebrano z OpenAI (RAW) – paczka " + batchNum, {
            rawType: typeof translatedLines,
            rawIsArray: Array.isArray(translatedLines),
            rawLength: Array.isArray(translatedLines) ? translatedLines.length : 0,
            rawContent: translatedLines
          });

          if (typeof translatedLines === "string") {
            translatedLines = translatedLines.split(/\r?\n/).map(function (s) { return s.trim(); });
          }
          if (!Array.isArray(translatedLines)) translatedLines = [];
          while (translatedLines.length < batch.length) translatedLines.push("");
          translatedLines = translatedLines.slice(0, batch.length);

          detailedLog("API_RESPONSE", "Odpowiedź z OpenAI (po normalizacji) – paczka " + batchNum, {
            lineCount: translatedLines.length,
            lines: translatedLines
          });

          var restoredList = batch.map(function (batchItem, j) {
            return restoreGlossaryTokens(translatedLines[j] || "", tokenMap);
          });
          detailedLog("RESTORE_GLOSSARY", "Paczka " + batchNum + ": po przywróceniu glosariusza", { restoredList: restoredList });

          if (API_DELAY_MS > 0) {
            detailedLog("DELAY", "Opóźnienie " + API_DELAY_MS + " ms przed kolejną paczką", {});
            await new Promise(function (r) { setTimeout(r, API_DELAY_MS); });
          }

          var isValid = function (j) {
            var res = (restoredList[j] || "").trim();
            if (!res) return false;
            return true;
          };
          var MAX_RETRY_PER_CELL = 5;
          var needRetry = [];
          for (var nj = 0; nj < batch.length; nj++) { if (!isValid(nj)) needRetry.push(nj); }
          var round = 0;
          while (needRetry.length > 0 && round < MAX_RETRY_PER_CELL) {
            round++;
            detailedLog("RETRY_START", "Ponowne wywołanie API dla " + needRetry.length + " komórek, runda " + round, { indices: needRetry });
            log("  Uzupełniam brakujące / błędne (" + needRetry.length + " komórek), próba " + round + "...");
            for (var ri = 0; ri < needRetry.length; ri++) {
              var j = needRetry[ri];
              await new Promise(function (r) { setTimeout(r, 250); });
              detailedLog("RETRY_SEND", "Ponowne wysłanie do API (indeks " + j + ")", { line: tokenizedLines[j] });
              try {
                var oneLine = await callOpenAI(lang, [tokenizedLines[j]], sourceLang);
                detailedLog("RETRY_RAW", "Odebrano (RAW) dla indeksu " + j, { raw: oneLine });
                var one = Array.isArray(oneLine) ? (oneLine[0] || "") : ("" + (oneLine || "")).trim();
                restoredList[j] = restoreGlossaryTokens(one || "", tokenMap);
                detailedLog("RETRY_RESTORED", "Po przywróceniu glosariusza [" + j + "]", { value: restoredList[j] });
              } catch (retryErr) {
                detailedLog("RETRY_ERROR", "Błąd API dla komórki " + j, { message: (retryErr && retryErr.message) || String(retryErr) });
                if (round < MAX_RETRY_PER_CELL) log("  Błąd API dla komórki – ponowię w następnej rundzie.");
              }
            }
            var stillInvalid = [];
            for (var nk = 0; nk < needRetry.length; nk++) { if (!isValid(needRetry[nk])) stillInvalid.push(needRetry[nk]); }
            needRetry = stillInvalid;
          }
          if (needRetry.length > 0) {
            log("  Uwaga: " + needRetry.length + " komórek nadal niepoprawnych po " + MAX_RETRY_PER_CELL + " próbach – zapisuję ostatni wynik.");
            detailedLog("RETRY", "Część komórek uzupełniona ponownym wywołaniem API", {
              stillInvalid: needRetry.length,
              maxRetries: MAX_RETRY_PER_CELL
            });
          }

          var writeBatch = function () {
            for (var wj = 0; wj < batch.length; wj++) {
              var restored = restoredList[wj] != null ? restoredList[wj] : "";
              var wit = batch[wj];
              var cellRange = sheet.getRangeByIndexes(wit.absRow, wit.absCol, 1, 1);
              cellRange.values = [[restored]];
            }
          };
          var writePlan = batch.map(function (wit, wj) {
            return {
              row: wit.absRow,
              col: wit.absCol,
              value: (restoredList[wj] != null ? restoredList[wj] : "").slice(0, 80)
            };
          });
          detailedLog("EXCEL_WRITE", "Zapis do arkusza – paczka " + batchNum, {
            cellsWritten: batch.length,
            done: done + batch.length,
            total: total,
            cells: writePlan
          });
          try {
            writeBatch();
            await ctx.sync();
            detailedLog("EXCEL_WRITE_OK", "Sync zapisu zakończony – paczka " + batchNum, {});
          } catch (writeErr) {
            if (is500(writeErr) && batch.length > 1) {
              detailedLog("EXCEL_WRITE_500", "Błąd 500 przy zapisie – zapisuję po 1 komórce", {
                error: (writeErr && writeErr.message) || String(writeErr)
              });
              log("  Błąd 500 przy zapisie – zapisuję po 1 komórce...");
              for (var wj2 = 0; wj2 < batch.length; wj2++) {
                var restored2 = restoredList[wj2] != null ? restoredList[wj2] : "";
                var wit2 = batch[wj2];
                sheet.getRangeByIndexes(wit2.absRow, wit2.absCol, 1, 1).values = [[restored2]];
                await ctx.sync();
                detailedLog("EXCEL_WRITE_ONE", "Zapisano komórkę (" + wit2.absRow + ", " + wit2.absCol + ")", { value: restored2.slice(0, 60) });
              }
            } else throw writeErr;
          }

          done += batch.length;
          setProgress(Math.round((done / total) * 100));
          if (done % 50 === 0 || done === total) {
            log("  Zapisano " + done + "/" + total);
          }
        }
      }

      setProgress(null);
      log("Gotowe.");
      detailedLog("DONE", "Tłumaczenie zakończone pomyślnie", { total: total });
    });
  } catch (err) {
    setProgress(null);
    var msg = (err && (err.message || err.code || err.toString())) || "";
    detailedLog("ERROR", "Wystąpił błąd", { message: msg });
    var isServerError = /500|Internal|RichApi\.Error|błąd wewnętrzny/i.test(msg);
    if (isServerError) {
      log("⚠️ Excel Online zwrócił błąd (500). Spróbuj mniejszego zaznaczenia (np. do 500–1000 komórek) lub podziel arkusz na mniejsze fragmenty.");
    } else {
      log("Błąd: " + (msg || "nieznany"));
    }
    console.error(err);
  }
}

// ======================== WIĄZANIE PRZYCISKU ========================

(function wireRunButton() {
  var detailedEl = document.getElementById("detailedLog");
  var mainEl = document.getElementById("log");
  if (detailedEl) _appendToLog(detailedEl, "[załadowano] Skrypt runTranslate-bulk-read-write.js – konsola logów aktywna.");
  if (mainEl) _appendToLog(mainEl, "[załadowano] Konsola logów (wersja bulk) gotowa.");

  function bindRun() {
    var runBtn = document.getElementById("runBtn");
    if (runBtn) {
      runBtn.onclick = runTranslate;
    }
  }

  bindRun();

  if (typeof Office !== "undefined" && Office.onReady) {
    Office.onReady(function () {
      setTimeout(bindRun, 300);
      setTimeout(bindRun, 1500);
    });
  }
})();
