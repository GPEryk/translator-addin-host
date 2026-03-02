/**
 * Zastępstwo funkcji runTranslate() – ta sama logika (zaznaczenie + nagłówki = języki), bez limitów.
 *
 * - Odczyt zaznaczenia w małych partiach (MAX_CELLS_PER_SYNC), żeby uniknąć 500. Przy błędzie 500 – retry z mniejszą partią.
 * - Tłumaczenie w paczkach po RECORDS_PER_BATCH (10): po każdej paczce od razu zapis do Excela (na bieżąco), bez jednego wielkiego zapisu na końcu.
 * - API_DELAY_MS: opóźnienie między requestami do OpenAI (mniej 429).
 */
const MAX_ROWS_PER_SYNC = 1000;   // Excel Online ~5MB limit na request; chunk po wierszach
const MAX_CELLS_PER_SYNC = 250;   // Maks. komórek na jeden ctx.sync() przy odczycie – małe partie, żeby uniknąć 500
const RECORDS_PER_BATCH = 10;     // Tłumaczymy paczkami po 10 i od razu zapisujemy do Excela (na bieżąco)
const API_DELAY_MS = 400;         // Opóźnienie między requestami (mniej 429)
const RETRY_DELAY_MS = 800;       // Pauza przed ponowną próbą po błędzie 500 (odczyt)
const API_429_RETRY_DELAY_MS = 3000; // Czekaj 3 s i ponów przy 429 (rate limit OpenAI)

// —— Logi szczegółowe (okienko dokładnych kroków) ——
function _detailedLogNow() {
  const d = new Date();
  return d.toTimeString().slice(0, 12);
}
function detailedLog(step, message, data) {
  const el = document.getElementById("detailedLog");
  if (!el) return;
  let line = `[${_detailedLogNow()}] [${step}] ${message}`;
  if (data !== undefined && data !== null) {
    const dataStr = typeof data === "object" ? JSON.stringify(data, null, 2) : String(data);
    line += "\n  " + dataStr.replace(/\n/g, "\n  ");
  }
  el.textContent += line + "\n";
  el.scrollTop = el.scrollHeight;
}
function clearDetailedLog() {
  const el = document.getElementById("detailedLog");
  if (el) el.textContent = "";
}

(function wireRunButton() {
  var runBtn = document.getElementById("runBtn");
  if (runBtn) runBtn.onclick = runTranslate;
})();

async function runTranslate() {
  const apiKey = getApiKey();
  if (!apiKey) {
    await customAlert("⚠️ Najpierw zapisz swój klucz OpenAI API w sekcji Konfiguracja!");
    return;
  }

  const skipFilled = document.getElementById("skipFilled").checked;
  const sourceLang = document.getElementById("sourceLangSelect").value || "EN";

  if (!glossaryCache) {
    await refreshGlossaryAll();
  }

  clearDetailedLog();
  detailedLog("START", "Rozpoczęto tłumaczenie", { sourceLang, skipFilled });

  log(`Czytam zaznaczenie... (źródło: ${sourceLang})`);

  try {
  await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    const sel = ctx.workbook.getSelectedRange();
    detailedLog("EXCEL", "Pobieranie zaznaczenia (load + sync)...", {});
    sel.load(["rowCount", "columnCount", "rowIndex", "columnIndex"]);
    await ctx.sync();
    detailedLog("EXCEL", "Sync zakończony – mam wymiary zaznaczenia", {});

    const rowCount = sel.rowCount;
    const columnCount = sel.columnCount;
    const rowIndex = sel.rowIndex;
    const columnIndex = sel.columnIndex;

    const totalCells = rowCount * columnCount;
    detailedLog("EXCEL_READ", "Odczytano zaznaczenie z arkusza", {
      rowCount,
      columnCount,
      totalCells,
      rowIndex,
      columnIndex
    });
    if (totalCells > 3000) {
      log(`⚠️ Duże zaznaczenie (${totalCells} komórek). Przy błędach spróbuj mniejszego zakresu.`);
    }

    // Nagłówek tylko dla zaznaczonych kolumn (bez getUsedRange – przy dużym arkuszu unikamy 500)
    detailedLog("EXCEL", "Ładowanie nagłówka (wiersz 1, zaznaczone kolumny)...", { columnIndex, columnCount });
    const headerRange = sheet.getRangeByIndexes(HEADER_ROW - 1, columnIndex, 1, columnCount);
    headerRange.load("values");
    await ctx.sync();

    const headerRaw = headerRange.values && headerRange.values[0] ? headerRange.values[0] : [];
    const header = headerRaw.map(h => normalizeHeader(h));
    detailedLog("HEADER", "Odczytano nagłówki kolumn (języki)", { headerRaw, header });
    const srcColInSel = header.indexOf(sourceLang);
    if (srcColInSel < 0) {
      log(`BŁĄD: brak kolumny ${sourceLang} w wierszu nagłówków (1).`);
      detailedLog("ERROR", "Brak kolumny źródłowej w nagłówkach", { sourceLang, header });
      return;
    }
    const srcCol = columnIndex + srcColInSel;

    const currentGlossary = sourceLang === "PL" ? plGlossaryCache : glossaryCache;

    // ----- BULK READ w partiach (max 250 komórek) + przy 500 ponowna próba z mniejszą partią -----
    const selValues = [];
    const sourceColValues = [];
    const maxRowsPerChunk = Math.max(1, Math.min(MAX_ROWS_PER_SYNC, Math.floor(MAX_CELLS_PER_SYNC / columnCount)));
    const totalChunks = Math.ceil(rowCount / maxRowsPerChunk);
    detailedLog("CHUNK_PLAN", "Plan odczytu partiami", {
      maxRowsPerChunk,
      totalChunks,
      maxCellsPerSync: MAX_CELLS_PER_SYNC
    });
    if (totalChunks > 1) {
      log(`Zaznaczenie: ${rowCount * columnCount} komórek → odczyt w partiach (po max ${MAX_CELLS_PER_SYNC} komórek).`);
    }
    const is500 = (e) => (e && (e.message || e.code || "" + e) && /500|Internal|RichApi|błąd wewnętrzny/i.test(e.message || e.code || "" + e));
    let rowOffset = 0, chunkNum = 0;
    while (rowOffset < rowCount) {
      let chunkRows = Math.min(maxRowsPerChunk, rowCount - rowOffset);
      let done = false;
      while (!done && chunkRows >= 1) {
        try {
          if (totalChunks > 1) {
            log(`  Czytam partię ${chunkNum + 1}...`);
            detailedLog("CHUNK_READ", `Partia odczytu ${chunkNum + 1}/${totalChunks}`, {
              chunkRows,
              rowOffset,
              cellsInChunk: chunkRows * columnCount
            });
          }
          const chunkSel = sheet.getRangeByIndexes(rowIndex + rowOffset, columnIndex, chunkRows, columnCount);
          const chunkSrc = sheet.getRangeByIndexes(rowIndex + rowOffset, srcCol, chunkRows, 1);
          chunkSel.load("values");
          chunkSrc.load("values");
          await ctx.sync();
          const cv = chunkSel.values || [];
          const sv = chunkSrc.values || [];
          for (let i = 0; i < cv.length; i++) selValues.push(cv[i] ? cv[i].slice() : []);
          for (let i = 0; i < sv.length; i++) sourceColValues.push(sv[i] ? sv[i].slice() : []);
          detailedLog("CHUNK_READ_OK", `Odebrano partię ${chunkNum + 1}: ${cv.length} wierszy`, {
            rowsRead: cv.length,
            sourceColSample: (sv.slice(0, 3) || []).map(r => r && r[0])
          });
          rowOffset += chunkRows;
          chunkNum++;
          done = true;
        } catch (e) {
          if (chunkRows > 1 && is500(e)) {
            chunkRows = Math.max(1, Math.floor(chunkRows / 2));
            detailedLog("CHUNK_500", "Błąd 500 przy odczycie – zmniejszam partię i ponawiam", {
              newChunkRows: chunkRows,
              waitMs: RETRY_DELAY_MS,
              error: (e && e.message) || String(e)
            });
            log(`  Błąd serwera (500) – czekam ${RETRY_DELAY_MS / 1000} s, ponawiam z mniejszą partią (${chunkRows} wierszy)...`);
            await new Promise(r => setTimeout(r, RETRY_DELAY_MS));
          } else throw e;
        }
      }
    }

    const items = [];
    for (let r = 0; r < rowCount; r++) {
      for (let c = 0; c < columnCount; c++) {
        const absRow = rowIndex + r;
        const absCol = columnIndex + c;
        const tgtLang = header[c];
        const src = (sourceColValues[r] && sourceColValues[r][0] != null ? sourceColValues[r][0] : "").toString().trim();
        const tgt = (selValues[r] && selValues[r][c] != null ? selValues[r][c] : "").toString().trim();

        items.push({ absRow, absCol, r, c, tgtLang, src, tgt });
      }
    }

    detailedLog("ITEMS", "Zbudowano listę par (wiersz, kolumna, język docelowy, źródło, cel)", {
      totalItems: items.length,
      sample: items.slice(0, 5).map(it => ({ absRow: it.absRow, absCol: it.absCol, tgtLang: it.tgtLang, src: (it.src || "").slice(0, 40), tgt: (it.tgt || "").slice(0, 40) }))
    });

    // group by language (te same filtry co wcześniej)
    const groups = new Map();
    for (const it of items) {
      if (!it.src) continue;
      if (skipFilled && it.tgt) continue;
      if (it.tgtLang === sourceLang) continue;

      if (!groups.has(it.tgtLang)) groups.set(it.tgtLang, []);
      groups.get(it.tgtLang).push({ ...it, src: it.src });
    }

    let total = 0;
    for (const arr of groups.values()) total += arr.length;

    const byLang = {};
    for (const [k, arr] of groups.entries()) byLang[k] = arr.length;
    detailedLog("GROUPS", "Przygotowano dane do tłumaczenia (grupy po językach)", { total, byLang });
    detailedLog("GROUPS_DETAIL", "Liczba rekordów per język docelowy", Object.fromEntries([...groups.entries()].map(([k, v]) => [k, v.length])));

    if (total === 0) {
      log("Brak danych do tłumaczenia.");
      detailedLog("DONE", "Brak danych do tłumaczenia – zakończono");
      return;
    }

    // Tłumaczenie w małych paczkach (RECORDS_PER_BATCH) i zapis do Excela co paczkę – na bieżąco, bez wielkiego zapisu na końcu
    let done = 0;
    const batchSize = RECORDS_PER_BATCH;

    for (const [lang, arr] of groups.entries()) {
      if (!lang) continue;
      log(`${sourceLang} → ${lang}: ${arr.length} rekordów (zapis co ${batchSize})`);
      detailedLog("LANG_START", `Start języka docelowego: ${lang}`, { recordCount: arr.length, batchSize });

      for (let i = 0; i < arr.length; i += batchSize) {
        const batch = arr.slice(i, i + batchSize);
        const lines = batch.map(x => x.src);
        const batchNum = Math.floor(i / batchSize) + 1;

        detailedLog("BATCH_SOURCE", `Paczka ${batchNum}: teksty z Excela (źródło)`, { lines });

        const { tokenizedLines, tokenMap } =
          applyGlossaryTokens(lines, lang, currentGlossary);

        detailedLog("BATCH_TOKENIZED", `Paczka ${batchNum}: po glosariuszu (wysyłane do API)`, {
          tokenizedLines,
          tokenMapKeys: Object.keys(tokenMap),
          tokenMap
        });

        detailedLog("API_SEND", `Wysyłam do OpenAI API (${sourceLang} → ${lang}), paczka ${batchNum}`, {
          url: "https://api.openai.com/v1/chat/completions",
          lineCount: tokenizedLines.length,
          linesSent: tokenizedLines
        });

        const is429 = (e) => (e && (e.message || "" + e) && /429|rate limit|Too Many Requests/i.test(e.message || "" + e));
        let translatedLines;
        try {
          translatedLines = await callOpenAI(lang, tokenizedLines, sourceLang);
        } catch (apiErr) {
          if (is429(apiErr)) {
            detailedLog("API_429", "Limit 429 – czekam i ponawiam", { waitMs: API_429_RETRY_DELAY_MS });
            log(`  Limit API (429) – czekam ${API_429_RETRY_DELAY_MS / 1000} s, ponawiam...`);
            await new Promise(r => setTimeout(r, API_429_RETRY_DELAY_MS));
            translatedLines = await callOpenAI(lang, tokenizedLines, sourceLang);
          } else {
            detailedLog("API_ERROR", "Błąd wywołania OpenAI", { message: (apiErr && apiErr.message) || String(apiErr) });
            throw apiErr;
          }
        }

        detailedLog("API_RAW", `Odebrano z OpenAI (RAW) – paczka ${batchNum}`, {
          rawType: typeof translatedLines,
          rawIsArray: Array.isArray(translatedLines),
          rawLength: Array.isArray(translatedLines) ? translatedLines.length : 0,
          rawContent: translatedLines
        });

        // Upewnij się, że mamy tablicę o długości batch (API czasem zwraca jeden string lub mniej linii)
        if (typeof translatedLines === "string") {
          translatedLines = translatedLines.split(/\r?\n/).map(s => s.trim()).filter((_, idx) => idx < batch.length);
        }
        if (!Array.isArray(translatedLines)) translatedLines = [];
        while (translatedLines.length < batch.length) translatedLines.push("");
        translatedLines = translatedLines.slice(0, batch.length);

        detailedLog("API_RESPONSE", `Odpowiedź z OpenAI (po normalizacji) – paczka ${batchNum}`, {
          lineCount: translatedLines.length,
          lines: translatedLines
        });

        const restoredList = batch.map((it, j) => restoreGlossaryTokens(translatedLines[j] || "", tokenMap));
        detailedLog("RESTORE_GLOSSARY", `Paczka ${batchNum}: po przywróceniu glosariusza`, { restoredList });

        if (API_DELAY_MS > 0) {
          detailedLog("DELAY", `Opóźnienie ${API_DELAY_MS} ms przed kolejną paczką`, {});
          await new Promise(r => setTimeout(r, API_DELAY_MS));
        }

        // Zbuduj wynik; brak tłumaczenia lub „to samo co źródło” = ponawiamy do skutku (bez luk)
        const isValid = (j) => {
          const res = (restoredList[j] || "").trim();
          const src = (batch[j].src || "").trim();
          if (!res) return false;
          if (lang !== sourceLang && res === src) return false;
          return true;
        };
        const MAX_RETRY_PER_CELL = 5;
        let needRetry = batch.map((_, j) => j).filter(j => !isValid(j));
        let round = 0;
        while (needRetry.length > 0 && round < MAX_RETRY_PER_CELL) {
          round++;
          detailedLog("RETRY_START", `Ponowne wywołanie API dla ${needRetry.length} komórek, runda ${round}`, { indices: needRetry });
          log(`  Uzupełniam brakujące / błędne (${needRetry.length} komórek), próba ${round}...`);
          for (const j of needRetry) {
            await new Promise(r => setTimeout(r, 250));
            detailedLog("RETRY_SEND", `Ponowne wysłanie do API (indeks ${j})`, { line: tokenizedLines[j] });
            try {
              const oneLine = await callOpenAI(lang, [tokenizedLines[j]], sourceLang);
              detailedLog("RETRY_RAW", `Odebrano (RAW) dla indeksu ${j}`, { raw: oneLine });
              const one = Array.isArray(oneLine) ? (oneLine[0] ?? "") : ("" + (oneLine || "")).trim();
              restoredList[j] = restoreGlossaryTokens(one || "", tokenMap);
              detailedLog("RETRY_RESTORED", `Po przywróceniu glosariusza [${j}]`, { value: restoredList[j] });
            } catch (e) {
              detailedLog("RETRY_ERROR", `Błąd API dla komórki ${j}`, { message: (e && e.message) || String(e) });
              if (round < MAX_RETRY_PER_CELL) log(`  Błąd API dla komórki – ponowię w następnej rundzie.`);
            }
          }
          needRetry = needRetry.filter(j => !isValid(j));
        }
        if (needRetry.length > 0) {
          log(`  Uwaga: ${needRetry.length} komórek nadal niepoprawnych po ${MAX_RETRY_PER_CELL} próbach – zapisuję ostatni wynik.`);
          detailedLog("RETRY", "Część komórek uzupełniona ponownym wywołaniem API", {
            stillInvalid: needRetry.length,
            maxRetries: MAX_RETRY_PER_CELL
          });
        }

        // Zapis do Excela po paczce; przy 500 – ponowna próba po 1 komórce
        const writeBatch = () => {
          for (let j = 0; j < batch.length; j++) {
            const restored = restoredList[j] ?? "";
            const it = batch[j];
            const cellRange = sheet.getRangeByIndexes(it.absRow, it.absCol, 1, 1);
            cellRange.values = [[restored]];
          }
        };
        const writePlan = batch.map((it, j) => ({
          row: it.absRow,
          col: it.absCol,
          value: (restoredList[j] ?? "").slice(0, 80)
        }));
        detailedLog("EXCEL_WRITE", `Zapis do arkusza – paczka ${batchNum}`, {
          cellsWritten: batch.length,
          done: done + batch.length,
          total,
          cells: writePlan
        });
        try {
          writeBatch();
          await ctx.sync();
          detailedLog("EXCEL_WRITE_OK", `Sync zapisu zakończony – paczka ${batchNum}`, {});
        } catch (writeErr) {
          if (is500(writeErr) && batch.length > 1) {
            detailedLog("EXCEL_WRITE_500", "Błąd 500 przy zapisie – zapisuję po 1 komórce", {
              error: (writeErr && writeErr.message) || String(writeErr)
            });
            log(`  Błąd 500 przy zapisie – zapisuję po 1 komórce...`);
            for (let j = 0; j < batch.length; j++) {
              const restored = restoredList[j] ?? "";
              const it = batch[j];
              sheet.getRangeByIndexes(it.absRow, it.absCol, 1, 1).values = [[restored]];
              await ctx.sync();
              detailedLog("EXCEL_WRITE_ONE", `Zapisano komórkę (${it.absRow}, ${it.absCol})`, { value: restored.slice(0, 60) });
            }
          } else throw writeErr;
        }

        done += batch.length;
        setProgress(Math.round((done / total) * 100));
        if (done % 50 === 0 || done === total) {
          log(`  Zapisano ${done}/${total}`);
        }
      }
    }

    setProgress(null);
    log("Gotowe.");
    detailedLog("DONE", "Tłumaczenie zakończone pomyślnie", { total });
  });
  } catch (err) {
    setProgress(null);
    const msg = (err && (err.message || err.code || err.toString())) || "";
    detailedLog("ERROR", "Wystąpił błąd", { message: msg });
    const isServerError = /500|Internal|RichApi\.Error|błąd wewnętrzny/i.test(msg);
    if (isServerError) {
      log("⚠️ Excel Online zwrócił błąd (500). Spróbuj mniejszego zaznaczenia (np. do 500–1000 komórek) lub podziel arkusz na mniejsze fragmenty. Jeśli błąd pojawia się od razu – zaznacz mniejszy zakres.");
    } else {
      log("Błąd: " + (msg || "nieznany"));
    }
    console.error(err);
  }
}
