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

  log(`Czytam zaznaczenie... (źródło: ${sourceLang})`);

  try {
  await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    const sel = ctx.workbook.getSelectedRange();
    sel.load(["rowCount", "columnCount", "rowIndex", "columnIndex"]);
    await ctx.sync();

    const rowCount = sel.rowCount;
    const columnCount = sel.columnCount;
    const rowIndex = sel.rowIndex;
    const columnIndex = sel.columnIndex;

    const totalCells = rowCount * columnCount;
    if (totalCells > 3000) {
      log(`⚠️ Duże zaznaczenie (${totalCells} komórek). Przy błędach spróbuj mniejszego zakresu.`);
    }

    // Nagłówek tylko dla zaznaczonych kolumn (bez getUsedRange – przy dużym arkuszu unikamy 500)
    const headerRange = sheet.getRangeByIndexes(HEADER_ROW - 1, columnIndex, 1, columnCount);
    headerRange.load("values");
    await ctx.sync();

    const header = (headerRange.values && headerRange.values[0] ? headerRange.values[0] : []).map(h => normalizeHeader(h));
    const srcColInSel = header.indexOf(sourceLang);
    if (srcColInSel < 0) {
      log(`BŁĄD: brak kolumny ${sourceLang} w wierszu nagłówków (1).`);
      return;
    }
    const srcCol = columnIndex + srcColInSel;

    const currentGlossary = sourceLang === "PL" ? plGlossaryCache : glossaryCache;

    // ----- BULK READ w partiach (max 250 komórek) + przy 500 ponowna próba z mniejszą partią -----
    const selValues = [];
    const sourceColValues = [];
    const maxRowsPerChunk = Math.max(1, Math.min(MAX_ROWS_PER_SYNC, Math.floor(MAX_CELLS_PER_SYNC / columnCount)));
    const totalChunks = Math.ceil(rowCount / maxRowsPerChunk);
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
          if (totalChunks > 1) log(`  Czytam partię ${chunkNum + 1}...`);
          const chunkSel = sheet.getRangeByIndexes(rowIndex + rowOffset, columnIndex, chunkRows, columnCount);
          const chunkSrc = sheet.getRangeByIndexes(rowIndex + rowOffset, srcCol, chunkRows, 1);
          chunkSel.load("values");
          chunkSrc.load("values");
          await ctx.sync();
          const cv = chunkSel.values || [];
          const sv = chunkSrc.values || [];
          for (let i = 0; i < cv.length; i++) selValues.push(cv[i] ? cv[i].slice() : []);
          for (let i = 0; i < sv.length; i++) sourceColValues.push(sv[i] ? sv[i].slice() : []);
          rowOffset += chunkRows;
          chunkNum++;
          done = true;
        } catch (e) {
          if (chunkRows > 1 && is500(e)) {
            chunkRows = Math.max(1, Math.floor(chunkRows / 2));
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

    if (total === 0) {
      log("Brak danych do tłumaczenia.");
      return;
    }

    // Tłumaczenie w małych paczkach (RECORDS_PER_BATCH) i zapis do Excela co paczkę – na bieżąco, bez wielkiego zapisu na końcu
    let done = 0;
    const batchSize = RECORDS_PER_BATCH;

    for (const [lang, arr] of groups.entries()) {
      if (!lang) continue;
      log(`${sourceLang} → ${lang}: ${arr.length} rekordów (zapis co ${batchSize})`);

      for (let i = 0; i < arr.length; i += batchSize) {
        const batch = arr.slice(i, i + batchSize);
        const lines = batch.map(x => x.src);

        const { tokenizedLines, tokenMap } =
          applyGlossaryTokens(lines, lang, currentGlossary);

        const is429 = (e) => (e && (e.message || "" + e) && /429|rate limit|Too Many Requests/i.test(e.message || "" + e));
        let translatedLines;
        try {
          translatedLines = await callOpenAI(lang, tokenizedLines, sourceLang);
        } catch (apiErr) {
          if (is429(apiErr)) {
            log(`  Limit API (429) – czekam ${API_429_RETRY_DELAY_MS / 1000} s, ponawiam...`);
            await new Promise(r => setTimeout(r, API_429_RETRY_DELAY_MS));
            translatedLines = await callOpenAI(lang, tokenizedLines, sourceLang);
          } else throw apiErr;
        }

        if (API_DELAY_MS > 0) {
          await new Promise(r => setTimeout(r, API_DELAY_MS));
        }

        // Zapis do Excela po paczce; przy 500 – ponowna próba po 1 komórce
        const writeBatch = () => {
          for (let j = 0; j < batch.length; j++) {
            const restored = restoreGlossaryTokens(translatedLines[j] || "", tokenMap);
            const it = batch[j];
            const cellRange = sheet.getRangeByIndexes(it.absRow, it.absCol, 1, 1);
            cellRange.values = [[restored]];
          }
        };
        try {
          writeBatch();
          await ctx.sync();
        } catch (writeErr) {
          if (is500(writeErr) && batch.length > 1) {
            log(`  Błąd 500 przy zapisie – zapisuję po 1 komórce...`);
            for (let j = 0; j < batch.length; j++) {
              const restored = restoreGlossaryTokens(translatedLines[j] || "", tokenMap);
              const it = batch[j];
              sheet.getRangeByIndexes(it.absRow, it.absCol, 1, 1).values = [[restored]];
              await ctx.sync();
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
  });
  } catch (err) {
    setProgress(null);
    const msg = (err && (err.message || err.code || err.toString())) || "";
    const isServerError = /500|Internal|RichApi\.Error|błąd wewnętrzny/i.test(msg);
    if (isServerError) {
      log("⚠️ Excel Online zwrócił błąd (500). Spróbuj mniejszego zaznaczenia (np. do 500–1000 komórek) lub podziel arkusz na mniejsze fragmenty. Jeśli błąd pojawia się od razu – zaznacz mniejszy zakres.");
    } else {
      log("Błąd: " + (msg || "nieznany"));
    }
    console.error(err);
  }
}
