/**
 * Zastępstwo funkcji runTranslate() – ta sama logika (zaznaczenie + nagłówki = języki), bez limitów.
 *
 * - Przy zaznaczeniu >500 komórek: odczyt i zapis w partiach (MAX_CELLS_PER_SYNC), krok po kroku, z logiem "Czytam partię X z Y".
 * - API_DELAY_MS: opóźnienie po każdym request do OpenAI, żeby przy dużej liczbie zapytań nie wpaść na 429 (rate limit).
 */
const MAX_ROWS_PER_SYNC = 1000;   // Excel Online ~5MB limit na request; chunk po wierszach
const MAX_CELLS_PER_SYNC = 500;   // Maks. komórek na jeden ctx.sync() – przy >500 zaznaczenie dzielone na partie, krok po kroku
const API_DELAY_MS = 400;        // Opóźnienie między requestami (mniej 429); zwiększ przy dużej liczbie języków/komórek

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

    // ----- BULK READ w partiach po max 500 komórek – przy większym zaznaczeniu: krok po kroku -----
    const selValues = [];
    const sourceColValues = [];
    const maxRowsPerChunk = Math.max(1, Math.min(MAX_ROWS_PER_SYNC, Math.floor(MAX_CELLS_PER_SYNC / columnCount)));
    const totalChunks = Math.ceil(rowCount / maxRowsPerChunk);
    if (totalChunks > 1) {
      log(`Zaznaczenie: ${rowCount * columnCount} komórek → odczyt w ${totalChunks} partiach (po max ${MAX_CELLS_PER_SYNC} komórek).`);
    }
    for (let rowOffset = 0, chunkNum = 0; rowOffset < rowCount; rowOffset += maxRowsPerChunk, chunkNum++) {
      const chunkRows = Math.min(maxRowsPerChunk, rowCount - rowOffset);
      if (totalChunks > 1) log(`  Czytam partię ${chunkNum + 1} z ${totalChunks}...`);
      const chunkSel = sheet.getRangeByIndexes(rowIndex + rowOffset, columnIndex, chunkRows, columnCount);
      const chunkSrc = sheet.getRangeByIndexes(rowIndex + rowOffset, srcCol, chunkRows, 1);
      chunkSel.load("values");
      chunkSrc.load("values");
      await ctx.sync();
      const cv = chunkSel.values || [];
      const sv = chunkSrc.values || [];
      for (let i = 0; i < cv.length; i++) selValues.push(cv[i] ? cv[i].slice() : []);
      for (let i = 0; i < sv.length; i++) sourceColValues.push(sv[i] ? sv[i].slice() : []);
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

    // Siatka wynikowa – kopia aktualnych values zaznaczenia (żeby nie nadpisać komórek, których nie tłumaczymy)
    const resultGrid = selValues.map(row => row ? row.slice() : []);

    let done = 0;

    for (const [lang, arr] of groups.entries()) {
      if (!lang) continue;
      log(`${sourceLang} → ${lang}: ${arr.length} rekordów`);

      for (let i = 0; i < arr.length; i += BATCH_SIZE) {
        const batch = arr.slice(i, i + BATCH_SIZE);
        const lines = batch.map(x => x.src);

        const { tokenizedLines, tokenMap } =
          applyGlossaryTokens(lines, lang, currentGlossary);

        log(`  batch ${Math.floor(i / BATCH_SIZE) + 1}: wysyłam ${lines.length} linii`);
        const translatedLines = await callOpenAI(lang, tokenizedLines, sourceLang);

        if (API_DELAY_MS > 0) {
          await new Promise(r => setTimeout(r, API_DELAY_MS));
        }

        for (let j = 0; j < batch.length; j++) {
          const restored = restoreGlossaryTokens(translatedLines[j] || "", tokenMap);
          const it = batch[j];
          if (resultGrid[it.r]) resultGrid[it.r][it.c] = restored;
        }

        done += batch.length;
        setProgress(Math.round((done / total) * 100));
      }
    }

    // ----- BULK WRITE w partiach (ten sam rozmiar co przy odczycie) -----
    for (let rowOffset = 0, chunkNum = 0; rowOffset < rowCount; rowOffset += maxRowsPerChunk, chunkNum++) {
      const chunkRows = Math.min(maxRowsPerChunk, rowCount - rowOffset);
      if (totalChunks > 1) log(`  Zapisuję partię ${chunkNum + 1} z ${totalChunks}...`);
      const chunkData = resultGrid.slice(rowOffset, rowOffset + chunkRows);
      const outRange = sheet.getRangeByIndexes(rowIndex + rowOffset, columnIndex, chunkRows, columnCount);
      outRange.values = chunkData;
      await ctx.sync();
    }

    setProgress(null);
    log("Gotowe.");
  });
  } catch (err) {
    setProgress(null);
    const msg = (err && (err.message || err.code || err.toString())) || "";
    const isServerError = /500|Internal|RichApi\.Error|błąd wewnętrzny/i.test(msg);
    if (isServerError) {
      log("⚠️ Excel Online zwrócił błąd przy dużym zaznaczeniu. Spróbuj mniejszego zakresu (np. do ok. 2000 komórek) lub podziel arkusz na mniejsze fragmenty.");
    } else {
      log("Błąd: " + (msg || "nieznany"));
    }
    console.error(err);
  }
}
