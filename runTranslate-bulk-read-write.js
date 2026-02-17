/**
 * Zastępstwo funkcji runTranslate() – ta sama logika (zaznaczenie + nagłówki = języki), bez limitów.
 *
 * - Odczyt/zapis w chunkach po MAX_ROWS_PER_SYNC (limit ~5MB Excel Online – żeby nie było "Rozmiar ładunku przekroczył limit").
 * - API_DELAY_MS: opóźnienie po każdym request do OpenAI, żeby przy dużej liczbie zapytań nie wpaść na 429 (rate limit).
 */
const MAX_ROWS_PER_SYNC = 1000; // Excel Online ~5MB limit na request; chunk po wierszach
const API_DELAY_MS = 400;      // Opóźnienie między requestami (mniej 429); zwiększ przy dużej liczbie języków/komórek

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

  await Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getActiveWorksheet();
    const sel = ctx.workbook.getSelectedRange();
    sel.load(["rowCount", "columnCount", "rowIndex", "columnIndex"]);
    await ctx.sync();

    const used = sheet.getUsedRange();
    used.load("columnCount");
    await ctx.sync();

    const headerRange = sheet.getRangeByIndexes(HEADER_ROW - 1, 0, 1, used.columnCount);
    headerRange.load("values");
    await ctx.sync();

    const header = headerRange.values[0].map(h => normalizeHeader(h));
    const srcCol = header.indexOf(sourceLang);
    if (srcCol < 0) {
      log(`BŁĄD: brak kolumny ${sourceLang} w wierszu nagłówków (1).`);
      return;
    }

    const currentGlossary = sourceLang === "PL" ? plGlossaryCache : glossaryCache;

    const rowCount = sel.rowCount;
    const columnCount = sel.columnCount;
    const rowIndex = sel.rowIndex;
    const columnIndex = sel.columnIndex;

    // ----- BULK READ w chunkach (limit ~5MB na request w Excel Online) -----
    const selValues = [];
    const sourceColValues = [];
    for (let rowOffset = 0; rowOffset < rowCount; rowOffset += MAX_ROWS_PER_SYNC) {
      const chunkRows = Math.min(MAX_ROWS_PER_SYNC, rowCount - rowOffset);
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
        const tgtLang = header[absCol];
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

    // ----- BULK WRITE w chunkach (limit ~5MB na request w Excel Online) -----
    for (let rowOffset = 0; rowOffset < rowCount; rowOffset += MAX_ROWS_PER_SYNC) {
      const chunkRows = Math.min(MAX_ROWS_PER_SYNC, rowCount - rowOffset);
      const chunkData = resultGrid.slice(rowOffset, rowOffset + chunkRows);
      const outRange = sheet.getRangeByIndexes(rowIndex + rowOffset, columnIndex, chunkRows, columnCount);
      outRange.values = chunkData;
      await ctx.sync();
    }

    setProgress(null);
    log("Gotowe.");
  });
}
