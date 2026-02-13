/**
 * Zastępstwo funkcji runTranslate() – odczyt i zapis BULK (bez getCell w pętli).
 * Użycie: w pliku źródłowym taskpane (np. src/taskpane/taskpane.js)
 * zastąp całą funkcję runTranslate poniższą implementacją, potem zbuduj projekt.
 *
 * Zmiany:
 * - Odczyt: jedno sel.load("values") + jeden zakres kolumny źródłowej + jeden ctx.sync().
 * - Items budowane z tablic w pamięci (bez srcCell/tgtCell).
 * - Zapis: jedna siatka resultGrid, na końcu jeden zapis zakresu.
 */

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

    // ----- BULK READ: tylko 2 zakresy zamiast tysięcy getCell -----
    sel.load("values");
    const sourceColRange = sheet.getRangeByIndexes(sel.rowIndex, srcCol, sel.rowCount, 1);
    sourceColRange.load("values");
    await ctx.sync();

    const selValues = sel.values || [];
    const sourceColValues = sourceColRange.values || [];
    const rowCount = sel.rowCount;
    const columnCount = sel.columnCount;
    const rowIndex = sel.rowIndex;
    const columnIndex = sel.columnIndex;

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

        for (let j = 0; j < batch.length; j++) {
          const restored = restoreGlossaryTokens(translatedLines[j] || "", tokenMap);
          const it = batch[j];
          if (resultGrid[it.r]) resultGrid[it.r][it.c] = restored;
        }

        done += batch.length;
        setProgress(Math.round((done / total) * 100));
      }
    }

    // ----- BULK WRITE: jeden zapis całego zaznaczenia -----
    const outRange = sheet.getRangeByIndexes(rowIndex, columnIndex, rowCount, columnCount);
    outRange.values = resultGrid;
    await ctx.sync();

    setProgress(null);
    log("Gotowe.");
  });
}
