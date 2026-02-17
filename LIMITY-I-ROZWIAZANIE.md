# Limity wtyczki tłumaczenia – wyjaśnienie i rozwiązanie

## Co się dzieje przy ~5000 komórkach

Przy zaznaczeniu **ok. 5 tysięcy komórek** wtyczka często **w ogóle nie wczytuje danych** i nie startuje tłumaczenia.  
Przy mniejszym zaznaczeniu (np. 255 rekordów) działa: widać „Czytam zaznaczenie…”, „PL → EN: 255 rekordów”, batche po 80 linii.

## Skąd biorą się liczby 255 i 80

- **255 rekordów** – to po prostu liczba komórek do przetłumaczenia w Twoim zaznaczeniu (po odfiltrowaniu pustych / już wypełnionych i kolumny źródłowej). Nie jest to sztywny limit wtyczki.
- **80 linii na batch** – to **limit w jednym wywołaniu API OpenAI**: wtyczka wysyła do tłumaczenia maksymalnie 80 tekstów na jedno żądanie (`BATCH_SIZE = 80`). Większe zaznaczenie jest dzielone na wiele batchy po 80 (np. 255 = 3 batche: 80 + 80 + 95). Ten limit jest rozsądny i nie blokuje dużych zaznaczeń.

## Gdzie jest prawdziwy problem: odczyt z Excela

Wtyczka **nie ma sztucznego limitu** typu „max 500 komórek”.  
Problem jest w **sposobie odczytu zaznaczenia** w Excel JavaScript API:

- Dla **każdej komórki** z zaznaczenia kod wywołuje osobno:
  - `sheet.getCell(wiersz, kolumna)` (dla źródła i dla celu),
  - `.load("values")` na każdej takiej komórce.
- Przy **5000 komórkach** to **ok. 10 000** obiektów Range i wywołań `load` w jednym `ctx.sync()`.

Dokumentacja Microsoftu („Read or write to a large range”) mówi wprost:

- Przy **dużych zakresach** (miliony komórek) zaleca się **podział na mniejsze bloki** (np. 5k–20k wierszy).
- Nie należy ładować wszystkiego naraz przez bardzo wiele małych zakresów; lepiej **jedna lub kilka operacji na dużych zakresach** (np. `range.load("values")` na całym zaznaczeniu).

Czyli: **limit wynika z Excel API i z tego, że obecna wersja wtyczki ładuje każdą komórkę osobno**, zamiast załadować całe zaznaczenie (i kolumnę źródłową) jednym lub dwoma wywołaniami. Przy ~5000 komórkach ten sposób prowadzi do timeoutu / braku odpowiedzi i tłumaczenie w ogóle się nie rozpoczyna.

## Rozwiązanie: odczyt i zapis „bulk”

Żeby wtyczka radziła sobie z **dużo większym zaznaczeniem** (np. 5k+ komórek), trzeba zmienić logikę w funkcji `runTranslate` tak, aby:

1. **Odczyt**
   - Zamiast pętli `getCell(…)` + `load("values")` dla każdej komórki:
   - **Jednorazowo** załadować:
     - `getSelectedRange()` i na nim `range.load("values")`,
     - oraz jeden zakres kolumny źródłowej dla wybranych wierszy (np. `getRangeByIndexes(sel.rowIndex, srcCol, sel.rowCount, 1)` i na nim `load("values")`).
   - Jeden `ctx.sync()` po tych dwóch `load` daje dwie tablice (zaznaczenie + kolumna źródłowa).
   - Listę „items” (rekordów do tłumaczenia) buduje się **w pamięci** z tych tablic, bez kolejnych wywołań Excela.

2. **Zapis**
   - Zamiast zapisywania każdej komórki osobno po każdym batchu:
   - Trzymać w pamięci **jedną siatkę wynikową** (np. kopia `values` zaznaczenia).
   - Po przetłumaczeniu każdego batchu uzupełniać odpowiednie pola w tej siatce.
   - Na końcu **jednym** `range.values = resultGrid` i jednym `ctx.sync()` zapisać całe zaznaczenie.

Dodatkowo w **Excel Online** obowiązuje **limit ~5 MB na rozmiar jednego żądania** (request payload). Przekroczenie daje błąd: *„Rozmiar ładunku żądania przekroczył limit”* (RichApi.Error). Dlatego w pliku `runTranslate-bulk-read-write.js` odczyt i zapis są **podzielone na bloki** po `MAX_ROWS_PER_SYNC` wierszy (np. 1000) – każdy blok to osobny `ctx.sync()`, więc żaden request nie przekracza limitu.

Dzięki temu:

- **Odczyt** nie rośnie liniowo z liczbą komórek (bulk w chunkach zamiast tysięcy getCell),
- **Zapis** to wiele małych zapisów po 1000 wierszy zamiast jednego gigantycznego,
- Ograniczeniem stają się **limit API OpenAI** (batche po 80) i **limit ~5 MB na request w Excel Online** (stąd chunkowanie), a nie „max 255 rekordów” w samej wtyczce.

**Rate limit OpenAI (429):** Przy bardzo dużej liczbie requestów (np. 5k tłumaczeń = dziesiątki batchy) OpenAI może zwrócić 429 (Too Many Requests). W pliku `runTranslate-bulk-read-write.js` jest stała `API_DELAY_MS` (np. 400 ms) – po każdym wywołaniu API wstawiane jest opóźnienie, żeby requesty nie zatrzymywały się na 429. Możesz zwiększyć (np. 600–800 ms), jeśli nadal pojawiają się błędy 429. W kodzie źródłowym w `callOpenAI` warto przy statusie 429 zwiększyć czas oczekiwania przed retry (np. 60 s zamiast 2 s).

## Gdzie to wstawić w kodzie

W repozytorium masz zbudowany (zminifikowany) plik `taskpane.js` i chunk `2927fb4829fc176b7062.js` – **nie ma tu oryginalnego kodu źródłowego ani `package.json`**, więc nie da się w tym miejscu po prostu „przebudować” wtyczki.

Masz dwie ścieżki:

1. **Masz gdzieś projekt źródłowy** (np. `office-addin-taskpane-js` z webpackiem):
   - W pliku źródłowym taskpane (np. `src/taskpane/taskpane.js`) **zastąp** obecną funkcję `runTranslate` wersją z **pliku `runTranslate-bulk-read-write.js`** (załóżmy, że taki plik z poprawioną funkcją jest w repo lub do niego dodany).
   - Zbuduj projekt (npm run build / webpack) i wgraj nowe `taskpane.js` (i ewentualnie chunk) do hosta wtyczki.

2. **Nie masz projektu źródłowego**:
   - W pliku `taskpane.js.map` w polu `sourcesContent` jest oryginalny (niezminifikowany) kod taskpane. Można odtworzyć projekt (np. z szablonu Office Add-in), wkleić ten kod, **zastąpić** w nim funkcję `runTranslate` wersją z pliku **`runTranslate-bulk-read-write.js`**, zbudować (webpack) i wgrać nowe pliki JS. Albo uzyskać od autora wtyczki zaktualizowaną wersję z odczytem/zapisem bulk.

## Podsumowanie

- **255** i **80** to: liczba rekordów w zaznaczeniu oraz rozmiar batcha do API – nie są to sztywne limity „max komórek” wtyczki.
- **Przy ~5000 komórkach** wtyczka „nie wczytuje” i nie zaczyna tłumaczenia, bo **obecna implementacja** ładuje każdą komórkę osobno i Excel API przy takim obciążeniu nie daje rady (timeout / brak odpowiedzi).
- **Rozwiązanie** to zmiana `runTranslate` na **bulk read** (zaznaczenie + kolumna źródłowa w 2 wywołaniach) i **bulk write** (jedna siatka wynikowa i jeden zapis zakresu). Po takiej zmianie wtyczka może obsłużyć znacznie większe zaznaczenia (np. 5k+ komórek), z zachowaniem batchy po 80 do OpenAI.
