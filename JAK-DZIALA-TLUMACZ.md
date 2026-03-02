# Jak działa dodatek GASTROPARTS Translator

Prosty przewodnik krok po kroku. Co robi program, kiedy i w jakiej kolejności.

---

## 1. Po co jest ten dodatek?

Dodatek tłumaczy teksty z arkusza Excel z jednego języka na inne.  
Kolumny = języki (np. EN, PL, DE). Wiersz 1 to nagłówki z nazwami języków.  
Program bierze teksty z kolumny „źródłowej” (np. EN) i wpisuje tłumaczenia do kolumn innych języków (PL, DE itd.).

---

## 2. Zanim zaczniesz

- **Klucz API OpenAI** – w sekcji „Konfiguracja API” wklej klucz (zaczyna się od `sk-`) i kliknij „Zapisz klucz”. Bez tego tłumaczenie nie ruszy.
- **Kierunek tłumaczenia** – wybierz:
  - **EN → inne języki** – źródło to kolumna EN, tłumaczy do PL, DE, FR itd.
  - **PL → EN** – źródło to kolumna PL, tłumaczy do EN.
- **Zaznaczenie w Excelu** – zaznacz w arkuszu prostokąt: **wiersz 1 (nagłówki) + wszystkie wiersze z danymi**. Kolumny to języki. Program czyta nagłówki z wiersza 1 i wie, która kolumna to który język.

---

## 3. Co się dzieje po kliknięciu „Tłumacz zaznaczenie”

Kolejność jest taka:

### Krok 1: Sprawdzenie klucza i glosariusza

- Sprawdza, czy zapisano klucz API. Jeśli nie – pokazuje komunikat i kończy.
- Jeśli glosariusz nie był jeszcze wczytany – ładuje go z arkuszy (EN-PL, PL-EN, EN-DE itd.) w tym skoroszycie.

### Krok 2: Odczyt zaznaczenia z Excela

- Pobiera **wymiary zaznaczenia**: ile wierszy, ile kolumn, od którego wiersza/kolumny zaczyna się zakres.
- Czyta **wiersz 1** (tylko w zaznaczonych kolumnach) – to są nazwy języków (EN, PL, DE…).
- Szuka w nagłówkach **kolumny źródłowej** (np. EN lub PL – zależnie od wybranego kierunku).  
  Jeśli jej nie ma, pisze błąd i kończy.

### Krok 3: Odczyt danych komórek (partiami)

- Żeby uniknąć błędów przy dużym zakresie, program **nie czyta wszystkiego naraz**. Dzieli odczyt na mniejsze partie (np. po 250 komórek).
- Dla każdej partii: ładuje wartości zaznaczenia + wartości kolumny źródłowej, potem synchronizuje z Excelem (`sync`).
- Przy błędzie 500 (serwer) program **zmniejsza partię** i próbuje jeszcze raz.

### Krok 4: Zbudowanie listy „co tłumaczyć”

- Dla każdej komórki w zaznaczeniu program wie: wiersz, kolumna, język docelowy (z nagłówka), tekst źródłowy (z kolumny np. EN), aktualna zawartość komórki docelowej.
- **Filtrowanie:**
  - Pomija komórki, w których **nie ma tekstu źródłowego** (pusta kolumna źródłowa).
  - Jeśli masz włączone „Pomijaj, jeśli komórka docelowa jest już wypełniona” – pomija komórki, które już mają wpisany tekst.
  - Pomija kolumnę tego samego języka co źródło (np. nie tłumaczy EN → EN).
- Na tej podstawie buduje **grupy po językach docelowych**: np. „do PL – 50 rekordów”, „do DE – 50 rekordów”.

### Krok 5: Tłumaczenie (dla każdego języka i każdej paczki)

Dla każdego języka docelowego (PL, DE, FR…) program:

1. **Bierze paczkę tekstów** (np. 10 sztuk) z listy do tego języka.
2. **Glosariusz** – zamienia wybrane słowa na „tokeny” (np. `[[G1]]`, `[[G2]]`), żeby OpenAI ich nie zmieniała. Po odpowiedzi tokeny są zamieniane z powrotem na właściwe tłumaczenia z glosariusza.
3. **Wysyła do OpenAI** – jeden request z tą paczką linii (już z tokenami glosariusza). Prośba: „Przetłumacz z EN na PL, jedna linia = jedno tłumaczenie, zachowaj tokeny [[G1]] itd.”.
4. **Odbiera odpowiedź** – lista linii (tłumaczeń). Program dopasowuje je po kolei do komórek, przywraca słowa z glosariusza zamiast tokenów.
5. **Walidacja** – jeśli któraś komórka wyszła pusta albo identyczna ze źródłem, program **ponawia** tylko dla tej komórki (pojedyncze wywołanie API), do kilku prób.
6. **Zapis do Excela** – wpisuje wynik do odpowiednich komórek w arkuszu i robi `sync`. Przy błędzie 500 przy zapisie zapisuje **po jednej komórce**.
7. **Opóźnienie** – po każdej paczce jest krótka pauza (np. 400 ms), żeby nie przekroczyć limitów API (błąd 429).
8. Powtarza to dla następnej paczki (następne 10 tekstów), aż skończy wszystkie dla tego języka. Potem to samo dla kolejnego języka.

### Krok 6: Koniec

- Ukrywa pasek postępu, w logu pojawia się „Gotowe.”.

---

## 4. Gdzie co jest w programie

| Element | Znaczenie |
|--------|-----------|
| **Sekcja Konfiguracja API** | Wpisujesz klucz OpenAI (sk-…) i zapisujesz. Klucz jest trzymany u Ciebie (Office lub przeglądarka). |
| **Kierunek: EN → inne / PL → EN** | Określa, która kolumna jest „źródłem” (EN albo PL). Reszta kolumn to języki docelowe. |
| **Przycisk „Tłumacz zaznaczenie”** | Start całego procesu opisanego wyżej. |
| **„Pomijaj, jeśli komórka docelowa jest już wypełniona”** | Jeśli zaznaczone – program nie nadpisuje komórek, które już mają tekst. |
| **Górny log (szare pole)** | Krótkie komunikaty: co teraz robi, ile zapisano, błędy. **Od teraz tutaj dopisywane są też jednolinijkowe wpisy z konsoli logów.** |
| **Konsola logów (białe pole)** | Szczegółowe logi: każdy krok, dane wysyłane do API, odpowiedzi, zapis do Excela. To samo w wersji skróconej trafia też do górnego logu. |
| **Glosariusz** | Arkusze w tym samym skoroszycie: EN-PL, EN-DE, PL-EN itd. Wiersz 1 = nagłówki (np. EN, PL), kolejne wiersze = pary: słowo źródłowe → tłumaczenie. Program ich używa, żeby np. „heating element” zawsze szło do API i wracało w ustalony sposób. |

---

## 5. Ważne ograniczenia i zachowania

- **Wiersz 1 = nagłówki.** Program **zawsze** traktuje wiersz 1 zaznaczenia jako nazwy języków. Bez tego nie wie, która kolumna to który język.
- **Duże zaznaczenia** – przy bardzo dużym zakresie odczyt i zapis są dzielone na mniejsze kawałki; przy błędzie 500 partia jest zmniejszana i powtarzana.
- **Limity OpenAI** – przy zbyt wielu requestach możesz dostać błąd 429. Program wtedy czeka kilka sekund i ponawia. Opóźnienie między paczkami (400 ms) ma ograniczać 429.
- **Excel Online** – ma limit ok. 5 MB na jedno żądanie, stąd podział na partie.

---

## 6. Skrót „krok po kroku” (dla szybkiego ogarnięcia)

1. Wpisujesz klucz API i zapisujesz.  
2. Zaznaczasz w Excelu zakres: wiersz 1 (języki) + dane.  
3. Wybierasz kierunek (EN→inne lub PL→EN).  
4. Klikasz „Tłumacz zaznaczenie”.  
5. Program: czyta zaznaczenie i nagłówki → czyta dane partiami → grupuje po językach → dla każdej paczki: glosariusz → wysyła do OpenAI → odbiera → przywraca glosariusz → zapisuje do Excela → po każdej paczce krótka pauza.  
6. Log (górny) i konsola (biała) pokazują, co się dzieje. Na końcu: „Gotowe.”.

Jeśli coś jest niejasne, najpierw zerknij w **Konsolę logów** i górny **log** – tam widać dokładnie, na którym etapie jest program i co wysyła/odbiera.
