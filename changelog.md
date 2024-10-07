# Changelog

## [0.37] - 2024-10-04

### Naprawiono
- obsługę wielowątkowego generowania plików pdf (pogram powinien być bardzie responsywny).
- naprawiono zliczanie plików w katalogach. Na przyciskach do generowania pdf wyświetla się liczba plików DOCX w nie wszystkich.

### Zmieniono
- Ze względu na kompilacje do pliku exe wszystkie Foldery "Data" i "Szablony" oraz plik zawody.json zostały przeniesione do folder _internal
- W przypadku dwóch uczniów o tym samym imieniu, nazwisku, klasie i zawodzie tworzone są dwa oddzielne skierowania, 
- pliki skierowań nie są nadpisywane tylko tworzony jest plik z nowym numerem,
- zamieniono listy na tuple - poprawa szybkości działania,
- wczytywanie ścieżek do plików i folderów 
- eksplorator otwiera domyślnie folder Data podczas wybierania pliku z danym,
- zmieniono obsługę błędu w przypadku nieprawidłowego pliku z danymi.
- Zliczanie plików w katalogach. Na przyciskach do generowania pdf wyświetla się liczba plików DOCX w nie wszystkich.

### Dodano
- plik z błędnymi danymi testowymi: BłądneDaneTestoweCHATGPT.xlsx - zawiera on błąd w nazwie kolumny - PESLE,
- w repozytoriom utworzono nowy folder o nazwie Dodatki (Nie jest on potrzebny do  działania aplikacji) po kompilacji kopiuję z niego plik zawody.json oraz folder Data z przykładowymi danymi.


## [0.36] - 2024-10-03
### Zmieniono
- Lista zawodów przeniesiona do oddzielnego pliku zawody.json,

## [0.35] - 2024-10-01
### Dodano
- Dodano instrukcje obsługi aplikacji.
- Dodano instrukcję generowania pliku z danymi z dziennika elektronicznego Vulcan.

### Zmieniono
- Aplikacja używa nazw kolumn z pliku z danymi zamiast ich numerów.

## [0.34] - 2024-09-01
### Dodano
- Udostępniono publicznie aplikację.