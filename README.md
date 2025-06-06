# eDress
eDress to interaktywny skoroszyt Excel do zarządzania procesem wydawania odzieży pracowniczej. Stworzony w ramach współpracy mojej uczelni z firmą zewnętrzną powstał na potrzeby zastąpienia przestarzałej i nieużywanej aplikacji.

# Struktura
Plik składa się z pięciu arkuszy:
## 1. Ekran główny
Ekran ten zawiera tekst powitalny oraz cztery przyciski:
- Wydaj odzież
- Historia pracownika
- Wygeneruj raport
- Dodaj odzież
Każdy z nich wyświetla formularz z odpowiednimi polami do wypełnienia.
### 1.1. Wydaj odzież
Formularz wydania odzieży zawiera cztery ListBoxy oraz jeden TextBox. Listboxy pracownika oraz osoby wydającej zawierają dwie kolumny z identyfikatorem oraz imieniem i nazwiskiem pracowników, zaś ListBox odzieży zawiera pozycje z arkusza spisu odzieży. ListBox Rodzaj uzupełnia się automatycznie dostępnymi wariantami dla danej pozycji. TextBox zawiera datę wydania i automatycznie wypełnia się obecną datą i godziną.

### 1.2. Historia pracownika
Otwiera prosty formularz z wyborem, dla którego pracownika chcemy wygenerować raport, wyświetlać on będzie wszystkie sztuki odzieży, które zostały wydane temu pracownikowi. Wynik będzie się znajdował w świeżo wygenerowanym arkuszu.

### 1.3. Wygeneruj raport
Bardziej skomplikowana wersja funkcjonalności opisanej powyżej, pozwala na wygenerowanie raportu w podobny sposób, jednak z wybranymi przez użytkownika filtrami.

### 1.4. Dodaj odzież
Ostatnia część głównego ekranu, to formularz dodania odzieży do spisu. Aby poprawnie wprowadzić nową pozycję wystarczy wypełnić podane cztery pola.

## 2. Pracownicy
Arkusz ten zawiera tabelę z dwoma kolumnami: id oraz Imię i nazwisko.

## 3. Odzież
Ten arkusz zawiera tabelę zawierającą spis odzieży. Zawiera kolumny:
- Numer
- Nazwa
- Rozmiar
- Barcode
Barcode jest zarezerwowany dla kodu kreskowego EAN-13, który był używany w oryginalnym projekcie.

## 4. Historia wydań
Ten arkusz zawiera tabelę ze wszystkimi wydaniami. Znajdują się w nim następujące kolumny:
- Data
- ID pracownika
- Pracownik
- ID wydającego pracownika
- Wydający pracownik
- Odzież
- Rodzaj
- Numer odzieży

## 5. Licencja
Znajduje się tam kopia licencji tego projektu.
# Technologia
Skoroszyt ten korzysta z makr oraz poleceń napisanych w VBA.

# Podziękowania
Składam szczególne podziękowania Wicie Lis, która stworzyła schemat aplikacji oraz dopilnowała, by projekt się pomyślnie zakończył w terminie.