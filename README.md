# Bot do Wysyłania E-maili z GUI

## Opis projektu
Ten projekt to bot do automatycznego wysyłania e-maili, zarządzany za pomocą graficznego interfejsu użytkownika (GUI). Pozwala na:

- Wprowadzanie i edycję treści wiadomości.
- Wskazywanie pliku Excel z listą odbiorców.
- Ustawianie godziny wysyłki.
- Uruchamianie harmonogramu wysyłki w tle.

Aplikacja korzysta z biblioteki **Tkinter** do budowy GUI oraz **smtplib** do wysyłania wiadomości e-mail.

---

## Funkcjonalności

1. **Treść wiadomości:**
   - Treść wiadomości można wpisać w polu tekstowym aplikacji.
   - Przyciskiem "Zaktualizuj treść" można zapisać zmiany.

2. **Plik Excel:**
   - Użytkownik może wybrać plik Excel z listą odbiorców e-mail za pomocą eksploratora plików.
   - Plik musi zawierać kolumnę **Email**, w której znajdują się adresy odbiorców.

3. **Godzina wysyłki:**
   - Możliwość ustawienia godziny wysyłki za pomocą rozwijanego menu.

4. **Harmonogram:**
   - Przyciskiem "Uruchom bota" aplikacja zaplanuje wysyłkę wiadomości o wybranej godzinie i uruchomi harmonogram w tle.

---

## Wymagania

### Techniczne:
- Python 3.7 lub nowszy
- Biblioteki Python:
  - **Tkinter** (wbudowany w Python)
  - **pandas** (do odczytu plików Excel):
    ```bash
    pip install pandas
    ```
  - **openpyxl** (do obsługi plików Excel):
    ```bash
    pip install openpyxl
    ```

### Dane logowania do e-maila:
W kodzie należy podać dane logowania do konta e-mail (np. Gmail). **Zadbaj o bezpieczeństwo i nie zapisuj hasła bezpośrednio w kodzie!**
Możesz użyć zmiennych środowiskowych lub pliku konfiguracyjnego.

---

## Instrukcja uruchomienia

### 1. Pobranie projektu
Pobierz pliki projektu i upewnij się, że wszystkie wymagane pliki znajdują się w jednym folderze.

### 2. Uruchomienie aplikacji
W terminalu wpisz:
```bash
python nazwapliku.py
```

### 3. Używanie aplikacji
1. **Treść wiadomości:**
   - Wprowadź treść wiadomości w polu tekstowym.
   - Kliknij "Zaktualizuj treść", aby zapisać zmiany.

2. **Wybór pliku Excel:**
   - Kliknij "Wybierz plik Excel" i wskaż plik z listą adresów e-mail.

3. **Godzina wysyłki:**
   - Wybierz godzinę z rozwijanego menu (format HH:MM).

4. **Uruchomienie bota:**
   - Kliknij "Uruchom bota", aby zaplanować wysyłkę wiadomości o wybranej godzinie.

---

## Struktura pliku Excel
Plik Excel powinien zawierać kolumnę `Email` z adresami odbiorców.
Przykład:

| Email               |
|---------------------|
| test1@example.com   |
| test2@example.com   |

---

## Uwagi

1. **Bezpieczeństwo:**
   - Użyj protokołu OAuth2, jeśli to możliwe, zwłaszcza dla Gmaila.
   - Nie zapisuj danych logowania w kodzie.

2. **Obsługa błędów:**
   - Aplikacja informuje o błędach przy braku pliku Excel lub nieprawidłowym formacie danych.

---

## Rozwój projektu
W przyszłości można dodać:
- Podgląd logów wysyłki w GUI.
- Obsługę wielu tematów wiadomości.
- Zaawansowaną obsługę błędów i raporty.

