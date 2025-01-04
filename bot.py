import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pandas as pd
import schedule
import time
from datetime import datetime

# Dane logowania do e-maila
EMAIL = "twoj_email@gmail.com"
PASSWORD = "twoje_haslo"

# Ścieżki do plików
PLIK_EXCEL = "lista_maili.xlsx"  # Lista odbiorców
PLIK_LOGOW = "logi_wysylki.txt"
PLIK_RAPORTU = "raport_wysylki.xlsx"
PLIK_TRESCI = "tresc_wiadomosci.txt"  # Treść wiadomości

# Domyślna treść wiadomości
DOMYSLNA_TRESC = """
Dzień dobry,

To jest domyślna treść wiadomości. Możesz zmienić ją, edytując plik tresc_wiadomosci.txt.

Pozdrawiamy,
Zespół Twojej firmy
"""

# Funkcja do logowania
def zapisz_log(wiadomosc):
    with open(PLIK_LOGOW, "a") as log_file:
        log_file.write(f"{datetime.now()} - {wiadomosc}\n")

# Funkcja do odczytu treści z pliku
def odczytaj_tresc_z_pliku():
    try:
        with open(PLIK_TRESCI, "r") as file:
            return file.read().strip()
    except FileNotFoundError:
        # Jeśli plik nie istnieje, utwórz go z domyślną treścią
        with open(PLIK_TRESCI, "w") as file:
            file.write(DOMYSLNA_TRESC)
        zapisz_log("Plik tresci_wiadomosci.txt został utworzony z domyślną treścią.")
        return DOMYSLNA_TRESC

# Funkcja do wysyłania e-maili
def wyslij_email(odbiorca, temat, tresc):
    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL
        msg['To'] = odbiorca
        msg['Subject'] = temat
        msg.attach(MIMEText(tresc, 'plain'))

        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(EMAIL, PASSWORD)
            server.send_message(msg)

        zapisz_log(f"E-mail wysłany do: {odbiorca}")
        return "Sukces"
    except Exception as e:
        zapisz_log(f"Błąd podczas wysyłania e-maila do {odbiorca}: {e}")
        return "Błąd"

# Funkcja odczytu danych z pliku Excel i wysyłania e-maili
def wyslij_maile_z_excela():
    try:
        # Wczytaj dane z pliku Excel
        df = pd.read_excel(PLIK_EXCEL)

        # Sprawdź, czy wymagana kolumna istnieje
        if 'Email' not in df.columns:
            zapisz_log("Plik Excel musi zawierać kolumnę: 'Email'")
            return

        # Dodaj kolumny na status i czas wysyłki
        df['Status'] = ""
        df['Czas wysyłki'] = ""

        # Pobierz treść wiadomości
        tresc = odczytaj_tresc_z_pliku()

        # Iteracja przez wiersze i wysyłanie e-maili
        temat = "Powiadomienie od Twojej firmy"  # Stały temat
        for index, row in df.iterrows():
            odbiorca = row['Email']
            status = wyslij_email(odbiorca, temat, tresc)
            df.at[index, 'Status'] = status
            df.at[index, 'Czas wysyłki'] = datetime.now()

        # Zapisz raport wysyłki
        df.to_excel(PLIK_RAPORTU, index=False)
        zapisz_log("Raport wysyłki zapisano do pliku.")

    except FileNotFoundError:
        zapisz_log("Nie znaleziono pliku Excel.")
    except Exception as e:
        zapisz_log(f"Błąd: {e}")

# Funkcja do harmonogramowania
def uruchom_harmonogram():
    godzina_wysylki = "09:00"  # Godzina wysyłki w formacie HH:MM
    schedule.every().day.at(godzina_wysylki).do(wyslij_maile_z_excela)
    zapisz_log(f"Harmonogram ustawiony na {godzina_wysylki} codziennie.")

    while True:
        schedule.run_pending()
        time.sleep(1)

# Główna część programu
if __name__ == "__main__":
    uruchom_harmonogram()
