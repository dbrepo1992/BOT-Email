import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import schedule
import time
from datetime import datetime
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

# Dane logowania do e-maila
EMAIL = "twoj_email@gmail.com"
PASSWORD = "twoje_haslo"

# Globalne zmienne
plik_excel = ""
plik_zalacznik = ""
godzina_wysylki = "09:00"
tresc_wiadomosci = ""
domyslna_stopka = "\n\n---\nPozdrawiam,\nZespół Twojej firmy"

# Funkcja do wysyłania e-maili z załącznikami
def wyslij_email(odbiorca, temat, tresc, zalacznik=None):
    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL
        msg['To'] = odbiorca
        msg['Subject'] = temat

        # Dodanie treści wiadomości i stopki
        tresc_koncowa = tresc + domyslna_stopka
        msg.attach(MIMEText(tresc_koncowa, 'plain'))

        # Dodanie załącznika, jeśli istnieje
        if zalacznik:
            try:
                with open(zalacznik, 'rb') as attachment_file:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment_file.read())
                encoders.encode_base64(part)
                part.add_header(
                    'Content-Disposition',
                    f'attachment; filename={zalacznik.split("/")[-1]}'
                )
                msg.attach(part)
            except Exception as e:
                print(f"Błąd przy dodawaniu załącznika: {e}")

        # Wysyłanie wiadomości
        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(EMAIL, PASSWORD)
            server.send_message(msg)

        print(f"E-mail wysłany do: {odbiorca}")
    except Exception as e:
        print(f"Błąd podczas wysyłania e-maila do {odbiorca}: {e}")

# Funkcja do odczytu danych z pliku Excel i wysyłania e-maili
def wyslij_maile():
    if not plik_excel:
        messagebox.showerror("Błąd", "Nie wybrano pliku Excel!")
        return

    try:
        df = pd.read_excel(plik_excel)
        if 'Email' not in df.columns:
            messagebox.showerror("Błąd", "Plik Excel musi zawierać kolumnę 'Email'.")
            return

        temat = "Powiadomienie od Twojej firmy"  # Stały temat
        for _, row in df.iterrows():
            odbiorca = row['Email']
            wyslij_email(odbiorca, temat, tresc_wiadomosci, zalacznik=plik_zalacznik)

        messagebox.showinfo("Sukces", "Wszystkie e-maile zostały wysłane!")
    except Exception as e:
        messagebox.showerror("Błąd", f"Nie udało się wysłać wiadomości: {e}")

# Funkcja harmonogramowania
def uruchom_harmonogram():
    schedule.every().day.at(godzina_wysylki).do(wyslij_maile)
    while True:
        schedule.run_pending()
        time.sleep(1)

# Funkcja uruchamiająca bota w osobnym wątku
def start_bot():
    threading.Thread(target=uruchom_harmonogram, daemon=True).start()
    messagebox.showinfo("Bot uruchomiony", f"Bot wysyłający e-maile uruchomiony na godzinę {godzina_wysylki}.")

# Funkcja wyboru pliku Excel
def wybierz_plik():
    global plik_excel
    plik_excel = filedialog.askopenfilename(filetypes=[("Pliki Excel", "*.xlsx")])
    if plik_excel:
        plik_label.config(text=f"Wybrany plik: {plik_excel}")

# Funkcja wyboru załącznika
def wybierz_zalacznik():
    global plik_zalacznik
    plik_zalacznik = filedialog.askopenfilename()
    if plik_zalacznik:
        zalacznik_label.config(text=f"Wybrany załącznik: {plik_zalacznik}")

# Funkcja ustawiania godziny wysyłki
def ustaw_godzine(event):
    global godzina_wysylki
    godzina_wysylki = godzina_dropdown.get()

# Funkcja ustawiania treści wiadomości
def ustaw_tresc():
    global tresc_wiadomosci
    tresc_wiadomosci = tresc_text.get("1.0", tk.END).strip()
    messagebox.showinfo("Treść ustawiona", "Treść wiadomości została zaktualizowana.")

# Interfejs graficzny
root = tk.Tk()
root.title("Bot do wysyłania e-maili")

# Treść wiadomości
tk.Label(root, text="Treść wiadomości:").pack()
tresc_text = tk.Text(root, height=10, width=50)
tresc_text.insert("1.0", "Dzień dobry,\n\nTo jest przykładowa treść wiadomości.")
tresc_text.pack()

tk.Button(root, text="Zaktualizuj treść", command=ustaw_tresc).pack()

# Wybór pliku Excel
tk.Button(root, text="Wybierz plik Excel", command=wybierz_plik).pack()
plik_label = tk.Label(root, text="Nie wybrano pliku.")
plik_label.pack()

# Wybór załącznika
tk.Button(root, text="Wybierz załącznik", command=wybierz_zalacznik).pack()
zalacznik_label = tk.Label(root, text="Nie wybrano załącznika.")
zalacznik_label.pack()

# Wybór godziny wysyłki
tk.Label(root, text="Wybierz godzinę wysyłki:").pack()
godziny = [f"{i:02}:00" for i in range(24)]
godzina_dropdown = tk.StringVar(root)
godzina_dropdown.set(godzina_wysylki)
godzina_menu = tk.OptionMenu(root, godzina_dropdown, *godziny, command=ustaw_godzine)
godzina_menu.pack()

# Start bota
tk.Button(root, text="Uruchom bota", command=start_bot).pack()

# Uruchomienie aplikacji
root.mainloop()
