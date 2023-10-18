import tkinter as tk
from tkinter import WORD, filedialog, messagebox
from reportlab.pdfgen import canvas
import openpyxl
import os
import random

# Funktion zum Importieren von Daten aus einer Excel-Datei
def import_from_excel():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if file_path:
        try:
            workbook = openpyxl.load_workbook(file_path)
            sheet = workbook.active
            data = [str(sheet.cell(row=i, column=1).value).strip() for i in range(1, sheet.max_row + 1)]
            # Daten mit Komma trennen
            data = ','.join(data)
            # Daten im Textfeld anzeigen
            entry_namen.delete("1.0", tk.END)
            entry_namen.insert(tk.END, data)
        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler beim Importieren der Daten aus der Excel-Datei:\n{str(e)}")


# Funktion zum Speichern als PDF
def save_as_pdf():
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    file_path = os.path.join(desktop_path, "Sitzordnung.pdf")
    text_content = text_output.get("1.0", tk.END).strip()

    pdf = canvas.Canvas(file_path)
    pdf.setFont("Helvetica", 12)
    pdf.drawString(50, 750, "Sitzordnungen:")

    y_position = 730
    for line in text_content.split('\n'):
        pdf.drawString(50, y_position, line)
        y_position -= 15

    pdf.save()

    messagebox.showinfo("PDF gespeichert", f"Die Datei wurde als 'Sitzordnung.pdf' auf dem Desktop gespeichert:\n{file_path}")

# Funktion zum Generieren der Sitzordnungen
def generate_sitzordnungen():
    namen_liste = entry_namen.get("1.0", tk.END).strip()
    personen_pro_tisch = int(entry_personen_pro_tisch.get())

    namen = [name.strip() for name in namen_liste.split(',') if name.strip()]

    random.shuffle(namen)

    anzahl_tische = len(namen) // personen_pro_tisch

    sitzordnungen_text = ""
    for i in range(2):
        sitzordnungen_text += f"Sitzordnung {i+1}:\n"
        for j in range(anzahl_tische):
            tisch_nummer = j + 1
            personen_am_tisch = namen[j * personen_pro_tisch: (j + 1) * personen_pro_tisch]
            sitzordnungen_text += f"Tisch {tisch_nummer} → [ {' | '.join(personen_am_tisch)} ]\n"
        sitzordnungen_text += "\n"

    text_output.delete("1.0", tk.END)
    text_output.insert(tk.END, sitzordnungen_text)

# Funktion zum Schließen des Fensters
def close_window():
    root.destroy()

# GUI erstellen
root = tk.Tk()
root.title("Sitzordnungen Generator")
root.configure(background="#424242")

# Mindestgröße setzen
root.minsize(1600, 900)
root.attributes("-fullscreen", True)

# Eingabefeld für Namen
label_namen = tk.Label(root, font=("Arial", 16), text="Namen (getrennt durch Komma):")
label_namen.grid(row=0, column=0, padx=(15, 10), pady=(15, 10), sticky="w")

entry_namen = tk.Text(root, font=("Arial", 12), width=25, height=40, wrap=WORD)
entry_namen.grid(row=1, column=0, padx=(15, 10), pady=(0, 10), sticky="nsew")

# Ausgabefeld für Sitzordnungen
label_ausgabe = tk.Label(root, font=("Arial", 16), text="Ausgabe:")
label_ausgabe.grid(row=0, column=1, padx=(15, 10), pady=(15, 10), sticky="w",)

text_output = tk.Text(root, font=("Arial", 12), width=60, height=40)
text_output.grid(row=1, column=1, padx=(15, 10), pady=(0, 10), sticky="nsew")

# Eingabefeld für Personen pro Tisch
label_personen_pro_tisch = tk.Label(root, font=("Arial", 16), text="Personen pro Tisch:")
label_personen_pro_tisch.grid(row=2, column=0, padx=(15, 10), pady=(15, 10), sticky="w")

entry_personen_pro_tisch = tk.Entry(root, font=("Arial", 12), width=25, justify="center")
entry_personen_pro_tisch.grid(row=3, column=0, padx=(15, 10), pady=(0, 10), sticky="nswe")

# Import Excel-Sheet
importexcel_button = tk.Button(root, font=("Arial", 16), width=50, text="Infos aus Excel-Datei importieren", command=import_from_excel)
importexcel_button.grid(row=4, column=0, padx=(15, 10), pady=(10, 15))

# Button zum Generieren der Sitzordnungen
generate_button = tk.Button(root, font=("Arial", 16), text="Sitzordnungen generieren", command=generate_sitzordnungen, width=50)
generate_button.grid(row=5, column=0, padx=(15, 10), pady=(10, 15))

# Als PDF Speichern Button
saveaspdf_button = tk.Button(root, font=("Arial", 16), width=50, text="Als PDF Speichern", command=save_as_pdf)
saveaspdf_button.grid(row=4, column=1, padx=(15, 10), pady=(10, 15))

# Exit-Button
exit_button = tk.Button(root, font=("Arial", 16), width=50, text="Exit", command=close_window)
exit_button.grid(row=5, column=1, padx=(15, 10), pady=(10, 15))

# GUI starten
root.mainloop()
