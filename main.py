import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
from analysis import generate_evaluation_report
from preprocessing import process_excel_data


def open_excel_file_dialog():
    """Lässt den Nutzer eine Excel-Datei auswählen."""

    # Dialog zum Auswählen einer Datei öffnen
    chosen_file_path = filedialog.askopenfilename(
        title="Wähle eine Excel-Datei", filetypes=[("Excel-Dateien", "*.xlsx")]
    )

    # Wenn eine Datei ausgewählt wurde, wird der Dateipfad gespeichert
    if chosen_file_path:
        excel_input_path.set(chosen_file_path)


def export_processed_excel():
    """Verarbeitet die ausgewählte Excel-Datei und speichert das Ergebnis."""

    # Dateipfad auslesen
    input_path = excel_input_path.get()

    # Wenn keine Datei ausgewählt wurde, wird eine Fehlermeldung angezeigt
    if not input_path:
        messagebox.showerror("Fehler", "Keine Eingabedatei ausgewählt.")
        return

    # Dialog zum Speichern der Datei öffnen
    output_path = filedialog.asksaveasfilename(
        title="Wähle den Speicherort für die neue Excel-Datei",
        defaultextension=".xlsx",
        filetypes=[("Excel-Dateien", "*.xlsx")],
        initialfile="Auswertung FSL-7.xlsx",
    )

    # Wenn kein Speicherort ausgewählt wurde, wird die Verarbeitung abgebrochen
    if not output_path:
        return

    try:
        # Excel-Datei verarbeiten
        excel_document = process_excel_data(input_path)
        result = generate_evaluation_report(excel_document)

        # Verarbeitete Datei speichern
        result.save(output_path)

        # Datei nach dem Speichern direkt öffnen
        os.startfile(output_path)

        # Programm beenden
        sys.exit()

    # Falls ein Fehler auftritt, wird eine Fehlermeldung angezeigt
    except Exception as e:
        messagebox.showerror("Fehler", f"Ein Fehler ist aufgetreten:\n{e}")


# Beginn des Programms
# Hauptfenster erstellen
main_window = tk.Tk()
main_window.title("Auswertung FSL-7")

# Maximierung deaktivieren
main_window.resizable(False, False)

# Variable zum Speichern des Dateipfads
excel_input_path = tk.StringVar()

# Fenster konfigurieren
frame = tk.Frame(main_window, padx=10, pady=10)
frame.pack(fill=tk.BOTH, expand=True)

# Button zum Auswählen der Datei
select_file_button = tk.Button(frame, text="Excel-Datei auswählen", command=open_excel_file_dialog)
select_file_button.grid(row=0, column=0, padx=5, pady=5, sticky="ew")

# Anzeige des ausgewählten Dateipfads
selected_file_path_entry = tk.Entry(frame, textvariable=excel_input_path, state="readonly", width=50)
selected_file_path_entry.grid(row=0, column=1, padx=5, pady=5, sticky="ew")

# Button zum Verarbeiten und Speichern der Datei
export_results_button = tk.Button(frame, text="Auswerten und Speichern", command=export_processed_excel)
export_results_button.grid(row=1, column=0, columnspan=2, padx=5, pady=5, sticky="ew")

# Spaltenverhalten konfigurieren
frame.columnconfigure(0, weight=1)
frame.columnconfigure(1, weight=3)

# Programm starten
main_window.mainloop()
