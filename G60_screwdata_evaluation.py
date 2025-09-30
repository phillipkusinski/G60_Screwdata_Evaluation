"""
Author: Phillip Kusinski
GUI tool for analyzing and exporting screw assembly data for BMW G60 production reports
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
import os

class ScrewAnalysisApp:
    def __init__(self, root):
        self.root = root
        self.root.title("G60 Schraubauswertung")
        self.root.geometry("350x320")
        self.root.resizable(False, False)
        self.root.configure(padx=20, pady=20, bg="#f0f0f0")

        # App state
        self.file_paths = []
        self.save_path = ""
        self.calendarweek = 0
        self.year = 0
        self.rob_nums = ["Rob_8_1", "Rob_8_2", "Rob_8_3", "Rob_9_1", "Rob_9_2", "Rob_9_3"]
        self.variant = ""
        self.df = pd.DataFrame()

        # GUI setup
        self.setup_styles()
        self.setup_gui()

    def setup_styles(self):
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure("TFrame", background="#f0f0f0")
        style.configure("TLabel", background="#f0f0f0")
        style.configure("Export.TButton",
                        font=("Arial", 16, "bold"),
                        foreground="white",
                        background="#28a745")
        style.map("Export.TButton",
                  background=[("active", "#1e7e34")],
                  foreground=[("active", "white")])

    def setup_gui(self):
        frame = ttk.Frame(self.root)
        frame.grid(row=0, column=0, sticky="ew")
        self.root.columnconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)

        ttk.Button(frame, text="📂 xlsx-Datei öffnen", command=self.open_csv_files).grid(row=0, column=0, sticky="ew")
        self.lbl_status = ttk.Label(frame, text="0 Dateien ausgewählt")
        self.lbl_status.grid(row=0, column=1, sticky="w", padx=(20, 0))

        ttk.Button(frame, text="Erstelle Datenstruktur", command=self.build_dataframe).grid(row=1, column=0, columnspan=2, sticky="ew", pady=10)
        ttk.Separator(frame, orient="horizontal").grid(row=5, column=0, columnspan=2, sticky="ew", pady=15)

        ttk.Button(frame, text="📂 Speicherpfad auswählen", command=self.select_save_path).grid(row=6, column=0, columnspan=2, sticky="ew")
        ttk.Button(frame, text="Export starten", command=self.export_data, style="Export.TButton").grid(row=7, column=0, columnspan=2, sticky="ew", pady=20)
        ttk.Separator(frame, orient="horizontal").grid(row=8, column=0, columnspan=2, sticky="ew", pady=15)

        ttk.Label(frame, text="Phillip Kusinski, V1.0", style="TLabel").grid(row=9, column=1, sticky="e")

    def open_csv_files(self):
        folder = filedialog.askdirectory(title="Ordner auswählen mit csv-Dateien")
        if not folder:
            return

        paths = [os.path.join(root, file)
                 for root, _, files in os.walk(folder)
                 for file in files if file.endswith(".csv")]

        if len(paths) > 21:
            messagebox.showwarning("Zu viele Dateien", "Bitte wählen Sie maximal 32 .xlsx-Dateien aus")
            return

        self.file_paths = paths
        self.lbl_status.config(text=f"{len(paths)} Datei(en) gefunden")

    def build_dataframe(self):
        if not self.file_paths:
            messagebox.showerror("Keine Daten ausgewählt", "Es wurden keine Daten zur Auswertung ausgewählt!")
            return

        self.variant = self.detect_variant()
        list_of_df = []
        for file in self.file_paths:
            try:
                df = pd.read_csv(file, sep=',', usecols=[0, 1, 2, 3], header=None, skiprows=1)
                # if df.shape[1] != 10:
                #     raise ValueError(f"Datei '{os.path.basename(file)}' hat {df.shape[1]} Spalten, erwartet wurden 10.")
                print(df.head())
                rob_num = next((part for part in os.path.normpath(file).split(os.sep) if part.startswith("Rob_")), "Unbekannt")
                df["Roboternummer"] = rob_num
                list_of_df.append(df)
            except Exception:
                messagebox.showerror("Fehler beim Laden", f"❌ Datei konnte nicht verarbeitet werden: {file}")
                return

        self.df = pd.concat(list_of_df, ignore_index=True)
        self.df.columns = ["Datum", "Programmnummer", "Fehlernummer", "Gesamtlaufzeit",
                           "Schritt 3", "Drehmoment 3", "Drehwinkel 3", "Schritt NOK",
                           "Drehmoment NOK", "Drehwinkel NOK", "Roboternummer"]

        if self.check_calendarweek():
            messagebox.showinfo("Datenstruktur erfolgreich", f"Variante {self.variant}, KW{self.calendarweek}")
        else:
            messagebox.showerror("Fehler beim Aufbau der Datenstruktur", "Datensätze stammen nicht aus derselben Kalenderwoche!")
            self.df = pd.DataFrame()
            self.calendarweek = 0

    def check_calendarweek(self):
        self.df['Datum'] = pd.to_datetime(self.df['Datum'])
        iso = self.df['Datum'].dt.isocalendar()
        if iso['week'].nunique() == 1 and iso['year'].nunique() == 1:
            self.calendarweek = iso['week'].iloc[0]
            self.year = iso['year'].iloc[0]
            return True
        return False

    def detect_variant(self):
        keywords = ["Hintertür", "Vordertür"]
        parts = os.path.normpath(self.file_paths[0]).split(os.sep)
        return next((part for part in parts if part in keywords), "Unbekannt")

    def select_save_path(self):
        path = filedialog.askdirectory(title="Ordner zur Abspeicherung der Prüfergebnisse auswählen.")
        if not path:
            return
        self.save_path = path
        messagebox.showinfo("Ordnerwahl erfolgreich", "Speicherpfad erfolgreich ausgewählt.")

    def export_data(self):
        if self.df.empty or not self.save_path:
            messagebox.showerror("Export nicht möglich", "Bitte zuerst Daten laden und Speicherpfad wählen.")
            return
        # Hier kommt deine Exportlogik
        print(f"Exportiere Daten für KW{self.calendarweek} nach: {self.save_path}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ScrewAnalysisApp(root)
    root.mainloop()
