"""
Author: Phillip Kusinski
GUI tool for analyzing and exporting screw assembly data for BMW G60 production reports on Deprag AST11 system.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
import os

class ScrewAnalysisApp:
    def __init__(self, root):
        #setup root app features
        self.root = root
        self.root.title("G60 Schraubauswertung")
        self.root.geometry("350x320")
        self.root.resizable(False, False)
        self.root.configure(padx=20, pady=20, bg="#f0f0f0")

        #setup instance variables
        self.file_paths = []
        self.save_path = ""
        self.calendarweek = 0
        self.year = 0
        self.variant = ""
        self.df = pd.DataFrame()

        #GUI styling init
        self.setup_styles()
        self.setup_gui()

    def setup_styles(self):
        #setup GUI styling
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
        #setup GUI structure 
        frame = ttk.Frame(self.root)
        frame.grid(row=0, column=0, sticky="ew")
        self.root.columnconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)

        ttk.Button(frame, text="📂 csv-Datei auswählen", command=self.open_csv_files).grid(row=0, column=0, sticky="ew")
        self.lbl_status = ttk.Label(frame, text="0 Dateien ausgewählt")
        self.lbl_status.grid(row=0, column=1, sticky="w", padx=(20, 0))

        ttk.Button(frame, text="Erstelle Datenstruktur", command=self.build_dataframe).grid(row=1, column=0, columnspan=2, sticky="ew", pady=10)
        ttk.Separator(frame, orient="horizontal").grid(row=5, column=0, columnspan=2, sticky="ew", pady=15)

        ttk.Button(frame, text="📂 Speicherpfad auswählen", command=self.select_save_path).grid(row=6, column=0, columnspan=2, sticky="ew")
        ttk.Button(frame, text="Export starten", command=self.export_data, style="Export.TButton").grid(row=7, column=0, columnspan=2, sticky="ew", pady=20)
        ttk.Separator(frame, orient="horizontal").grid(row=8, column=0, columnspan=2, sticky="ew", pady=15)

        ttk.Label(frame, text="Phillip Kusinski, V1.0", style="TLabel").grid(row=9, column=1, sticky="e")

    def open_csv_files(self):
        #select folders with askdirectory
        folder = filedialog.askdirectory(title="Ordner auswählen mit csv-Dateien")
        #if no folder was selected return
        if not folder:
            return
        #save selected paths in variable
        paths = [os.path.join(root, file)
                 for root, _, files in os.walk(folder)
                 for file in files if file.endswith(".csv")]
        #not possible to select more than 3 Robs * 7 days = 21
        if len(paths) > 21:
            messagebox.showwarning("Zu viele Dateien", "Bitte wählen Sie maximal 32 .xlsx-Dateien aus")
            return
        #set instance and status 
        self.file_paths = paths
        self.lbl_status.config(text=f"{len(paths)} Datei(en) gefunden")

    def build_dataframe(self):
        expected_col = 5
        if not self.file_paths:
            messagebox.showerror("Keine Daten ausgewählt", "Es wurden keine Daten zur Auswertung ausgewählt!")
            return

        self.variant = self.detect_variant()
        list_of_df = []
        for file in self.file_paths:
            try:
                df = pd.read_csv(file, sep=',', usecols=[0, 1, 2, 3], header=None, skiprows=1)
                # if df.shape[1] != expected_col:
                #     raise ValueError(f"Datei '{os.path.basename(file)}' hat {df.shape[1]} Spalten, erwartet wurden {expected_col}.")
                rob_num = next((part for part in os.path.normpath(file).split(os.sep) if part.startswith("Rob_")), "Unbekannt")
                df["Roboternummer"] = rob_num
                list_of_df.append(df)
            except Exception:
                messagebox.showerror("Fehler beim Laden", f"❌ Datei konnte nicht verarbeitet werden: {file}")
                return

        self.df = pd.concat(list_of_df, ignore_index=True)
        cols = ["Datum", "Uhrzeit", "Programmnummer", "Status", "Roboternummer"]
        self.df.columns = cols
        
        if self.check_calendarweek() == True:          
            messagebox.showinfo("Datenstruktur erfolgreich", f"Variante {self.variant}, KW{self.calendarweek}")
        else:
            messagebox.showerror("Fehler beim Aufbau der Datenstruktur", "Datensätze stammen nicht aus derselben Kalenderwoche!")
            self.df = pd.DataFrame()
            self.calendarweek = 0

    def check_calendarweek(self):
        self.df['Datum'] = pd.to_datetime(self.df['Datum'], dayfirst = True)
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

    def create_failure_plot(self):
        df_failure = (self.df.groupby(["Datum", "Roboternummer"], group_keys = False)
        .apply(lambda df_lambda: (df_lambda["Status"] != 0).sum() / len(df_lambda) * 100)
        .reset_index(name="Fehleranteil in %")
        )
        #set date without timestamps
        df_failure["Datum"] = df_failure["Datum"].dt.date

        #pivot df into correct form
        pivot_df = df_failure.pivot(index="Datum", columns="Roboternummer", values="Fehleranteil in %")

        #calculate weekly failure
        #1.         .groupby: "Roboternummer": all dates will be set together per robot 
        #2.         .apply: lambda func that iterates through set groupby filters and calculates the "Fehler in %" value of the sum of all days per robot
        #3.         .round(2): for better visualization
        weekly_failure = (
            self.df.groupby("Roboternummer")
            .apply(lambda x: (x["Status"] != 0).sum() / len(x) * 100)
            .round(2)
        )

        #set data for plot df
        pivot_df.loc["Ø Woche"] = weekly_failure

        #plot data
        ax = pivot_df.plot(kind="bar", figsize=(12, 6))
        plt.axhline(0.2, color='red', linestyle='--', linewidth = 2)
        plt.ylabel("Fehleranteil in %")
        plt.title(f"Variante = {self.variant}, Kalenderwoche = {self.calendarweek}, Absoluter Fehleranteil in % pro Roboter")
        plt.xticks(rotation=0)
        plt.legend(title="Roboternummer", framealpha = 1)

        sep_index = len(pivot_df) - 2
        plt.axvline(x=sep_index + 0.5, color="gray", linestyle="--", linewidth=1)

        plt.tight_layout()
        fig = ax.figure
        return fig
    
    def main_filter_func(self):
        fig = self.create_failure_plot()
        fig.show()
        # df_grouped_detailed = create_detailed_dataframe()
        # df_grouped_detailed_weekly = create_detailed_dataframe_weekly()
        return
    
    def export_data(self):
        if self.df.empty or not self.save_path:
            messagebox.showerror("Export nicht möglich", "Bitte zuerst Daten laden und Speicherpfad wählen.")
            return
        self.main_filter_func()
        # Hier kommt deine Exportlogik
        print(f"Exportiere Daten für KW{self.calendarweek} nach: {self.save_path}")


if __name__ == "__main__":
    root = tk.Tk()
    app = ScrewAnalysisApp(root)
    root.mainloop()
