"""
Author: Phillip Kusinski
GUI tool for analyzing and exporting screw assembly data for BMW G60 production reports on Deprag AST11 system.
"""

import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
import pandas as pd
import matplotlib.pyplot as plt
import os
from io import BytesIO

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
    
    def create_detailed_dataframe(self):
        #.groupby:group filtered dataframe into "Datum", "Roboternummer", "Fehlernummer"
        #.size(): counts number of entrys i nevery group
        #.unstack(): "Fehlernummer" will be changed to the different failure nums as a own col
        df_grouped_detailed = self.df.groupby([self.df["Datum"].dt.date, "Roboternummer", "Status"]).size().unstack(fill_value=0)
        #sum num of all entrys of axis 1
        df_grouped_detailed["Gesamtverschraubungen"] = df_grouped_detailed.sum(axis=1)
        #get all the cols with failures except 0 and "Gesamtverschraubungen"
        fail_cols = [col for col in df_grouped_detailed.columns if col not in [0, "Gesamtverschraubungen"]]
        #Calculate "Fehler in %"
        df_grouped_detailed["Fehler in %"] = (df_grouped_detailed[fail_cols].sum(axis=1) / df_grouped_detailed["Gesamtverschraubungen"] * 100).round(2)
        return df_grouped_detailed
    
    def create_detailed_dataframe_weekly(self):
        #.groupby:group filtered dataframe into "Roboternummer", "Fehlernummer"
        #.size(): counts number of entrys i nevery group
        #.unstack(): "Fehlernummer" will be changed to the different failure nums as a own col
        df_grouped_detailed_weekly = self.df.groupby(["Roboternummer", "Status"]).size().unstack(fill_value=0)
        #sum num of all entrys of axis 1
        df_grouped_detailed_weekly["Gesamtverschraubungen"] = df_grouped_detailed_weekly.sum(axis=1)
        #get all the cols with failures except 0 and "Gesamtverschraubungen"
        fail_cols = [col for col in df_grouped_detailed_weekly.columns if col not in [0, "Gesamtverschraubungen"]]
        #Calculate "Fehler in %"
        df_grouped_detailed_weekly["Fehler in %"] = (df_grouped_detailed_weekly[fail_cols].sum(axis=1) / df_grouped_detailed_weekly["Gesamtverschraubungen"] * 100).round(2)
        return df_grouped_detailed_weekly
    
    def excel_export(self, fig, df_daily, df_weekly):
        #select save name for the file with correct path settings
        save_name = f"{self.save_path}/Schraubreport_{self.variant}_KW{self.calendarweek}_{self.year}.xlsx"
        
        with pd.ExcelWriter(save_name) as writer:
            df_daily.to_excel(writer, sheet_name = "G60 daily")
            df_weekly.to_excel(writer, sheet_name = "G60 weekly")

            #format seetings for failure percent coloring
            workbook = writer.book
            green_format = workbook.add_format({'bg_color': '#C6EFCE', 'font_color': '#006100'})
            red_format = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#9C0006'})
            yellow_format = workbook.add_format({'bg_color': '#FFEB9C', 'font_color': '#9C6500'})
            
            worksheet = writer.sheets["G60 weekly"]  
            try:
                #code snippet created with AI
                col_index = df_weekly.columns.get_loc("Fehler in %")  
                excel_col_letter = chr(ord('A') + col_index + 1)   
                cell_range = f"{excel_col_letter}2:{excel_col_letter}3"
                worksheet.conditional_format(cell_range, {
                    'type': 'cell',
                    'criteria': '>=',
                    'value': 0.5,
                    'format': red_format
                })
                worksheet.conditional_format(cell_range, {
                    'type': 'cell',
                    'criteria': 'between',
                    'minimum': 0.2001,
                    'maximum': 0.4999,
                    'format': yellow_format
                })
                worksheet.conditional_format(cell_range, {
                    'type': 'cell',
                    'criteria': '<=',
                    'value': 0.2,
                    'format': green_format
                })
            except KeyError:
                print(f"Spalte 'Fehler in %' nicht gefunden weekly.")
                
            worksheet = writer.sheets["G60 daily"]
            try:
                #code snippet created with AI
                col_index = df_daily.columns.get_loc("Fehler in %")  
                excel_col_letter = chr(ord('A') + col_index + 2)   
                row_start = 2
                row_end = len(df_daily) + 1
                cell_range = f"{excel_col_letter}{row_start}:{excel_col_letter}{row_end}"
                worksheet.conditional_format(cell_range, {
                    'type': 'cell',
                    'criteria': '>=',
                    'value': 0.5,
                    'format': red_format
                })
                worksheet.conditional_format(cell_range, {
                    'type': 'cell',
                    'criteria': 'between',
                    'minimum': 0.2001,
                    'maximum': 0.4999,
                    'format': yellow_format
                })
                worksheet.conditional_format(cell_range, {
                    'type': 'cell',
                    'criteria': '<=',
                    'value': 0.2,
                    'format': green_format
                })
            except KeyError:
                print(f"Spalte 'Fehler in %' nicht gefunden daily.")   

            #BytesIO that the image is buffered in the RAM rather than saved on the desktop
            image_stream = BytesIO()
            fig.savefig(image_stream, format='png', dpi=300, bbox_inches='tight')
            #reset RAM buffer
            image_stream.seek(0)

            #insert fig into selected worksheet
            worksheet = writer.sheets["G60 weekly"]
            worksheet.insert_image("A7", "", {
                "image_data": image_stream,
                "x_offset": 5,
                "y_offset": 5,
                "x_scale": 0.5,
                "y_scale": 0.5
            })  
                
    def main_filter_func(self):
        fig = self.create_failure_plot()
        df_grouped_detailed = self.create_detailed_dataframe()
        df_grouped_detailed_weekly = self.create_detailed_dataframe_weekly()
        return fig, df_grouped_detailed, df_grouped_detailed_weekly
    
    def export_data(self):
        if self.df.empty or not self.save_path:
            messagebox.showerror("Export nicht möglich", "Bitte zuerst Daten laden und Speicherpfad wählen.")
            return
        fig, df_grouped_detailed, df_grouped_detailed_weekly = self.main_filter_func()
        self.excel_export(fig, df_grouped_detailed, df_grouped_detailed_weekly)
        messagebox.showinfo("Export erfolgreich", "Der Export wurde erfolgreich durchgeführt.")


if __name__ == "__main__":
    root = tk.Tk()
    app = ScrewAnalysisApp(root)
    root.mainloop()
