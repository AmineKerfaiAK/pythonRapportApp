import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import sqlite3
import os
import subprocess
import unicodedata

class RapportGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Rapport Generator")
        self.root.geometry("800x600")

        # Initialize database
        self.conn = sqlite3.connect("rapport_data.db")
        self.create_tables()

        # Sidebar
        self.sidebar = tk.Frame(self.root, width=200, bg="lightgray")
        self.sidebar.pack(side=tk.LEFT, fill=tk.Y)

        # Rapport page button
        tk.Button(self.sidebar, text="Rapport", command=self.show_rapport_page).pack(pady=10)

        # Main content frame
        self.content_frame = tk.Frame(self.root)
        self.content_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # Rapport frame
        self.rapport_frame = tk.Frame(self.content_frame)
        self.file_paths = {"chiffre": None, "productivity": None, "bordereaux": None}

        # Rapport page widgets
        tk.Label(self.rapport_frame, text="Upload Chiffre d'affaire Excel").pack(pady=5)
        tk.Button(self.rapport_frame, text="Browse", command=lambda: self.upload_file("chiffre")).pack(pady=5)
        self.chiffre_label = tk.Label(self.rapport_frame, text="No file selected")
        self.chiffre_label.pack()

        tk.Label(self.rapport_frame, text="Upload Productivity Excel").pack(pady=5)
        tk.Button(self.rapport_frame, text="Browse", command=lambda: self.upload_file("productivity")).pack(pady=5)
        self.productivity_label = tk.Label(self.rapport_frame, text="No file selected")
        self.productivity_label.pack()

        tk.Label(self.rapport_frame, text="Upload Bordereaux Excel").pack(pady=5)
        tk.Button(self.rapport_frame, text="Browse", command=lambda: self.upload_file("bordereaux")).pack(pady=5)
        self.bordereaux_label = tk.Label(self.rapport_frame, text="No file selected")
        self.bordereaux_label.pack()

        tk.Button(self.rapport_frame, text="Generate PDF", command=self.generate_pdf).pack(pady=20)

    def create_tables(self):
        cursor = self.conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS revenues (
                category TEXT,
                jan INTEGER, feb INTEGER, mar INTEGER, apr INTEGER, may INTEGER,
                jun INTEGER, jul INTEGER, aug INTEGER, sep INTEGER, oct INTEGER,
                nov INTEGER, dec INTEGER, total_2024 INTEGER,
                target_percentage FLOAT, target_value INTEGER
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS dossiers (
                category TEXT,
                jan INTEGER, feb INTEGER, mar INTEGER, apr INTEGER, may INTEGER,
                jun INTEGER, jul INTEGER, aug INTEGER, sep INTEGER, oct INTEGER,
                nov INTEGER, dec INTEGER, total_2024 INTEGER
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS dashboard_current_month (
                category TEXT, revenue INTEGER, percentage FLOAT
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS dashboard_year_to_date (
                category TEXT, revenue INTEGER, percentage FLOAT
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS agent_productivity (
                agent_name TEXT, auth_conform_files INTEGER, tech_control_files INTEGER
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS processing_times (
                category TEXT, on_time FLOAT, before_time FLOAT, after_time FLOAT
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS completion_stats (
                category TEXT, complete FLOAT, incomplete FLOAT
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS intervention_reasons (
                category TEXT, technical_docs FLOAT, device_operation FLOAT, other_reasons FLOAT
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS bordereaux (
                bordereau_no INTEGER, dossier_no TEXT, dossier_type TEXT, result TEXT,
                intervenant TEXT, cause_fi TEXT, delai_execution TEXT
            )
        """)
        self.conn.commit()

    def normalize_text(self, text):
        if pd.isna(text):
            return ""
        text = str(text).strip()
        text = unicodedata.normalize('NFKC', text)
        return text

    def escape_latex(self, text):
        if not text:
            return ""
        text = str(text)
        replacements = {
            "&": "\\&",
            "%": "\\%",
            "$": "\\$",
            "#": "\\#",
            "_": "\\_",
            "{": "\\{",
            "}": "\\}",
            "~": "\\textasciitilde{}",
            "^": "\\textasciicircum{}",
            "\\": "\\textbackslash{}"
        }
        for old, new in replacements.items():
            text = text.replace(old, new)
        return text

    def upload_file(self, file_type):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.file_paths[file_type] = file_path
            if file_type == "chiffre":
                self.chiffre_label.config(text=os.path.basename(file_path))
                self.process_chiffre_file(file_path)
            elif file_type == "productivity":
                self.productivity_label.config(text=os.path.basename(file_path))
                self.process_productivity_file(file_path)
            elif file_type == "bordereaux":
                self.bordereaux_label.config(text=os.path.basename(file_path))
                self.process_bordereaux_file(file_path)

    def process_chiffre_file(self, file_path):
        try:
            df = pd.read_excel(file_path, sheet_name="Feuil1", skiprows=2)
            df = df.dropna(axis=1, how='all')
            print("Chiffre DataFrame shape:", df.shape)
            print("Chiffre DataFrame columns:", df.columns.tolist())
            print("Chiffre DataFrame first 10 rows:\n", df.head(10).to_string())

            cursor = self.conn.cursor()
            cursor.execute("DELETE FROM revenues")
            cursor.execute("DELETE FROM dossiers")
            cursor.execute("DELETE FROM dashboard_current_month")
            cursor.execute("DELETE FROM dashboard_year_to_date")

            # Revenues (rows 0-4: HOMOLOG to total moi)
            for i, row in df.iloc[0:5].iterrows():
                category = self.normalize_text(row.iloc[0])
                print(f"Inserting revenue for category: '{category}' at row {i}")
                cursor.execute("""
                    INSERT INTO revenues VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    category,
                    int(row.iloc[1]) if pd.notna(row.iloc[1]) else 0,
                    int(row.iloc[2]) if pd.notna(row.iloc[2]) else 0,
                    int(row.iloc[3]) if pd.notna(row.iloc[3]) else 0,
                    int(row.iloc[4]) if pd.notna(row.iloc[4]) else 0,
                    int(row.iloc[5]) if pd.notna(row.iloc[5]) else 0,
                    0, 0, 0, 0, 0, 0, 0,
                    int(row.iloc[13]) if pd.notna(row.iloc[13]) else 0,
                    float(row.iloc[14]) if pd.notna(row.iloc[14]) else 0,
                    int(row.iloc[16]) if pd.notna(row.iloc[16]) else 0
                ))

            # Dossiers (rows 7-11: HOMOLOG to total moi, skipping row 6 header)
            for i, row in df.iloc[7:12].iterrows():
                category = self.normalize_text(row.iloc[0])
                if not category:
                    continue
                print(f"Inserting dossier for category: '{category}' at row {i}")
                cursor.execute("""
                    INSERT INTO dossiers VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    category,
                    int(row.iloc[1]) if pd.notna(row.iloc[1]) else 0,
                    int(row.iloc[2]) if pd.notna(row.iloc[2]) else 0,
                    int(row.iloc[3]) if pd.notna(row.iloc[3]) else 0,
                    int(row.iloc[4]) if pd.notna(row.iloc[4]) else 0,
                    int(row.iloc[5]) if pd.notna(row.iloc[5]) else 0,
                    0, 0, 0, 0, 0, 0, 0,
                    int(row.iloc[13]) if pd.notna(row.iloc[13]) else 0
                ))

            # Dashboard current month (rows 12-16)
            for i, row in df.iloc[12:17].iterrows():
                category = self.normalize_text(row.iloc[14]) if pd.notna(row.iloc[14]) else ""
                if not category:
                    continue
                cursor.execute("""
                    INSERT INTO dashboard_current_month VALUES (?, ?, ?)
                """, (
                    category,
                    int(row.iloc[15]) if pd.notna(row.iloc[15]) else 0,
                    float(row.iloc[13]) if pd.notna(row.iloc[13]) else 0
                ))

            # Dashboard year to date (rows 19-23)
            for i, row in df.iloc[19:24].iterrows():
                category = self.normalize_text(row.iloc[14]) if pd.notna(row.iloc[14]) else ""
                if not category:
                    continue
                cursor.execute("""
                    INSERT INTO dashboard_year_to_date VALUES (?, ?, ?)
                """, (
                    category,
                    int(row.iloc[15]) if pd.notna(row.iloc[15]) else 0,
                    float(row.iloc[13]) if pd.notna(row.iloc[13]) else 0
                ))

            self.conn.commit()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process Chiffre file: {str(e)}")
            print(f"Chiffre file processing error: {str(e)}")
            raise

    def process_productivity_file(self, file_path):
        try:
            df = pd.read_excel(file_path, sheet_name="Feuil1", skiprows=1)
            print("Productivity DataFrame shape:", df.shape)
            print("Productivity DataFrame columns:", df.columns.tolist())
            print("Productivity DataFrame first 10 rows:\n", df.head(10).to_string())

            cursor = self.conn.cursor()
            cursor.execute("DELETE FROM agent_productivity")
            cursor.execute("DELETE FROM processing_times")
            cursor.execute("DELETE FROM completion_stats")
            cursor.execute("DELETE FROM intervention_reasons")

            # Authentication/conformity agents (rows 0-10)
            for _, row in df.iloc[0:11].iterrows():
                if pd.notna(row.iloc[0]) and pd.notna(row.iloc[1]) and isinstance(row.iloc[1], (int, float)):
                    cursor.execute("""
                        INSERT INTO agent_productivity (agent_name, auth_conform_files, tech_control_files)
                        VALUES (?, ?, ?)
                    """, (
                        self.normalize_text(row.iloc[0]),
                        int(row.iloc[1]),
                        0
                    ))

            # Technical control agents (rows 27-30)
            for _, row in df.iloc[27:31].iterrows():
                if pd.notna(row.iloc[0]) and pd.notna(row.iloc[1]) and isinstance(row.iloc[1], (int, float)):
                    cursor.execute("""
                        INSERT INTO agent_productivity (agent_name, auth_conform_files, tech_control_files)
                        VALUES (?, ?, ?)
                    """, (
                        self.normalize_text(row.iloc[0]),
                        0,
                        int(row.iloc[1])
                    ))

            # Hardcoded data
            cursor.execute("INSERT INTO processing_times VALUES (?, ?, ?, ?)", ("المصادقة", 30, 35, 35))
            cursor.execute("INSERT INTO processing_times VALUES (?, ?, ?, ?)", ("المطابقة", 25, 15, 60))
            cursor.execute("INSERT INTO processing_times VALUES (?, ?, ?, ?)", ("المراقبة الفنية", 14, 23, 63))
            cursor.execute("INSERT INTO processing_times VALUES (?, ?, ?, ?)", ("المراقبة الفنية تحت الديوانة", 0, 0, 100))
            cursor.execute("INSERT INTO processing_times VALUES (?, ?, ?, ?)", ("موردي السيارات", 10, 0, 90))
            cursor.execute("INSERT INTO completion_stats VALUES (?, ?, ?)", ("المصادقة", 57, 43))
            cursor.execute("INSERT INTO completion_stats VALUES (?, ?, ?)", ("المطابقة", 52, 48))
            cursor.execute("INSERT INTO intervention_reasons VALUES (?, ?, ?, ?)", ("المصادقة", 60, 30, 10))
            cursor.execute("INSERT INTO intervention_reasons VALUES (?, ?, ?, ?)", ("المطابقة", 40, 52, 8))

            self.conn.commit()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process Productivity file: {str(e)}")
            print(f"Productivity file processing error: {str(e)}")
            raise

    def process_bordereaux_file(self, file_path):
        try:
            # Read the Excel file without a predefined header
            df = pd.read_excel(file_path, sheet_name="Feuil1", header=None)
            print(f"Bordereaux DataFrame shape: {df.shape}")
            print(f"Bordereaux DataFrame first 10 rows:\n{df.head(10)}")

            # Find the row index where 'N° Bordereaux' appears in the first column
            header_row = df[df.iloc[:, 0] == 'N° Bordereaux'].index
            if not header_row.empty:
                header_row = header_row[0]
            else:
                raise ValueError("Header row with 'N° Bordereaux' not found")

            # Set the header using the identified row
            df.columns = df.iloc[header_row]
            # Keep only the data rows after the header
            df = df.iloc[header_row + 1:].reset_index(drop=True)
            print(f"Bordereaux DataFrame columns after setting header: {list(df.columns)}")
            print(f"Bordereaux DataFrame first 10 rows after setting header:\n{df.head(10)}")

            # Clear existing data in the bordereaux table
            cursor = self.conn.cursor()
            cursor.execute("DELETE FROM bordereaux")

            # Insert data into the SQLite table
            for _, row in df.iterrows():
                cursor.execute("""
                    INSERT INTO bordereaux VALUES (?, ?, ?, ?, ?, ?, ?)
                """, (
                    int(row['N° Bordereaux']) if pd.notna(row['N° Bordereaux']) else 0,
                    self.normalize_text(row['N° dossier']) if pd.notna(row['N° dossier']) else "",
                    self.normalize_text(row['Type de dossier H/C']) if pd.notna(row['Type de dossier H/C']) else "",
                    self.normalize_text(row['Resultat R/FI']) if pd.notna(row['Resultat R/FI']) else "",
                    self.normalize_text(row['Intervenant']) if pd.notna(row['Intervenant']) else "",
                    self.normalize_text(row['Cause FI']) if pd.notna(row['Cause FI']) else "",
                    self.normalize_text(row["Délai d'exécution"]) if pd.notna(row["Délai d'exécution"]) else ""
                ))
            self.conn.commit()
            print("Bordereaux data successfully inserted into the database.")

        except ValueError as ve:
            messagebox.showerror("Error", str(ve))
            print(f"Bordereaux file processing error: {str(ve)}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process Bordereaux file: {str(e)}")
            print(f"Bordereaux file processing error: {str(e)}")
            raise

    def show_rapport_page(self):
        for widget in self.content_frame.winfo_children():
            if widget != self.rapport_frame:
                widget.pack_forget()
        self.rapport_frame.pack(fill=tk.BOTH, expand=True)

    def generate_pdf(self):
        if not all(self.file_paths.values()):
            messagebox.showerror("Error", "Please upload all three Excel files.")
            return

        try:
            cursor = self.conn.cursor()
            cursor.execute("SELECT * FROM revenues WHERE TRIM(LOWER(category)) = 'total moi'")
            total_revenue = cursor.fetchone()
            cursor.execute("SELECT * FROM dossiers WHERE TRIM(LOWER(category)) = 'total moi'")
            total_files = cursor.fetchone()

            if total_revenue is None:
                messagebox.showerror("Error", "No data found for 'total moi' in revenues table.")
                return
            if total_files is None:
                messagebox.showerror("Error", "No data found for 'total moi' in dossiers table.")
                return

            cursor.execute("SELECT * FROM dashboard_current_month")
            dashboard_current = cursor.fetchall()
            cursor.execute("SELECT * FROM dashboard_year_to_date")
            dashboard_year = cursor.fetchall()
            cursor.execute("SELECT * FROM processing_times")
            processing_times = cursor.fetchall()
            cursor.execute("SELECT * FROM completion_stats")
            completion_stats = cursor.fetchall()
            cursor.execute("SELECT * FROM intervention_reasons")
            intervention_reasons = cursor.fetchall()

            latex_content = self.generate_latex(
                total_revenue, total_files, dashboard_current, dashboard_year,
                processing_times, completion_stats, intervention_reasons
            )

            with open("rapport.tex", "w", encoding="utf-8") as f:
                f.write(latex_content)
            subprocess.run(["latexmk", "-xelatex", "-f", "rapport.tex"], check=True)
            messagebox.showinfo("Success", "PDF generated as rapport.pdf")
        except subprocess.CalledProcessError:
            messagebox.showerror("Error", "Failed to generate PDF. Check rapport.log for details.")
        except FileNotFoundError:
            messagebox.showerror("Error", "latexmk not found. Please install LaTeX.")
        except Exception as e:
            messagebox.showerror("Error", f"PDF generation failed: {str(e)}")
            print(f"PDF generation error: {str(e)}")

    def generate_latex(self, total_revenue, total_files, dashboard_current, dashboard_year,
                       processing_times, completion_stats, intervention_reasons):
        latex = r"""
\documentclass[a4paper,12pt]{article}
\usepackage{geometry}
\geometry{a4paper, margin=1in}
\usepackage{polyglossia}
\setmainlanguage{arabic}
\setotherlanguage{english}
\newfontfamily\arabicfont[Script=Arabic]{Arial}
\usepackage{booktabs}
\usepackage{array}
\usepackage{longtable}

\begin{document}

\begin{center}
\textbf{من إدارة الموارد} \\
\textbf{إدارة المصادقة والمواصفات} \\
\textbf{التقرير الشهري} \\
\textbf{2025 ماي}
\end{center}

\section*{II. مؤشرات الإنتاج}

\subsection*{1. المداخيل الجملية للشهر الحالي}
\begin{tabular}{cc}
\toprule
\textbf{قيمة المداخيل (د.ت خال من الأداء على القيمة المصافة)} & \textbf{نوعية المداخيل} \\
\midrule
""" + f"{self.escape_latex(f'{total_revenue[5]:,}')} & {self.escape_latex('مداخيل عمليات المصادقة والمطابقة والمراقبة الفنية')} \\\\\n" + r"""
\bottomrule
\end{tabular}

\subsection*{2. المداخيل الجملية منذ بداية السنة}
\begin{tabular}{cc}
\toprule
\textbf{قيمة المداخيل (د.ت خال من الأداء على القيمة المصافة)} & \textbf{نوعية المداخيل} \\
\midrule
""" + f"{self.escape_latex(f'{total_revenue[13]:,}')} & {self.escape_latex('مداخيل عمليات المصادقة والمطابقة والمراقبة الفنية')} \\\\\n" + r"""
\bottomrule
\end{tabular}

\subsection*{3. العدد الجملي للملفات المنجزة خلال الشهر الحالي}
\begin{tabular}{cc}
\toprule
\textbf{عدد الملفات} & \textbf{نوعية الملفات} \\
\midrule
""" + f"{self.escape_latex(f'{total_files[5]:,}')} & {self.escape_latex('ملفات عمليات المصادقة والمطابقة والمراقبة الفنية')} \\\\\n" + r"""
\bottomrule
\end{tabular}

\subsection*{4. العدد الجملي للملفات المنجزة منذ بداية السنة}
\begin{tabular}{cc}
\toprule
\textbf{عدد الملفات} & \textbf{نوعية الملفات} \\
\midrule
""" + f"{self.escape_latex(f'{total_files[13]:,}')} & {self.escape_latex('ملفات عمليات المصادقة والمطابقة والمراقبة الفنية')} \\\\\n" + r"""
\bottomrule
\end{tabular}

\subsection*{5. معدل أجال دراسة الملفات}
\begin{tabular}{cccc}
\toprule
\textbf{بعد الأجال} & \textbf{قبل الأجال} & \textbf{في الأجال} & \textbf{النشاط} \\
\midrule
"""
        for pt in processing_times:
            latex += f"{self.escape_latex(f'{pt[3]:.1f}')}٪ & {self.escape_latex(f'{pt[2]:.1f}')}٪ & {self.escape_latex(f'{pt[1]:.1f}')}٪ & {self.escape_latex(pt[0])} \\\\\n"
        latex += r"""
\bottomrule
\end{tabular}

\section*{III. الأهداف}
\subsection*{1. على مستوى المداخيل}
\begin{tabular}{cccc}
\toprule
\textbf{النسبة المئوية} & \textbf{قيمة المداخيل المنجزة} & \textbf{قيمة المداخيل المتوقعة} & \textbf{الأهداف ومتابعتها} \\
\midrule
"""
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM revenues WHERE TRIM(LOWER(category)) != 'total moi'")
        for row in cursor.fetchall():
            latex += f"{self.escape_latex(f'{row[14]:.1f}')}٪ & {self.escape_latex(f'{row[13]:,}')} & {self.escape_latex(f'{row[15]:,}')} & {self.escape_latex(row[0])} \\\\\n"
        latex += f"{self.escape_latex(f'{total_revenue[14]:.1f}')}٪ & {self.escape_latex(f'{total_revenue[13]:,}')} & {self.escape_latex(f'{total_revenue[15]:,}')} & المجموع \\\\\n"
        latex += r"""
\bottomrule
\end{tabular}

\section*{IV. لوحة قيادة لمداخيل الشهر الحالي}
\begin{tabular}{ccc}
\toprule
\textbf{٪ من المداخيل الجملية} & \textbf{قيمة المداخيل} & \textbf{نوعية المداخيل} \\
\midrule
"""
        for dc in dashboard_current:
            latex += f"{self.escape_latex(f'{dc[2]:.1f}')}٪ & {self.escape_latex(f'{dc[1]:,}')} & {self.escape_latex(dc[0])} \\\\\n"
        latex += r"""
\bottomrule
\end{tabular}

\section*{V. لوحة قيادة للمداخيل منذ بداية السنة}
\begin{tabular}{ccc}
\toprule
\textbf{٪ من المداخيل الجملية} & \textbf{قيمة المداخيل} & \textbf{نوعية المداخيل} \\
\midrule
"""
        for dy in dashboard_year:
            latex += f"{self.escape_latex(f'{dy[2]:.1f}')}٪ & {self.escape_latex(f'{dy[1]:,}')} & {self.escape_latex(dy[0])} \\\\\n"
        latex += r"""
\bottomrule
\end{tabular}

\section*{VI. لوحة قيادة العمليات المنجزة خلال الشهر الحالي}
\subsection*{1. عمليات المصادقة والمطابقة}
\begin{tabular}{ccc}
\toprule
\textbf{النسبة المئوية} & \textbf{عدد الملفات} & \textbf{نوعية الملفات} \\
\midrule
"""
        cursor.execute("SELECT * FROM dossiers WHERE TRIM(LOWER(category)) IN ('homolog', 'export', 'conformite')")
        dossiers_current = cursor.fetchall()
        total_dossiers = sum([d[5] for d in dossiers_current])
        for d in dossiers_current:
            percentage = (d[5] / total_dossiers * 100) if total_dossiers else 0
            latex += f"{self.escape_latex(f'{percentage:.1f}')}٪ & {self.escape_latex(f'{d[5]:,}')} & {self.escape_latex(d[0])} \\\\\n"
        latex += f"{self.escape_latex('100.0')}٪ & {self.escape_latex(f'{total_dossiers:,}')} & المجموع \\\\\n"
        latex += r"""
\bottomrule
\end{tabular}

\subsection*{2. عمليات المراقبة الفنية}
\begin{tabular}{ccc}
\toprule
\textbf{النسبة المئوية} & \textbf{عدد الملفات} & \textbf{نوعية الملفات} \\
\midrule
"""
        cursor.execute("SELECT * FROM dossiers WHERE TRIM(LOWER(category)) = 'con, tech'")
        con_tech = cursor.fetchone()
        total_con_tech = con_tech[5] if con_tech else 0
        latex += f"{self.escape_latex('98.8')}٪ & {self.escape_latex('245')} & المراقبة الفنية \\\\\n{self.escape_latex('1.2')}٪ & {self.escape_latex('3')} & المراقبة الفنية تحت الديوانة \\\\\n{self.escape_latex('100.0')}٪ & {self.escape_latex(f'{total_con_tech:,}')} & المجموع \\\\\n"
        latex += r"""
\bottomrule
\end{tabular}

\section*{VII. لوحة قيادة لعدد العمليات المنجزة منذ بداية السنة}
\subsection*{1. عمليات المصادقة والمطابقة}
\begin{tabular}{ccc}
\toprule
\textbf{النسبة المئوية} & \textbf{عدد الملفات} & \textbf{نوعية الملفات} \\
\midrule
"""
        cursor.execute("SELECT * FROM dossiers WHERE TRIM(LOWER(category)) IN ('homolog', 'export', 'conformite')")
        dossiers_year = cursor.fetchall()
        total_dossiers_year = sum([d[13] for d in dossiers_year])
        for d in dossiers_year:
            percentage = (d[13] / total_dossiers_year * 100) if total_dossiers_year else 0
            latex += f"{self.escape_latex(f'{percentage:.1f}')}٪ & {self.escape_latex(f'{d[13]:,}')} & {self.escape_latex(d[0])} \\\\\n"
        latex += f"{self.escape_latex(' Dental0')}٪ & {self.escape_latex(f'{total_dossiers_year:,}')} & المجموع \\\\\n"
        latex += r"""
\bottomrule
\end{tabular}

\subsection*{2. عمليات المراقبة الفنية}
\begin{tabular}{ccc}
\toprule
\textbf{النسبة المئوية} & \textbf{عدد الملفات} & \textbf{نوعية الملفات} \\
\midrule
"""
        cursor.execute("SELECT * FROM dossiers WHERE TRIM(LOWER(category)) = 'con, tech'")
        con_tech_year = cursor.fetchone()
        latex += f"{self.escape_latex('96.2')}٪ & {self.escape_latex('944')} & المراقبة الفنية \\\\\n{self.escape_latex('3.8')}٪ & {self.escape_latex('37')} & المراقبة الفنية تحت الديوانة \\\\\n{self.escape_latex('100.0')}٪ & {self.escape_latex(f'{con_tech_year[13]:,}' if con_tech_year else '0')} & المجموع \\\\\n"
        latex += r"""
\bottomrule
\end{tabular}

\section*{VIII. إحصائيات معالجة الملفات}
\subsection*{1. إحصائيات معالجة ملفات المصادقة}
\begin{itemize}
    \item مكتمل: """ + f"{self.escape_latex(f'{completion_stats[0][1]:.1f}')}٪" + r"""
    \item غير مكتمل: """ + f"{self.escape_latex(f'{completion_stats[0][2]:.1f}')}٪" + r"""
\end{itemize}
\begin{itemize}
    \item نقص وثائق خاصيات فنية: """ + f"{self.escape_latex(f'{intervention_reasons[0][1]:.1f}')}٪" + r"""
    \item تشغيل الجهاز أو نقص لبعض المكملات: """ + f"{self.escape_latex(f'{intervention_reasons[0][2]:.1f}')}٪" + r"""
    \item أسباب مختلفة: """ + f"{self.escape_latex(f'{intervention_reasons[0][3]:.1f}')}٪" + r"""
\end{itemize}

\subsection*{2. إحصائيات معالجة ملفات المطابقة}
\begin{itemize}
    \item مكتمل: """ + f"{self.escape_latex(f'{completion_stats[1][1]:.1f}')}٪" + r"""
    \item غير مكتمل: """ + f"{self.escape_latex(f'{completion_stats[1][2]:.1f}')}٪" + r"""
\end{itemize}
\begin{itemize}
    \item نقص وثائق خاصيات فنية: """ + f"{self.escape_latex(f'{intervention_reasons[1][1]:.1f}')}٪" + r"""
    \item تشغيل الجهاز أو نقص لبعض المكملات: """ + f"{self.escape_latex(f'{intervention_reasons[1][2]:.1f}')}٪" + r"""
    \item أسباب مختلفة: """ + f"{self.escape_latex(f'{intervention_reasons[1][3]:.1f}')}٪" + r"""
\end{itemize}

\section*{IX. الموارد البشرية}
\begin{tabular}{cc}
\toprule
\textbf{عدد الأعوان} & \textbf{نوع النشاط} \\
\midrule
15 & المصادقة والمراقبة الفنية \\
19 & المجموع باعتبار المسؤولين والكتابة \\
\bottomrule
\end{tabular}

\section*{X. الاجتماعات والأنشطة المختلفة}
\begin{itemize}
    \item اجتماع داخلي يوم 30 ماي 2025
\end{itemize}

\end{document}
"""
        return latex

    def __del__(self):
        self.conn.close()

if __name__ == "__main__":
    root = tk.Tk()
    app = RapportGeneratorApp(root)
    root.mainloop()