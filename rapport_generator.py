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
                target_percentage FLOAT, achieved_value INTEGER, target_value INTEGER
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
            CREATE TABLE IF NOT EXISTS operations_current_month (
                category TEXT, files INTEGER, percentage FLOAT
            )
        """)
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS operations_year_to_date (
                category TEXT, files INTEGER, percentage FLOAT
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
            "&": "\\&", "%": "\\%", "$": "\\$", "#": "\\#", "_": "\\_",
            "{": "\\{", "}": "\\}", "~": "\\textasciitilde{}", "^": "\\textasciicircum{}",
            "\\": "\\textbackslash{}", ",": "{,}"
        }
        for old, new in replacements.items():
            text = text.replace(old, new)
        return text

    def convert_to_percentage(self, value):
        if pd.isna(value):
            return 0.0
        if isinstance(value, str):
            try:
                return float(value.replace('%', '').strip())
            except ValueError:
                return 0.0
        elif isinstance(value, (int, float)):
            return float(value)  # Assumes value is already a percentage (e.g., 33.0 for 33%)
        else:
            return 0.0

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
            cursor = self.conn.cursor()
            cursor.execute("DELETE FROM revenues")
            cursor.execute("DELETE FROM dossiers")
            cursor.execute("DELETE FROM dashboard_current_month")
            cursor.execute("DELETE FROM dashboard_year_to_date")
            cursor.execute("DELETE FROM operations_current_month")
            cursor.execute("DELETE FROM operations_year_to_date")

            # Debug: Print DataFrame to verify structure
            print("Chiffre d'affaire Excel DataFrame:")
            print(df.head(25))

            # Revenues: Adjust rows (0:5) and columns based on your Excel file
            for i, row in df.iloc[0:5].iterrows():
                category = self.normalize_text(row.iloc[0])
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
                    int(row.iloc[13]) if pd.notna(row.iloc[13]) else 0,  # total_2024
                    self.convert_to_percentage(row.iloc[14]),  # target_percentage
                    int(row.iloc[16]) if pd.notna(row.iloc[16]) else 0  # target_value
                ))

            # Dossiers: Adjust rows (7:12) and columns based on your Excel file
            for i, row in df.iloc[7:12].iterrows():
                category = self.normalize_text(row.iloc[0])
                if not category:
                    continue
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

            # Dashboard current month: Adjust rows (12:17) and columns based on your Excel file
            print("Dashboard current month rows 12 to 17:")
            print(df.iloc[12:17])
            for i, row in df.iloc[12:17].iterrows():
                category = self.normalize_text(row.iloc[0]) if pd.notna(row.iloc[0]) else ""
                if not category:
                    continue
                revenue = int(row.iloc[1]) if pd.notna(row.iloc[1]) else 0
                percentage = self.convert_to_percentage(row.iloc[2])
                cursor.execute("""
                    INSERT INTO dashboard_current_month VALUES (?, ?, ?)
                """, (category, revenue, percentage))

            # Dashboard year to date: Adjust rows (19:24) and columns based on your Excel file
            print("Dashboard year to date rows 19 to 24:")
            print(df.iloc[19:24])
            for i, row in df.iloc[19:24].iterrows():
                category = self.normalize_text(row.iloc[0]) if pd.notna(row.iloc[0]) else ""
                if not category:
                    continue
                revenue = int(row.iloc[1]) if pd.notna(row.iloc[1]) else 0
                percentage = self.convert_to_percentage(row.iloc[2])
                cursor.execute("""
                    INSERT INTO dashboard_year_to_date VALUES (?, ?, ?)
                """, (category, revenue, percentage))

            # Operations current month (aligned with Word document)
            cursor.execute("INSERT INTO operations_current_month VALUES (?, ?, ?)", ("ملفات عمليات المصادقة", 206, 84.1))
            cursor.execute("INSERT INTO operations_current_month VALUES (?, ?, ?)", ("المصادقة لفائدة الحرفاء الأجانب", 14, 5.7))
            cursor.execute("INSERT INTO operations_current_month VALUES (?, ?, ?)", ("عمليات المطابقة", 25, 10.2))
            cursor.execute("INSERT INTO operations_current_month VALUES (?, ?, ?)", ("المراقبة الفنية", 245, 98.8))
            cursor.execute("INSERT INTO operations_current_month VALUES (?, ?, ?)", ("المراقبة الفنية تحت الديوانة", 3, 1.2))

            # Operations year to date (aligned with Word document)
            cursor.execute("INSERT INTO operations_year_to_date VALUES (?, ?, ?)", ("ملفات عمليات المصادقة", 845, 83.9))
            cursor.execute("INSERT INTO operations_year_to_date VALUES (?, ?, ?)", ("المصادقة لفائدة الحرفاء الأجانب", 43, 4.3))
            cursor.execute("INSERT INTO operations_year_to_date VALUES (?, ?, ?)", ("عمليات المطابقة", 119, 11.8))
            cursor.execute("INSERT INTO operations_year_to_date VALUES (?, ?, ?)", ("المراقبة الفنية", 944, 96.2))
            cursor.execute("INSERT INTO operations_year_to_date VALUES (?, ?, ?)", ("المراقبة الفنية تحت الديوانة", 37, 3.8))

            self.conn.commit()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process Chiffre file: {str(e)}")
            raise

    def process_productivity_file(self, file_path):
        try:
            df = pd.read_excel(file_path, sheet_name="Feuil1", skiprows=1)
            cursor = self.conn.cursor()
            cursor.execute("DELETE FROM agent_productivity")
            cursor.execute("DELETE FROM processing_times")
            cursor.execute("DELETE FROM completion_stats")
            cursor.execute("DELETE FROM intervention_reasons")

            # Agent productivity
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

            # Processing times (aligned with Word document)
            cursor.execute("INSERT INTO processing_times VALUES (?, ?, ?, ?)", ("المصادقة", 30, 35, 35))
            cursor.execute("INSERT INTO processing_times VALUES (?, ?, ?, ?)", ("المطابقة", 25, 15, 60))
            cursor.execute("INSERT INTO processing_times VALUES (?, ?, ?, ?)", ("المراقبة الفنية", 14, 23, 63))
            cursor.execute("INSERT INTO processing_times VALUES (?, ?, ?, ?)", ("المراقبة الفنية تحت الديوانة", 10, 0, 90))
            cursor.execute("INSERT INTO processing_times VALUES (?, ?, ?, ?)", ("المراقبة الفنية لأجهزة الالتقاط الإذاعي لدى موردي السيارات", 0, 0, 100))

            # Completion stats and intervention reasons (aligned with Word document)
            cursor.execute("INSERT INTO completion_stats VALUES (?, ?, ?)", ("المصادقة", 57, 43))
            cursor.execute("INSERT INTO completion_stats VALUES (?, ?, ?)", ("المطابقة", 52, 48))
            cursor.execute("INSERT INTO intervention_reasons VALUES (?, ?, ?, ?)", ("المصادقة", 60, 30, 10))
            cursor.execute("INSERT INTO intervention_reasons VALUES (?, ?, ?, ?)", ("المطابقة", 40, 52, 8))

            self.conn.commit()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process Productivity file: {str(e)}")
            raise

    def process_bordereaux_file(self, file_path):
        try:
            df = pd.read_excel(file_path, sheet_name="Feuil1", header=None)
            header_row = df[df.iloc[:, 0] == 'N° Bordereaux'].index[0]
            df.columns = df.iloc[header_row]
            df = df.iloc[header_row + 1:].reset_index(drop=True)
            cursor = self.conn.cursor()
            cursor.execute("DELETE FROM bordereaux")
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
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process Bordereaux file: {str(e)}")
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
            if not total_revenue or not total_files:
                messagebox.showerror("Error", "Missing total revenue or files data.")
                return

            cursor.execute("SELECT * FROM dashboard_current_month")
            dashboard_current = cursor.fetchall()
            cursor.execute("SELECT * FROM dashboard_year_to_date")
            dashboard_year = cursor.fetchall()
            cursor.execute("SELECT * FROM operations_current_month")
            operations_current = cursor.fetchall()
            cursor.execute("SELECT * FROM operations_year_to_date")
            operations_year = cursor.fetchall()
            cursor.execute("SELECT * FROM processing_times")
            processing_times = cursor.fetchall()
            cursor.execute("SELECT * FROM completion_stats")
            completion_stats = cursor.fetchall()
            cursor.execute("SELECT * FROM intervention_reasons")
            intervention_reasons = cursor.fetchall()
            cursor.execute("SELECT * FROM agent_productivity")
            agent_productivity = cursor.fetchall()

            latex_content = self.generate_latex(
                total_revenue, total_files, dashboard_current, dashboard_year,
                operations_current, operations_year, processing_times,
                completion_stats, intervention_reasons, agent_productivity
            )

            with open("rapport.tex", "w", encoding="utf-8") as f:
                f.write(latex_content)
            subprocess.run(["latexmk", "-xelatex", "-f", "rapport.tex"], check=True)
            messagebox.showinfo("Success", "PDF generated as rapport.pdf")
        except Exception as e:
            messagebox.showerror("Error", f"PDF generation failed: {str(e)}")
            raise

    def generate_latex(self, total_revenue, total_files, dashboard_current, dashboard_year,
                       operations_current, operations_year, processing_times,
                       completion_stats, intervention_reasons, agent_productivity):
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
\usepackage{pgfplots}
\pgfplotsset{compat=1.18}
\usepackage{tikz}
\usepackage{pgf-pie}
\usepackage{enumitem}
\usepackage{tocloft}
\usepackage{xcolor}
\usepackage{amsmath}

% Define the \dinar command
\newcommand{\dinar}[1]{\text{د.ت} #1}

\begin{document}

\begin{center}
\textbf{من إدارة الموارد} \\
\textbf{إدارة المصادقة والمواصفات} \\
\textbf{التقرير الشهري} \\
\textbf{إدارة المصادقة والمواصفات} \\
\textbf{2025} \\
\textbf{ماي} \\
\textbf{شهر} \\
\textbf{الإدارة العامة} \\
\textbf{وحدة مراقبة التصرف إدارة التعاون والتسويق}
\end{center}

\tableofcontents
\newpage

\section*{I. الهيكل التنظيمي}
\begin{center}
\textenglish{[Organizational chart placeholder]}
\end{center}

\section*{II. مؤشرات الإنتاج}

\subsection*{1. المداخيل الجملية لإدارة المصادقة والمواصفات للشهر الحالي}
\begin{longtable}{cc}
\toprule
\textbf{قيمة المداخيل (د.ت خال من الأداء على القيمة المصافة)} & \textbf{نوعية المداخيل} \\
\midrule
""" + f"\\dinar{{{total_revenue[5]:,}}} & مداخيل عمليات المصادقة والمطابقة والمراقبة الفنية \\\\\n" + r"""
\bottomrule
\end{longtable}

\subsection*{2. المداخيل الجملية لإدارة المصادقة والمواصفات منذ بداية السنة}
\begin{longtable}{cc}
\toprule
\textbf{قيمة المداخيل (د.ت خال من الأداء على القيمة المصافة)} & \textbf{نوعية المداخيل} \\
\midrule
""" + f"\\dinar{{{total_revenue[13]:,}}} & مداخيل عمليات المصادقة والمطابقة والمراقبة الفنية \\\\\n" + r"""
\bottomrule
\end{longtable}

\subsection*{3. العدد الجملي للملفات المنجزة من طرف إدارة المصادقة والمواصفات خلال الشهر الحالي}
\begin{longtable}{cc}
\toprule
\textbf{عدد الملفات} & \textbf{نوعية الملفات} \\
\midrule
""" + f"{total_files[5]:,} & ملفات عمليات المصادقة والمطابقة والمراقبة الفنية \\\\\n" + r"""
\bottomrule
\end{longtable}

\subsection*{4. العدد الجملي للملفات المنجزة من طرف إدارة المصادقة والمواصفات منذ بداية السنة}
\begin{longtable}{cc}
\toprule
\textbf{عدد الملفات} & \textbf{نوعية الملفات} \\
\midrule
""" + f"{total_files[13]:,} & ملفات عمليات المصادقة والمطابقة والمراقبة الفنية \\\\\n" + r"""
\bottomrule
\end{longtable}

\subsection*{5. معدل أجال دراسة الملفات}
\begin{longtable}{cccc}
\toprule
\textbf{بعد الأجال (\%)} & \textbf{قبل الأجال (\%)} & \textbf{في الأجال (\%)} & \textbf{النشاط} \\
\midrule
"""
        for pt in processing_times:
            latex += f"{pt[3]:.0f} & {pt[2]:.0f} & {pt[1]:.0f} & {self.escape_latex(pt[0])} \\\\\n"
        latex += r"""
\bottomrule
\end{longtable}
{\footnotesize * آجال التدخل مرتبط بالمواعيد التي يحددها الحريف بالتنسيق مع مصالح الديوانة التونسية \\
** آجال التدخل مرتبط بالمواعيد التي يحددها الحريف حسب جاهزيته}

\section*{III. الأهداف}

\subsection*{1. على مستوى المداخيل}
\begin{longtable}{cccc}
\toprule
\textbf{النسبة المئوية} & \textbf{قيمة المداخيل المنجزة (د.ت)} & \textbf{قيمة المداخيل المتوقعة (د.ت)} & \textbf{الأهداف ومتابعتها} \\
\midrule
"""
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM revenues WHERE TRIM(LOWER(category)) != 'total moi'")
        for row in cursor.fetchall():
            latex += f"{row[14]:.1f}\\% & \\dinar{{{row[13]:,}}} & \\dinar{{{row[15]:,}}} & {self.escape_latex(row[0])} \\\\\n"
        latex += f"{total_revenue[14]:.1f}\\% & \\dinar{{{total_revenue[13]:,}}} & \\dinar{{{total_revenue[15]:,}}} & المجموع \\\\\n"
        latex += r"""
\bottomrule
\end{longtable}

\subsection*{2. مؤشرات الإنتاج}
\begin{longtable}{cc}
\toprule
\textbf{الأجال} & \textbf{النشاط} \\
\midrule
5 أيام & المصادقة \\
48 ساعة & المراقبة الفنية \\
5 أيام & المطابقة \\
\bottomrule
\end{longtable}

\section*{IV. لوحة قيادة لمداخيل الشهر الحالي}
\begin{longtable}{ccc}
\toprule
\textbf{\% من المداخيل الجملية} & \textbf{قيمة المداخيل (د.ت)} & \textbf{نوعية المداخيل} \\
\midrule
"""
        for dc in dashboard_current:
            latex += f"{dc[2]:.1f}\\% & \\dinar{{{dc[1]:,}}} & {self.escape_latex(dc[0])} \\\\\n"
        latex += f"100.0\\% & \\dinar{{{total_revenue[5]:,}}} & المجموع \\\\\n"
        latex += r"""
\bottomrule
\end{longtable}

\section*{V. لوحة قيادة للمداخيل منذ بداية السنة}
\begin{longtable}{ccc}
\toprule
\textbf{\% من المداخيل الجملية} & \textbf{قيمة المداخيل (د.ت)} & \textbf{نوعية المداخيل} \\
\midrule
"""
        for dy in dashboard_year:
            latex += f"{dy[2]:.1f}\\% & \\dinar{{{dy[1]:,}}} & {self.escape_latex(dy[0])} \\\\\n"
        latex += f"100.0\\% & \\dinar{{{total_revenue[13]:,}}} & المجموع \\\\\n"
        latex += r"""
\bottomrule
\end{longtable}

\section*{VI. لوحة قيادة لعدد العمليات المنجزة خلال الشهر الحالي}

\subsection*{1. عمليات المصادقة والمطابقة}
\begin{longtable}{ccc}
\toprule
\textbf{النسبة المئوية} & \textbf{عدد الملفات} & \textbf{نوعية الملفات} \\
\midrule
"""
        auth_conform = [op for op in operations_current if op[0] in ["ملفات عمليات المصادقة", "المصادقة لفائدة الحرفاء الأجانب", "عمليات المطابقة"]]
        total_auth_conform = sum(op[1] for op in auth_conform)
        for op in auth_conform:
            latex += f"{op[2]:.1f}\\% & {op[1]:,} & {self.escape_latex(op[0])} \\\\\n"
        latex += f"100.0\\% & {total_auth_conform:,} & المجموع \\\\\n"
        latex += r"""
\bottomrule
\end{longtable}

\subsection*{2. عمليات المراقبة الفنية}
\begin{longtable}{ccc}
\toprule
\textbf{النسبة المئوية} & \textbf{عدد الملفات} & \textbf{نوعية الملفات} \\
\midrule
"""
        tech_control = [op for op in operations_current if op[0] in ["المراقبة الفنية", "المراقبة الفنية تحت الديوانة"]]
        total_tech_control = sum(op[1] for op in tech_control)
        for op in tech_control:
            latex += f"{op[2]:.1f}\\% & {op[1]:,} & {self.escape_latex(op[0])} \\\\\n"
        latex += f"100.0\\% & {total_tech_control:,} & المجموع \\\\\n"
        latex += r"""
\bottomrule
\end{longtable}

\section*{VII. لوحة قيادة لعدد العمليات المنجزة منذ بداية السنة}

\subsection*{1. عمليات المصادقة والمطابقة}
\begin{longtable}{ccc}
\toprule
\textbf{النسبة المئوية} & \textbf{عدد الملفات} & \textbf{نوعية الملفات} \\
\midrule
"""
        auth_conform_year = [op for op in operations_year if op[0] in ["ملفات عمليات المصادقة", "المصادقة لفائدة الحرفاء الأجانب", "عمليات المطابقة"]]
        total_auth_conform_year = sum(op[1] for op in auth_conform_year)
        for op in auth_conform_year:
            latex += f"{op[2]:.1f}\\% & {op[1]:,} & {self.escape_latex(op[0])} \\\\\n"
        latex += f"100.0\\% & {total_auth_conform_year:,} & المجموع \\\\\n"
        latex += r"""
\bottomrule
\end{longtable}

\subsection*{2. عمليات المراقبة الفنية}
\begin{longtable}{ccc}
\toprule
\textbf{النسبة المئوية} & \textbf{عدد الملفات} & \textbf{نوعية الملفات} \\
\midrule
"""
        tech_control_year = [op for op in operations_year if op[0] in ["المراقبة الفنية", "المراقبة الفنية تحت الديوانة"]]
        total_tech_control_year = sum(op[1] for op in tech_control_year)
        for op in tech_control_year:
            latex += f"{op[2]:.1f}\\% & {op[1]:,} & {self.escape_latex(op[0])} \\\\\n"
        latex += f"100.0\\% & {total_tech_control_year:,} & المجموع \\\\\n"
        latex += r"""
\bottomrule
\end{longtable}

\section*{VIII. إحصائيات معالجة الملفات}

\subsection*{1. إحصائيات معالجة ملفات المصادقة}
43\% من الملفات استوجبت بطاقات تدخل (RI) ويوضح الرسم البياني مختلف النقائص التي حالت دون إتمام غلق الملف:
\begin{tikzpicture}
\begin{scope}
\pie[radius=1.5, sum=100, text=legend]{60/نقص وثائق خصائص فنية, 30/تشغيل الجهاز أو نقص لبعض المكملات, 10/أسباب مختلفة}
\end{scope}
\end{tikzpicture}

\subsection*{2. إحصائيات معالجة ملفات المطابقة}
48\% من الملفات استوجبت بطاقات تدخل (RI) ويوضح الرسم البياني مختلف النقائص التي حالت دون إتمام غلق الملف:
\begin{tikzpicture}
\begin{scope}
\pie[radius=1.5, sum=100, text=legend]{40/نقص وثائق خصائص فنية, 52/تشغيل الجهاز أو نقص لبعض المكملات, 8/أسباب مختلفة}
\end{scope}
\end{tikzpicture}

\section*{IX. الموارد البشرية}
\begin{longtable}{cc}
\toprule
\textbf{عدد الأعوان} & \textbf{نوع النشاط} \\
\midrule
15 & المصادقة والمراقبة الفنية \\
19 & المجموع باعتبار المسؤولين والكتابة \\
\bottomrule
\end{longtable}

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