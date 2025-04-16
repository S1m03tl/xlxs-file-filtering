import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.platypus import Table, TableStyle
from reportlab.lib import colors
from pandastable import Table as PandasTable
from tkinter import Frame
import webbrowser
import os

# Variables globales
df = None
filtered_df = None

class FilterWindow(tk.Toplevel):
    def __init__(self, parent, dataframe):
        global filtered_df
        super().__init__(parent)
        self.title("Filtrer les Données")
        self.geometry("1000x700")
        self.configure(bg="#f5f5f5")
        self.dataframe = dataframe
        
        # Style
        self.style = ttk.Style()
        self.style.configure('TFrame', background='#f5f5f5')
        self.style.configure('TLabel', background='#f5f5f5', font=('Arial', 10))
        self.style.configure('TButton', font=('Arial', 10), padding=5)
        
        # Frame principale
        main_frame = ttk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Frame de recherche rapide
        quick_search_frame = ttk.LabelFrame(main_frame, text="Recherche Rapide", padding=10)
        quick_search_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(quick_search_frame, text="Rechercher dans toutes les colonnes:").pack(side=tk.LEFT)
        self.quick_search_var = tk.StringVar()
        quick_search_entry = ttk.Entry(quick_search_frame, textvariable=self.quick_search_var, width=40)
        quick_search_entry.pack(side=tk.LEFT, padx=5)
        quick_search_entry.bind('<KeyRelease>', self.quick_search)
        
        # Frame de filtres avancés
        filter_frame = ttk.LabelFrame(main_frame, text="Filtres Avancés", padding=10)
        filter_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Sélection de colonne
        ttk.Label(filter_frame, text="Colonne:").grid(row=0, column=0, padx=5, pady=5)
        self.column_var = tk.StringVar()
        self.column_dropdown = ttk.Combobox(filter_frame, textvariable=self.column_var, 
                                          values=list(dataframe.columns), state="readonly")
        self.column_dropdown.grid(row=0, column=1, padx=5, pady=5)
        self.column_dropdown.bind("<<ComboboxSelected>>", self.update_filter_options)
        
        # Opérateur
        ttk.Label(filter_frame, text="Opérateur:").grid(row=0, column=2, padx=5, pady=5)
        self.operator_var = tk.StringVar()
        self.operator_dropdown = ttk.Combobox(filter_frame, textvariable=self.operator_var, 
                                             values=[], state="readonly")
        self.operator_dropdown.grid(row=0, column=3, padx=5, pady=5)
        
        # Valeur
        ttk.Label(filter_frame, text="Valeur:").grid(row=0, column=4, padx=5, pady=5)
        self.value_entry = ttk.Entry(filter_frame)
        self.value_entry.grid(row=0, column=5, padx=5, pady=5)
        
        # Boutons
        btn_frame = ttk.Frame(filter_frame)
        btn_frame.grid(row=1, column=0, columnspan=6, pady=10)
        
        ttk.Button(btn_frame, text="Ajouter Filtre", command=self.add_filter).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Appliquer Filtres", command=self.apply_filters).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Réinitialiser", command=self.reset_filters).pack(side=tk.LEFT, padx=5)
        
        # Liste des filtres
        self.filters_listbox = tk.Listbox(filter_frame, height=4)
        self.filters_listbox.grid(row=2, column=0, columnspan=6, sticky='ew', pady=(0, 10))
        
        # Tableau de données
        table_frame = ttk.Frame(main_frame)
        table_frame.pack(fill=tk.BOTH, expand=True)
        
        self.table = PandasTable(table_frame, dataframe=dataframe)
        self.table.show()
        
        # Boutons d'export
        export_frame = ttk.Frame(main_frame)
        export_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Button(export_frame, text="Télécharger PDF", command=self.download_pdf).pack(side=tk.LEFT, padx=5)
        ttk.Button(export_frame, text="Fermer", command=self.destroy).pack(side=tk.RIGHT, padx=5)
        
        self.filters = []
    
    def quick_search(self, event=None):
        search_term = self.quick_search_var.get().lower()
        if not search_term:
            self.table.model.df = self.dataframe
            self.table.redraw()
            return
        
        try:
            # Recherche dans toutes les colonnes de type string
            mask = pd.concat([self.dataframe[col].astype(str).str.lower().str.contains(search_term) 
                            for col in self.dataframe.columns 
                            if pd.api.types.is_string_dtype(self.dataframe[col])], axis=1).any(axis=1)
            filtered = self.dataframe[mask]
            self.table.model.df = filtered
            self.table.redraw()
            
            filtered_df = filtered
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de la recherche:\n{str(e)}")
    
    def update_filter_options(self, event=None):
        col = self.column_var.get()
        if col in self.dataframe.columns:
            dtype = self.dataframe[col].dtype
            if pd.api.types.is_numeric_dtype(dtype):
                operators = ["==", "!=", ">", "<", ">=", "<="]
            else:
                operators = ["==", "!=", "contient", "ne contient pas"]
            self.operator_dropdown['values'] = operators
            if operators:
                self.operator_dropdown.current(0)
    
    def add_filter(self):
        col = self.column_var.get()
        operator = self.operator_var.get()
        value = self.value_entry.get()
        
        if not col or not operator:
            messagebox.showwarning("Attention", "Veuillez sélectionner une colonne et un opérateur")
            return
        
        filter_str = f"{col} {operator} {value}"
        self.filters_listbox.insert(tk.END, filter_str)
        self.filters.append((col, operator, value))
        
        # Effacer les champs
        self.value_entry.delete(0, tk.END)
    
    def apply_filters(self):
        global filtered_df
        if not self.filters and not self.quick_search_var.get():
            self.table.model.df = self.dataframe
            self.table.redraw()
            
            filtered_df = None
            return
        
        filtered = self.dataframe.copy()
        
        # Appliquer les filtres avancés
        for col, operator, value in self.filters:
            try:
                if operator == "==":
                    if pd.api.types.is_numeric_dtype(filtered[col]):
                        filtered = filtered[filtered[col] == float(value)]
                    else:
                        filtered = filtered[filtered[col] == value]
                elif operator == "!=":
                    if pd.api.types.is_numeric_dtype(filtered[col]):
                        filtered = filtered[filtered[col] != float(value)]
                    else:
                        filtered = filtered[filtered[col] != value]
                elif operator == ">":
                    filtered = filtered[filtered[col] > float(value)]
                elif operator == "<":
                    filtered = filtered[filtered[col] < float(value)]
                elif operator == ">=":
                    filtered = filtered[filtered[col] >= float(value)]
                elif operator == "<=":
                    filtered = filtered[filtered[col] <= float(value)]
                elif operator == "contient":
                    filtered = filtered[filtered[col].astype(str).str.contains(value, case=False)]
                elif operator == "ne contient pas":
                    filtered = filtered[~filtered[col].astype(str).str.contains(value, case=False)]
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur lors de l'application du filtre: {str(e)}")
                return
        
        self.table.model.df = filtered
        self.table.redraw()
        filtered_df = filtered
    
    def reset_filters(self):
        self.filters = []
        self.filters_listbox.delete(0, tk.END)
        self.quick_search_var.set("")
        self.table.model.df = self.dataframe
        self.table.redraw()
        global filtered_df
        filtered_df = None
    
    def download_pdf(self):
        data = self.table.model.df
        if data.empty:
            messagebox.showwarning("Attention", "Aucune donnée à exporter")
            return
        
        try:
            # Créer un dossier temporaire si nécessaire
            if not os.path.exists("temp"):
                os.makedirs("temp")
            
            pdf_path = os.path.join("temp", "rapport_filtre.pdf")
            generate_pdf(data, pdf_path)
            
            # Ouvrir le PDF dans le navigateur par défaut
            webbrowser.open(pdf_path)
            messagebox.showinfo("Succès", f"PDF généré avec succès:\n{pdf_path}")
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors de la génération du PDF:\n{str(e)}")

def generate_pdf(data, filename="rapport.pdf"):
    try:
        c = canvas.Canvas(filename, pagesize=letter)
        width, height = letter
        
        # Logo RMA (placeholder - remplacer par votre propre logo)
        logo_path = "logo_RMA.png"  # Remplacez par le chemin de votre logo
        if os.path.exists(logo_path):
            c.drawImage(logo_path, 50, height - 100, width=100, preserveAspectRatio=True)
        
        # En-tête
        c.setFont("Helvetica-Bold", 16)
        c.drawString(170, height - 70, "Rapport RMA")
        c.setFont("Helvetica", 10)
        c.drawString(170, height - 90, f"Nombre d'enregistrements: {len(data)} | Généré le: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M')}")
        
        # Ligne de séparation
        c.line(50, height - 110, width - 50, height - 110)
        
        # Tableau de données
        data_list = [data.columns.tolist()] + data.values.tolist()
        
        table = Table(data_list)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1a5fb4')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#f5f5f5')),
            ('GRID', (0, 0), (-1, -1), 1, colors.lightgrey),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
        ]))
        
        table.wrapOn(c, width - 100, height)
        table.drawOn(c, 50, height - 130 - (20 * min(20, len(data))))
        
        # Pied de page
        c.setFont("Helvetica", 8)
        c.drawString(50, 30, "Rapport Généré Automatiquement")
        c.drawRightString(width - 50, 30, f"Page 1")
        
        c.save()
    except Exception as e:
        messagebox.showerror("Erreur PDF", f"Erreur lors de la génération du PDF:\n{str(e)}")
        raise

def load_file():
    global df, filtered_df
    file_path = filedialog.askopenfilename(
        title="Sélectionner un fichier Excel", 
        filetypes=[("Excel files", "*.xlsx;*.xls"), ("All files", "*.*")]
    )
    if file_path:
        try:
            df = pd.read_excel(file_path)
            filtered_df = None
            messagebox.showinfo("Succès", f"Fichier chargé avec succès!\n{len(df)} lignes chargées.")
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de charger le fichier:\n{str(e)}")

def open_filter_window():
    if df is None:
        messagebox.showerror("Erreur", "Veuillez d'abord charger un fichier Excel.")
        return
    FilterWindow(root, df)

# Interface Tkinter améliorée
root = tk.Tk()
root.title("Générateur de Rapport")
root.geometry("900x600")
root.configure(bg="#f5f5f5")

# Style
style = ttk.Style()
style.theme_use('clam')
style.configure('TFrame', background='#f5f5f5')
style.configure('TButton', font=('Arial', 10), padding=8)
style.configure('TLabel', background='#f5f5f5', font=('Arial', 10))
style.configure('Accent.TButton', font=('Arial', 12, 'bold'), foreground='white', background='#1a5fb4')
style.map('Accent.TButton', background=[('pressed', '#1a5fb4'), ('active', '#1a5fb4')])

# Barre d'en-tête
header_frame = tk.Frame(root, bg="#1a5fb4", height=100)
header_frame.pack(fill=tk.X)

# Logo RMA (placeholder)
logo_label = tk.Label(header_frame, text="RMA", fg="white", bg="#1a5fb4", font=("Arial", 24, "bold"))
logo_label.pack(side=tk.LEFT, padx=20)

label_title = tk.Label(header_frame, 
                      text="Générateur de Rapport ", 
                      fg="white", 
                      bg="#1a5fb4", 
                      font=("Arial", 16))
label_title.pack(side=tk.LEFT, padx=10, pady=30)

# Contenu principal
main_frame = ttk.Frame(root)
main_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=20)

# Section de chargement
load_frame = ttk.LabelFrame(main_frame, text="Chargement des Données", padding=15)
load_frame.pack(fill=tk.X, pady=(0, 20))

btn_load = ttk.Button(load_frame, 
                     text="📂 Charger un fichier Excel", 
                     command=load_file,
                     style='Accent.TButton')
btn_load.pack(side=tk.LEFT, padx=10)

btn_filter = ttk.Button(load_frame, 
                       text="🔍 Ouvrir l'outil de filtrage", 
                       command=open_filter_window,
                       style='Accent.TButton')
btn_filter.pack(side=tk.LEFT, padx=10)

# Instructions
instructions = tk.Label(main_frame, 
                       text="1. Chargez un fichier Excel\n2. Utilisez l'outil de filtrage pour sélectionner vos données\n3. Téléchargez le PDF directement depuis l'outil",
                       bg="#f5f5f5",
                       justify=tk.LEFT,
                       font=("Arial", 10))
instructions.pack(fill=tk.X, pady=(10, 20))

# Pied de page
footer_frame = tk.Frame(root, bg="#1a5fb4", height=40)
footer_frame.pack(fill=tk.X, side=tk.BOTTOM)

footer_label = tk.Label(footer_frame, 
                       text="©  Tous droits réservés", 
                       fg="white", 
                       bg="#1a5fb4",
                       font=("Arial", 9))
footer_label.pack(pady=10)

root.mainloop()