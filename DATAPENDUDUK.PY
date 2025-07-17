import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from datetime import datetime
import os
from ttkthemes import ThemedTk
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import numpy as np
from PIL import Image, ImageTk


class DataPendudukApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Sistem Informasi Kependudukan Kecamatan")
        self.root.geometry("1200x700")
        
        # Data storage
        self.data_penduduk = pd.DataFrame(columns=[
            "No KK", "Nama Kepala Keluarga", "Desa", "RT/RW", "Jumlah Anggota"
        ])
        self.data_anggota = pd.DataFrame(columns=[
            "No KK", "NIK", "Nama", "Hubungan", "Jenis_Kelamin",
            "Tanggal_Lahir", "Umur", "Pendidikan", "Pekerjaan"
        ])
        self.data_kelahiran = pd.DataFrame(columns=[
            "NIK", "Nama_Bayi", "Tanggal_Lahir", "Jenis_Kelamin",
            "Nama_Ayah", "Nama_Ibu", "Desa", "No_KK"
        ])
        self.data_kematian = pd.DataFrame(columns=[
            "NIK", "Nama", "Tanggal_Lahir", "Tanggal_Meninggal",
            "Usia_Meninggal", "Penyebab", "Desa", "No_KK"
        ])
        
        # Excel file path
        self.excel_file = None
        
        # Create main notebook
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Create styles
        self.style = ttk.Style()
        self.style.configure("TFrame", background="#f0f0f0")
        self.style.configure("Header.TLabel", font=("Arial", 16, "bold"), background="#f0f0f0")
        self.style.configure("TButton", font=("Arial", 11))
        self.style.map("Accent.TButton",
                      background=[("active", "#007bff"), ("!active", "#0069d9")],
                      foreground=[("active", "white"), ("!active", "white")])
        
        # Create tabs
        self.create_dashboard_tab()
        self.create_kartu_keluarga_tab()
        self.create_anggota_keluarga_tab()
        self.create_kelahiran_tab()
        self.create_kematian_tab()
        self.create_laporan_tab()
        
        # Create menu
        self.create_menu()
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Siap")
        status_bar = ttk.Label(root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        
    def create_menu(self):
        menubar = tk.Menu(self.root)
        
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Buka Excel", command=self.load_excel)
        file_menu.add_command(label="Simpan Excel", command=self.save_excel)
        file_menu.add_separator()
        file_menu.add_command(label="Keluar", command=self.root.quit)
        menubar.add_cascade(label="File", menu=file_menu)
        
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="Petunjuk", command=self.show_help)
        help_menu.add_command(label="Tentang", command=self.show_about)
        menubar.add_cascade(label="Bantuan", menu=help_menu)
        
        self.root.config(menu=menubar)
    
    def create_dashboard_tab(self):
        dashboard_frame = ttk.Frame(self.notebook)
        self.notebook.add(dashboard_frame, text="Dashboard")
        
        # Header
        header_frame = ttk.Frame(dashboard_frame)
        header_frame.pack(fill="x", padx=20, pady=20)
        
        ttk.Label(header_frame, text="DASHBOARD KEPENDUDUKAN KECAMATAN",
                 style="Header.TLabel").pack()
        
        nama_kecamatan_frame = ttk.Frame(header_frame)
        nama_kecamatan_frame.pack(pady=10)
        ttk.Label(nama_kecamatan_frame, text="Nama Kecamatan:").pack(side="left")
        self.kecamatan_var = tk.StringVar(value="CONTOH")
        kecamatan_entry = ttk.Entry(nama_kecamatan_frame, textvariable=self.kecamatan_var, width=20)
        kecamatan_entry.pack(side="left", padx=5)
        ttk.Button(nama_kecamatan_frame, text="Perbarui", command=self.update_dashboard).pack(side="left", padx=5)
        
        # Main content
        content_frame = ttk.Frame(dashboard_frame)
        content_frame.pack(fill="both", expand=True, padx=20)
        
        # Left frame - stats
        stats_frame = ttk.LabelFrame(content_frame, text="Statistik Penduduk")
        stats_frame.pack(side="left", fill="both", expand=True, padx=5, pady=5)
        
        self.total_penduduk_var = tk.StringVar(value="0")
        self.total_kk_var = tk.StringVar(value="0")
        self.total_laki_var = tk.StringVar(value="0")
        self.total_perempuan_var = tk.StringVar(value="0")
        
        stats_items = [
            ("Total Penduduk:", self.total_penduduk_var),
            ("Total KK:", self.total_kk_var),
            ("Laki-laki:", self.total_laki_var),
            ("Perempuan:", self.total_perempuan_var)
        ]
        
        for i, (label_text, var) in enumerate(stats_items):
            frame = ttk.Frame(stats_frame)
            frame.pack(fill="x", pady=10)
            ttk.Label(frame, text=label_text).pack(side="left", padx=5)
            ttk.Label(frame, textvariable=var, font=("Arial", 14, "bold")).pack(side="left")
        
        # Right frame - charts
        charts_frame = ttk.LabelFrame(content_frame, text="Grafik")
        charts_frame.pack(side="right", fill="both", expand=True, padx=5, pady=5)
        
        # Sample chart (will be updated with real data later)
        self.fig, self.ax = plt.subplots(figsize=(5, 4))
        self.canvas = FigureCanvasTkAgg(self.fig, master=charts_frame)
        self.canvas_widget = self.canvas.get_tk_widget()
        self.canvas_widget.pack(fill="both", expand=True)
        
        # Default chart
        self.ax.bar(['Balita', 'Anak', 'Dewasa', 'Lansia'], [10, 25, 50, 15])
        self.ax.set_title('Distribusi Penduduk berdasarkan Usia')
        self.canvas.draw()
    
    def create_kartu_keluarga_tab(self):
        kk_frame = ttk.Frame(self.notebook)
        self.notebook.add(kk_frame, text="Data KK")
        
        # Form Frame
        form_frame = ttk.LabelFrame(kk_frame, text="Form Data Kartu Keluarga")
        form_frame.pack(fill="x", padx=20, pady=10)
        
        # Create a grid layout for the form
        form_fields = [
            ("No KK:", "no_kk_var"),
            ("Nama Kepala Keluarga:", "nama_kk_var"),
            ("Desa/Kelurahan:", "desa_kk_var"),
            ("RT/RW:", "rtrw_kk_var"),
            ("Jumlah Anggota:", "jumlah_anggota_var")
        ]
        
        # Create form fields
        row = 0
        for label_text, var_name in form_fields:
            ttk.Label(form_frame, text=label_text).grid(row=row, column=0, padx=5, pady=5, sticky="w")
            
            # Create StringVar
            setattr(self, var_name, tk.StringVar())
            
            # Create Entry widget
            ttk.Entry(form_frame, textvariable=getattr(self, var_name), width=30).grid(
                row=row, column=1, padx=5, pady=5, sticky="w")
            
            row += 1
        
        # Create a dropdown for desa selection
        desa_options = ["Desa A", "Desa B", "Desa C", "Desa D", "Desa E"]
        self.desa_kk_var.set(desa_options[0])
        ttk.Combobox(form_frame, textvariable=self.desa_kk_var, values=desa_options, state="readonly").grid(
            row=2, column=1, padx=5, pady=5, sticky="w")
        
        # Buttons
        buttons_frame = ttk.Frame(form_frame)
        buttons_frame.grid(row=row, column=0, columnspan=2, pady=10)
        
        ttk.Button(buttons_frame, text="Simpan", command=self.save_kk).pack(side="left", padx=5)
        ttk.Button(buttons_frame, text="Reset", command=self.reset_kk_form).pack(side="left", padx=5)
        
        # Table Frame
        table_frame = ttk.LabelFrame(kk_frame, text="Data Kartu Keluarga")
        table_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # Create TreeView for KK data
        self.kk_tree = ttk.Treeview(table_frame, columns=("No KK", "Nama KK", "Desa", "RT/RW", "Jumlah Anggota"))
        self.kk_tree.heading("#0", text="No")
        self.kk_tree.heading("No KK", text="No KK")
        self.kk_tree.heading("Nama KK", text="Nama Kepala Keluarga")
        self.kk_tree.heading("Desa", text="Desa/Kelurahan")
        self.kk_tree.heading("RT/RW", text="RT/RW")
        self.kk_tree.heading("Jumlah Anggota", text="Jumlah Anggota")
        
        self.kk_tree.column("#0", width=50)
        self.kk_tree.column("No KK", width=150)
        self.kk_tree.column("Nama KK", width=200)
        self.kk_tree.column("Desa", width=150)
        self.kk_tree.column("RT/RW", width=100)
        self.kk_tree.column("Jumlah Anggota", width=120)
        
        # Add scrollbar to treeview
        kk_scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.kk_tree.yview)
        self.kk_tree.configure(yscrollcommand=kk_scrollbar.set)
        
        kk_scrollbar.pack(side="right", fill="y")
        self.kk_tree.pack(fill="both", expand=True)
        
        # Bind events
        self.kk_tree.bind("<ButtonRelease-1>", self.load_kk_to_form)
    
    def create_anggota_keluarga_tab(self):
        anggota_frame = ttk.Frame(self.notebook)
        self.notebook.add(anggota_frame, text="Anggota Keluarga")
        
        # Form Frame
        form_frame = ttk.LabelFrame(anggota_frame, text="Form Data Anggota Keluarga")
        form_frame.pack(fill="x", padx=20, pady=10)
        
        # Create a grid layout for the form
        form_fields = [
            ("No KK:", "anggota_kk_var"),
            ("NIK:", "anggota_nik_var"),
            ("Nama:", "anggota_nama_var"),
            ("Hubungan:", "anggota_hubungan_var"),
            ("Jenis Kelamin:", "anggota_jk_var"),
            ("Tanggal Lahir:", "anggota_tgl_lahir_var"),
            ("Pendidikan:", "anggota_pendidikan_var"),
            ("Pekerjaan:", "anggota_pekerjaan_var")
        ]
        
        # Create form fields
        row = 0
        col = 0
        for label_text, var_name in form_fields:
            ttk.Label(form_frame, text=label_text).grid(row=row, column=col, padx=5, pady=5, sticky="w")
            
            # Create StringVar
            setattr(self, var_name, tk.StringVar())
            
            # Handle special cases
            if var_name == "anggota_jk_var":
                # Radio buttons for gender
                rb_frame = ttk.Frame(form_frame)
                rb_frame.grid(row=row, column=col + 1, padx=5, pady=5, sticky="w")
                
                ttk.Radiobutton(rb_frame, text="Laki-laki", variable=self.anggota_jk_var, value="Laki-laki").pack(side="left", padx=5)
                ttk.Radiobutton(rb_frame, text="Perempuan", variable=self.anggota_jk_var, value="Perempuan").pack(side="left", padx=5)
                self.anggota_jk_var.set("Laki-laki")
            elif var_name == "anggota_hubungan_var":
                # Dropdown for relationship
                hubungan_options = ["Kepala Keluarga", "Istri", "Anak", "Orang Tua", "Mertua", "Menantu", "Cucu", "Lainnya"]
                ttk.Combobox(form_frame, textvariable=self.anggota_hubungan_var, values=hubungan_options, state="readonly", width=28).grid(
                    row=row, column=col + 1, padx=5, pady=5, sticky="w")
                self.anggota_hubungan_var.set(hubungan_options[0])
            elif var_name == "anggota_pendidikan_var":
                # Dropdown for education
                pendidikan_options = ["Tidak Sekolah", "SD", "SMP", "SMA/SMK", "D1", "D2", "D3", "S1", "S2", "S3"]
                ttk.Combobox(form_frame, textvariable=self.anggota_pendidikan_var, values=pendidikan_options, state="readonly", width=28).grid(
                    row=row, column=col + 1, padx=5, pady=5, sticky="w")
                self.anggota_pendidikan_var.set(pendidikan_options[0])
            elif var_name == "anggota_pekerjaan_var":
                # Dropdown for job
                pekerjaan_options = ["Tidak Bekerja", "Petani", "Nelayan", "PNS", "TNI/Polri", "Swasta", "Wiraswasta", "Pelajar/Mahasiswa", "Pensiunan", "Lainnya"]
                ttk.Combobox(form_frame, textvariable=self.anggota_pekerjaan_var, values=pekerjaan_options, state="readonly", width=28).grid(
                    row=row, column=col + 1, padx=5, pady=5, sticky="w")
                self.anggota_pekerjaan_var.set(pekerjaan_options[0])
            elif var_name == "anggota_tgl_lahir_var":
                # Date field
                date_frame = ttk.Frame(form_frame)
                date_frame.grid(row=row, column=col + 1, padx=5, pady=5, sticky="w")
                
                # Date components
                self.anggota_tgl_day = ttk.Combobox(date_frame, width=3, values=[str(i).zfill(2) for i in range(1, 32)], state="readonly")
                self.anggota_tgl_month = ttk.Combobox(date_frame, width=3, values=[str(i).zfill(2) for i in range(1, 13)], state="readonly")
                self.anggota_tgl_year = ttk.Combobox(date_frame, width=5, values=[str(i) for i in range(1940, datetime.now().year + 1)], state="readonly")
                
                self.anggota_tgl_day.pack(side="left", padx=2)
                ttk.Label(date_frame, text="/").pack(side="left")
                self.anggota_tgl_month.pack(side="left", padx=2)
                ttk.Label(date_frame, text="/").pack(side="left")
                self.anggota_tgl_year.pack(side="left", padx=2)
                
                # Set default to today
                today = datetime.now()
                self.anggota_tgl_day.set(str(today.day).zfill(2))
                self.anggota_tgl_month.set(str(today.month).zfill(2))
                self.anggota_tgl_year.set(str(today.year))
            elif var_name == "anggota_kk_var":
                # Dropdown for KK numbers
                self.kk_dropdown = ttk.Combobox(form_frame, textvariable=self.anggota_kk_var, width=28)
                self.kk_dropdown.grid(row=row, column=col + 1, padx=5, pady=5, sticky="w")
                # Will be populated when KK data is loaded
            else:
                # Regular entry field
                ttk.Entry(form_frame, textvariable=getattr(self, var_name), width=30).grid(
                    row=row, column=col + 1, padx=5, pady=5, sticky="w")
            
            # Handle layout
            row += 1
            if row > 3 and col == 0:
                row = 0
                col = 2
        
        # Buttons
        buttons_frame = ttk.Frame(form_frame)
        buttons_frame.grid(row=row + 1, column=0, columnspan=4, pady=10)
        
        ttk.Button(buttons_frame, text="Simpan", command=self.save_anggota).pack(side="left", padx=5)
        ttk.Button(buttons_frame, text="Reset", command=self.reset_anggota_form).pack(side="left", padx=5)
        ttk.Button(buttons_frame, text="Hitung Umur", command=self.calculate_age).pack(side="left", padx=5)
        
        # Table Frame
        table_frame = ttk.LabelFrame(anggota_frame, text="Data Anggota Keluarga")
        table_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # Create TreeView for Anggota data
        columns = ("No KK", "NIK", "Nama", "Hubungan", "Jenis Kelamin", "Tanggal Lahir", "Umur", "Pendidikan", "Pekerjaan")
        self.anggota_tree = ttk.Treeview(table_frame, columns=columns)
        self.anggota_tree.heading("#0", text="No")
        
        for col in columns:
            self.anggota_tree.heading(col, text=col)
            width = 100 if col not in ["Nama", "Tanggal Lahir"] else 150
            self.anggota_tree.column(col, width=width)
        
        self.anggota_tree.column("#0", width=50)
        
        # Add scrollbar to treeview
        anggota_scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.anggota_tree.yview)
        self.anggota_tree.configure(yscrollcommand=anggota_scrollbar.set)
        
        anggota_scrollbar.pack(side="right", fill="y")
        self.anggota_tree.pack(fill="both", expand=True)
        
        # Bind events
        self.anggota_tree.bind("<ButtonRelease-1>", self.load_anggota_to_form)
    
    def create_kelahiran_tab(self):
        kelahiran_frame = ttk.Frame(self.notebook)
        self.notebook.add(kelahiran_frame, text="Data Kelahiran")
        
        # Form Frame
        form_frame = ttk.LabelFrame(kelahiran_frame, text="Form Data Kelahiran")
        form_frame.pack(fill="x", padx=20, pady=10)
        
        # Create form fields (simplified version)
        fields_layout = [
            ("NIK Bayi:", "lahir_nik_var", 0, 0),
            ("Nama Bayi:", "lahir_nama_var", 1, 0),
            ("Tanggal Lahir:", "lahir_tgl_var", 2, 0),
            ("Jenis Kelamin:", "lahir_jk_var", 3, 0),
            ("Nama Ayah:", "lahir_ayah_var", 0, 2),
            ("Nama Ibu:", "lahir_ibu_var", 1, 2),
            ("Desa/Kelurahan:", "lahir_desa_var", 2, 2),
            ("No KK:", "lahir_kk_var", 3, 2)
        ]
        
        for label_text, var_name, row, col in fields_layout:
            ttk.Label(form_frame, text=label_text).grid(row=row, column=col, padx=5, pady=5, sticky="w")
            
            # Create StringVar
            setattr(self, var_name, tk.StringVar())
            
            # Special cases
            if var_name == "lahir_jk_var":
                rb_frame = ttk.Frame(form_frame)
                rb_frame.grid(row=row, column=col + 1, padx=5, pady=5, sticky="w")
                
                ttk.Radiobutton(rb_frame, text="Laki-laki", variable=self.lahir_jk_var, value="Laki-laki").pack(side="left", padx=5)
                ttk.Radiobutton(rb_frame, text="Perempuan", variable=self.lahir_jk_var, value="Perempuan").pack(side="left", padx=5)
                self.lahir_jk_var.set("Laki-laki")
            elif var_name == "lahir_tgl_var":
                date_frame = ttk.Frame(form_frame)
                date_frame.grid(row=row, column=col + 1, padx=5, pady=5, sticky="w")
                
                self.lahir_tgl_day = ttk.Combobox(date_frame, width=3, values=[str(i).zfill(2) for i in range(1, 32)], state="readonly")
                self.lahir_tgl_month = ttk.Combobox(date_frame, width=3, values=[str(i).zfill(2) for i in range(1, 13)], state="readonly")
                self.lahir_tgl_year = ttk.Combobox(date_frame, width=5, values=[str(i) for i in range(2000, datetime.now().year + 1)], state="readonly")
                
                self.lahir_tgl_month.pack(side="left", padx=2)
                ttk.Label(date_frame, text="/").pack(side="left")
                self.lahir_tgl_year.pack(side="left", padx=2)
                
                # Set default to today
                today = datetime.now()
                self.lahir_tgl_day.set(str(today.day).zfill(2))
                self.lahir_tgl_month.set(str(today.month).zfill(2))
                self.lahir_tgl_year.set(str(today.year))
            elif var_name == "lahir_desa_var":
                desa_options = ["Desa A", "Desa B", "Desa C", "Desa D", "Desa E"]
                ttk.Combobox(form_frame, textvariable=self.lahir_desa_var, values=desa_options, state="readonly", width=28).grid(
                    row=row, column=col + 1, padx=5, pady=5, sticky="w")
                self.lahir_desa_var.set(desa_options[0])
            elif var_name == "lahir_kk_var":
                self.lahir_kk_dropdown = ttk.Combobox(form_frame, textvariable=self.lahir_kk_var, width=28)
                self.lahir_kk_dropdown.grid(row=row, column=col + 1, padx=5, pady=5, sticky="w")
                # Will be populated when KK data is loaded
            else:
                ttk.Entry(form_frame, textvariable=getattr(self, var_name), width=30).grid(
                    row=row, column=col + 1, padx=5, pady=5, sticky="w")
        
        # Buttons
        buttons_frame = ttk.Frame(form_frame)
        buttons_frame.grid(row=5, column=0, columnspan=4, pady=10)
        
        ttk.Button(buttons_frame, text="Simpan", command=self.save_kelahiran).pack(side="left", padx=5)
        ttk.Button(buttons_frame, text="Reset", command=self.reset_kelahiran_form).pack(side="left", padx=5)
        
        # Table Frame
        table_frame = ttk.LabelFrame(kelahiran_frame, text="Data Kelahiran")
        table_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # Create TreeView for birth data
        columns = ("NIK", "Nama Bayi", "Tanggal Lahir", "Jenis Kelamin", "Nama Ayah", "Nama Ibu", "Desa", "No KK")
        self.kelahiran_tree = ttk.Treeview(table_frame, columns=columns)
        self.kelahiran_tree.heading("#0", text="No")
        
        for col in columns:
            self.kelahiran_tree.heading(col, text=col)
            width = 120 if col not in ["Nama Bayi", "Nama Ayah", "Nama Ibu"] else 150
            self.kelahiran_tree.column(col, width=width)
        
        self.kelahiran_tree.column("#0", width=50)
        
        # Add scrollbar to treeview
        kelahiran_scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.kelahiran_tree.yview)
        self.kelahiran_tree.configure(yscrollcommand=kelahiran_scrollbar.set)
        
        kelahiran_scrollbar.pack(side="right", fill="y")
        self.kelahiran_tree.pack(fill="both", expand=True)
        
        # Bind events
        self.kelahiran_tree.bind("<ButtonRelease-1>", self.load_kelahiran_to_form)
    
    def create_kematian_tab(self):
        kematian_frame = ttk.Frame(self.notebook)
        self.notebook.add(kematian_frame, text="Data Kematian")
        
        # Form Frame
        form_frame = ttk.LabelFrame(kematian_frame, text="Form Data Kematian")
        form_frame.pack(fill="x", padx=20, pady=10)
        
        # Create form fields
        fields_layout = [
            ("NIK:", "mati_nik_var", 0, 0),
            ("Nama:", "mati_nama_var", 1, 0),
            ("Tanggal Lahir:", "mati_tgl_lahir_var", 2, 0),
            ("Tanggal Meninggal:", "mati_tgl_meninggal_var", 3, 0),
            ("Usia Meninggal:", "mati_usia_var", 0, 2),
            ("Penyebab:", "mati_penyebab_var", 1, 2),
            ("Desa/Kelurahan:", "mati_desa_var", 2, 2),
            ("No KK:", "mati_kk_var", 3, 2)
        ]
        
        for label_text, var_name, row, col in fields_layout:
            ttk.Label(form_frame, text=label_text).grid(row=row, column=col, padx=5, pady=5, sticky="w")
            
            # Create StringVar
            setattr(self, var_name, tk.StringVar())
            
            # Special cases
            if var_name in ["mati_tgl_lahir_var", "mati_tgl_meninggal_var"]:
                date_frame = ttk.Frame(form_frame)
                date_frame.grid(row=row, column=col + 1, padx=5, pady=5, sticky="w")
                
                day_var = var_name.replace("var", "day")
                month_var = var_name.replace("var", "month")
                year_var = var_name.replace("var", "year")
                
                setattr(self, day_var, ttk.Combobox(date_frame, width=3, values=[str(i).zfill(2) for i in range(1, 32)], state="readonly"))
                setattr(self, month_var, ttk.Combobox(date_frame, width=3, values=[str(i).zfill(2) for i in range(1, 13)], state="readonly"))
                setattr(self, year_var, ttk.Combobox(date_frame, width=5, values=[str(i) for i in range(1940, datetime.now().year + 1)], state="readonly"))
                
                getattr(self, day_var).pack(side="left", padx=2)
                ttk.Label(date_frame, text="/").pack(side="left")
                getattr(self, month_var).pack(side="left", padx=2)
                ttk.Label(date_frame, text="/").pack(side="left")
                getattr(self, year_var).pack(side="left", padx=2)
                
                # Set default to today
                today = datetime.now()
                getattr(self, day_var).set(str(today.day).zfill(2))
                getattr(self, month_var).set(str(today.month).zfill(2))
                getattr(self, year_var).set(str(today.year))
            elif var_name == "mati_desa_var":
                desa_options = ["Desa A", "Desa B", "Desa C", "Desa D", "Desa E"]
                ttk.Combobox(form_frame, textvariable=self.mati_desa_var, values=desa_options, state="readonly", width=28).grid(
                    row=row, column=col + 1, padx=5, pady=5, sticky="w")
                self.mati_desa_var.set(desa_options[0])
            elif var_name == "mati_kk_var":
                self.mati_kk_dropdown = ttk.Combobox(form_frame, textvariable=self.mati_kk_var, width=28)
                self.mati_kk_dropdown.grid(row=row, column=col + 1, padx=5, pady=5, sticky="w")
                # Will be populated when KK data is loaded
            elif var_name == "mati_usia_var":
                # This will be calculated, set it as readonly
                ttk.Entry(form_frame, textvariable=self.mati_usia_var, width=30, state="readonly").grid(
                    row=row, column=col + 1, padx=5, pady=5, sticky="w")
            else:
                ttk.Entry(form_frame, textvariable=getattr(self, var_name), width=30).grid(
                    row=row, column=col + 1, padx=5, pady=5, sticky="w")
        
        # Buttons
        buttons_frame = ttk.Frame(form_frame)
        buttons_frame.grid(row=5, column=0, columnspan=4, pady=10)
        
        ttk.Button(buttons_frame, text="Hitung Usia", command=self.calculate_death_age).pack(side="left", padx=5)
        ttk.Button(buttons_frame, text="Simpan", command=self.save_kematian).pack(side="left", padx=5)
        ttk.Button(buttons_frame, text="Reset", command=self.reset_kematian_form).pack(side="left", padx=5)
        
        # Table Frame
        table_frame = ttk.LabelFrame(kematian_frame, text="Data Kematian")
        table_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # Create TreeView for death data
        columns = ("NIK", "Nama", "Tanggal Lahir", "Tanggal Meninggal", "Usia Meninggal", "Penyebab", "Desa", "No KK")
        self.kematian_tree = ttk.Treeview(table_frame, columns=columns)
        self.kematian_tree.heading("#0", text="No")
        
        for col in columns:
            self.kematian_tree.heading(col, text=col)
            width = 100 if col not in ["Nama", "Penyebab", "Tanggal Lahir", "Tanggal Meninggal"] else 150
            self.kematian_tree.column(col, width=width)
        
        self.kematian_tree.column("#0", width=50)
        
        # Add scrollbar to treeview
        kematian_scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self.kematian_tree.yview)
        self.kematian_tree.configure(yscrollcommand=kematian_scrollbar.set)
        
        kematian_scrollbar.pack(side="right", fill="y")
        self.kematian_tree.pack(fill="both", expand=True)
        
        # Bind events
        self.kematian_tree.bind("<ButtonRelease-1>", self.load_kematian_to_form)
    
    def create_laporan_tab(self):
        laporan_frame = ttk.Frame(self.notebook)
        self.notebook.add(laporan_frame, text="Laporan & Ekspor")
        
        # Options Frame
        options_frame = ttk.LabelFrame(laporan_frame, text="Opsi Laporan")
        options_frame.pack(fill="x", padx=20, pady=10)
        
        # Report Type
        ttk.Label(options_frame, text="Jenis Laporan:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.report_type_var = tk.StringVar(value="Rekapitulasi Penduduk")
        report_types = [
            "Rekapitulasi Penduduk",
            "Statistik per Desa",
            "Data Kelahiran",
            "Data Kematian",
            "Pendidikan",
            "Pekerjaan",
            "Semua Data"
        ]
        ttk.Combobox(options_frame, textvariable=self.report_type_var, values=report_types, state="readonly", width=30).grid(
            row=0, column=1, padx=5, pady=5, sticky="w")
        
        # Filter by Desa
        ttk.Label(options_frame, text="Filter Desa:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.report_desa_var = tk.StringVar(value="Semua Desa")
        desa_options = ["Semua Desa", "Desa A", "Desa B", "Desa C", "Desa D", "Desa E"]
        ttk.Combobox(options_frame, textvariable=self.report_desa_var, values=desa_options, state="readonly", width=30).grid(
            row=1, column=1, padx=5, pady=5, sticky="w")
        
        # Buttons Frame
        buttons_frame = ttk.Frame(options_frame)
        buttons_frame.grid(row=2, column=0, columnspan=2, pady=10)
        
        ttk.Button(buttons_frame, text="Buat Laporan", command=self.generate_report).pack(side="left", padx=5)
        ttk.Button(buttons_frame, text="Ekspor ke Excel", command=self.export_to_excel).pack(side="left", padx=5)
        ttk.Button(buttons_frame, text="Cetak Grafik", command=self.print_chart).pack(side="left", padx=5)
        
        # Preview Frame
        preview_frame = ttk.LabelFrame(laporan_frame, text="Pratinjau Laporan")
        preview_frame.pack(fill="both", expand=True, padx=20, pady=10)
        
        # Create a canvas for displaying report content
        self.report_canvas = tk.Canvas(preview_frame, bg="white")
        
        # Add scrollbar to canvas
        report_scrollbar = ttk.Scrollbar(preview_frame, orient="vertical", command=self.report_canvas.yview)
        self.report_canvas.configure(yscrollcommand=report_scrollbar.set)
        
        report_scrollbar.pack(side="right", fill="y")
        self.report_canvas.pack(fill="both", expand=True)
        
        # Create an inner frame for report contents
        self.report_content_frame = ttk.Frame(self.report_canvas)
        self.report_canvas.create_window((0, 0), window=self.report_content_frame, anchor="nw")
        
        # Configure the canvas
        self.report_content_frame.bind("<Configure>", lambda e: self.report_canvas.configure(
            scrollregion=self.report_canvas.bbox("all")))
        
        # Add initial content to preview
        ttk.Label(self.report_content_frame, text="Silahkan pilih jenis laporan dan klik 'Buat Laporan'",
                 font=("Arial", 14)).pack(pady=50)
    
    def load_excel(self):
        """Load data from Excel file"""
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls")])
        if not file_path:
            return
        
        try:
            # Read Excel file using pandas
            self.excel_file = file_path
            
            # Load Kartu Keluarga data
            kk_df = pd.read_excel(file_path, sheet_name='Data Penduduk', skiprows=3, nrows=6,
                             usecols=range(1, 6), header=0)
            kk_df.columns = ["No KK", "Nama Kepala Keluarga", "Desa", "RT/RW", "Jumlah Anggota"]
            self.data_penduduk = kk_df.dropna(subset=["No KK"])
            
            # Load Anggota Keluarga data
            anggota_df = pd.read_excel(file_path, sheet_name='Data Penduduk', skiprows=10,
                                  usecols=range(1, 10), header=0)
            anggota_df.columns = ["NIK", "Nama", "Hubungan", "Jenis_Kelamin",
                               "Tanggal_Lahir", "Umur", "Pendidikan", "Pekerjaan", "No KK"]
            self.data_anggota = anggota_df.dropna(subset=["NIK"])
            
            # Load Kelahiran data
            lahir_df = pd.read_excel(file_path, sheet_name='Data Kelahiran', skiprows=2,
                                usecols=range(1, 9), header=0)
            lahir_df.columns = ["NIK", "Nama_Bayi", "Tanggal_Lahir", "Jenis_Kelamin",
                             "Nama_Ayah", "Nama_Ibu", "Desa", "No_KK"]
            self.data_kelahiran = lahir_df.dropna(subset=["NIK"])
            
            # Load Kematian data
            mati_df = pd.read_excel(file_path, sheet_name='Data Kematian', skiprows=2,
                               usecols=range(1, 9), header=0)
            mati_df.columns = ["NIK", "Nama", "Tanggal_Lahir", "Tanggal_Meninggal",
                            "Usia_Meninggal", "Penyebab", "Desa", "No_KK"]
            self.data_kematian = mati_df.dropna(subset=["NIK"])
            
            # Update UI with loaded data
            self.refresh_all_data()
            self.update_dashboard()
            
            # Update status bar and show message
            self.status_var.set(f"Data berhasil dimuat dari {os.path.basename(file_path)}")
            messagebox.showinfo("Sukses", "Data berhasil dimuat")
            
        except Exception as e:
            messagebox.showerror("Error", f"Gagal memuat data: {str(e)}")
            print(f"Error: {str(e)}")
    
    def save_excel(self):
        """Save data to Excel file"""
        if self.excel_file:
            file_path = self.excel_file
        else:
            file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                filetypes=[("Excel files", "*.xlsx")])
            if not file_path:
                return
        
        try:
            # Create a writer object
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                # Sheet 1: Data Penduduk
                # Add header
                kecamatan_name = self.kecamatan_var.get()
                header_df = pd.DataFrame([["FORM DATA KEPENDUDUKAN KECAMATAN " + kecamatan_name]],
                                        columns=["A"])
                header_df.to_excel(writer, sheet_name='Data Penduduk', index=False, header=False)
                
                # Add Data KK section header
                kk_header = pd.DataFrame([["", ""], ["", ""], ["Data KK", "", "", "", ""]],
                                     columns=["A", "B", "C", "D", "E"])
                kk_header.to_excel(writer, sheet_name='Data Penduduk', startrow=1, index=False, header=False)
                
                # Add KK data
                self.data_penduduk.to_excel(writer, sheet_name='Data Penduduk', startrow=4, index=True,
                                      header=True)
                
                # Add Data Anggota section header
                anggota_header = pd.DataFrame([["Data Anggota Keluarga", "", "", "", "", "", "", "", ""]],
                                          columns=["A", "B", "C", "D", "E", "F", "G", "H", "I"])
                anggota_header.to_excel(writer, sheet_name='Data Penduduk', startrow=10, index=False,
                                   header=False)
                
                # Add Anggota data
                self.data_anggota.to_excel(writer, sheet_name='Data Penduduk', startrow=11, index=True,
                                     header=True)
                
                # Sheet 2: Rekapitulasi per Desa
                # Create summary data
                desa_list = self.data_penduduk["Desa"].unique()
                summary_data = []
                
                for desa in desa_list:
                    desa_kk = len(self.data_penduduk[self.data_penduduk["Desa"] == desa])
                    desa_anggota = len(self.data_anggota[self.data_anggota["No KK"].isin(
                        self.data_penduduk[self.data_penduduk["Desa"] == desa]["No KK"])])
                    
                    # Gender counts
                    anggota_in_desa = self.data_anggota[self.data_anggota["No KK"].isin(
                        self.data_penduduk[self.data_penduduk["Desa"] == desa]["No KK"])]
                    laki = len(anggota_in_desa[anggota_in_desa["Jenis_Kelamin"] == "Laki-laki"])
                    perempuan = len(anggota_in_desa[anggota_in_desa["Jenis_Kelamin"] == "Perempuan"])
                    
                    # Age categories
                    balita = len(anggota_in_desa[anggota_in_desa["Umur"] < 5])
                    anak = len(anggota_in_desa[(anggota_in_desa["Umur"] >= 5) & (anggota_in_desa["Umur"] < 18)])
                    dewasa = len(anggota_in_desa[(anggota_in_desa["Umur"] >= 18) & (anggota_in_desa["Umur"] < 60)])
                    lansia = len(anggota_in_desa[anggota_in_desa["Umur"] >= 60])
                    
                    summary_data.append([desa, desa_kk, desa_anggota, laki, perempuan, balita, anak, dewasa, lansia])
                
                rekapitulasi_df = pd.DataFrame(summary_data, columns=[
                    "Desa/Kelurahan", "Jumlah KK", "Jumlah Penduduk", "Laki-laki", "Perempuan",
                    "Balita", "Anak", "Dewasa", "Lansia"
                ])
                
                # Add header
                recap_header = pd.DataFrame([["REKAPITULASI DATA PENDUDUK PER DESA"], [""], ["", "", "", "", "", "", "", "", ""]],
                                         columns=["A"])
                recap_header.to_excel(writer, sheet_name='Rekapitulasi per Desa', index=False, header=False)
                
                # Add rekapitulasi data
                rekapitulasi_df.to_excel(writer, sheet_name='Rekapitulasi per Desa', startrow=2, index=False)
                
                # Education summary
                education_header = pd.DataFrame([[""], [""], ["Pendidikan"], ["", "", "", "", "", "", "", "", ""]],
                                             columns=["A"])
                education_header.to_excel(writer, sheet_name='Rekapitulasi per Desa', startrow=len(desa_list) + 5,
                                       index=False, header=False)
                
                edu_summary = []
                for desa in desa_list:
                    anggota_in_desa = self.data_anggota[self.data_anggota["No KK"].isin(
                        self.data_penduduk[self.data_penduduk["Desa"] == desa]["No KK"])]
                    
                    tidak_sekolah = len(anggota_in_desa[anggota_in_desa["Pendidikan"] == "Tidak Sekolah"])
                    sd = len(anggota_in_desa[anggota_in_desa["Pendidikan"] == "SD"])
                    smp = len(anggota_in_desa[anggota_in_desa["Pendidikan"] == "SMP"])
                    sma = len(anggota_in_desa[anggota_in_desa["Pendidikan"] == "SMA/SMK"])
                    d1_d3 = len(anggota_in_desa[(anggota_in_desa["Pendidikan"] == "D1") | 
                                            (anggota_in_desa["Pendidikan"] == "D2") | 
                                            (anggota_in_desa["Pendidikan"] == "D3")])
                    s1 = len(anggota_in_desa[anggota_in_desa["Pendidikan"] == "S1"])
                    s2 = len(anggota_in_desa[anggota_in_desa["Pendidikan"] == "S2"])
                    s3 = len(anggota_in_desa[anggota_in_desa["Pendidikan"] == "S3"])
                    
                    edu_summary.append([desa, tidak_sekolah, sd, smp, sma, d1_d3, s1, s2, s3])
                
                edu_df = pd.DataFrame(edu_summary, columns=[
                    "Desa/Kelurahan", "Tidak Sekolah", "SD", "SMP", "SMA", "D1-D3", "S1", "S2", "S3"
                ])
                
                edu_df.to_excel(writer, sheet_name='Rekapitulasi per Desa', startrow=len(desa_list) + 9,
                             index=False)
                
                # Job summary
                job_header = pd.DataFrame([[""], [""], ["Pekerjaan"], ["", "", "", "", "", "", "", "", ""]],
                                       columns=["A"])
                job_header.to_excel(writer, sheet_name='Rekapitulasi per Desa',
                                 startrow=len(desa_list) + 9 + len(desa_list) + 5,
                                 index=False, header=False)
                
                job_summary = []
                for desa in desa_list:
                    anggota_in_desa = self.data_anggota[self.data_anggota["No KK"].isin(
                        self.data_penduduk[self.data_penduduk["Desa"] == desa]["No KK"])]
                    
                    petani = len(anggota_in_desa[anggota_in_desa["Pekerjaan"] == "Petani"])
                    nelayan = len(anggota_in_desa[anggota_in_desa["Pekerjaan"] == "Nelayan"])
                    pns = len(anggota_in_desa[anggota_in_desa["Pekerjaan"] == "PNS"])
                    swasta = len(anggota_in_desa[anggota_in_desa["Pekerjaan"] == "Swasta"])
                    wiraswasta = len(anggota_in_desa[anggota_in_desa["Pekerjaan"] == "Wiraswasta"])
                    pelajar = len(anggota_in_desa[anggota_in_desa["Pekerjaan"] == "Pelajar/Mahasiswa"])
                    tidak_bekerja = len(anggota_in_desa[anggota_in_desa["Pekerjaan"] == "Tidak Bekerja"])
                    lainnya = len(anggota_in_desa[(anggota_in_desa["Pekerjaan"] == "Lainnya") | 
                                             (anggota_in_desa["Pekerjaan"] == "TNI/Polri") | 
                                             (anggota_in_desa["Pekerjaan"] == "Pensiunan")])
                    
                    job_summary.append([desa, petani, nelayan, pns, swasta, wiraswasta, pelajar, tidak_bekerja, lainnya])
                
                job_df = pd.DataFrame(job_summary, columns=[
                    "Desa/Kelurahan", "Petani", "Nelayan", "PNS", "Swasta", "Wiraswasta", "Pelajar", "Tidak Bekerja", "Lainnya"
                ])
                
                job_df.to_excel(writer, sheet_name='Rekapitulasi per Desa',
                             startrow=len(desa_list) + 9 + len(desa_list) + 9,
                             index=False)
                
                # Sheet 3: Data Kelahiran
                birth_header = pd.DataFrame([["DATA KELAHIRAN"], [""], []], columns=["A"])
                birth_header.to_excel(writer, sheet_name='Data Kelahiran', index=False, header=False)
                
                # Add birth data
                self.data_kelahiran.to_excel(writer, sheet_name='Data Kelahiran', startrow=2, index=True)
                
                # Sheet 4: Data Kematian
                death_header = pd.DataFrame([["DATA KEMATIAN"], [""], []], columns=["A"])
                death_header.to_excel(writer, sheet_name='Data Kematian', index=False, header=False)
                
                # Add death data
                self.data_kematian.to_excel(writer, sheet_name='Data Kematian', startrow=2, index=True)
                
                # Sheet 5: Dashboard
                dash_header = pd.DataFrame([["DASHBOARD KEPENDUDUKAN KECAMATAN " + kecamatan_name]],
                                        columns=["A"])
                dash_header.to_excel(writer, sheet_name='Dashboard', index=False, header=False)
                
                # Here we would add chart images if needed
            
            # Update status bar and show message
            self.excel_file = file_path
            self.status_var.set(f"Data berhasil disimpan ke {os.path.basename(file_path)}")
            messagebox.showinfo("Sukses", "Data berhasil disimpan")
            
        except Exception as e:
            messagebox.showerror("Error", f"Gagal menyimpan data: {str(e)}")
    
    def refresh_all_data(self):
        """Refresh all UI elements with current data"""
        # Clear and reload KK Treeview
        for item in self.kk_tree.get_children():
            self.kk_tree.delete(item)
        
        for i, row in self.data_penduduk.iterrows():
            self.kk_tree.insert("", "end", text=str(i + 1),
                              values=(row["No KK"], row["Nama Kepala Keluarga"],
                                     row["Desa"], row["RT/RW"], row["Jumlah Anggota"]))
        
        # Clear and reload Anggota Treeview
        for item in self.anggota_tree.get_children():
            self.anggota_tree.delete(item)
        
        for i, row in self.data_anggota.iterrows():
            self.anggota_tree.insert("", "end", text=str(i + 1),
                                   values=(row["No KK"], row["NIK"], row["Nama"],
                                          row["Hubungan"], row["Jenis_Kelamin"],
                                          row["Tanggal_Lahir"], row["Umur"],
                                          row["Pendidikan"], row["Pekerjaan"]))
        
        # Clear and reload Kelahiran Treeview
        for item in self.kelahiran_tree.get_children():
            self.kelahiran_tree.delete(item)
        
        for i, row in self.data_kelahiran.iterrows():
            self.kelahiran_tree.insert("", "end", text=str(i + 1),
                                     values=(row["NIK"], row["Nama_Bayi"],
                                            row["Tanggal_Lahir"], row["Jenis_Kelamin"],
                                            row["Nama_Ayah"], row["Nama_Ibu"],
                                            row["Desa"], row["No_KK"]))
        
        # Clear and reload Kematian Treeview
        for item in self.kematian_tree.get_children():
            self.kematian_tree.delete(item)
        
        for i, row in self.data_kematian.iterrows():
            self.kematian_tree.insert("", "end", text=str(i + 1),
                                    values=(row["NIK"], row["Nama"],
                                           row["Tanggal_Lahir"], row["Tanggal_Meninggal"],
                                           row["Usia_Meninggal"], row["Penyebab"],
                                           row["Desa"], row["No_KK"]))
        
        # Update KK dropdowns in all tabs
        kk_list = self.data_penduduk["No KK"].tolist()
        self.kk_dropdown['values'] = kk_list
        self.lahir_kk_dropdown['values'] = kk_list
        self.mati_kk_dropdown['values'] = kk_list
        
        # Update status bar
        self.status_var.set(f"Data diperbarui: {len(self.data_penduduk)} KK, {len(self.data_anggota)} anggota keluarga")
    
    def update_dashboard(self):
        """Update dashboard with current statistics"""
        try:
            # Calculate statistics
            total_penduduk = len(self.data_anggota)
            total_kk = len(self.data_penduduk)
            total_laki = len(self.data_anggota[self.data_anggota["Jenis_Kelamin"] == "Laki-laki"])
            total_perempuan = len(self.data_anggota[self.data_anggota["Jenis_Kelamin"] == "Perempuan"])
            
            # Calculate age categories
            self.data_anggota["Umur"] = self.data_anggota["Umur"].fillna(0).astype(int)
            balita = len(self.data_anggota[self.data_anggota["Umur"] < 5])
            anak = len(self.data_anggota[(self.data_anggota["Umur"] >= 5) & (self.data_anggota["Umur"] < 18)])
            dewasa = len(self.data_anggota[(self.data_anggota["Umur"] >= 18) & (self.data_anggota["Umur"] < 60)])
            lansia = len(self.data_anggota[self.data_anggota["Umur"] >= 60])
            
            # Update dashboard values
            self.total_penduduk_var.set(str(total_penduduk))
            self.total_kk_var.set(str(total_kk))
            self.total_laki_var.set(str(total_laki))
            self.total_perempuan_var.set(str(total_perempuan))
            
            # Update age distribution chart
            self.ax.clear()
            categories = ['Balita (0-4)', 'Anak (5-17)', 'Dewasa (18-59)', 'Lansia (60+)']
            values = [balita, anak, dewasa, lansia]
            colors = ['#3498db', '#2ecc71', '#e74c3c', '#f39c12']
            
            bars = self.ax.bar(categories, values, color=colors)
            
            # Add data labels on top of each bar
            for bar in bars:
                height = bar.get_height()
                self.ax.text(bar.get_x() + bar.get_width() / 2., height + 0.5,
                           f'{int(height)}',
                           ha='center', va='bottom', fontweight='bold')
            
            self.ax.set_title(f'Distribusi Penduduk Kecamatan {self.kecamatan_var.get()}', fontweight='bold')
            self.ax.set_ylabel('Jumlah Penduduk')
            
            # Set y limit to add some space for the labels
            max_value = max(values) if max(values) > 0 else 10
            self.ax.set_ylim(0, max_value * 1.15)
            
            # Add a light grid for better readability
            self.ax.grid(axis='y', linestyle='--', alpha=0.7)
            
            # Adjust layout and draw
            self.fig.tight_layout()
            self.canvas.draw()
            
        except Exception as e:
            messagebox.showerror("Error", f"Gagal update dashboard: {str(e)}")
            print(f"Error updating dashboard: {str(e)}")
    
    def save_kk(self):
        """Save Kartu Keluarga data"""
        try:
            # Get data from form
            no_kk = self.no_kk_var.get().strip()
            nama_kk = self.nama_kk_var.get().strip()
            desa = self.desa_kk_var.get()
            rtrw = self.rtrw_kk_var.get().strip()
            jumlah_anggota = self.jumlah_anggota_var.get().strip()
            
            # Validation
            if not all([no_kk, nama_kk, desa, rtrw, jumlah_anggota]):
                messagebox.showwarning("Perhatian", "Semua field harus diisi!")
                return
            
            # Check if no_kk is numeric and has correct length
            if not no_kk.isdigit() or len(no_kk) != 16:
                messagebox.showwarning("Perhatian", "No KK harus berupa 16 digit angka!")
                return
            
            # Check if jumlah_anggota is numeric
            if not jumlah_anggota.isdigit():
                messagebox.showwarning("Perhatian", "Jumlah anggota harus berupa angka!")
                return
            
            # Check if KK already exists
            if no_kk in self.data_penduduk["No KK"].values:
                # Update existing record
                idx = self.data_penduduk[self.data_penduduk["No KK"] == no_kk].index[0]
                self.data_penduduk.loc[idx] = [no_kk, nama_kk, desa, rtrw, int(jumlah_anggota)]
                status_msg = "Data KK berhasil diperbarui"
            else:
                # Add new record
                new_row = pd.DataFrame([[no_kk, nama_kk, desa, rtrw, int(jumlah_anggota)]],
                                     columns=["No KK", "Nama Kepala Keluarga", "Desa", "RT/RW", "Jumlah Anggota"])
                self.data_penduduk = pd.concat([self.data_penduduk, new_row], ignore_index=True)
                status_msg = "Data KK baru berhasil ditambahkan"
            
            # Refresh UI
            self.refresh_all_data()
            self.update_dashboard()
            self.reset_kk_form()
            
            # Update status bar
            self.status_var.set(status_msg)
            messagebox.showinfo("Sukses", status_msg)
            
        except Exception as e:
            messagebox.showerror("Error", f"Gagal menyimpan data KK: {str(e)}")
            print(f"Error saving KK: {str(e)}")
    
    def reset_kk_form(self):
        """Reset Kartu Keluarga form"""
        self.no_kk_var.set("")
        self.nama_kk_var.set("")
        self.desa_kk_var.set("Desa A")
        self.rtrw_kk_var.set("")
        self.jumlah_anggota_var.set("")
    
    def load_kk_to_form(self, event):
        """Load selected KK data to form"""
        # Get selected item
        selected_item = self.kk_tree.selection()
        if not selected_item:
            return
        
        # Get values
        values = self.kk_tree.item(selected_item[0], "values")
        
        # Set form values
        self.no_kk_var.set(values[0])
        self.nama_kk_var.set(values[1])
        self.desa_kk_var.set(values[2])
        self.rtrw_kk_var.set(values[3])
        self.jumlah_anggota_var.set(values[4])
    
    def save_anggota(self):
        """Save Anggota Keluarga data"""
        try:
            # Get data from form
            no_kk = self.anggota_kk_var.get().strip()
            nik = self.anggota_nik_var.get().strip()
            nama = self.anggota_nama_var.get().strip()
            hubungan = self.anggota_hubungan_var.get()
            jenis_kelamin = self.anggota_jk_var.get()
            
            # Construct date string
            try:
                tanggal_lahir = f"{self.anggota_tgl_day.get()}/{self.anggota_tgl_month.get()}/{self.anggota_tgl_year.get()}"
                # Calculate age
                tgl_lahir = datetime.strptime(tanggal_lahir, "%d/%m/%Y")
                umur = (datetime.now() - tgl_lahir).days // 365
            except:
                messagebox.showwarning("Perhatian", "Format tanggal lahir tidak valid!")
                return
            
            pendidikan = self.anggota_pendidikan_var.get()
            pekerjaan = self.anggota_pekerjaan_var.get()
            
            # Validation
            if not all([no_kk, nik, nama, hubungan, jenis_kelamin, tanggal_lahir, pendidikan, pekerjaan]):
                messagebox.showwarning("Perhatian", "Semua field harus diisi!")
                return
            
            # Check if no_kk exists in KK data
            if no_kk not in self.data_penduduk["No KK"].values:
                messagebox.showwarning("Perhatian", "No KK tidak ditemukan dalam data KK!")
                return
            
            # Check if NIK is numeric and has correct length
            if not nik.isdigit() or len(nik) != 16:
                messagebox.showwarning("Perhatian", "NIK harus berupa 16 digit angka!")
                return
            
            # Check if NIK already exists
            if nik in self.data_anggota["NIK"].values:
                # Update existing record
                idx = self.data_anggota[self.data_anggota["NIK"] == nik].index[0]
                self.data_anggota.loc[idx] = [no_kk, nik, nama, hubungan, jenis_kelamin,
                                           tanggal_lahir, umur, pendidikan, pekerjaan]
                status_msg = "Data anggota keluarga berhasil diperbarui"
            else:
                # Add new record
                new_row = pd.DataFrame([[no_kk, nik, nama, hubungan, jenis_kelamin,
                                        tanggal_lahir, umur, pendidikan, pekerjaan]],
                                     columns=["No KK", "NIK", "Nama", "Hubungan", "Jenis_Kelamin",
                                           "Tanggal_Lahir", "Umur", "Pendidikan", "Pekerjaan"])
                self.data_anggota = pd.concat([self.data_anggota, new_row], ignore_index=True)
                status_msg = "Data anggota keluarga baru berhasil ditambahkan"
                
                # Update jumlah anggota in KK data
                kk_idx = self.data_penduduk[self.data_penduduk["No KK"] == no_kk].index[0]
                current_count = self.data_penduduk.loc[kk_idx, "Jumlah Anggota"]
                if isinstance(current_count, str):
                    current_count = int(current_count.strip() or "0")
                self.data_penduduk.loc[kk_idx, "Jumlah Anggota"] = current_count + 1
            
            # Refresh UI
            self.refresh_all_data()
            self.update_dashboard()
            self.reset_anggota_form()
            
            # Update status bar
            self.status_var.set(status_msg)
            messagebox.showinfo("Sukses", status_msg)
            
        except Exception as e:
            messagebox.showerror("Error", f"Gagal menyimpan data anggota: {str(e)}")
            print(f"Error saving anggota: {str(e)}")
    
    def reset_anggota_form(self):
        """Reset Anggota Keluarga form"""
        self.anggota_kk_var.set("")
        self.anggota_nik_var.set("")
        self.anggota_nama_var.set("")
        self.anggota_hubungan_var.set("Kepala Keluarga")
        self.anggota_jk_var.set("Laki-laki")
        
        # Reset date to today
        today = datetime.now()
        self.anggota_tgl_day.set(str(today.day).zfill(2))
        self.anggota_tgl_month.set(str(today.month).zfill(2))
        self.anggota_tgl_year.set(str(today.year))
        
        self.anggota_pendidikan_var.set("Tidak Sekolah")
        self.anggota_pekerjaan_var.set("Tidak Bekerja")
    
    def load_anggota_to_form(self, event):
        """Load selected anggota data to form"""
        # Get selected item
        selected_item = self.anggota_tree.selection()
        if not selected_item:
            return
        
        # Get values
        values = self.anggota_tree.item(selected_item[0], "values")
        
        # Set form values
        self.anggota_kk_var.set(values[0])
        self.anggota_nik_var.set(values[1])
        self.anggota_nama_var.set(values[2])
        self.anggota_hubungan_var.set(values[3])
        self.anggota_jk_var.set(values[4])
        
        # Handle date
        try:
            tgl_parts = values[5].split('/')
            if len(tgl_parts) == 3:
                self.anggota_tgl_day.set(tgl_parts[0])
                self.anggota_tgl_month.set(tgl_parts[1])
                self.anggota_tgl_year.set(tgl_parts[2])
        except:
            # Reset to today if error
            today = datetime.now()
            self.anggota_tgl_day.set(str(today.day).zfill(2))
            self.anggota_tgl_month.set(str(today.month).zfill(2))
            self.anggota_tgl_year.set(str(today.year))
        
        self.anggota_pendidikan_var.set(values[7])
        self.anggota_pekerjaan_var.set(values[8])
    
    def calculate_age(self):
        """Calculate age based on birth date"""
        try:
            # Get date components
            day = self.anggota_tgl_day.get()
            month = self.anggota_tgl_month.get()
            year = self.anggota_tgl_year.get()
            
            # Create date object
            birth_date = datetime.strptime(f"{day}/{month}/{year}", "%d/%m/%Y")
            
            # Calculate age
            today = datetime.now()
            age = today.year - birth_date.year - ((today.month, today.day) < (birth_date.month, birth_date.day))
            
            # Show in a message box
            messagebox.showinfo("Hasil Perhitungan", f"Umur: {age} tahun")
            
        except Exception as e:
            messagebox.showerror("Error", f"Gagal menghitung umur: {str(e)}")
    
    def save_kelahiran(self):
        """Save Kelahiran data"""
        try:
            # Get data from form
            nik = self.lahir_nik_var.get().strip()
            nama_bayi = self.lahir_nama_var.get().strip()
            jenis_kelamin = self.lahir_jk_var.get()
            nama_ayah = self.lahir_ayah_var.get().strip()
            nama_ibu = self.lahir_ibu_var.get().strip()
            desa = self.lahir_desa_var.get()
            no_kk = self.lahir_kk_var.get().strip()
            
            # Construct date string
            try:
                tanggal_lahir = f"{self.lahir_tgl_day.get()}/{self.lahir_tgl_month.get()}/{self.lahir_tgl_year.get()}"
            except:
                messagebox.showwarning("Perhatian", "Format tanggal lahir tidak valid!")
                return
            
            # Validation
            if not all([nik, nama_bayi, tanggal_lahir, jenis_kelamin, nama_ayah, nama_ibu, desa, no_kk]):
                messagebox.showwarning("Perhatian", "Semua field harus diisi!")
                return
            
            # Check if no_kk exists in KK data
            if no_kk not in self.data_penduduk["No KK"].values:
                messagebox.showwarning("Perhatian", "No KK tidak ditemukan dalam data KK!")
                return
            
            # Check if NIK is numeric and has correct length
            if not nik.isdigit() or len(nik) != 16:
                messagebox.showwarning("Perhatian", "NIK harus berupa 16 digit angka!")
                return
            
            # Check if NIK already exists
            if nik in self.data_kelahiran["NIK"].values:
                # Update existing record
                idx = self.data_kelahiran[self.data_kelahiran["NIK"] == nik].index[0]
                self.data_kelahiran.loc[idx] = [nik, nama_bayi, tanggal_lahir, jenis_kelamin,
                                             nama_ayah, nama_ibu, desa, no_kk]
                status_msg = "Data kelahiran berhasil diperbarui"
            else:
                # Add new record
                new_row = pd.DataFrame([[nik, nama_bayi, tanggal_lahir, jenis_kelamin,
                                        nama_ayah, nama_ibu, desa, no_kk]],
                                     columns=["NIK", "Nama_Bayi", "Tanggal_Lahir", "Jenis_Kelamin",
                                           "Nama_Ayah", "Nama_Ibu", "Desa", "No_KK"])
                self.data_kelahiran = pd.concat([self.data_kelahiran, new_row], ignore_index=True)
                
                # Also add to anggota keluarga
                tgl_lahir = datetime.strptime(tanggal_lahir, "%d/%m/%Y")
                umur = (datetime.now() - tgl_lahir).days // 365
                
                new_anggota = pd.DataFrame([[no_kk, nik, nama_bayi, "Anak", jenis_kelamin,
                                           tanggal_lahir, umur, "Tidak Sekolah", "Tidak Bekerja"]],
                                         columns=["No KK", "NIK", "Nama", "Hubungan", "Jenis_Kelamin",
                                               "Tanggal_Lahir", "Umur", "Pendidikan", "Pekerjaan"])
                self.data_anggota = pd.concat([self.data_anggota, new_anggota], ignore_index=True)
                
                # Update jumlah anggota in KK data
                kk_idx = self.data_penduduk[self.data_penduduk["No KK"] == no_kk].index[0]
                current_count = self.data_penduduk.loc[kk_idx, "Jumlah Anggota"]
                if isinstance(current_count, str):
                    current_count = int(current_count.strip() or "0")
                self.data_penduduk.loc[kk_idx, "Jumlah Anggota"] = current_count + 1
                
                status_msg = "Data kelahiran baru berhasil ditambahkan"
            
            # Refresh UI
            self.refresh_all_data()
            self.update_dashboard()
            self.reset_kelahiran_form()
            
            # Update status bar
            self.status_var.set(status_msg)
            messagebox.showinfo("Sukses", status_msg)
            
        except Exception as e:
            messagebox.showerror("Error", f"Gagal menyimpan data kelahiran: {str(e)}")
            print(f"Error saving kelahiran: {str(e)}")
    
    def reset_kelahiran_form(self):
        """Reset Kelahiran form"""
        self.lahir_nik_var.set("")
        self.lahir_nama_var.set("")
        self.lahir_jk_var.set("Laki-laki")
        self.lahir_ayah_var.set("")
        self.lahir_ibu_var.set("")
        self.lahir_desa_var.set("Desa A")
        self.lahir_kk_var.set("")
        
        # Reset date to today
        today = datetime.now()
        self.lahir_tgl_day.set(str(today.day).zfill(2))
        self.lahir_tgl_month.set(str(today.month).zfill(2))
        self.lahir_tgl_year.set(str(today.year))
    
    def load_kelahiran_to_form(self, event):
        """Load selected kelahiran data to form"""
        # Get selected item
        selected_item = self.kelahiran_tree.selection()
        if not selected_item:
            return
        
        # Get values
        values = self.kelahiran_tree.item(selected_item[0], "values")
        
        # Set form values
        self.lahir_nik_var.set(values[0])
        self.lahir_nama_var.set(values[1])
        self.lahir_jk_var.set(values[3])
        self.lahir_ayah_var.set(values[4])
        self.lahir_ibu_var.set(values[5])
        self.lahir_desa_var.set(values[6])
        self.lahir_kk_var.set(values[7])
        
        # Handle date
        try:
            tgl_parts = values[2].split('/')
            if len(tgl_parts) == 3:
                self.lahir_tgl_day.set(tgl_parts[0])
                self.lahir_tgl_month.set(tgl_parts[1])
                self.lahir_tgl_year.set(tgl_parts[2])
        except:
            # Reset to today if error
            today = datetime.now()
            self.lahir_tgl_day.set(str(today.day).zfill(2))
            self.lahir_tgl_month.set(str(today.month).zfill(2))
            self.lahir_tgl_year.set(str(today.year))
    
    def save_kematian(self):
        """Save Kematian data"""

        try:
            # Get data from form
            nik = self.mati_nik_var.get().strip()
            nama = self.mati_nama_var.get().strip()
            penyebab = self.mati_penyebab_var.get().strip()
            desa = self.mati_desa_var.get()
            no_kk = self.mati_kk_var.get().strip()
            usia = self.mati_usia_var.get().strip()
            
            # Construct date strings
            try:
                tanggal_lahir = f"{self.mati_tgl_lahir_day.get()}/{self.mati_tgl_lahir_month.get()}/{self.mati_tgl_lahir_year.get()}"
                tanggal_meninggal = f"{self.mati_tgl_meninggal_day.get()}/{self.mati_tgl_meninggal_month.get()}/{self.mati_tgl_meninggal_year.get()}"
            except:
                messagebox.showwarning("Perhatian", "Format tanggal tidak valid!")
                return
            
            # Validation
            if not all([nik, nama, tanggal_lahir, tanggal_meninggal, usia, penyebab, desa, no_kk]):
                messagebox.showwarning("Perhatian", "Semua field harus diisi!")
                return
            
            # Check if no_kk exists in KK data
            if no_kk not in self.data_penduduk["No KK"].values:
                messagebox.showwarning("Perhatian", "No KK tidak ditemukan dalam data KK!")
                return
            
            # Check if NIK is numeric and has correct length
            if not nik.isdigit() or len(nik) != 16:
                messagebox.showwarning("Perhatian", "NIK harus berupa 16 digit angka!")
                return
            
            # Add/update data
            new_row = pd.DataFrame([[nik, nama, tanggal_lahir, tanggal_meninggal, usia, penyebab, desa, no_kk]],
                                 columns=["NIK", "Nama", "Tanggal_Lahir", "Tanggal_Meninggal", 
                                        "Usia_Meninggal", "Penyebab", "Desa", "No_KK"])
            self.data_kematian = pd.concat([self.data_kematian, new_row], ignore_index=True)
            
            # Refresh data
            self.refresh_all_data()
            messagebox.showinfo("Sukses", "Data kematian berhasil disimpan!")
            
        except Exception as e:
            messagebox.showerror("Error", f"Gagal menyimpan data: {str(e)}")
    
    def reset_kematian_form(self):
        """Reset Kematian form"""
        self.mati_nik_var.set("")
        self.mati_nama_var.set("")
        # Reset tanggal lahir dan meninggal ke hari ini
        today = datetime.now()
        self.mati_tgl_lahir_day.set(str(today.day).zfill(2))
        self.mati_tgl_lahir_month.set(str(today.month).zfill(2))
        self.mati_tgl_lahir_year.set(str(today.year))
        self.mati_tgl_meninggal_day.set(str(today.day).zfill(2))
        self.mati_tgl_meninggal_month.set(str(today.month).zfill(2))
        self.mati_tgl_meninggal_year.set(str(today.year))
        self.mati_usia_var.set("")
        self.mati_penyebab_var.set("")
        self.mati_desa_var.set("Desa A")
        self.mati_kk_var.set("")
    
    def load_kematian_to_form(self, event):
        """Load selected kematian data to form"""
        # Ambil item yang dipilih
        selected_item = self.kematian_tree.selection()
        if not selected_item:
            return
        
        # Ambil nilai dari treeview
        values = self.kematian_tree.item(selected_item[0], "values")
        
        # Set nilai ke form
        self.mati_nik_var.set(values[0])
        self.mati_nama_var.set(values[1])
        self.mati_penyebab_var.set(values[5])
        self.mati_desa_var.set(values[6])
        self.mati_kk_var.set(values[7])
        self.mati_usia_var.set(values[4])
        
        # Handle tanggal lahir
        try:
            tgl_lahir_parts = values[2].split('/')
            if len(tgl_lahir_parts) == 3:
                self.mati_tgl_lahir_day.set(tgl_lahir_parts[0])
                self.mati_tgl_lahir_month.set(tgl_lahir_parts[1])
                self.mati_tgl_lahir_year.set(tgl_lahir_parts[2])
        except:
            today = datetime.now()
            self.mati_tgl_lahir_day.set(str(today.day).zfill(2))
            self.mati_tgl_lahir_month.set(str(today.month).zfill(2))
            self.mati_tgl_lahir_year.set(str(today.year))
        
        # Handle tanggal meninggal
        try:
            tgl_mati_parts = values[3].split('/')
            if len(tgl_mati_parts) == 3:
                self.mati_tgl_meninggal_day.set(tgl_mati_parts[0])
                self.mati_tgl_meninggal_month.set(tgl_mati_parts[1])
                self.mati_tgl_meninggal_year.set(tgl_mati_parts[2])
        except:
            today = datetime.now()
            self.mati_tgl_meninggal_day.set(str(today.day).zfill(2))
            self.mati_tgl_meninggal_month.set(str(today.month).zfill(2))
            self.mati_tgl_meninggal_year.set(str(today.year))
    
    def calculate_death_age(self):
        """Hitung usia saat meninggal berdasarkan tanggal lahir dan meninggal"""
        try:
            # Ambil komponen tanggal lahir
            birth_day = self.mati_tgl_lahir_day.get()
            birth_month = self.mati_tgl_lahir_month.get()
            birth_year = self.mati_tgl_lahir_year.get()
            
            # Ambil komponen tanggal meninggal
            death_day = self.mati_tgl_meninggal_day.get()
            death_month = self.mati_tgl_meninggal_month.get()
            death_year = self.mati_tgl_meninggal_year.get()
            
            # Buat objek tanggal
            birth_date = datetime.strptime(f"{birth_day}/{birth_month}/{birth_year}", "%d/%m/%Y")
            death_date = datetime.strptime(f"{death_day}/{death_month}/{death_year}", "%d/%m/%Y")
            
            # Hitung usia
            age = death_date.year - birth_date.year
            if (death_date.month, death_date.day) < (birth_date.month, birth_date.day):
                age -= 1
            
            # Update field usia
            self.mati_usia_var.set(str(age))
            
        except Exception as e:
            messagebox.showerror("Error", f"Gagal menghitung usia: {str(e)}")
    
    def generate_report(self):
        """Generate laporan berdasarkan opsi yang dipilih"""
        report_type = self.report_type_var.get()
        desa_filter = self.report_desa_var.get()
        
        try:
            # Hapus konten sebelumnya
            for widget in self.report_content_frame.winfo_children():
                widget.destroy()
            
            # Buat konten berdasarkan jenis laporan
            if report_type == "Rekapitulasi Penduduk":
                self.generate_population_summary(desa_filter)
            elif report_type == "Statistik per Desa":
                self.generate_desa_stats(desa_filter)
            elif report_type == "Data Kelahiran":
                self.generate_birth_report(desa_filter)
            elif report_type == "Data Kematian":
                self.generate_death_report(desa_filter)
            elif report_type == "Pendidikan":
                self.generate_education_report(desa_filter)
            elif report_type == "Pekerjaan":
                self.generate_job_report(desa_filter)
            elif report_type == "Semua Data":
                self.generate_full_report(desa_filter)
            
            # Update canvas
            self.report_canvas.update_idletasks()
            self.report_canvas.configure(scrollregion=self.report_canvas.bbox("all"))
            
        except Exception as e:
            messagebox.showerror("Error", f"Gagal membuat laporan: {str(e)}")
    
    def generate_population_summary(self, desa_filter):
        """Buat ringkasan populasi penduduk"""
        # Filter data berdasarkan desa
        if desa_filter == "Semua Desa":
            data_penduduk = self.data_penduduk
            data_anggota = self.data_anggota
        else:
            kk_in_desa = self.data_penduduk[self.data_penduduk["Desa"] == desa_filter]["No KK"]
            data_penduduk = self.data_penduduk[self.data_penduduk["Desa"] == desa_filter]
            data_anggota = self.data_anggota[self.data_anggota["No KK"].isin(kk_in_desa)]
        
        # Buat elemen laporan
        title = ttk.Label(self.report_content_frame,
                         text="REKAPITULASI PENDUDUK",
                         font=("Arial", 16, "bold"))
        title.pack(pady=10)
        
        # Statistik total
        stats_frame = ttk.Frame(self.report_content_frame)
        stats_frame.pack(pady=10)
        
        total_kk = len(data_penduduk)
        total_penduduk = len(data_anggota)
        total_laki = len(data_anggota[data_anggota["Jenis_Kelamin"] == "Laki-laki"])
        total_perempuan = total_penduduk - total_laki
        
        stats_text = f"""
        Total Kartu Keluarga: {total_kk}
        Total Penduduk: {total_penduduk}
        Laki-laki: {total_laki}
        Perempuan: {total_perempuan}
        """
        stats_label = ttk.Label(stats_frame, text=stats_text, font=("Arial", 12))
        stats_label.pack()
        
        # Distribusi usia
        age_frame = ttk.Frame(self.report_content_frame)
        age_frame.pack(pady=10)
        
        ttk.Label(age_frame, text="Distribusi Usia:",
                 font=("Arial", 12, "bold")).pack(anchor="w")
        
        data_anggota["Umur"] = data_anggota["Umur"].fillna(0).astype(int)
        balita = len(data_anggota[data_anggota["Umur"] < 5])
        anak = len(data_anggota[(data_anggota["Umur"] >= 5) & (data_anggota["Umur"] < 18)])
        dewasa = len(data_anggota[(data_anggota["Umur"] >= 18) & (data_anggota["Umur"] < 60)])
        lansia = len(data_anggota[data_anggota["Umur"] >= 60])
        
        age_text = f"""
        Balita (0-4 tahun): {balita}
        Anak (5-17 tahun): {anak}
        Dewasa (18-59 tahun): {dewasa}
        Lansia (60+ tahun): {lansia}
        """
        age_label = ttk.Label(age_frame, text=age_text, font=("Arial", 12))
        age_label.pack(anchor="w")
    
    def generate_birth_report(self, desa_filter):
        """Buat laporan kelahiran"""
        # Filter data
        if desa_filter == "Semua Desa":
            data = self.data_kelahiran
        else:
            data = self.data_kelahiran[self.data_kelahiran["Desa"] == desa_filter]
        
        # Buat tabel
        columns = ["No", "NIK Bayi", "Nama Bayi", "Tanggal Lahir", "Jenis Kelamin", "Desa"]
        self.create_report_table(columns, data[["NIK", "Nama_Bayi", "Tanggal_Lahir", "Jenis_Kelamin", "Desa"]])
    
    def create_report_table(self, columns, data):
        """Bantu membuat tabel dalam laporan"""
        # Frame untuk tabel
        table_frame = ttk.Frame(self.report_content_frame)
        table_frame.pack(fill="x", padx=10, pady=10)
        
        # Header
        for i, col in enumerate(columns):
            header = ttk.Label(table_frame, text=col,
                             font=("Arial", 10, "bold"),
                             borderwidth=1, relief="solid", width=15)
            header.grid(row=0, column=i, sticky="nsew")
        
        # Baris data
        for row_idx, (_, row) in enumerate(data.iterrows(), start=1):
            for col_idx, col in enumerate(columns[1:]):  # Skip kolom "No"
                cell_value = row[col] if col in row else ""
                cell = ttk.Label(table_frame, text=cell_value,
                                font=("Arial", 10),
                                borderwidth=1, relief="solid", width=15)
                cell.grid(row=row_idx, column=col_idx + 1, sticky="nsew")
            
            # Nomor urut
            no_label = ttk.Label(table_frame, text=str(row_idx),
                                font=("Arial", 10),
                                borderwidth=1, relief="solid", width=5)
            no_label.grid(row=row_idx, column=0, sticky="nsew")
        
        # Konfigurasi grid
        for i in range(len(columns)):
            table_frame.grid_columnconfigure(i, weight=1)
    
    def export_to_excel(self):
        """Ekspor data laporan ke Excel"""
        try:
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")]
            )
            if not file_path:
                return
            
            # Buat writer Excel
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                # Ekspor data sesuai jenis laporan
                report_type = self.report_type_var.get()
                desa_filter = self.report_desa_var.get()
                
                if report_type == "Rekapitulasi Penduduk":
                    # ... (implementasi serupa dengan save_excel)
                    pass
                elif report_type == "Data Kelahiran":
                    if desa_filter == "Semua Desa":
                        self.data_kelahiran.to_excel(writer, sheet_name='Laporan Kelahiran')
                    else:
                        filtered_data = self.data_kelahiran[self.data_kelahiran["Desa"] == desa_filter]
                        filtered_data.to_excel(writer, sheet_name='Laporan Kelahiran')
                
                # ... (implementasi untuk tipe laporan lainnya)
                
            messagebox.showinfo("Sukses", "Data berhasil diekspor ke Excel")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal mengekspor data: {str(e)}")
    
    def print_chart(self):
        """Cetak grafik ke printer"""
        try:
            # Simpan grafik ke file sementara
            file_path = "temp_chart.png"
            self.fig.savefig(file_path, dpi=300)
            
            # Buka gambar untuk dicetak
            img = Image.open(file_path)
            img.show()  # Bergantung pada sistem untuk mencetak
            
            messagebox.showinfo("Info", "Grafik siap dicetak melalui preview sistem")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal mencetak grafik: {str(e)}")
    
    def show_help(self):
        """Tampilkan petunjuk penggunaan"""
        help_text = """
        PETUNJUK PENGGUNAAN:
        1. Buka/Simpan file Excel melalui menu File
        2. Input data melalui tab yang relevan
        3. Gunakan tombol Simpan/Reset di setiap form
        4. Lihat data di tabel masing-masing tab
        5. Buat laporan di tab Laporan & Ekspor
        6. Ekspor data ke Excel atau cetak grafik
        """
        messagebox.showinfo("Petunjuk", help_text)
        def generate_desa_stats(self, desa_filter):
            """Buat statistik detail per desa"""
        if desa_filter != "Semua Desa":
            messagebox.showinfo("Info", "Filter desa tidak berlaku untuk laporan ini")
            return

        # Frame utama
        main_frame = ttk.Frame(self.report_content_frame)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Data untuk semua desa
        desa_list = self.data_penduduk["Desa"].unique()
        
        # Buat notebook untuk tampilan tab
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill="both", expand=True)

        # Tab Demografi
        demografi_frame = ttk.Frame(notebook)
        self.create_demografi_tab(demografi_frame, desa_list)
        notebook.add(demografi_frame, text="Demografi")

        # Tab Pendidikan
        pendidikan_frame = ttk.Frame(notebook)
        self.create_pendidikan_tab(pendidikan_frame, desa_list)
        notebook.add(pendidikan_frame, text="Pendidikan")

        # Tab Pekerjaan
        pekerjaan_frame = ttk.Frame(notebook)
        self.create_pekerjaan_tab(pekerjaan_frame, desa_list)
        notebook.add(pekerjaan_frame, text="Pekerjaan")

    def create_demografi_tab(self, parent, desa_list):
        # Data preparation
        data = []
        for desa in desa_list:
            kk_count = len(self.data_penduduk[self.data_penduduk["Desa"] == desa])
            anggota_count = len(self.data_anggota[self.data_anggota["No KK"].isin(
                self.data_penduduk[self.data_penduduk["Desa"] == desa]["No KK"]
            )])
            
            # Gender distribution
            laki = len(self.data_anggota[
                (self.data_anggota["No KK"].isin(self.data_penduduk[self.data_penduduk["Desa"] == desa]["No KK"])) &
                (self.data_anggota["Jenis_Kelamin"] == "Laki-laki")
            ])
            perempuan = anggota_count - laki
            
            # Age groups
            anggota_desa = self.data_anggota[self.data_anggota["No KK"].isin(
                self.data_penduduk[self.data_penduduk["Desa"] == desa]["No KK"]
            )]
            balita = len(anggota_desa[anggota_desa["Umur"] < 5])
            anak = len(anggota_desa[(anggota_desa["Umur"] >= 5) & (anggota_desa["Umur"] < 18)])
            dewasa = len(anggota_desa[(anggota_desa["Umur"] >= 18) & (anggota_desa["Umur"] < 60)])
            lansia = len(anggota_desa[anggota_desa["Umur"] >= 60])
            
            data.append([
                desa, kk_count, anggota_count, laki, perempuan,
                balita, anak, dewasa, lansia
            ])

        # Create table
        columns = ["Desa", "KK", "Penduduk", "Laki", "Perempuan", 
                 "Balita", "Anak", "Dewasa", "Lansia"]
        self.create_detailed_table(parent, columns, data)

    def create_pendidikan_tab(self, parent, desa_list):
        pendidikan_levels = ["Tidak Sekolah", "SD", "SMP", "SMA/SMK", 
                           "D1-D3", "S1", "S2", "S3"]
        data = []
        
        for desa in desa_list:
            anggota_desa = self.data_anggota[self.data_anggota["No KK"].isin(
                self.data_penduduk[self.data_penduduk["Desa"] == desa]["No KK"]
            )]
            counts = [
                len(anggota_desa[anggota_desa["Pendidikan"] == level])
                for level in pendidikan_levels
            ]
            data.append([desa] + counts)
        
        columns = ["Desa"] + pendidikan_levels
        self.create_detailed_table(parent, columns, data)

    def create_pekerjaan_tab(self, parent, desa_list):
        pekerjaan_types = ["Tidak Bekerja", "Petani", "Nelayan", "PNS",
                          "Swasta", "Wiraswasta", "Pelajar/Mahasiswa"]
        data = []
        
        for desa in desa_list:
            anggota_desa = self.data_anggota[self.data_anggota["No KK"].isin(
                self.data_penduduk[self.data_penduduk["Desa"] == desa]["No KK"]
            )]
            counts = [
                len(anggota_desa[anggota_desa["Pekerjaan"] == job])
                for job in pekerjaan_types
            ]
            data.append([desa] + counts)
        
        columns = ["Desa"] + pekerjaan_types
        self.create_detailed_table(parent, columns, data)

    def create_detailed_table(self, parent, columns, data):
        """Membuat tabel detail dengan scrollbar"""
        # Frame untuk canvas dan scrollbar
        container = ttk.Frame(parent)
        container.pack(fill="both", expand=True)

        # Canvas
        canvas = tk.Canvas(container)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        # Konfigurasi scroll
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Header
        for col_idx, col in enumerate(columns):
            header = ttk.Label(scrollable_frame, text=col, 
                             font=("Arial", 10, "bold"), 
                             borderwidth=1, relief="solid", width=15)
            header.grid(row=0, column=col_idx, sticky="nsew")

        # Data rows
        for row_idx, row_data in enumerate(data, start=1):
            for col_idx, value in enumerate(row_data):
                cell = ttk.Label(scrollable_frame, text=value,
                                font=("Arial", 10), 
                                borderwidth=1, relief="solid", width=15)
                cell.grid(row=row_idx, column=col_idx, sticky="nsew")

        # Layout
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def generate_death_report(self, desa_filter):
        """Buat laporan kematian"""
        # Filter data
        if desa_filter == "Semua Desa":
            data = self.data_kematian
        else:
            data = self.data_kematian[self.data_kematian["Desa"] == desa_filter]
        
        # Buat tabel
        columns = ["No", "NIK", "Nama", "Usia Meninggal", 
                 "Tanggal Meninggal", "Penyebab", "Desa"]
        self.create_report_table(columns, data[["NIK", "Nama", "Usia_Meninggal", 
                                             "Tanggal_Meninggal", "Penyebab", "Desa"]])

    def generate_education_report(self, desa_filter):
        """Buat laporan pendidikan"""
        # Filter data
        if desa_filter == "Semua Desa":
            data = self.data_anggota
        else:
            kk_in_desa = self.data_penduduk[self.data_penduduk["Desa"] == desa_filter]["No KK"]
            data = self.data_anggota[self.data_anggota["No KK"].isin(kk_in_desa)]
        
        # Hitung distribusi pendidikan
        pendidikan_counts = data["Pendidikan"].value_counts().reset_index()
        pendidikan_counts.columns = ["Pendidikan", "Jumlah"]
        
        # Buat tabel
        columns = ["No", "Pendidikan", "Jumlah"]
        self.create_report_table(columns, pendidikan_counts)

    def generate_job_report(self, desa_filter):
        """Buat laporan pekerjaan"""
        # Filter data
        if desa_filter == "Semua Desa":
            data = self.data_anggota
        else:
            kk_in_desa = self.data_penduduk[self.data_penduduk["Desa"] == desa_filter]["No KK"]
            data = self.data_anggota[self.data_anggota["No KK"].isin(kk_in_desa)]
        
        # Hitung distribusi pekerjaan
        pekerjaan_counts = data["Pekerjaan"].value_counts().reset_index()
        pekerjaan_counts.columns = ["Pekerjaan", "Jumlah"]
        
        # Buat tabel
        columns = ["No", "Pekerjaan", "Jumlah"]
        self.create_report_table(columns, pekerjaan_counts)

    def generate_full_report(self, desa_filter):
        """Buat laporan komprehensif"""
        main_frame = ttk.Frame(self.report_content_frame)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        notebook = ttk.Notebook(main_frame)
        notebook.pack(fill="both", expand=True)
        
        # Tab Penduduk
        penduduk_frame = ttk.Frame(notebook)
        self.generate_population_summary(desa_filter)
        notebook.add(penduduk_frame, text="Penduduk")
        
        # Tab Kelahiran
        lahir_frame = ttk.Frame(notebook)
        self.generate_birth_report(desa_filter)
        notebook.add(lahir_frame, text="Kelahiran")
        
        # Tab Kematian
        mati_frame = ttk.Frame(notebook)
        self.generate_death_report(desa_filter)
        notebook.add(mati_frame, text="Kematian")

    def export_to_excel(self):
        """Ekspor data laporan ke Excel"""
        try:
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")]
            )
            if not file_path:
                return

            report_type = self.report_type_var.get()
            desa_filter = self.report_desa_var.get()

            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                if report_type == "Rekapitulasi Penduduk":
                    # Ekspor data penduduk
                    if desa_filter == "Semua Desa":
                        self.data_penduduk.to_excel(writer, sheet_name='Data Penduduk')
                        self.data_anggota.to_excel(writer, sheet_name='Data Anggota')
                    else:
                        filtered_kk = self.data_penduduk[self.data_penduduk["Desa"] == desa_filter]
                        filtered_anggota = self.data_anggota[self.data_anggota["No KK"].isin(filtered_kk["No KK"])]
                        filtered_kk.to_excel(writer, sheet_name='Data Penduduk')
                        filtered_anggota.to_excel(writer, sheet_name='Data Anggota')
                
                elif report_type == "Data Kelahiran":
                    if desa_filter == "Semua Desa":
                        self.data_kelahiran.to_excel(writer, sheet_name='Laporan Kelahiran')
                    else:
                        filtered_data = self.data_kelahiran[self.data_kelahiran["Desa"] == desa_filter]
                        filtered_data.to_excel(writer, sheet_name='Laporan Kelahiran')
                
                elif report_type == "Data Kematian":
                    if desa_filter == "Semua Desa":
                        self.data_kematian.to_excel(writer, sheet_name='Laporan Kematian')
                    else:
                        filtered_data = self.data_kematian[self.data_kematian["Desa"] == desa_filter]
                        filtered_data.to_excel(writer, sheet_name='Laporan Kematian')
                
                elif report_type == "Semua Data":
                    self.data_penduduk.to_excel(writer, sheet_name='Data KK')
                    self.data_anggota.to_excel(writer, sheet_name='Data Anggota')
                    self.data_kelahiran.to_excel(writer, sheet_name='Data Kelahiran')
                    self.data_kematian.to_excel(writer, sheet_name='Data Kematian')

            messagebox.showinfo("Sukses", "Data berhasil diekspor ke Excel")
        except Exception as e:
            messagebox.showerror("Error", f"Gagal mengekspor data: {str(e)}")

    def show_about(self):
        """Tampilkan informasi tentang aplikasi"""
        about_text = """
        Sistem Informasi Kependudukan Kecamatan
        Versi 1.0
        Dibangun dengan Python dan Tkinter
        © 2024 Dinas Kependudukan dan Catatan Sipil
        """
        messagebox.showinfo("Tentang Aplikasi", about_text)

if __name__ == "__main__":
    root = ThemedTk(theme="arc")
    app = DataPendudukApp(root)
    root.mainloop()