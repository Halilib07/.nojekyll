import os
import math
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
import re
import threading
import time
import subprocess
import sys
from PIL import Image
import shutil
import concurrent.futures
from threading import Lock
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib
matplotlib.use('TkAgg')
from tkcalendar import DateEntry
import piexif
import folium
import datetime
from datetime import datetime as dt
from folium.plugins import MarkerCluster



# ==============================================
# GÖMÜLÜ HARİTA ANALİZ EKRANI SINIFI (RaporEkrani'den ÖNCE tanımlanmalı)
# ==============================================

class GömülüHaritaAnalizEkrani(tk.Frame):
    def __init__(self, parent, excel_dosyasi):
        super().__init__(parent)
        self.parent = parent
        self.excel_dosyasi = excel_dosyasi
        self.qt_app = None
        self.pyqt_widget = None  # PyQt widget referansı
        
        try:
            self.configure(bg="#2c3e50")
            
            # Değişkenler
            self.df_veri = None
            self.df_filtreli = None
            self.map_temp_html = None
            
            # PyQt6 kontrolü - LAZY olarak
            self.pyqt_available = False  # Başlangıçta False
            self._pyqt_checked = False  # Kontrol edildi mi?
            
            # Rapor klasörü ve HTML dosya yolu
            self.rapor_klasoru = self.get_rapor_klasoru()
            self.harita_html_path = os.path.join(self.rapor_klasoru, "RaporHarita.html")
            self.oncelik_toplam_var = tk.StringVar(value="0")  # YENİ EKLENDİ
            
            self.arayuz_olustur()
            self.verileri_yukle()
            
        except Exception as e:
            messagebox.showerror("Hata", f"Harita ekranı oluşturulamadı: {str(e)}")
            import traceback
            print(f"GömülüHaritaAnalizEkrani hatası: {traceback.format_exc()}")
    
    def check_pyqt_availability(self):
        """PyQt6'nın kurulu olup olmadığını kontrol eder - LAZY"""
        if self._pyqt_checked:
            return self.pyqt_available
            
        try:
            import sys
            from PyQt6.QtCore import Qt, QUrl
            from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout
            from PyQt6.QtWebEngineWidgets import QWebEngineView
            from PyQt6.QtWebEngineCore import QWebEngineSettings
            
            self.pyqt_available = True
            self._pyqt_checked = True
            return True
        except ImportError:
            self.pyqt_available = False
            self._pyqt_checked = True
            print("PyQt6 kütüphaneleri yüklü değil. Harita tarayıcıda açılacak.")
            return False
    
    def get_rapor_klasoru(self):
        """Masaüstündeki Rapor klasörünün yolunu alır"""
        try:
            desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
            rapor_klasoru = os.path.join(desktop_path, "Rapor")
            
            if not os.path.exists(rapor_klasoru):
                os.makedirs(rapor_klasoru)
            
            return rapor_klasoru
        except:
            import tempfile
            return tempfile.gettempdir()
    
    def arayuz_olustur(self):
        """Gömülü harita arayüzünü oluşturur - Filtreler ve İstatistikler SOLDA, Harita SAĞDA"""
        # Ana konteyner
        main_frame = tk.Frame(self, bg="#2c3e50")
        main_frame.pack(fill='both', expand=True, padx=0, pady=0)
        
        # Grid yapısı: sol panel ve sağ panel
        main_frame.grid_rowconfigure(0, weight=1)
        main_frame.grid_columnconfigure(0, weight=0)  # Sol panel: sabit genişlik
        main_frame.grid_columnconfigure(1, weight=1)  # Sağ panel: tüm kalan alan
        
        # === SOL PANEL: Filtreler ve İstatistikler -
        left_panel = tk.Frame(main_frame, bg="#2c3e50", width=320) 
        left_panel.grid(row=0, column=0, sticky='ns', padx=0, pady=0)
        left_panel.grid_propagate(False)
        
        # Sol panel içinde dikey düzen - SATIR EKLENDİ
        left_panel.grid_rowconfigure(0, weight=0)  # Filtreler
        left_panel.grid_rowconfigure(1, weight=0)  # İstatistikler
        left_panel.grid_rowconfigure(2, weight=0)  # Link Paneli - YENİ EKLENDİ
        left_panel.grid_rowconfigure(3, weight=1)  # Boşluk
        
        # 1. FİLTRELER FRAME - DAHA DAR
        filter_frame = tk.LabelFrame(left_panel, text="🔍 FİLTRELER", 
                                    font=('Segoe UI', 10, 'bold'),  # Yazı boyutu küçültüldü
                                    fg='white', bg="#34495e", padx=12, pady=12)  # Padding azaltıldı
        filter_frame.grid(row=0, column=0, sticky='ew', pady=(5, 5), padx=5)
        
        # İlçe filtresi - KOMBOKS GENİŞLİĞİ DARALTILDI
        tk.Label(filter_frame, text="İLÇE:", font=('Segoe UI', 9, 'bold'),  # Font küçültüldü
                bg="#34495e", fg='#3498db').pack(anchor='w', pady=(0, 3))  # Padding azaltıldı
        
        self.ilce_var = tk.StringVar(value="TÜMÜ")
        self.ilce_combo = ttk.Combobox(filter_frame, textvariable=self.ilce_var, 
                                      state="readonly", width=22)  # Genişlik 25'ten 22'ye
        self.ilce_combo.pack(fill='x', pady=(0, 8))
        
        # Hat filtresi - KOMBOKS GENİŞLİĞİ DARALTILDI
        tk.Label(filter_frame, text="HAT ADI:", font=('Segoe UI', 9, 'bold'),
                bg="#34495e", fg='#3498db').pack(anchor='w', pady=(0, 3))
        
        self.hat_var = tk.StringVar(value="TÜMÜ")
        self.hat_combo = ttk.Combobox(filter_frame, textvariable=self.hat_var,
                                     state="readonly", width=22)  # Genişlik 25'ten 22'ye
        self.hat_combo.pack(fill='x', pady=(0, 8))

        # Yapıldı mı? filtresi - KOMBOKS GENİŞLİĞİ DARALTILDI
        tk.Label(filter_frame, text="YAPILDI MI?:", font=('Segoe UI', 9, 'bold'),
                bg="#34495e", fg='#3498db').pack(anchor='w', pady=(0, 3))
        
        self.yapildi_var = tk.StringVar(value="TÜMÜ")
        self.yapildi_combo = ttk.Combobox(filter_frame, textvariable=self.yapildi_var,
                                         state="readonly", width=22)  # Genişlik 25'ten 22'ye
        self.yapildi_combo['values'] = ['TÜMÜ', 'YAPILDI', 'YAPILMADI']
        self.yapildi_combo.pack(fill='x', pady=(0, 10))

        
        # Harita ayarları - DAHA KOMPAKT
        settings_frame = tk.LabelFrame(filter_frame, text="AYARLAR",
                                      font=('Segoe UI', 9, 'bold'),  # Font küçültüldü
                                      bg="#2c3e50", fg='white', padx=8, pady=5)  # Padding azaltıldı
        settings_frame.pack(fill='x', pady=(5, 8))
        
        self.kumelenme_var = tk.BooleanVar(value=True)
        kumelenme_check = tk.Checkbutton(settings_frame, text="Noktaları kümelenmiş göster",
                                        variable=self.kumelenme_var,
                                        font=('Segoe UI', 8),  # Font küçültüldü
                                        bg="#2c3e50", fg='white',
                                        selectcolor="#2c3e50")
        kumelenme_check.pack(anchor='w', pady=1)
        
        self.oncelik_dagilimi_var = tk.BooleanVar(value=True)
        oncelik_check = tk.Checkbutton(settings_frame, text="Sadece öncelikli direkleri göster",
                                      variable=self.oncelik_dagilimi_var,
                                      font=('Segoe UI', 8),  # Font küçültüldü
                                      bg="#2c3e50", fg='white',
                                      selectcolor="#2c3e50")
        oncelik_check.pack(anchor='w', pady=1)
        
        # Butonlar - DAHA KOMPAKT
        button_frame = tk.Frame(filter_frame, bg="#34495e")
        button_frame.pack(fill='x', pady=(8, 0))
        
        tk.Button(button_frame, text="🔍 FİLTRELE", 
                 command=self.filtrele,
                 bg="#27ae60", fg='white', font=('Segoe UI', 9, 'bold'),  # Font küçültüldü
                 height=1).pack(side='left', fill='x', expand=True, padx=(0, 2))  # Height 2'den 1'e
        
        tk.Button(button_frame, text="🔄 SIFIRLA", 
                 command=self.filtreleri_sifirla,
                 bg="#e74c3c", fg='white', font=('Segoe UI', 9, 'bold'),
                 height=1).pack(side='left', fill='x', expand=True, padx=2)
        
        # 2. İSTATİSTİKLER FRAME - DAHA KOMPAKT
        stat_frame = tk.LabelFrame(left_panel, text="📊 İSTATİSTİKLER", 
                                  font=('Segoe UI', 10, 'bold'),  # Font küçültüldü
                                  fg='white', bg="#34495e", padx=8, pady=8)  # Padding azaltıldı
        stat_frame.grid(row=1, column=0, sticky='ew', pady=(0, 5), padx=5)
        
        # İstatistik grid
        stats_grid = tk.Frame(stat_frame, bg="#34495e")
        stats_grid.pack(fill='x')
        
        # Grid yapısı
        for i in range(11):
            stats_grid.grid_rowconfigure(i, weight=0)
        stats_grid.grid_columnconfigure(0, weight=1)
        stats_grid.grid_columnconfigure(1, weight=0)
        
        row = 0
        
        # Genel İstatistikler
        tk.Label(stats_grid, text="GENEL İSTATİSTİKLER", 
                font=('Segoe UI', 9, 'bold', 'underline'),  # Font küçültüldü
                bg="#34495e", fg="#f39c12").grid(row=row, column=0, columnspan=2, sticky='w', pady=(0, 3))
        row += 1
        
        # Toplam kayıt
        tk.Label(stats_grid, text="Toplam Direk Sayısı:", font=('Segoe UI', 8, 'bold'),  # Font küçültüldü
                bg="#34495e", fg='white', anchor='w').grid(row=row, column=0, sticky='w', pady=1)
        self.toplam_kayit_var = tk.StringVar(value="0")
        tk.Label(stats_grid, textvariable=self.toplam_kayit_var, 
                font=('Segoe UI', 8, 'bold'), bg="#34495e", fg="#3498db",
                anchor='w').grid(row=row, column=1, sticky='w', pady=1)
        row += 1
        
        # ÖNCELİK TOPLAMI
        tk.Label(stats_grid, text="Öncelik Toplamı:", font=('Segoe UI', 8, 'bold'),
                bg="#34495e", fg='white', anchor='w').grid(row=row, column=0, sticky='w', pady=1)
        self.oncelik_toplam_var = tk.StringVar(value="0")
        tk.Label(stats_grid, textvariable=self.oncelik_toplam_var, 
                font=('Segoe UI', 8, 'bold'), bg="#34495e", fg="#f39c12",
                anchor='w').grid(row=row, column=1, sticky='e', pady=1)
        row += 1
        
        # Ayırıcı
        tk.Frame(stats_grid, height=1, bg="#7f8c8d").grid(row=row, column=0, columnspan=2, sticky='ew', pady=(5, 5))
        row += 1
        
        # Öncelik Dağılımı
        tk.Label(stats_grid, text="ÖNCELİK DAĞILIMI", 
                font=('Segoe UI', 9, 'bold', 'underline'),
                bg="#34495e", fg="#f39c12").grid(row=row, column=0, columnspan=2, sticky='w', pady=(0, 3))
        row += 1
        
        # Çok acil
        tk.Label(stats_grid, text="🔴 Çok Acil:", font=('Segoe UI', 8, 'bold'),  # Font küçültüldü
                bg="#34495e", fg='white', anchor='w').grid(row=row, column=0, sticky='w', pady=1)
        self.cok_acil_var = tk.StringVar(value="0")
        tk.Label(stats_grid, textvariable=self.cok_acil_var, 
                font=('Segoe UI', 8, 'bold'), bg="#34495e", fg="#e74c3c",
                anchor='e').grid(row=row, column=1, sticky='e', pady=1)
        row += 1

        # Acil
        tk.Label(stats_grid, text="🟠 Acil:", font=('Segoe UI', 8, 'bold'),
                bg="#34495e", fg='white', anchor='w').grid(row=row, column=0, sticky='w', pady=1)
        self.acil_var = tk.StringVar(value="0")
        tk.Label(stats_grid, textvariable=self.acil_var, 
                font=('Segoe UI', 8, 'bold'), bg="#34495e", fg="#f39c12",
                anchor='e').grid(row=row, column=1, sticky='e', pady=1)
        row += 1

        # Normal
        tk.Label(stats_grid, text="🔵 Normal:", font=('Segoe UI', 8, 'bold'),
                bg="#34495e", fg='white', anchor='w').grid(row=row, column=0, sticky='w', pady=1)
        self.normal_var = tk.StringVar(value="0")
        tk.Label(stats_grid, textvariable=self.normal_var, 
                font=('Segoe UI', 8, 'bold'), bg="#34495e", fg="#3498db",
                anchor='e').grid(row=row, column=1, sticky='e', pady=1)
        row += 1

        # Bekleyebilir
        tk.Label(stats_grid, text="🟢 Bekleyebilir:", font=('Segoe UI', 8, 'bold'),
                bg="#34495e", fg='white', anchor='w').grid(row=row, column=0, sticky='w', pady=1)
        self.bekleyebilir_var = tk.StringVar(value="0")
        tk.Label(stats_grid, textvariable=self.bekleyebilir_var, 
                font=('Segoe UI', 8, 'bold'), bg="#34495e", fg="#27ae60",
                anchor='e').grid(row=row, column=1, sticky='e', pady=1)
        row += 1
        
        # Ayırıcı
        tk.Frame(stats_grid, height=1, bg="#7f8c8d").grid(row=row, column=0, columnspan=2, sticky='ew', pady=(5, 5))
        row += 1
        
        # Durum bilgisi
        self.bilgi_var = tk.StringVar(value="📁 Veriler yükleniyor...")
        bilgi_label = tk.Label(stats_grid, textvariable=self.bilgi_var,
                              font=('Segoe UI', 8), bg="#34495e", fg="#f39c12",  # Font küçültüldü
                              wraplength=250, justify='left', anchor='w')  # Wraplength azaltıldı
        bilgi_label.grid(row=row, column=0, columnspan=2, pady=(3, 0), sticky='w')
        
        # 3. LİNK PANELİ - YENİ EKLENDİ
        link_frame = tk.LabelFrame(left_panel, text="🔗 WEB LİNK", 
                                  font=('Segoe UI', 10, 'bold'),
                                  fg='white', bg="#34495e", padx=8, pady=8)
        link_frame.grid(row=2, column=0, sticky='ew', pady=(0, 5), padx=5)
        
        # Link oluşturma fonksiyonunu bağla
        self.hat_var.trace('w', self.guncelle_link)  # Hat değiştiğinde link güncellenecek
        
        # Link etiketi
        self.link_var = tk.StringVar(value="https://halilib07.github.io/.nojekyll/")
        link_label = tk.Label(link_frame, textvariable=self.link_var,
                             font=('Segoe UI', 8), bg="#34495e", fg="#3498db",
                             wraplength=260, justify='left', anchor='w', cursor="hand2")
        link_label.pack(fill='x', pady=(0, 5))
        
        # Link'e tıklama özelliği ekle
        def link_tikla(event):
            link = self.link_var.get()
            if link:
                import webbrowser
                webbrowser.open(link)
        
        link_label.bind("<Button-1>", link_tikla)
        
        # Butonlar frame
        link_btn_frame = tk.Frame(link_frame, bg="#34495e")
        link_btn_frame.pack(fill='x')
        
        # Kopyala butonu
        copy_btn = tk.Button(link_btn_frame, text="📋 Kopyala", 
                            command=self.link_kopyala,
                            bg="#f39c12", fg='white', font=('Segoe UI', 8, 'bold'),
                            height=1, width=10)
        copy_btn.pack(side='left', padx=(0, 5))
        
        # Tarayıcıda aç butonu
        open_btn = tk.Button(link_btn_frame, text="🌐 Aç", 
                            command=self.link_ac,
                            bg="#3498db", fg='white', font=('Segoe UI', 8, 'bold'),
                            height=1, width=10)
        open_btn.pack(side='left')
        
        # Boşluk
        empty_space = tk.Frame(left_panel, bg="#2c3e50")
        empty_space.grid(row=3, column=0, sticky='nsew')
        
        # === SAĞ PANEL: Harita Görünümü - DAHA GENİŞ ===
        right_panel = tk.Frame(main_frame, bg="#2c3e50")
        right_panel.grid(row=0, column=1, sticky='nsew', padx=0, pady=0)
        right_panel.grid_rowconfigure(0, weight=1)
        right_panel.grid_columnconfigure(0, weight=1)
        
        # Harita frame
        harita_frame = tk.Frame(right_panel, bg="#34495e")
        harita_frame.grid(row=0, column=0, sticky='nsew', padx=0, pady=0)
        harita_frame.grid_rowconfigure(0, weight=1)
        harita_frame.grid_columnconfigure(0, weight=1)
        
        # Harita gömme çerçevesi
        self.map_frame = tk.Frame(harita_frame, bg="black")
        self.map_frame.grid(row=0, column=0, sticky='nsew', padx=2, pady=2)
        self.map_frame.grid_rowconfigure(0, weight=1)
        self.map_frame.grid_columnconfigure(0, weight=1)
        
        # Başlangıçta bilgi göster
        self.harita_bilgi_label = tk.Label(self.map_frame,
                                         text="Harita yükleniyor...\n\n"
                                              "Veriler işleniyor, lütfen bekleyin.",
                                         font=('Segoe UI', 12),
                                         bg="black", fg="white",
                                         justify='center')
        self.harita_bilgi_label.pack(expand=True)



    def guncelle_link(self, *args):
        """Linki günceller - seçili hat adına göre"""
        try:
            # Sabit URL kısmı
            base_url = "https://halilib07.github.io/.nojekyll/"
            
            # Seçili hat adını al
            secili_hat = self.hat_var.get()
            
            # Link oluştur
            if secili_hat != "TÜMÜ" and secili_hat != "GENEL":
                # Dosya adı için uygun formata çevir
                dosya_adi = self.hat_to_filename(secili_hat)
                link = f"{base_url}{dosya_adi}"
            else:
                link = base_url  # Sadece base URL göster
                
            self.link_var.set(link)
            
        except Exception as e:
            print(f"Link güncelleme hatası: {e}")
            self.link_var.set("https://halilib07.github.io/.nojekyll/")

    def hat_to_filename(self, hat_adi):
        """Hat adını dosya adı formatına çevirir"""
        try:
            # Özel karakterleri temizle
            import re
            
            # Tarih formatını bul (örn: 2025.12.03_724280_H02_Cerdin Dm Cerdin Erdaş çıkışı)
            # Sadece tarih ve diğer bilgiler varsa olduğu gibi kullan
            if hat_adi and len(hat_adi) > 10:
                # Boşlukları koru, sadece URL'de sorun çıkaracak karakterleri temizle
                cleaned = hat_adi.strip()
                # URL uygun karakterlere çevir
                cleaned = re.sub(r'[<>:"/\\|?*]', '', cleaned)  # Dosya sistemi için geçersiz karakterleri kaldır
                return cleaned
            else:
                return hat_adi.strip()
                
        except Exception as e:
            print(f"Hat adı dönüştürme hatası: {e}")
            return hat_adi

    def link_kopyala(self):
        """Linki panoya kopyalar"""
        try:
            link = self.link_var.get()
            if link:
                self.clipboard_clear()
                self.clipboard_append(link)
                self.update()  # Clipboard'ı güncelle
                
                # Geçici bilgi mesajı
                original_text = self.bilgi_var.get()
                self.bilgi_var.set("✅ Link panoya kopyalandı!")
                
                # 3 saniye sonra orijinal metne dön
                self.after(3000, lambda: self.bilgi_var.set(original_text))
        except Exception as e:
            print(f"Link kopyalama hatası: {e}")
            self.bilgi_var.set(f"❌ Kopyalama hatası: {str(e)}")

    def link_ac(self):
        """Linki tarayıcıda açar"""
        try:
            link = self.link_var.get()
            if link and link != "https://halilib07.github.io/.nojekyll/":
                import webbrowser
                webbrowser.open(link)
            else:
                # Sadece base URL ise uyarı göster
                if link == "https://halilib07.github.io/.nojekyll/":
                    self.bilgi_var.set("⚠️ Lütfen önce bir hat seçin!")
                else:
                    self.bilgi_var.set("⚠️ Geçerli bir link yok!")
        except Exception as e:
            print(f"Link açma hatası: {e}")
            self.bilgi_var.set(f"❌ Link açma hatası: {str(e)}")


    
    def verileri_yukle(self):
        """Birleştirilmiş_Tespit excel'inden verileri yükler - DÜZELTİLMİŞ"""
        try:
            # Fotoğraf Klasörleri yerine Birleştirilmiş_Tespit sheet'ini oku
            self.df_veri = pd.read_excel(self.excel_dosyasi, sheet_name='Birleştirilmiş_Tespit')
            
            print(f"DEBUG: Yüklenen sütunlar: {list(self.df_veri.columns)}")
            
            # Sütun isimlerini temizle (boşlukları kaldır)
            self.df_veri.columns = self.df_veri.columns.str.strip()
            
            # Sütun isimlerini standartlaştır
            rename_dict = {}
            
            # İlçe sütunu (Aob)
            if 'İlçe' in self.df_veri.columns:
                rename_dict['İlçe'] = 'Aob'
            elif 'İLÇE' in self.df_veri.columns:
                rename_dict['İLÇE'] = 'Aob'
            
            # Enlem sütunu -> LAT
            if 'Enlem' in self.df_veri.columns:
                rename_dict['Enlem'] = 'LAT'
            elif 'ENLEM' in self.df_veri.columns:
                rename_dict['ENLEM'] = 'LAT'
            
            # Boylam sütunu -> LON
            if 'Boylam' in self.df_veri.columns:
                rename_dict['Boylam'] = 'LON'
            elif 'BOYLAM' in self.df_veri.columns:
                rename_dict['BOYLAM'] = 'LON'
            
            # Direk No sütunu
            if 'Direk No' in self.df_veri.columns:
                rename_dict['Direk No'] = 'Direk No'
            elif 'DİREK NO' in self.df_veri.columns:
                rename_dict['DİREK NO'] = 'Direk No'
            
            # Öncelik sütunu
            if 'Öncelik' in self.df_veri.columns:
                rename_dict['Öncelik'] = 'Öncelik'
            elif 'ÖNCELİK' in self.df_veri.columns:
                rename_dict['ÖNCELİK'] = 'Öncelik'
            
            # Tespit Notu sütunu
            if 'Tespit Notu' in self.df_veri.columns:
                rename_dict['Tespit Notu'] = 'Tespit Notu'
            elif 'TESPİT NOTU' in self.df_veri.columns:
                rename_dict['TESPİT NOTU'] = 'Tespit Notu'
            
            # Yapıldı mı? sütunu
            if 'Yapıldı mı?' in self.df_veri.columns:
                rename_dict['Yapıldı mı?'] = 'Yapıldı mı?'
            elif 'YAPILDI MI?' in self.df_veri.columns:
                rename_dict['YAPILDI MI?'] = 'Yapıldı mı?'
            
            # 🟢 EN ÖNEMLİ DÜZELTME: Hat Adı sütunu
            # "Hat Adı" sütununu kontrol et, yoksa oluştur
            if 'Hat Adı' not in self.df_veri.columns:
                # Alternatif isimlerle kontrol et
                if 'Hat adı' in self.df_veri.columns:
                    rename_dict['Hat adı'] = 'Hat Adı'
                    print("DEBUG: 'Hat adı' sütunu 'Hat Adı' olarak değiştirildi")
                elif 'HAT ADI' in self.df_veri.columns:
                    rename_dict['HAT ADI'] = 'Hat Adı'
                    print("DEBUG: 'HAT ADI' sütunu 'Hat Adı' olarak değiştirildi")
                else:
                    # Hiçbir hat sütunu yoksa, ilk satırı kullanarak oluştur
                    print("DEBUG: 'Hat Adı' sütunu bulunamadı, oluşturuluyor...")
                    # Boş bir Hat Adı sütunu oluştur
                    self.df_veri['Hat Adı'] = ''
            
            # Sütunları yeniden adlandır
            if rename_dict:
                self.df_veri.rename(columns=rename_dict, inplace=True)
                print(f"DEBUG: Yeniden adlandırılan sütunlar: {rename_dict}")
            
            # 🟢 HAT ADI SÜTUNUNU KONTROL ET
            print(f"DEBUG: Mevcut sütunlar: {list(self.df_veri.columns)}")
            
            # Hat Adı sütununu doldur (eğer boşsa)
            if 'Hat Adı' in self.df_veri.columns:
                # Boş değerleri kontrol et
                empty_hat_count = self.df_veri['Hat Adı'].isna().sum()
                if empty_hat_count > 0:
                    print(f"DEBUG: {empty_hat_count} boş Hat Adı değeri var")
                    
                    # Boş değerleri "GENEL" ile doldur
                    self.df_veri['Hat Adı'] = self.df_veri['Hat Adı'].fillna('GENEL')
                    
                    # String olmayan değerleri string'e çevir
                    self.df_veri['Hat Adı'] = self.df_veri['Hat Adı'].astype(str)
                    
                    # "nan" stringlerini "GENEL" ile değiştir
                    self.df_veri['Hat Adı'] = self.df_veri['Hat Adı'].replace('nan', 'GENEL')
                
                # Unique hat adlarını kontrol et
                unique_hats = self.df_veri['Hat Adı'].unique()
                print(f"DEBUG: {len(unique_hats)} unique hat adı bulundu")
                print(f"DEBUG: Hat adı örnekleri: {list(unique_hats[:10])}")
            else:
                print("DEBUG: 'Hat Adı' sütunu hala bulunamadı, oluşturuluyor...")
                self.df_veri['Hat Adı'] = 'GENEL'
            
            # Koordinat temizleme
            self.df_veri['LAT'] = pd.to_numeric(self.df_veri['LAT'], errors='coerce')
            self.df_veri['LON'] = pd.to_numeric(self.df_veri['LON'], errors='coerce')
            
            # İlçe verilerini temizle
            if 'Aob' in self.df_veri.columns:
                self.df_veri['Aob'] = self.df_veri['Aob'].astype(str).str.strip()
            
            # Direk No temizle
            if 'Direk No' in self.df_veri.columns:
                self.df_veri['Direk No'] = self.df_veri['Direk No'].astype(str).str.strip()
            
            # Öncelik temizle
            if 'Öncelik' in self.df_veri.columns:
                self.df_veri['Öncelik'] = self.df_veri['Öncelik'].astype(str).str.strip()
            
            # Veri kalitesi kontrolü
            koordinatli_kayit = self.df_veri['LAT'].notna().sum()
            print(f"DEBUG: Koordinatlı kayıt sayısı: {koordinatli_kayit}/{len(self.df_veri)}")
            print(f"DEBUG: İlçe sayısı: {self.df_veri['Aob'].nunique()}")
            print(f"DEBUG: Öncelik dağılımı: {self.df_veri['Öncelik'].value_counts()}")
            
            self.df_filtreli = self.df_veri.copy()
            
            # Listeleri doldur
            self.listeleri_doldur()
            
            # İstatistikleri güncelle
            self.istatistikleri_guncelle()
            
            # Başlangıç haritasını oluştur
            self.harita_olustur()
            
        except Exception as e:
            self.bilgi_var.set(f"❌ Veri yükleme hatası: {str(e)}")
            import traceback
            print(f"Veri yükleme hatası detayı: {traceback.format_exc()}")
    
    def listeleri_doldur(self):
        """Filtre listelerini doldurur - DÜZELTİLMİŞ"""
        try:
            # İlçe listesi (Aob sütunu)
            if 'Aob' in self.df_veri.columns:
                # NaN değerleri kaldır ve benzersiz değerleri al
                ilce_values = self.df_veri['Aob'].dropna().unique()
                # Boş string'leri ve 'nan' değerlerini filtrele
                ilce_list = ['TÜMÜ'] + sorted([str(x) for x in ilce_values if str(x).strip() and str(x).lower() != 'nan'])
                self.ilce_combo['values'] = ilce_list
                print(f"DEBUG: İlçe listesi oluşturuldu: {len(ilce_list)-1} ilçe")
            else:
                print("HATA: 'Aob' sütunu bulunamadı!")
                self.ilce_combo['values'] = ['TÜMÜ']
            
            # 🟢 HAT ADI LİSTESİ - DÜZELTİLMİŞ
            if 'Hat Adı' in self.df_veri.columns:
                # NaN değerleri 'GENEL' ile doldur
                self.df_veri['Hat Adı'] = self.df_veri['Hat Adı'].fillna('GENEL')
                
                # String olmayan değerleri string'e çevir
                self.df_veri['Hat Adı'] = self.df_veri['Hat Adı'].astype(str)
                
                # "nan" stringlerini "GENEL" ile değiştir
                self.df_veri['Hat Adı'] = self.df_veri['Hat Adı'].replace('nan', 'GENEL')
                
                # Unique değerleri al
                hat_values = self.df_veri['Hat Adı'].dropna().unique()
                
                # Boş string'leri filtrele
                hat_list = ['TÜMÜ'] + sorted([str(x) for x in hat_values if str(x).strip() and str(x).lower() != 'nan'])
                
                print(f"DEBUG: Hat listesi oluşturuldu: {len(hat_list)-1} hat")
                print(f"DEBUG: Hat örnekleri: {hat_list[:10]}")
            else:
                # Hat bilgisi yoksa, "GENEL" olarak işaretle
                hat_list = ['TÜMÜ', 'GENEL']
                if 'Hat Adı' in self.df_veri.columns:
                    self.df_veri['Hat Adı'] = 'GENEL'
            
            self.hat_combo['values'] = hat_list
            
            # Yapıldı mı? listesini ayarla
            self.yapildi_combo['values'] = ['TÜMÜ', 'YAPILDI', 'YAPILMADI']
            
            # İlçe değiştiğinde hatları güncelle
            self.ilce_var.trace('w', self.ilce_degisti)
            
        except Exception as e:
            print(f"Liste doldurma hatası: {e}")
            import traceback
            traceback.print_exc()
    
    def ilce_degisti(self, *args):
        """İlçe değiştiğinde hat listesini günceller"""
        try:
            secili_ilce = self.ilce_var.get()
            
            if secili_ilce == "TÜMÜ":
                # Tüm hatları göster
                if 'Hat Adı' in self.df_veri.columns and self.df_veri['Hat Adı'].notna().any():
                    hat_values = self.df_veri['Hat Adı'].dropna().unique()
                    hat_list = ['TÜMÜ'] + sorted([str(x) for x in hat_values if str(x).strip()])
                else:
                    hat_list = ['TÜMÜ']  # GENEL'i kaldırdım
            else:
                # Seçili ilçeye ait hatları göster
                if 'Aob' in self.df_veri.columns and 'Hat Adı' in self.df_veri.columns:
                    df_filtreli = self.df_veri[self.df_veri['Aob'] == secili_ilce]
                    
                    # Hat adı sütunu yoksa oluştur
                    if 'Hat Adı' not in df_filtreli.columns:
                        df_filtreli['Hat Adı'] = ''
                    
                    # NaN değerleri filtrele
                    hat_values = df_filtreli['Hat Adı'].dropna().unique()
                    
                    if len(hat_values) > 0:
                        hat_list = ['TÜMÜ'] + sorted([str(x) for x in hat_values if str(x).strip()])
                    else:
                        hat_list = ['TÜMÜ']  # GENEL'i kaldırdım
                else:
                    hat_list = ['TÜMÜ']  # GENEL'i kaldırdım
            
            # Hat listesini güncelle
            self.hat_combo['values'] = hat_list
            
            # Eğer mevcut seçim yeni listede yoksa, TÜMÜ yap
            current_selection = self.hat_var.get()
            if current_selection not in hat_list:
                self.hat_var.set("TÜMÜ")
            
            print(f"DEBUG: İlçe değişti: {secili_ilce}, Hat listesi: {hat_list}")
            
        except Exception as e:
            print(f"İlçe değişim hatası: {e}")
            import traceback
            traceback.print_exc()
    
    def istatistikleri_guncelle(self):
        """İstatistikleri günceller - Fotoğraf Klasörleri sayfasından direk sayısını alır"""
        try:
            if self.df_veri is None:
                return
                
            # Filtre değerlerini al
            secili_ilce = self.ilce_var.get()
            secili_hat = self.hat_var.get()
            
            # 1. FOTOĞRAF KLASÖRLERİ SAYFASINI OKU
            df_fotograf = pd.read_excel(self.excel_dosyasi, sheet_name='Fotoğraf Klasörleri')
            
            # Sütun isimlerini temizle
            df_fotograf.columns = df_fotograf.columns.str.strip()
            
            # Sütun isimlerini standartlaştır
            if 'Aob' not in df_fotograf.columns:
                # Alternatif sütun isimlerini kontrol et
                if 'İlçe' in df_fotograf.columns:
                    df_fotograf = df_fotograf.rename(columns={'İlçe': 'Aob'})
                elif 'İLÇE' in df_fotograf.columns:
                    df_fotograf = df_fotograf.rename(columns={'İLÇE': 'Aob'})
            
            if 'Hat Adı' not in df_fotograf.columns:
                # Alternatif sütun isimlerini kontrol et
                if 'Hat adı' in df_fotograf.columns:
                    df_fotograf = df_fotograf.rename(columns={'Hat adı': 'Hat Adı'})
                elif 'HAT ADI' in df_fotograf.columns:
                    df_fotograf = df_fotograf.rename(columns={'HAT ADI': 'Hat Adı'})
            
            if 'Direk No' not in df_fotograf.columns:
                # Alternatif sütun isimlerini kontrol et
                if 'Direk no' in df_fotograf.columns:
                    df_fotograf = df_fotograf.rename(columns={'Direk no': 'Direk No'})
                elif 'DİREK NO' in df_fotograf.columns:
                    df_fotograf = df_fotograf.rename(columns={'DİREK NO': 'Direk No'})
            
            # Verileri temizle
            df_fotograf['Aob'] = df_fotograf['Aob'].astype(str).str.strip()
            df_fotograf['Hat Adı'] = df_fotograf['Hat Adı'].astype(str).str.strip()
            df_fotograf['Direk No'] = df_fotograf['Direk No'].astype(str).str.strip()
            
            # 2. FİLTRELEME UYGULA
            df_filtreli_fotograf = df_fotograf.copy()
            
            # İLÇE filtresi
            if secili_ilce != "TÜMÜ" and 'Aob' in df_filtreli_fotograf.columns:
                df_filtreli_fotograf = df_filtreli_fotograf[df_filtreli_fotograf['Aob'] == secili_ilce]
            
            # HAT ADI filtresi
            if secili_hat != "TÜMÜ" and 'Hat Adı' in df_filtreli_fotograf.columns:
                if secili_hat != "GENEL":
                    df_filtreli_fotograf = df_filtreli_fotograf[df_filtreli_fotograf['Hat Adı'] == secili_hat]
            
            # 3. TOPLAM DİREK SAYISINI HESAPLA - TÜM SATIRLARI SAY
            if 'Direk No' in df_filtreli_fotograf.columns:
                # Boş olmayan tüm Direk No'ları say
                direk_nolari = df_filtreli_fotograf['Direk No'].dropna().astype(str).str.strip()
                direk_nolari = direk_nolari[direk_nolari != '']
                
                # TÜM satırları say (benzersiz olmasın, aynı numaralı direkler de tekrar saysın)
                total_direk = len(direk_nolari)
            else:
                total_direk = len(df_filtreli_fotograf)
            
            # TOPLAM DİREK SAYISINI GÜNCELLE (BULGU SAYISI OLARAK KULLANILACAK)
            self.toplam_kayit_var.set(str(total_direk))
            
            # 4. BİRLEŞTİRİLMİŞ TESPİT SAYFASINDAN ÖNCELİK DAĞILIMINI HESAPLA
            df_veri_filtreli = self.df_veri.copy()
            
            # İLÇE filtresi
            if secili_ilce != "TÜMÜ" and 'Aob' in df_veri_filtreli.columns:
                df_veri_filtreli = df_veri_filtreli[df_veri_filtreli['Aob'] == secili_ilce]
            
            # HAT ADI filtresi
            if secili_hat != "TÜMÜ" and 'Hat Adı' in df_veri_filtreli.columns:
                if secili_hat != "GENEL":
                    df_veri_filtreli = df_veri_filtreli[df_veri_filtreli['Hat Adı'] == secili_hat]
            
            # ÖNCELİK DAĞILIMI HESAPLA - TÜM SIRALARI SAY
            if 'Öncelik' in df_veri_filtreli.columns and 'Direk No' in df_veri_filtreli.columns:
                # Tüm satırları say
                cok_acil_count = 0
                acil_count = 0
                normal_count = 0
                bekleyebilir_count = 0
                
                for idx, row in df_veri_filtreli.iterrows():
                    direk_no = str(row.get('Direk No', '')).strip()
                    oncelik = str(row.get('Öncelik', '')).lower()
                    
                    if direk_no and direk_no != '' and direk_no.lower() != 'nan':
                        if 'çok acil' in oncelik or 'cok acil' in oncelik:
                            cok_acil_count += 1
                        elif 'acil' in oncelik:
                            acil_count += 1
                        elif 'normal' in oncelik:
                            normal_count += 1
                        elif 'bekleyebilir' in oncelik:
                            bekleyebilir_count += 1
                
                # ÖNCELİK TOPLAMI
                oncelik_toplam = cok_acil_count + acil_count + normal_count + bekleyebilir_count
                
                # DEĞERLERİ ATA
                self.cok_acil_var.set(f"{cok_acil_count:,}")
                self.acil_var.set(f"{acil_count:,}")
                self.normal_var.set(f"{normal_count:,}")
                self.bekleyebilir_var.set(f"{bekleyebilir_count:,}")
                self.oncelik_toplam_var.set(str(oncelik_toplam))
            else:
                # Öncelik sütunu yoksa
                self.cok_acil_var.set("0")
                self.acil_var.set("0")
                self.normal_var.set("0")
                self.bekleyebilir_var.set("0")
                self.oncelik_toplam_var.set("0")
                oncelik_toplam = 0
            
            print(f"İSTATİSTİK DEBUG: Direk={total_direk}, Öncelik Toplam={oncelik_toplam}")
            
        except Exception as e:
            print(f"İstatistik güncelleme hatası: {e}")
            import traceback
            print(traceback.format_exc())
            
            # Hata durumunda varsayılan değerleri ayarla
            self.toplam_kayit_var.set("0")
            self.cok_acil_var.set("0")
            self.acil_var.set("0")
            self.normal_var.set("0")
            self.bekleyebilir_var.set("0")
            self.oncelik_toplam_var.set("0")
            
            self.bilgi_var.set("❌ İstatistik hesaplama hatası")

    def filtrele(self):
        """Filtreleri uygular"""
        try:
            # Orijinal veriden başla
            filtered_df = self.df_veri.copy()
            
            print(f"DEBUG: Filtreleme başlıyor. Toplam kayıt: {len(filtered_df)}")
            
            # İlçe filtresi
            secili_ilce = self.ilce_var.get()
            if secili_ilce != "TÜMÜ" and 'Aob' in filtered_df.columns:
                filtered_df = filtered_df[filtered_df['Aob'] == secili_ilce]
                print(f"DEBUG: İlçe filtresi uygulandı: {secili_ilce}, Kalan: {len(filtered_df)}")
            
            # Hat filtresi
            secili_hat = self.hat_var.get()
            if secili_hat != "TÜMÜ" and 'Hat Adı' in filtered_df.columns:
                if secili_hat == "GENEL":
                    # GENEL seçilirse, tüm kayıtları bırak
                    pass
                else:
                    filtered_df = filtered_df[filtered_df['Hat Adı'] == secili_hat]
                print(f"DEBUG: Hat filtresi uygulandı: {secili_hat}, Kalan: {len(filtered_df)}")
            
            # Harita için "Yapıldı mı?" filtresi
            harita_df = filtered_df.copy()  # Harita için ayrı dataframe
            
            secili_yapildi = self.yapildi_var.get()
            if secili_yapildi != "TÜMÜ" and 'Yapıldı mı?' in harita_df.columns:
                if secili_yapildi == "YAPILDI":
                    # Yapıldı olanları filtrele
                    harita_df = harita_df[harita_df['Yapıldı mı?'].astype(str).str.lower().str.contains(
                        'evet|yapıldı|x|✓|✅|tamamlandı|bitti|yes', na=False)]
                    print(f"DEBUG: Harita için YAPILDI filtresi uygulandı, Kalan: {len(harita_df)}")
                elif secili_yapildi == "YAPILMADI":
                    # Yapılmadı olanları filtrele
                    harita_df = harita_df[~harita_df['Yapıldı mı?'].astype(str).str.lower().str.contains(
                        'evet|yapıldı|x|✓|✅|tamamlandı|bitti|yes', na=False)]
                    print(f"DEBUG: Harita için YAPILMADI filtresi uygulandı, Kalan: {len(harita_df)}")
            
            # Harita için filtrelenmiş veri
            self.df_filtreli = harita_df
            
            print(f"DEBUG: Filtreleme tamamlandı. Son kayıt sayısı: {len(self.df_filtreli)}")
            
            # 1. ÖNCE İSTATİSTİKLERİ GÜNCELLE
            self.istatistikleri_guncelle()
            
            # 2. FİLTRELEME BİLGİSİNİ GÖSTER
            if self.df_filtreli is not None:
                # Fotoğraf Klasörleri'nden gelen Bulgu (direk) sayısı
                total_direk = self.toplam_kayit_var.get()
                
                # Filtrelenmiş verideki koordinatlı kayıt sayısı
                valid_coords = len(self.df_filtreli[self.df_filtreli['LAT'].notna() & self.df_filtreli['LON'].notna()])
                
                # Öncelik toplamını al
                oncelik_toplam = self.oncelik_toplam_var.get()
                
                # FİLTRELEME BİLGİSİ - SIFIRLA METODUYLA AYNI FORMATTA
                self.bilgi_var.set(f" {total_direk} Direk\n {len(self.df_filtreli)} Bulgu\n {valid_coords} Konumlu Buldu\n Öncelik Toplamı: {oncelik_toplam}")
            else:
                self.bilgi_var.set("🔍 Filtreleme yapıldı")
            
            # 3. SONRA HARİTAYI OLUŞTUR (HARİTA_OLUSTUR MARKER EKLER)
            self.harita_olustur()
            
        except Exception as e:
            self.bilgi_var.set(f"❌ Filtreleme hatası: {str(e)}")
            print(f"Filtreleme hatası: {e}")
            import traceback
            print(traceback.format_exc())
        
    def filtreleri_sifirla(self):
        """Filtreleri sıfırlar"""
        try:
            self.ilce_var.set("TÜMÜ")
            self.hat_var.set("TÜMÜ")
            self.yapildi_var.set("TÜMÜ")
            
            # Filtreleri sıfırla (orijinal veriye dön)
            self.df_filtreli = self.df_veri.copy() if self.df_veri is not None else None
            
            # 1. ÖNCE İSTATİSTİKLERİ GÜNCELLE
            self.istatistikleri_guncelle()
            
            # 2. SIFIRLAMA BİLGİSİNİ GÖSTER
            if self.df_filtreli is not None:
                # Fotoğraf Klasörleri'nden gelen Bulgu (direk) sayısı
                total_direk = self.toplam_kayit_var.get()
                
                # Koordinatlı kayıt sayısını hesapla
                valid_coords = len(self.df_filtreli[self.df_filtreli['LAT'].notna() & self.df_filtreli['LON'].notna()])
                
                # Öncelik toplamını al
                oncelik_toplam = self.oncelik_toplam_var.get()
                
                # SIFIRLAMA BİLGİSİ - FİLTRELEMEYLE AYNI FORMATTA
                self.bilgi_var.set(f" {total_direk} Direk\n {len(self.df_filtreli)} Bulgu\n {valid_coords} Konumlu Buldu\n Öncelik Toplamı: {oncelik_toplam}")
            else:
                self.bilgi_var.set(f"🔄 Filtreler sıfırlandı")
            
            # 3. SONRA HARİTAYI OLUŞTUR (HARİTA_OLUSTUR MARKER EKLER)
            self.harita_olustur()
            
        except Exception as e:
            error_msg = str(e)
            self.bilgi_var.set(f"❌ Filtre sıfırlama hatası: {error_msg[:50]}")
            print(f"Filtre sıfırlama hatası: {e}")
            import traceback
            print(f"Hata detayı: {traceback.format_exc()}")
    def direk_no_temizle(self, direk_no):
        """Direk numarasını temizler (AnaUygulama'daki ile aynı mantık)"""
        if not direk_no or str(direk_no).lower() in ['nan', 'none', 'null', '']:
            return ""
        
        direk_no = str(direk_no).strip()
        
        if not direk_no:
            return ""
        
        # "0" kontrolü
        if direk_no.replace('0', '') == '':
            return '0'
        
        try:
            if '.' in direk_no:
                cleaned = str(int(float(direk_no)))
            else:
                cleaned = str(int(direk_no))
            
            cleaned = cleaned.lstrip('0') or '0'
            return cleaned
            
        except (ValueError, TypeError):
            sadece_rakamlar = ''.join(filter(str.isdigit, direk_no))
            if sadece_rakamlar:
                sadece_rakamlar = sadece_rakamlar.lstrip('0') or '0'
                return sadece_rakamlar
            
            return direk_no

        
    def harita_olustur(self):
        """Harita oluşturur - Aynı direkler gruplanacak"""
        try:
            if self.df_filtreli is None or self.df_filtreli.empty:
                self.bilgi_var.set("❌ Harita için veri yok")
                return
            
            # Koordinatlı verileri filtrele
            valid = self.df_filtreli['LAT'].notna() & self.df_filtreli['LON'].notna()
            df_map = self.df_filtreli[valid].copy()
            
            if df_map.empty:
                self.bilgi_var.set("❌ Koordinatlı veri yok")
                return
            
            # Merkez hesapla
            try:
                center_lat = df_map['LAT'].astype(float).mean()
                center_lon = df_map['LON'].astype(float).mean()
            except:
                center_lat = 39.9334
                center_lon = 32.8597
            
            # Folium harita oluştur
            harita = folium.Map(location=[center_lat, center_lon], 
                              zoom_start=12, 
                              control_scale=True,
                              tiles=None)
            
            # Tile layer'lar
            folium.TileLayer(
                tiles='https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',
                name='🗺️ Normal Harita',
                attr='© OpenStreetMap contributors',
                max_zoom=19,
                control=True
            ).add_to(harita)
            
            folium.TileLayer(
                tiles='https://mt1.google.com/vt/lyrs=y&x={x}&y={y}&z={z}',
                attr='Google Hybrid',
                name='🛰️🔤 Uydu + Etiketler',
                max_zoom=22,
                control=True
            ).add_to(harita)
            
            # 1. FOTOĞRAF KLASÖRLERİ SAYFASINI OKU - TARİH İÇİN
            try:
                df_fotograf = pd.read_excel(self.excel_dosyasi, sheet_name='Fotoğraf Klasörleri')
                df_fotograf.columns = df_fotograf.columns.str.strip()
                
                if 'Direk No' in df_fotograf.columns:
                    df_fotograf['Direk_No_Temiz'] = df_fotograf['Direk No'].apply(lambda x: self.direk_no_temizle(str(x)))
                
            except Exception as e:
                print(f"⚠️ Fotoğraf Klasörleri sayfası yüklenemedi: {e}")
                df_fotograf = None
            
            # Öncelik dağılımına göre marker ekle
            use_cluster = self.kumelenme_var.get()
            oncelik_dagilimi = self.oncelik_dagilimi_var.get()
            
            # Marker layer oluştur
            if use_cluster:
                if oncelik_dagilimi:
                    marker_layer = MarkerCluster(name='🚨 Öncelikli Direkler').add_to(harita)
                else:
                    marker_layer = MarkerCluster(name='📍 Tüm Direkler').add_to(harita)
            else:
                if oncelik_dagilimi:
                    marker_layer = folium.FeatureGroup(name='🚨 Öncelikli Direkler').add_to(harita)
                else:
                    marker_layer = folium.FeatureGroup(name='📍 Tüm Direkler').add_to(harita)
            
            # ✅ DEĞİŞİKLİK: AYNI DİREKLERİ GRUPLAMA
            # İlçe, Hat Adı ve Direk No'ya göre grupla
            df_map['Direk_No_Temiz'] = df_map['Direk No'].apply(lambda x: self.direk_no_temizle(str(x)))
            
            # Grup anahtarı oluştur (ilçe + hat + direk no)
            df_map['Grup_Anahtari'] = df_map['Aob'].fillna('') + '|' + df_map['Hat Adı'].fillna('') + '|' + df_map['Direk_No_Temiz']
            
            # Grupları işle
            gruplanmis_markerlar = {}
            
            for grup_anahtari, grup_df in df_map.groupby('Grup_Anahtari'):
                if len(grup_df) > 0:
                    # İlk satırı temel al (ortanca koordinatlar için de hesaplayabilirsiniz)
                    ilk_satir = grup_df.iloc[0]
                    
                    # Ortalama koordinatları hesapla
                    try:
                        lat = float(grup_df['LAT'].astype(float).mean())
                        lon = float(grup_df['LON'].astype(float).mean())
                    except:
                        lat = float(ilk_satir['LAT'])
                        lon = float(ilk_satir['LON'])
                    
                    # Grup bilgilerini topla
                    grup_bilgileri = {
                        'lat': lat,
                        'lon': lon,
                        'ilce': ilk_satir.get('Aob', ''),
                        'hat_adi': ilk_satir.get('Hat Adı', ''),
                        'direk_no': ilk_satir.get('Direk No', ''),
                        'direk_no_temiz': grup_anahtari.split('|')[-1],
                        'satirlar': [],
                        'tespit_notlari': []
                    }
                    
                    # Tüm satırları işle
                    for idx, row in grup_df.iterrows():
                        try:
                            # Tespit notu
                            tespit_notu = str(row.get('Tespit Notu', '')).strip()
                            if tespit_notu and tespit_notu != 'nan':
                                grup_bilgileri['tespit_notlari'].append(tespit_notu)
                            
                            # Yapıldı mı? bilgisini satırla birlikte kaydet
                            yapildi_mi = str(row.get('Yapıldı mı?', '')).strip()
                            yapildi_durum = 'Yapılmadı'
                            if yapildi_mi and yapildi_mi.lower() in ['evet', 'yapıldı', 'x', '✓', '✅', 'tamamlandı', 'bitti', 'yes']:
                                yapildi_durum = 'Yapıldı'
                            
                            # Öncelik rengi (sadece marker rengi için)
                            oncelik = str(row.get('Öncelik', '')).lower()
                            renk = 'gray'
                            if 'çok acil' in oncelik or 'cok acil' in oncelik:
                                renk = 'red'
                            elif 'acil' in oncelik:
                                renk = 'orange'
                            elif 'normal' in oncelik:
                                renk = 'blue'
                            elif 'bekleyebilir' in oncelik:
                                renk = 'green'
                            
                            # Satır bilgisini ekle
                            grup_bilgileri['satirlar'].append({
                                'renk': renk,
                                'yapildi_mi': yapildi_mi,
                                'yapildi_durum': yapildi_durum,
                                'tespit_notu': tespit_notu
                            })
                            
                        except Exception as e:
                            print(f"Grup işleme hatası satır {idx}: {e}")
                    
                    # Grup bilgilerini kaydet
                    gruplanmis_markerlar[grup_anahtari] = grup_bilgileri
            
            # Marker'ları ekle
            eklenen = 0
            for grup_anahtari, grup_bilgi in gruplanmis_markerlar.items():
                try:
                    lat = grup_bilgi['lat']
                    lon = grup_bilgi['lon']
                    ilce = grup_bilgi['ilce']
                    hat_adi = grup_bilgi['hat_adi']
                    direk_no = grup_bilgi['direk_no']
                    direk_no_temiz = grup_bilgi['direk_no_temiz']
                    
                    # ✅ ÖNCELİK DAĞILIMI FİLTRESİ
                    if oncelik_dagilimi:
                        # Sadece öncelikli olanları kontrol et
                        oncelik_renkleri = [satir['renk'] for satir in grup_bilgi['satirlar']]
                        # Eğer tüm satırlar gri ise (önceliksiz) atla
                        if all(renk == 'gray' for renk in oncelik_renkleri):
                            continue
                    
                    # EN ÖNCELİKLİ RENGİ BELİRLE (marker rengi için)
                    renk_oncelik_sirasi = {'red': 1, 'orange': 2, 'blue': 3, 'green': 4, 'gray': 5}
                    en_oncelikli_renk = 'gray'
                    en_oncelikli_derece = 5
                    
                    for satir in grup_bilgi['satirlar']:
                        renk_derece = renk_oncelik_sirasi.get(satir['renk'], 5)
                        if renk_derece < en_oncelikli_derece:
                            en_oncelikli_derece = renk_derece
                            en_oncelikli_renk = satir['renk']
                    
                    # 📌 DÜZELTME: Birden fazla kayıt varsa ikonu 'list' yap
                    kayit_sayisi = len(grup_bilgi['satirlar'])
                    icon_name = 'list' if kayit_sayisi > 1 else 'info-sign'
                    
                    # Eğer tüm kayıtlar yapıldıysa check ikonu kullan (sadece 1 kayıt varsa)
                    yapildi_sayisi = sum(1 for satir in grup_bilgi['satirlar'] if satir['yapildi_durum'] == 'Yapıldı')
                    if yapildi_sayisi == kayit_sayisi and kayit_sayisi == 1:
                        icon_name = 'check'
                    
                    # TARİH BİLGİSİ (FOTOĞRAF KLASÖRLERİ'NDEN)
                    en_yeni_tarih = "Belirsiz"
                    if df_fotograf is not None and direk_no_temiz:
                        filtered_fotograf = df_fotograf[
                            (df_fotograf['Direk_No_Temiz'] == direk_no_temiz) &
                            (df_fotograf['Aob'].astype(str).str.strip() == ilce)
                        ]
                        
                        if not filtered_fotograf.empty:
                            hat_eslesen = filtered_fotograf[
                                filtered_fotograf['Hat adı'].astype(str).str.strip() == hat_adi
                            ]
                            
                            if not hat_eslesen.empty:
                                tarih_verisi = hat_eslesen.iloc[0]
                            else:
                                tarih_verisi = filtered_fotograf.iloc[0]
                            
                            if 'En Yeni Fotoğraf Tarihi' in tarih_verisi:
                                tarih = str(tarih_verisi['En Yeni Fotoğraf Tarihi'])
                                if tarih != 'nan' and tarih != 'NaT' and pd.notna(tarih):
                                    if isinstance(tarih, pd.Timestamp):
                                        en_yeni_tarih = tarih.strftime('%d/%m/%Y %H:%M:%S')
                                    elif isinstance(tarih, datetime.datetime):
                                        en_yeni_tarih = tarih.strftime('%d/%m/%Y %H:%M:%S')
                                    elif isinstance(tarih, str):
                                        if tarih not in ['Tarihi Yok', 'Belirsiz', 'nan', 'NaT', '']:
                                            en_yeni_tarih = tarih
                    
                    # TESPİT DETAYLARI - Her tespitin yanında yapıldı/yapılmadı durumu
                    tespit_detaylari = []
                    for i, satir in enumerate(grup_bilgi['satirlar'], 1):
                        yapildi_icon = "✅" if satir['yapildi_durum'] == 'Yapıldı' else "❌"
                        
                        # Tespit notu
                        not_metni = ""
                        if i <= len(grup_bilgi['tespit_notlari']):
                            not_metni = grup_bilgi['tespit_notlari'][i-1]
                        else:
                            not_metni = satir.get('tespit_notu', '') or "Tespit notu girilmemiş"
                        
                        # Öncelik rengine göre emoji (her tespit için)
                        oncelik_renk = satir['renk']
                        oncelik_emoji = ""
                        if oncelik_renk == 'red':
                            oncelik_emoji = "🔴"
                        elif oncelik_renk == 'orange':
                            oncelik_emoji = "🟠"
                        elif oncelik_renk == 'blue':
                            oncelik_emoji = "🔵"
                        elif oncelik_renk == 'green':
                            oncelik_emoji = "🟢"
                        else:
                            oncelik_emoji = "⚪"
                        
                        tespit_detaylari.append({
                            'no': i,
                            'notu': not_metni,
                            'yapildi_icon': yapildi_icon,
                            'yapildi_durum': satir['yapildi_durum'],
                            'oncelik_emoji': oncelik_emoji,
                            'renk': oncelik_renk
                        })
                    
                    # TESPİT LİSTESİ HTML
                    tespit_listesi_html = ""
                    if tespit_detaylari:
                        tespit_listesi_html = f"""
                        <div style='margin-top:10px;'>
                            <div style='font-weight:bold; color:#333; margin-bottom:8px; padding-bottom:3px; border-bottom:2px solid #3498db; font-size:12px;'>
                                🔍 {len(tespit_detaylari)} TESPİT BULUNDU:
                            </div>
                            <div style='max-height:220px; overflow-y:auto; padding-right:3px;'>
                        """
                        
                        for tespit in tespit_detaylari:
                            durum_rengi = "#27ae60" if tespit['yapildi_durum'] == 'Yapıldı' else "#e74c3c"
                            oncelik_renk_text = ""
                            if tespit['renk'] == 'red':
                                oncelik_renk_text = "Çok Acil"
                            elif tespit['renk'] == 'orange':
                                oncelik_renk_text = "Acil"
                            elif tespit['renk'] == 'blue':
                                oncelik_renk_text = "Normal"
                            elif tespit['renk'] == 'green':
                                oncelik_renk_text = "Bekleyebilir"
                            else:
                                oncelik_renk_text = "Belirsiz"
                            
                            tespit_listesi_html += f"""
                                <div style='margin-bottom:10px; padding:8px; background:#f8f9fa; border-radius:4px; border-left:3px solid {durum_rengi};'>
                                    <div style='display:flex; justify-content:space-between; align-items:center; margin-bottom:5px;'>
                                        <div style='display:flex; align-items:center; gap:5px;'>
                                            <span style='font-size:12px;'>{tespit['oncelik_emoji']}</span>
                                            <div style='font-weight:bold; color:#2c3e50; font-size:11px;'>
                                                Tespit {tespit['no']}:
                                            </div>
                                        </div>
                                        <div style='display:flex; align-items:center; gap:5px;'>
                                            <span style='font-size:12px;'>{tespit['yapildi_icon']}</span>
                                            <span style='font-size:10px; color:{durum_rengi}; font-weight:bold;'>
                                                {tespit['yapildi_durum']}
                                            </span>
                                        </div>
                                    </div>
                                    <div style='margin-bottom:5px; font-size:10px; color:#7f8c8d; padding-left:18px;'>
                                        {oncelik_renk_text}
                                    </div>
                                    <div style='font-size:10px; color:#34495e; padding:5px; background:white; border-radius:3px; border:1px solid #ecf0f1;'>
                                        {tespit['notu']}
                                    </div>
                                </div>
                            """
                        
                        tespit_listesi_html += "</div></div>"
                    else:
                        tespit_listesi_html = """
                        <div style='margin-top:10px; padding:10px; background:#f8f9fa; border-radius:4px;'>
                            <div style='font-weight:bold; color:#7f8c8d; font-size:11px; text-align:center;'>
                                📝 TESPİT BULUNAMADI
                            </div>
                        </div>
                        """
                    
                    # Popup HTML - SADELESTIRILMIS TASARIM
                    google_maps_directions = f"https://www.google.com/maps/dir/?api=1&destination={lat},{lon}"
                    
                    # Ana başlık kısmı (sadece ilçe, hat)
                    header_html = f"""
                    <div style='margin-bottom:10px;'>
                        <div style='font-weight:bold; color:#2c3e50; font-size:14px; margin-bottom:3px;'>
                            {ilce} - {hat_adi}
                        </div>
                    </div>
                    """
                    
                    # Bilgi kutusu (çerçeveli) - Direk No ve Tarih
                    bilgi_kutusu_html = f"""
                    <div style='background:#ecf0f1; padding:10px; border-radius:4px; margin-bottom:10px; border:1px solid #bdc3c7;'>
                        <div style='font-size:11px; color:#34495e;'>
                            <div style='margin-bottom:3px;'>
                                <strong>Direk No:</strong> <span style='color:#3498db; font-weight:bold;'>{direk_no}</span>
                                <span style='margin-left:8px; font-size:10px; color:#95a5a6;'>
                                    ({kayit_sayisi} kayıt)
                                </span>
                            </div>
                            <div>
                                <strong>Tespit Tarihi:</strong> {en_yeni_tarih}
                            </div>
                        </div>
                    </div>
                    """
                    
                    # Tüm içeriği birleştir
                    popup_html = f"""
                    <div style='font-family: Arial; width: 300px; max-height: 400px; overflow-y: auto; font-size: 12px; padding:5px;'>
                        {header_html}
                        {bilgi_kutusu_html}
                        {tespit_listesi_html}
                        
                        <!-- BUTONLAR -->
                        <div style='margin-top:15px;'>
                            <a href='{google_maps_directions}' target='_blank' 
                               style='display:flex; align-items:center; justify-content:center;
                                      background:#3498db; color:white; text-decoration:none; 
                                      padding:8px 15px; border-radius:5px; font-weight:bold;
                                      transition:all 0.2s; font-size:11px; border:none;'>
                                <span style='margin-right:6px; font-size:14px;'>🚗</span>
                                Yol Tarifi Al
                            </a>
                        </div>
                        
                        <!-- KOORDİNAT -->
                        <div style='margin-top:10px; font-size:10px; color:#7f8c8d; text-align:center; padding-top:8px; border-top:1px dashed #bdc3c7;'>
                            Koordinat: 
                            <span style='font-family:monospace; background:#f8f9fa; padding:3px 6px; 
                                         border-radius:3px; cursor:pointer; font-size:9px; border:1px solid #ecf0f1;' 
                                 onclick='navigator.clipboard.writeText("{lat:.6f}, {lon:.6f}")'>
                                {lat:.6f}, {lon:.6f}
                            </span>
                        </div>
                    </div>
                    """
                    
                    # Popup boyutu
                    max_popup_width = 320
                    
                    marker = folium.Marker(
                        location=[lat, lon],
                        popup=folium.Popup(popup_html, max_width=max_popup_width),
                        icon=folium.Icon(color=en_oncelikli_renk, icon=icon_name, prefix='fa')
                    )
                    
                    marker.add_to(marker_layer)
                    eklenen += 1
                    
                except Exception as e:
                    print(f"Grup marker ekleme hatası {grup_anahtari}: {e}")
                    continue
            
            # Layer kontrol
            folium.LayerControl(collapsed=False, position='topright').add_to(harita)
            
            # Haritayı kaydet
            harita.save(self.harita_html_path)
            self.map_temp_html = self.harita_html_path
            
            # HTML dosyasına JavaScript ekle
            self.html_dosyasina_javascript_ekle_konumlu(self.harita_html_path)
            
            # Haritayı göster
            self.harita_goster_embed(self.map_temp_html)
            
            # İstatistikleri güncelle
            current_info = self.bilgi_var.get()
            new_info = f"{current_info} | ✅ {eklenen} grup marker eklendi ({len(df_map)} kayıt gruplandı)"
            self.bilgi_var.set(new_info)
            
            print(f"✅ {eklenen} grup marker eklendi (Toplam {len(df_map)} kayıt)")
            
        except Exception as e:
            self.bilgi_var.set(f"❌ Harita oluşturma hatası: {str(e)}")
            print(f"Harita oluşturma hatası: {e}")
            import traceback
            traceback.print_exc()

    def html_dosyasina_javascript_ekle_konumlu(self, html_yolu):
        """HTML dosyasına anlık konum özelliği ekler (Folium uyumlu)"""
        try:
            with open(html_yolu, 'r', encoding='utf-8') as file:
                html_icerik = file.read()

            javascript_ekle = """
    <script>
    var userLocationMarker = null;
    var userLocationCircle = null;
    var watchId = null;
    var map = null;

    // Folium haritasını BUL (EN KRİTİK DÜZELTME)
    document.addEventListener("DOMContentLoaded", function () {
        for (var key in window) {
            if (key.startsWith("map_") && window[key] instanceof L.Map) {
                map = window[key];
                console.log("Folium haritası bulundu:", key);
                createLocationButton();
                return;
            }
        }
        console.error("Folium haritası bulunamadı!");
    });

    // Konum butonu
    function createLocationButton() {
        var LocationControl = L.Control.extend({
            options: { position: 'topleft' },

            onAdd: function () {
                var container = L.DomUtil.create('div', 'leaflet-bar');
                container.style.background = 'white';

                var btn = L.DomUtil.create('a', '', container);
                btn.innerHTML = '📍';
                btn.style.width = '36px';
                btn.style.height = '36px';
                btn.style.lineHeight = '36px';
                btn.style.textAlign = 'center';
                btn.style.cursor = 'pointer';
                btn.title = 'Konumumu Göster';

                btn.onclick = function (e) {
                    e.preventDefault();
                    getUserLocation();
                };
                return container;
            }
        });

        map.addControl(new LocationControl());
    }

    // Konumu al
    function getUserLocation() {
        if (!navigator.geolocation) {
            alert("Tarayıcı konum desteklemiyor");
            return;
        }

        navigator.geolocation.getCurrentPosition(
            function (pos) {
                showUserLocation(
                    pos.coords.latitude,
                    pos.coords.longitude,
                    pos.coords.accuracy
                );
                startTracking();
            },
            function (err) {
                alert("Konum alınamadı: " + err.message);
            },
            { enableHighAccuracy: true }
        );
    }

    // Haritada göster
    function showUserLocation(lat, lng, acc) {
        if (userLocationMarker) map.removeLayer(userLocationMarker);
        if (userLocationCircle) map.removeLayer(userLocationCircle);

        userLocationMarker = L.circleMarker([lat, lng], {
            radius: 8,
            color: '#4285F4',
            fillColor: '#4285F4',
            fillOpacity: 1
        }).addTo(map);

        userLocationCircle = L.circle([lat, lng], {
            radius: acc,
            color: '#4285F4',
            fillOpacity: 0.15
        }).addTo(map);

        map.setView([lat, lng], 16);
    }

    // Canlı takip
    function startTracking() {
        watchId = navigator.geolocation.watchPosition(function (pos) {
            var lat = pos.coords.latitude;
            var lng = pos.coords.longitude;
            var acc = pos.coords.accuracy;

            userLocationMarker.setLatLng([lat, lng]);
            userLocationCircle.setLatLng([lat, lng]);
            userLocationCircle.setRadius(acc);
        });
    }
    </script>
    """

            if "</body>" in html_icerik:
                html_icerik = html_icerik.replace("</body>", javascript_ekle + "</body>")
            else:
                html_icerik += javascript_ekle

            with open(html_yolu, 'w', encoding='utf-8') as file:
                file.write(html_icerik)

            print("✅ Konum JS başarıyla eklendi")

        except Exception as e:
            print("❌ Hata:", e)


    
    def harita_goster_embed(self, html_yolu):
        """Haritayı PyQt6 ile gömülü olarak gösterir - GÜVENLİ"""
        try:
            # 🔴 1) ÖNCE ESKİ HARİTA WIDGET’LARINI TEMİZLE
            for w in self.map_frame.winfo_children():
                w.destroy()

            # 🔴 2) ESKİ PyQt WIDGET VARSA KAPAT
            if hasattr(self, 'pyqt_widget') and self.pyqt_widget:
                try:
                    self.pyqt_widget.close()
                except:
                    pass
                self.pyqt_widget = None

            # 🔴 3) PyQt6 KONTROLÜ (LAZY)
            if not self.check_pyqt_availability():
                self.harita_basit_hata_goster(
                    "Harita gömülü gösterimi için 'PyQt6' gerekli.\n\n"
                    "Kurulum için:\n"
                    "1. Terminali açın\n"
                    "2. Şu komutu çalıştırın: pip install PyQt6 PyQt6-WebEngine\n"
                    "3. Uygulamayı yeniden başlatın",
                    html_yolu
                )
                return

            # 🔴 4) PyQt6'YI GÜVENLİ BAŞLAT
            self.guvenli_pyqt_baslat(html_yolu)

        except Exception as e:
            self.bilgi_var.set(f"❌ Harita gösterim hatası: {str(e)}")
            import traceback
            print(f"Harita gösterme hatası: {traceback.format_exc()}")

            # 🔴 5) HATA DURUMUNDA FALLBACK
            self.harita_basit_hata_goster(str(e), html_yolu)

    
    def guvenli_pyqt_baslat(self, html_yolu):
        """PyQt6'yı güvenli şekilde başlatır"""
        try:
            from PyQt6.QtWidgets import QApplication
            from PyQt6.QtCore import QTimer
            
            # Eğer QApplication yoksa oluştur
            if QApplication.instance() is None:
                import sys
                self.qt_app = QApplication(sys.argv)
                self.qt_app.setQuitOnLastWindowClosed(False)  # Önemli!
            
            # Widget'ı oluştur
            self.create_pyqt_map_widget(html_yolu)
            
            # Event loop'u güvenli başlat
            self.start_safe_qt_timer()
            
        except Exception as e:
            print(f"PyQt6 güvenli başlatma hatası: {e}")
            raise
    
    def start_safe_qt_timer(self):
        """Güvenli Qt event loop timer'ı"""
        def process_qt_events():
            try:
                from PyQt6.QtWidgets import QApplication
                if QApplication.instance():
                    QApplication.instance().processEvents()
            except:
                pass
            
            # Timer'ı devam ettir
            if hasattr(self, 'pyqt_widget') and self.pyqt_widget:
                self.after(100, process_qt_events)
        
        # İlk timer'ı başlat
        self.after(100, process_qt_events)
    
    def create_pyqt_map_widget(self, html_yolu):
        """PyQt6 tabanlı harita widget'ını oluşturur"""
        try:
            import sys
            from PyQt6.QtCore import Qt, QUrl, QTimer, QSize, QPoint
            from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QSizePolicy, QHBoxLayout
            from PyQt6.QtWebEngineWidgets import QWebEngineView
            from PyQt6.QtWebEngineCore import QWebEngineSettings
            
            # PyQt6 uygulaması yoksa oluştur
            if QApplication.instance() is None:
                self.qt_app = QApplication(sys.argv)
            
            # PyQt widget'ını oluştur
            class PyQtMapWidget(QWidget):
                def __init__(self, parent=None, html_path=None):
                    super().__init__(parent, Qt.WindowType.Window)
                    self.html_path = html_path
                    self.setup_ui()
                    # Pencerenin border'ını kaldır
                    self.setWindowFlag(Qt.WindowType.FramelessWindowHint)
                    
                def setup_ui(self):
                    # Ana layout
                    main_layout = QVBoxLayout(self)
                    main_layout.setContentsMargins(0, 0, 0, 0)
                    main_layout.setSpacing(0)
                    
                    # WebView oluştur
                    self.view = QWebEngineView()
                    self.view.settings().setAttribute(QWebEngineSettings.WebAttribute.LocalContentCanAccessRemoteUrls, True)
                    self.view.settings().setAttribute(QWebEngineSettings.WebAttribute.LocalContentCanAccessFileUrls, True)
                    self.view.settings().setAttribute(QWebEngineSettings.WebAttribute.JavascriptEnabled, True)
                    self.view.settings().setAttribute(QWebEngineSettings.WebAttribute.AllowRunningInsecureContent, True)
                    
                    # HTML dosyasını yükle
                    if self.html_path:
                        self.view.load(QUrl.fromLocalFile(os.path.abspath(self.html_path)))
                    
                    # Layout'a ekle
                    main_layout.addWidget(self.view)
                    
                def sizeHint(self):
                    return QSize(800, 600)
            
            # PyQt widget'ını oluştur ve tkinter'e göm
            container = tk.Frame(self.map_frame, bg="black")
            container.pack(fill='both', expand=True, padx=0, pady=0)
            
            # Widget ID'sini al
            wid = container.winfo_id()
            
            # PyQt widget'ını oluştur ve göm
            self.pyqt_widget = PyQtMapWidget(None, html_yolu)
            self.pyqt_widget.show()
            
            # Windows için pencere tanıtıcıyı ayarla
            if sys.platform == "win32":
                import ctypes
                
                # PyQt penceresini Tkinter konteynerına göm
                try:
                    # Get the window handles
                    pyqt_hwnd = int(self.pyqt_widget.winId())
                    tk_hwnd = int(wid)
                    
                    # Set parent window
                    ctypes.windll.user32.SetParent(pyqt_hwnd, tk_hwnd)
                    
                    # Pencerenin boyutunu güncelle (container'ın tamamını kapla)
                    self.pyqt_widget.resize(container.winfo_width(), container.winfo_height())
                    
                    # Window style'ı ayarla (border kaldır)
                    GWL_STYLE = -16
                    WS_CHILD = 0x40000000
                    WS_VISIBLE = 0x10000000
                    style = WS_CHILD | WS_VISIBLE
                    ctypes.windll.user32.SetWindowLongW(pyqt_hwnd, GWL_STYLE, style)
                    
                    # Pencerenin pozisyonunu (0,0) yap
                    SWP_NOZORDER = 0x0004
                    SWP_NOACTIVATE = 0x0010
                    SWP_FRAMECHANGED = 0x0020
                    ctypes.windll.user32.SetWindowPos(
                        pyqt_hwnd, 0, 0, 0, 
                        container.winfo_width(), container.winfo_height(),
                        SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED
                    )
                    
                    print(f"✅ Widget başarıyla gömüldü. Boyut: {container.winfo_width()}x{container.winfo_height()}")
                    
                except Exception as win_error:
                    print(f"⚠️ Windows gömme hatası: {win_error}")
                    # Fallback: pencerenin boyutunu ayarla
                    self.pyqt_widget.resize(container.winfo_width(), container.winfo_height())
            
            # Boyut değişikliklerini takip et
            def update_size(event=None):
                try:
                    if hasattr(self, 'pyqt_widget') and self.pyqt_widget:
                        width = container.winfo_width()
                        height = container.winfo_height()
                        
                        if width > 0 and height > 0:
                            # PyQt widget boyutunu güncelle
                            self.pyqt_widget.resize(width, height)
                            
                            # Windows için ekstra ayar
                            if sys.platform == "win32":
                                try:
                                    pyqt_hwnd = int(self.pyqt_widget.winId())
                                    SWP_NOZORDER = 0x0004
                                    SWP_NOACTIVATE = 0x0010
                                    ctypes.windll.user32.SetWindowPos(
                                        pyqt_hwnd, 0, 0, 0, width, height,
                                        SWP_NOZORDER | SWP_NOACTIVATE
                                    )
                                except:
                                    pass
                            
                            print(f"🔄 Boyut güncellendi: {width}x{height}")
                except Exception as size_error:
                    print(f"⚠️ Boyut güncelleme hatası: {size_error}")
            
            # İlk boyutu ayarla
            container.after(100, lambda: update_size())
            
            # Configure event'ini bağla
            container.bind('<Configure>', update_size)
            
            # Container yeniden boyutlandırıldığında
            def on_container_configure(event):
                update_size(event)
            
            container.bind('<Configure>', on_container_configure)
            
            # Qt event loop'unu çalıştırmak için timer
            self.start_qt_timer()
            
        except Exception as e:
            print(f"PyQt widget oluşturma hatası: {e}")
            import traceback
            print(f"🔍 Hata detayı: {traceback.format_exc()}")
            self.harita_basit_hata_goster(f"PyQt6 hatası: {str(e)}", html_yolu)
    
    def start_qt_timer(self):
        """PyQt event loop'unu çalıştırmak için timer başlat"""
        try:
            from PyQt6.QtWidgets import QApplication
            
            def process_qt_events():
                try:
                    if QApplication.instance():
                        QApplication.instance().processEvents()
                except:
                    pass
                
                # Timer'ı sürekli çalıştır
                self.after(50, process_qt_events)
            
            # İlk timer'ı başlat
            self.after(50, process_qt_events)
            
        except Exception as e:
            print(f"Qt timer hatası: {e}")
    
    def harita_basit_hata_goster(self, hata_mesaji, html_yolu=None):
        """Basit hata mesajı gösterimi"""
        # Önceki widget'ları temizle
        for w in self.map_frame.winfo_children():
            w.destroy()
        
        yolu = html_yolu if html_yolu else self.map_temp_html
        
        # Bilgi paneli oluştur
        info_frame = tk.Frame(self.map_frame, bg="white")
        info_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        tk.Label(info_frame, text="🌍 HARİTA BİLGİSİ", 
                 font=('Segoe UI', 14, 'bold'), bg="white").pack(pady=10)
        
        # Hata mesajını göster
        tk.Label(info_frame, text="Harita görüntülenemiyor:", 
                font=('Segoe UI', 10, 'bold'), bg="white").pack()
        
        hata_label = tk.Label(info_frame, text=hata_mesaji, 
                             font=('Segoe UI', 9), bg="#ffebee", fg="#c62828",
                             relief='sunken', padx=10, pady=5,
                             wraplength=600, justify='left')
        hata_label.pack(pady=10, fill='x')
        
        # HTML yolunu göster
        if yolu and os.path.exists(yolu):
            tk.Label(info_frame, text=f"Harita dosyası:", 
                    font=('Segoe UI', 10), bg="white").pack()
            
            path_label = tk.Label(info_frame, text=yolu, 
                                 font=('Consolas', 9), bg="#f0f0f0",
                                 relief='sunken', padx=10, pady=5,
                                 wraplength=600, justify='left')
            path_label.pack(pady=10, fill='x')
        
        # Butonlar
        btn_frame = tk.Frame(info_frame, bg="white")
        btn_frame.pack(pady=20)
        
        if yolu and os.path.exists(yolu):
            tk.Button(btn_frame, text="🌐 Haritayı Tarayıcıda Aç",
                     command=lambda: self.harita_tarayiciya_ac(yolu),
                     bg="#3498db", fg='white', font=('Segoe UI', 10),
                     width=20, height=2).pack(side='left', padx=10)
        
        tk.Button(btn_frame, text="🔄 Haritayı Yenile",
                 command=self.harita_olustur,
                 bg="#27ae60", fg='white', font=('Segoe UI', 10),
                 width=20, height=2).pack(side='left', padx=10)
    
    def harita_tarayiciya_ac(self, html_yolu=None):
        """Haritayı tarayıcıda açar"""
        try:
            yolu = html_yolu if html_yolu else self.map_temp_html
            if yolu and os.path.exists(yolu):
                import webbrowser
                webbrowser.open(f'file://{os.path.abspath(yolu)}')
            else:
                messagebox.showwarning("Uyarı", "Harita dosyası bulunamadı!")
        except Exception as e:
            messagebox.showerror("Hata", f"Tarayıcı açma hatası: {str(e)}")


class ExcelSecimDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("Excel Dosyalarını Seçin")
        self.geometry("900x400")
        self.configure(bg="#2c3e50")
        self.resizable(False, False)
        
        self.transient(parent)
        self.grab_set()
        
        self.direk_musterek_yolu = None
        self.direk_og_yolu = None
        
        
        self.arayuz_olustur()
        
        self.geometry("+%d+%d" % (
            parent.winfo_rootx() + parent.winfo_width() // 2 - 300,
            parent.winfo_rooty() + parent.winfo_height() // 2 - 150
        ))
    
    def arayuz_olustur(self):
        baslik_frame = tk.Frame(self, bg="#3498db", height=50)
        baslik_frame.pack(fill='x', padx=1, pady=1)
        baslik_frame.pack_propagate(False)
        
        baslik_label = tk.Label(baslik_frame, text="📊 Excel Dosyalarını Seçin", 
                               font=('Segoe UI', 14, 'bold'),
                               fg='white', bg="#3498db")
        baslik_label.pack(expand=True, pady=10)
        
        aciklama_label = tk.Label(baslik_frame, 
                                 text="Lütfen aşağıdaki Excel dosyalarını seçin. Bu dosyalardaki direk numaraları ana Excel'de işaretlenecektir.",
                                 font=('Segoe UI', 10),
                                 fg='white', bg="#3498db", wraplength=550)
        aciklama_label.pack(expand=True, pady=(0, 10))
        
        main_frame = tk.Frame(self, bg="#2c3e50", padx=20, pady=20)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Müşterek Direk Excel seçimi
        musterek_frame = tk.LabelFrame(main_frame, text="🔗 Müşterek Direk Excel Dosyası", 
                                      font=('Segoe UI', 11, 'bold'),
                                      fg='white', bg="#34495e", padx=15, pady=15)
        musterek_frame.pack(fill='x', pady=(0, 15))
        
        musterek_inner = tk.Frame(musterek_frame, bg="#34495e")
        musterek_inner.pack(fill='x')
        
        self.musterek_var = tk.StringVar(value="Seçilmedi")
        musterek_label = tk.Label(musterek_inner, textvariable=self.musterek_var,
                                 font=('Segoe UI', 9), fg="#ecf0f1", bg="#34495e",
                                 wraplength=400, justify='left')
        musterek_label.pack(side='left', fill='x', expand=True, padx=(0, 10))
        
        musterek_btn = tk.Button(musterek_inner, text="Dosya Seç", 
                                font=('Segoe UI', 10, 'bold'),
                                bg="#3498db", fg='white', relief='raised',
                                command=self.musterek_sec, width=12)
        musterek_btn.pack(side='right')
        
        # OG Direk Excel seçimi
        og_frame = tk.LabelFrame(main_frame, text="⚡ OG Direk Excel Dosyası", 
                                font=('Segoe UI', 11, 'bold'),
                                fg='white', bg="#34495e", padx=15, pady=15)
        og_frame.pack(fill='x', pady=(0, 20))
        
        og_inner = tk.Frame(og_frame, bg="#34495e")
        og_inner.pack(fill='x')
        
        self.og_var = tk.StringVar(value="Seçilmedi")
        og_label = tk.Label(og_inner, textvariable=self.og_var,
                           font=('Segoe UI', 9), fg="#ecf0f1", bg="#34495e",
                           wraplength=400, justify='left')
        og_label.pack(side='left', fill='x', expand=True, padx=(0, 10))
        
        og_btn = tk.Button(og_inner, text="Dosya Seç", 
                          font=('Segoe UI', 10, 'bold'),
                          bg="#3498db", fg='white', relief='raised',
                          command=self.og_sec, width=12)
        og_btn.pack(side='right')
        
        # Bilgi notu
        info_frame = tk.Frame(main_frame, bg="#2c3e50")
        info_frame.pack(fill='x', pady=(10, 0))
        
        info_text = ("• Müşterek Direk Excel'inde bulunan direkler 'Müşterek Direk' olarak işaretlenecek\n"
                    "• OG Direk Excel'inde bulunan direkler 'Og Direk' olarak işaretlenecek\n"
                    "• Her iki dosyada da bulunmayan direkler boş bırakılacak")
        
        info_label = tk.Label(info_frame, text=info_text,
                             font=('Segoe UI', 9), fg="#bdc3c7", bg="#2c3e50",
                             justify='left')
        info_label.pack(anchor='w')
        
        # Butonlar
        buton_frame = tk.Frame(self, bg="#2c3e50", pady=15)
        buton_frame.pack(fill='x', padx=20)
        
        tk.Button(buton_frame, text="İşlemi Başlat", font=('Segoe UI', 11, 'bold'),
                 bg="#27ae60", fg='white', relief='raised',
                 command=self.onayla, width=15).pack(side='left', padx=10)
        
        tk.Button(buton_frame, text="İptal", font=('Segoe UI', 11, 'bold'),
                bg="#e74c3c", fg='white', relief='raised',
                command=self.iptal, width=15).pack(side='right', padx=10)
    
    def musterek_sec(self):
        dosya = filedialog.askopenfilename(
            title="Müşterek Direk Excel Dosyasını Seçin",
            filetypes=[("Excel dosyaları", "*.xlsx *.xls"), ("Tüm dosyalar", "*.*")]
        )
        if dosya:
            self.direk_musterek_yolu = dosya
            self.musterek_var.set(os.path.basename(dosya))
    
    def og_sec(self):
        dosya = filedialog.askopenfilename(
            title="OG Direk Excel Dosyasını Seçin",
            filetypes=[("Excel dosyaları", "*.xlsx *.xls"), ("Tüm dosyalar", "*.*")]
        )
        if dosya:
            self.direk_og_yolu = dosya
            self.og_var.set(os.path.basename(dosya))
    
    def onayla(self):
        if not self.direk_musterek_yolu and not self.direk_og_yolu:
            messagebox.showwarning("Uyarı", "En az bir Excel dosyası seçmelisiniz!")
            return
        
        self.destroy()
    
    def iptal(self):
        self.direk_musterek_yolu = None
        self.direk_og_yolu = None
        self.destroy()

class TakvimDialog(tk.Toplevel):
    def __init__(self, parent, baslik="Tarih Seçin"):
        super().__init__(parent)
        self.parent = parent
        self.baslik = baslik
        self.secili_tarih = None
        
        self.title(baslik)
        self.geometry("300x280")
        self.configure(bg="#2c3e50")
        self.resizable(False, False)
        
        self.transient(parent)
        self.grab_set()
        
        self.arayuz_olustur()
        
        self.geometry("+%d+%d" % (
            parent.winfo_rootx() + parent.winfo_width() // 2 - 150,
            parent.winfo_rooty() + parent.winfo_height() // 2 - 140
        ))
        
    def arayuz_olustur(self):
        baslik_frame = tk.Frame(self, bg="#3498db", height=40)
        baslik_frame.pack(fill='x', padx=1, pady=1)
        baslik_frame.pack_propagate(False)
        
        baslik_label = tk.Label(baslik_frame, text=self.baslik, 
                               font=('Segoe UI', 12, 'bold'),
                               fg='white', bg="#3498db")
        baslik_label.pack(expand=True)
        
        takvim_frame = tk.Frame(self, bg="#2c3e50", padx=10, pady=10)
        takvim_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        self.takvim = Calendar(takvim_frame, 
                              selectmode='day',
                              date_pattern='dd/mm/yyyy',
                              font=('Segoe UI', 10),
                              background='#34495e',
                              foreground='white',
                              selectbackground='#3498db',
                              normalbackground='#34495e',
                              weekendbackground='#2c3e50',
                              weekendforeground='white',
                              bordercolor='#7f8c8d',
                              headersbackground='#3498db',
                              headersforeground='white',
                              showweeknumbers=False,
                              firstweekday='monday')
        self.takvim.pack(fill='both', expand=True)
        
        buton_frame = tk.Frame(self, bg="#2c3e50", pady=10)
        buton_frame.pack(fill='x', padx=20)
        
        tk.Button(buton_frame, text="Seç", font=('Segoe UI', 10, 'bold'),
                 bg="#27ae60", fg='white', relief='raised',
                 command=self.tarih_sec, width=10).pack(side='left', padx=5)
        
        tk.Button(buton_frame, text="İptal", font=('Segoe UI', 10, 'bold'),
                 bg="#e74c3c", fg='white', relief='raised',
                 command=self.iptal, width=10).pack(side='right', padx=5)
        
        tk.Button(buton_frame, text="Bugün", font=('Segoe UI', 10, 'bold'),
                 bg="#3498db", fg='white', relief='raised',
                 command=self.bugun_sec, width=10).pack(side='right', padx=5)
    
    def tarih_sec(self):
        self.secili_tarih = self.takvim.get_date()
        self.destroy()
    
    def iptal(self):
        self.secili_tarih = None
        self.destroy()
    
    def bugun_sec(self):
        bugun = datetime.date.today()
        self.takvim.selection_set(bugun)

class ExifDateEditor(tk.Toplevel):
    def __init__(self, parent, klasor_yolu=None, secim_tipi="tumunu"):
        super().__init__(parent)
        self.parent = parent
        self.klasor_yolu = klasor_yolu
        self.secim_tipi = secim_tipi  # "tumunu" veya "tarihsizler"
        self.title("EXIF Tarih Düzenleyici - Tüm Fotoğraflar")
        self.geometry("800x870")
        self.configure(bg="#2c3e50")
        
        # Değişkenler
        self.files_list = []
        self.tum_fotograflar = []
        self.secilen_fotograflar = []
        self.manuel_klasor_yolu = None  # YENİ: Manuel klasör değişkeni
        
        self.setup_ui()
        
        if klasor_yolu:
            self.folder_path.set(klasor_yolu)
            self.manuel_klasor_yolu = klasor_yolu  # YENİ: Manuel klasörü ayarla
            self.tum_fotograflari_tara()
            
    def manuel_klasor_degistir(self):
        """Manuel olarak klasör değiştirme - YENİ EKLENDİ"""
        yeni_klasor = filedialog.askdirectory(title="Yeni Klasör Seçin")
        if yeni_klasor:
            self.manuel_klasor_yolu = yeni_klasor
            self.folder_path.set(yeni_klasor)
            self.klasor_yolu = yeni_klasor  # Orijinal klasör yolunu da güncelle
            self.log(f"📁 Klasör değiştirildi: {yeni_klasor}")
            self.tum_fotograflari_tara()
            
    def klasoru_yenile(self):
        """Klasörü yeniden tarar ve mevcut modda kalır - GÜNCELLENDİ"""
        if self.manuel_klasor_yolu:
            self.log("🔄 Klasör yeniden taranıyor...")
            # MEVCUT MODU KORUYARAK yeniden tarama yap
            self.tum_fotograflari_tara()
            
            # Mod bilgisini güncelle
            if self.secim_tipi == "tarihsizler":
                self.selection_info_var.set("🔍 Sadece tarihsiz fotoğraflar gösteriliyor")
            else:
                self.selection_info_var.set("🔍 Tüm fotoğraflar gösteriliyor")
                
        else:
            messagebox.showwarning("Uyarı", "Lütfen önce bir klasör seçin!")
            
    def listbox_sol_tik_olayi(self, event):
        """Listbox'ta sol tık olayını işler - DÜZELTİLDİ"""
        # Bu metod sadece seçim değişikliklerini takip etmek için
        # Asıl seçim işlemini Listbox'ın kendisi yapacak
        self.after(10, self.secilenleri_guncelle)  # 10ms sonra seçimleri güncelle

        
    def fotograf_detay_goster(self, event):
        """Fotoğraf detaylarını gösterir - YENİ EKLENDİ"""
        secili_indeksler = self.foto_listbox.curselection()
        if not secili_indeksler:
            return
        
        ilk_secili = secili_indeksler[0]
        if ilk_secili < len(self.tum_fotograflar):
            foto_path = self.tum_fotograflar[ilk_secili]
            self.fotograf_detay_penceresi_ac(foto_path)

    def secilenleri_guncelle(self):
        """Seçilen fotoğrafları günceller"""
        secili_indeksler = self.foto_listbox.curselection()
        self.secilen_fotograflar = [self.tum_fotograflar[i] for i in secili_indeksler]
        self.selected_var.set(f"Seçilen: {len(secili_indeksler)}")
        
    def tumunu_goster(self):
        """Tüm fotoğrafları göster - GÜNCELLENDİ"""
        self.secim_tipi = "tumunu"
        self.selection_info_var.set("🔍 Tüm fotoğraflar gösteriliyor")
        self.log("🔄 Tüm fotoğraflar gösteriliyor...")
        self.tum_fotograflari_tara()

    def sadece_tarihsizleri_goster(self):
        """Sadece tarihsiz fotoğrafları göster - GÜNCELLENDİ"""
        self.secim_tipi = "tarihsizler"
        self.selection_info_var.set("🔍 Sadece tarihsiz fotoğraflar gösteriliyor")
        self.log("🔄 Sadece tarihsiz fotoğraflar gösteriliyor...")
        self.tum_fotograflari_tara()
        
        # Eğer tarihsiz fotoğraf yoksa bilgi ver
        tarihsiz_sayisi = len([f for f in self.tum_fotograflar if not self.exif_tarihi_var_mi(f)])
        if tarihsiz_sayisi == 0:
            self.log("ℹ️ Tarihsiz fotoğraf bulunamadı!")
        
    def setup_ui(self):
        # Başlık - SEÇİM TİPİNE GÖRE DEĞİŞSİN
        baslik_metni = "📅 EXİF TARİH DÜZENLEYİCİ - "
        if self.secim_tipi == "tarihsizler":
            baslik_metni += "SADECE TARİHSİZ FOTOĞRAFLAR"
        else:
            baslik_metni += "TÜM FOTOĞRAFLAR"
        
        title_label = tk.Label(self, text=baslik_metni, 
                             font=("Arial", 12, "bold"), bg="#2c3e50", fg="white")
        title_label.pack(pady=10)
        
        
        # Klasör bilgisi - YENİ: Manuel klasör değiştirme eklendi
        folder_frame = tk.LabelFrame(self, text="📂 Klasör Seçimi", bg="#34495e", fg="white")
        folder_frame.pack(pady=10, fill="x", padx=20)
        
        folder_inner_frame = tk.Frame(folder_frame, bg="#34495e")
        folder_inner_frame.pack(pady=10, padx=10, fill="x")
        
        self.folder_path = tk.StringVar()
        folder_entry = tk.Entry(folder_inner_frame, textvariable=self.folder_path, width=80, 
                               bg="#ecf0f1", fg="#2c3e50", state='readonly')
        folder_entry.pack(side="left", padx=5, fill="x", expand=True)
        
        # Manuel klasör değiştirme butonu - YENİ EKLENDİ
        change_folder_btn = tk.Button(folder_inner_frame, text="Klasör Değiştir", 
                                    command=self.manuel_klasor_degistir,
                                    bg="#3498db", fg="white", font=("Arial", 9, "bold"))
        change_folder_btn.pack(side="left", padx=5)
        
        # Yenile butonu - YENİ EKLENDİ
        refresh_btn = tk.Button(folder_inner_frame, text="Yenile", 
                              command=self.klasoru_yenile,
                              bg="#27ae60", fg="white", font=("Arial", 9, "bold"))
        refresh_btn.pack(side="left", padx=5)
        
        # İstatistikler
        stats_frame = tk.Frame(self, bg="#2c3e50")
        stats_frame.pack(pady=10, fill="x", padx=20)
        
        self.total_var = tk.StringVar(value="Toplam Fotoğraf: 0")
        self.no_date_var = tk.StringVar(value="Tarihsiz Fotoğraf: 0")
        self.selected_var = tk.StringVar(value="Seçilen: 0")
        
        tk.Label(stats_frame, textvariable=self.total_var, bg="#2c3e50", fg="#3498db", 
                font=("Arial", 10, "bold")).pack(side="left", padx=20)
        tk.Label(stats_frame, textvariable=self.no_date_var, bg="#2c3e50", fg="#e74c3c", 
                font=("Arial", 10, "bold")).pack(side="left", padx=20)
        tk.Label(stats_frame, textvariable=self.selected_var, bg="#2c3e50", fg="#27ae60", 
                font=("Arial", 10, "bold")).pack(side="left", padx=20)
        
        # Seçim durumu bilgisi
        selection_info_frame = tk.Frame(self, bg="#2c3e50")
        selection_info_frame.pack(pady=5, fill="x", padx=20)
        
        self.selection_info_var = tk.StringVar()
        if self.secim_tipi == "tarihsizler":
            self.selection_info_var.set("🔍 Sadece tarihsiz fotoğraflar gösteriliyor")
        else:
            self.selection_info_var.set("🔍 Tüm fotoğraflar gösteriliyor")
        
        selection_info_label = tk.Label(selection_info_frame, textvariable=self.selection_info_var,
                                      font=("Arial", 9, "bold"), bg="#2c3e50", fg="#f39c12")
        selection_info_label.pack()
        
        # Yeni tarih seçimi ve butonlar için ana çerçeve
        date_buttons_frame = tk.Frame(self, bg="#2c3e50")
        date_buttons_frame.pack(pady=10, fill="x", padx=20)
        
        # Sol taraf: Tarih seçimi
        date_left_frame = tk.Frame(date_buttons_frame, bg="#2c3e50")
        date_left_frame.pack(side="left", fill="x", expand=True)
        
        date_frame = tk.LabelFrame(date_left_frame, text="🕒 Yeni Çekim Tarihi", bg="#34495e", fg="white")
        date_frame.pack(fill="x")
        
        inner_date_frame = tk.Frame(date_frame, bg="#34495e")
        inner_date_frame.pack(pady=15, padx=10)
        
        self.date_var = tk.StringVar(value=datetime.datetime.now().strftime("%Y:%m:%d"))
        self.time_var = tk.StringVar(value=datetime.datetime.now().strftime("%H:%M:%S"))
        
        # Tarih satırı
        date_row = tk.Frame(inner_date_frame, bg="#34495e")
        date_row.pack(pady=5, fill="x")
        tk.Label(date_row, text="Tarih:", width=8, bg="#34495e", fg="white", 
                font=("Arial", 10, "bold")).pack(side="left")
        tk.Entry(date_row, textvariable=self.date_var, width=20, font=("Arial", 10),
                bg="#ecf0f1", fg="#2c3e50").pack(side="left", padx=5)
        tk.Label(date_row, text="(YYYY:AA:GG - Örnek: 2024:12:31)", bg="#34495e", fg="white").pack(side="left", padx=5)
        
        # Saat satırı
        time_row = tk.Frame(inner_date_frame, bg="#34495e")
        time_row.pack(pady=5, fill="x")
        tk.Label(time_row, text="Saat:", width=8, bg="#34495e", fg="white",
                font=("Arial", 10, "bold")).pack(side="left")
        tk.Entry(time_row, textvariable=self.time_var, width=20, font=("Arial", 10),
                bg="#ecf0f1", fg="#2c3e50").pack(side="left", padx=5)
        tk.Label(time_row, text="(SS:DD:SS - Örnek: 14:30:25)", bg="#34495e", fg="white").pack(side="left", padx=5)
        
        
        # Sağ taraf: Butonlar
        buttons_right_frame = tk.Frame(date_buttons_frame, bg="#2c3e50")
        buttons_right_frame.pack(side="right", padx=(20, 0))
        
        buttons_frame = tk.LabelFrame(buttons_right_frame, text="⚡ İşlemler", bg="#34495e", fg="white")
        buttons_frame.pack(fill="both", expand=True)
        
        buttons_inner_frame = tk.Frame(buttons_frame, bg="#34495e")
        buttons_inner_frame.pack(pady=15, padx=15)
        
        tk.Button(buttons_inner_frame, text="🔄 Tümüne Tarih Ekle", 
                  command=self.tum_fotograflara_tarih_ekle, width=25, 
                  bg="#27ae60", fg="white", font=("Arial", 10, "bold")).pack(pady=5)
        
        tk.Button(buttons_inner_frame, text="🔄 Seçililere Tarih Ekle", 
                  command=self.secili_fotograflara_tarih_ekle, width=25,
                  bg="#3498db", fg="white", font=("Arial", 10, "bold")).pack(pady=5)
        
        # Seçim değiştirme butonu - GÜNCELLENDİ
        self.switch_button = tk.Button(buttons_inner_frame, text="", 
                                      command=self.secim_modunu_degistir, width=25,
                                      bg="#3498db", fg="white", font=("Arial", 10))
        self.switch_button.pack(pady=5)

        # Buton metnini ilk kez güncelle
        self.secim_buton_metnini_guncelle()

        tk.Button(buttons_inner_frame, text="📊 İstatistikler", 
                  command=self.istatistikleri_goster, width=25,
                  bg="#f39c12", fg="white", font=("Arial", 10)).pack(pady=5)
        
        
        # Fotoğraf listesi
        list_frame = tk.LabelFrame(self, text="📷 Fotoğraf Listesi", bg="#34495e", fg="white", height=150)
        list_frame.pack(pady=10, fill="x", padx=20)
        list_frame.pack_propagate(False)

        list_inner_frame = tk.Frame(list_frame, bg="#34495e")
        list_inner_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Listbox ve scrollbar
        self.listbox_frame = tk.Frame(list_inner_frame, bg="#34495e")
        self.listbox_frame.pack(fill="both", expand=True)
        
        self.foto_listbox = tk.Listbox(self.listbox_frame, selectmode=tk.EXTENDED,  # DEĞİŞTİ: EXTENDED modu
                                      bg="#2c3e50", fg="#ecf0f1", font=("Consolas", 9),
                                      selectbackground="#3498db", selectforeground="white")
        
        scrollbar_y = ttk.Scrollbar(self.listbox_frame, orient="vertical", command=self.foto_listbox.yview)
        scrollbar_x = ttk.Scrollbar(self.listbox_frame, orient="horizontal", command=self.foto_listbox.xview)
        
        self.foto_listbox.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        
        self.foto_listbox.pack(side="left", fill="both", expand=True)
        scrollbar_y.pack(side="right", fill="y")
        scrollbar_x.pack(side="bottom", fill="x")
        
        # YENİ: Çift tıklama olayı - fotoğraf detayı göstermek için
        self.foto_listbox.bind('<Double-1>', self.fotograf_detay_goster)
        
        # YENİ: Sol tık olayı - seçimleri güncellemek için
        self.foto_listbox.bind('<Button-1>', self.listbox_sol_tik_olayi)
        
        # Seçim butonları - GÜNCELLENDİ
        selection_frame = tk.Frame(list_inner_frame, bg="#34495e")
        selection_frame.pack(fill="x", pady=5)
        
        tk.Button(selection_frame, text="Tümünü Seç", command=self.tumunu_sec,
                 bg="#3498db", fg="white", width=12).pack(side="left", padx=5)
        tk.Button(selection_frame, text="Seçimi Kaldır", command=self.secimi_kaldir,
                 bg="#e74c3c", fg="white", width=12).pack(side="left", padx=5)
        tk.Button(selection_frame, text="Tarihsizleri Seç", command=self.tarihsizleri_sec,
                 bg="#f39c12", fg="white", width=14).pack(side="left", padx=5)
        tk.Button(selection_frame, text="Tarihlileri Seç", command=self.tarihlileri_sec,
                 bg="#27ae60", fg="white", width=14).pack(side="left", padx=5)
        
        # İlerleme çubuğu
        progress_frame = tk.Frame(self, bg="#2c3e50")
        progress_frame.pack(pady=10, fill="x", padx=20)
        
        tk.Label(progress_frame, text="İlerleme:", bg="#2c3e50", fg="white").pack(anchor="w")
        self.progress = ttk.Progressbar(progress_frame, orient="horizontal", length=760, mode="determinate")
        self.progress.pack(pady=5, fill="x")
        
        # Durum bilgisi
        self.status_var = tk.StringVar(value="Tüm fotoğraflar taranıyor...")
        status_label = tk.Label(self, textvariable=self.status_var, foreground="#3498db", bg="#2c3e50")
        status_label.pack(pady=5)
        
        # Log alanı
        log_frame = tk.LabelFrame(self, text="📋 İşlem Logları", bg="#34495e", fg="white", height=200)
        log_frame.pack(pady=10, fill="x", padx=20)
        log_frame.pack_propagate(False)

        log_inner_frame = tk.Frame(log_frame, bg="#34495e")
        log_inner_frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.log_text = tk.Text(log_inner_frame, height=8, width=85, font=("Consolas", 9),
                               bg="#2c3e50", fg="#ecf0f1")

        scrollbar_log = ttk.Scrollbar(log_inner_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar_log.set)

        self.log_text.pack(side="left", fill="both", expand=True)
        scrollbar_log.pack(side="right", fill="y")
        
        # Başlangıç logu
        if self.secim_tipi == "tarihsizler":
            self.log("EXIF tarih düzenleyici başlatıldı. Sadece tarihsiz fotoğraflar işlenecek.")
        else:
            self.log("EXIF tarih düzenleyici başlatıldı. Tüm fotoğraflar işlenecek.")


    def secim_buton_metnini_guncelle(self):
        """Seçim değiştirme butonunun metnini günceller"""
        if self.secim_tipi == "tarihsizler":
            self.switch_button.config(text="📷 Tümünü Göster", bg="#3498db")
        else:
            self.switch_button.config(text="📅 Sadece Tarihsizleri Göster", bg="#f39c12")

    def secim_modunu_degistir(self):
        """Seçim modunu değiştirir"""
        if self.secim_tipi == "tarihsizler":
            self.tumunu_goster()
        else:
            self.sadece_tarihsizleri_goster()
        
        self.secim_buton_metnini_guncelle()


    def tum_fotograflari_tara(self):
        """Klasördeki TÜM fotoğrafları tarar - SEÇİM TİPİNE GÖRE"""
        # Manuel klasör yolunu kullan
        tarama_klasoru = self.manuel_klasor_yolu if self.manuel_klasor_yolu else self.klasor_yolu
        
        if not tarama_klasoru or not os.path.exists(tarama_klasoru):
            messagebox.showerror("Hata", "Klasör bulunamadı!")
            return
        
        self.log("🔍 Tüm fotoğraflar taranıyor (alt klasörler dahil)...")
        self.status_var.set("Tüm fotoğraflar taranıyor...")
        
        # Tüm fotoğrafları bul - ALT KLASÖRLER DAHİL
        foto_uzantilari = ('.jpg', '.jpeg', '.png', '.tiff', '.tif', '.bmp')
        tum_fotograflar = []
        tarihsiz_fotograflar = []
        tarihli_fotograflar = []
        
        # ALT KLASÖRLERİ DE TARA
        for root, dirs, files in os.walk(tarama_klasoru):
            for dosya in files:
                if any(dosya.lower().endswith(ext) for ext in foto_uzantilari):
                    full_path = os.path.join(root, dosya)
                    tum_fotograflar.append(full_path)
                    
                    # Tarihsiz olanları ve tarihli olanları ayır
                    if not self.exif_tarihi_var_mi(full_path):
                        tarihsiz_fotograflar.append(full_path)
                    else:
                        tarihli_fotograflar.append(full_path)
        
        self.tum_fotograflar = tum_fotograflar
        
        # SEÇİM TİPİNE GÖRE FOTOĞRAFLARI BELİRLE
        if self.secim_tipi == "tarihsizler":
            islenecek_fotograflar = tarihsiz_fotograflar
            self.log(f"🎯 Sadece tarihsiz fotoğraflar seçildi")
            
            # EĞER TARİHSİZ FOTOĞRAF YOKSA OTOMATİK TÜMÜNE DÖN
            if not tarihsiz_fotograflar:
                self.log("ℹ️ Tarihsiz fotoğraf bulunamadı, tüm fotoğraflar gösteriliyor...")
                islenecek_fotograflar = tum_fotograflar
                self.secim_tipi = "tumunu"  # Modu değiştir
                self.selection_info_var.set("🔍 Tüm fotoğraflar gösteriliyor (tarihsiz fotoğraf yok)")
        else:
            islenecek_fotograflar = tum_fotograflar
            self.log(f"🎯 Tüm fotoğraflar seçildi")
        
        self.total_var.set(f"Toplam Fotoğraf: {len(tum_fotograflar)}")
        self.no_date_var.set(f"Tarihsiz Fotoğraf: {len(tarihsiz_fotograflar)}")
        self.selected_var.set(f"Seçilen: {len(islenecek_fotograflar)}")
        
        # Listbox'ı doldur
        self.foto_listbox.delete(0, tk.END)
        for foto_path in islenecek_fotograflar:
            dosya_adi = os.path.basename(foto_path)
            # Göreli yolu göster (ana klasöre göre)
            try:
                goreli_yol = os.path.relpath(foto_path, tarama_klasoru)
                list_item = f"{goreli_yol}"
            except:
                list_item = dosya_adi
            
            # Fotoğrafın tarih bilgisini de göster
            tarih = self.fotograf_tarihini_al(foto_path)
            if tarih:
                tarih_str = tarih.strftime('%d/%m/%Y %H:%M')
                list_item += f" - {tarih_str}"
            else:
                list_item += " - TARİHSİZ"
            
            self.foto_listbox.insert(tk.END, list_item)
        
        # SEÇİM TİPİNE GÖRE OTOMATİK SEÇ
        self.foto_listbox.selection_set(0, tk.END)
        self.secilen_fotograflar = islenecek_fotograflar.copy()
        
        self.log(f"✅ Tarama tamamlandı!")
        self.log(f"📊 Toplam: {len(tum_fotograflar)} fotoğraf")
        self.log(f"📅 Tarihli: {len(tarihli_fotograflar)} fotoğraf")
        self.log(f"📅 Tarihsiz: {len(tarihsiz_fotograflar)} fotoğraf")
        self.log(f"🎯 İşlenecek: {len(islenecek_fotograflar)} fotoğraf")
        self.log(f"📁 Ana Klasör: {tarama_klasoru}")
        
        # Eğer tarihsiz modda ama tarihsiz fotoğraf yoksa, buton metnini güncelle
        if self.secim_tipi == "tarihsizler" and not tarihsiz_fotograflar:
            self.selection_info_var.set("ℹ️ Tarihsiz fotoğraf bulunamadı, tüm fotoğraflar gösteriliyor")
        
        self.status_var.set(f"Hazır - {len(islenecek_fotograflar)} fotoğraf seçildi")


    def tumunu_sec(self):
        """Tüm fotoğrafları seçer"""
        self.foto_listbox.selection_set(0, tk.END)
        secilen_sayisi = len(self.foto_listbox.curselection())
        self.secilen_fotograflar = [self.tum_fotograflar[i] for i in self.foto_listbox.curselection()]
        self.selected_var.set(f"Seçilen: {secilen_sayisi}")
        self.log(f"✓ {secilen_sayisi} fotoğraf seçildi")

    def secimi_kaldir(self):
        """Seçimi kaldırır"""
        self.foto_listbox.selection_clear(0, tk.END)
        self.secilen_fotograflar = []
        self.selected_var.set("Seçilen: 0")
        self.log("✗ Seçim kaldırıldı")

    def fotograf_tarihini_al(self, dosya_yolu):
        """Fotoğrafın EXIF tarihini alır"""
        try:
            with Image.open(dosya_yolu) as img:
                exif_data = img.getexif()
                if exif_data:
                    # Öncelikle çekilme tarihini ara
                    datetime_original = exif_data.get(36867)  # DateTimeOriginal
                    if not datetime_original:
                        datetime_original = exif_data.get(306)  # DateTime
                    
                    if datetime_original:
                        try:
                            return datetime.datetime.strptime(datetime_original, '%Y:%m:%d %H:%M:%S')
                        except ValueError:
                            try:
                                return datetime.datetime.strptime(datetime_original, '%Y-%m:%d %H:%M:%S')
                            except:
                                return None
                
                # Piexif ile de dene
                try:
                    exif_dict = piexif.load(dosya_yolu)
                    if piexif.ExifIFD.DateTimeOriginal in exif_dict['Exif']:
                        tarih_bytes = exif_dict['Exif'][piexif.ExifIFD.DateTimeOriginal]
                        tarih_str = tarih_bytes.decode('utf-8')
                        return datetime.datetime.strptime(tarih_str, '%Y:%m:%d %H:%M:%S')
                except:
                    pass
                
                return None
                
        except Exception as e:
            print(f"EXIF okuma hatası {dosya_yolu}: {e}")
            return None

    def exif_tarihi_var_mi(self, dosya_yolu):
        """Dosyada EXIF tarihi olup olmadığını kontrol eder"""
        try:
            with Image.open(dosya_yolu) as img:
                exif_data = img.getexif()
                if exif_data:
                    # EXIF tarih tag'lerini kontrol et
                    datetime_original = exif_data.get(36867)  # DateTimeOriginal
                    datetime_digitized = exif_data.get(36868)  # DateTimeDigitized
                    datetime_normal = exif_data.get(306)       # DateTime
                    
                    # Herhangi bir tarih varsa True döndür
                    if datetime_original or datetime_digitized or datetime_normal:
                        return True
                
                # Piexif ile de kontrol et
                try:
                    exif_dict = piexif.load(dosya_yolu)
                    if exif_dict['Exif']:
                        if piexif.ExifIFD.DateTimeOriginal in exif_dict['Exif']:
                            return True
                        if piexif.ExifIFD.DateTimeDigitized in exif_dict['Exif']:
                            return True
                    if exif_dict['0th']:
                        if piexif.ImageIFD.DateTime in exif_dict['0th']:
                            return True
                except:
                    pass
                
                return False
                
        except Exception as e:
            self.log(f"⚠️ Tarih kontrol hatası {os.path.basename(dosya_yolu)}: {str(e)}")
            return False

    def tum_fotograflara_tarih_ekle(self):
        """Tüm seçili fotoğraflara tarih ekler - GÜNCELLENMİŞ"""
        # Manuel klasör yolunu kullan
        tarama_klasoru = self.manuel_klasor_yolu if self.manuel_klasor_yolu else self.klasor_yolu
        
        if not self.secilen_fotograflar:
            messagebox.showwarning("Uyarı", "Lütfen tarih eklemek için fotoğraf seçin!")
            return
        
        if not self.date_var.get() or not self.time_var.get():
            messagebox.showwarning("Uyarı", "Lütfen tarih ve saat bilgisini girin!")
            return
        
        # Tarih formatı kontrolü
        try:
            datetime.datetime.strptime(self.date_var.get(), "%Y:%m:%d")
            datetime.datetime.strptime(self.time_var.get(), "%H:%M:%S")
        except ValueError:
            messagebox.showerror("Hata", "Tarih veya saat formatı hatalı!\nTarih: YYYY:AA:GG\nSaat: SS:DD:SS")
            return
        
        new_datetime = f"{self.date_var.get().strip()} {self.time_var.get().strip()}"
        
        result = messagebox.askyesno("Onay", 
                                   f"{len(self.secilen_fotograflar)} fotoğrafa tarih eklenecek:\n"
                                   f"Yeni Tarih: {new_datetime}\n\n"
                                   f"Devam etmek istiyor musunuz?")
        if not result:
            return
        
        self.log("🔄 Tüm fotoğraflara tarih ekleniyor...")
        self.log(f"📅 Yeni tarih: {new_datetime}")
        self.log(f"📁 İşlem yapılan klasör: {tarama_klasoru}")
        
        # İlerleme çubuğunu ayarla
        self.progress['maximum'] = len(self.secilen_fotograflar)
        self.progress['value'] = 0
        
        success_count = 0
        error_count = 0
        verified_count = 0
        
        for i, foto_path in enumerate(self.secilen_fotograflar):
            try:
                # Önceki tarihi kaydet (log için)
                previous_date = self.fotograf_tarihini_al(foto_path)
                previous_str = previous_date.strftime('%Y:%m:%d %H:%M:%S') if previous_date else "Tarihsiz"
                
                # Göreli yolu al (daha okunabilir log için)
                try:
                    goreli_yol = os.path.relpath(foto_path, tarama_klasoru)
                    log_adi = goreli_yol
                except:
                    log_adi = os.path.basename(foto_path)
                
                # Tarihi güncelle
                if self.update_exif_date_advanced(foto_path, new_datetime):
                    # Güncellemeyi doğrula
                    if self.verify_exif_date(foto_path, new_datetime):
                        verified_count += 1
                        self.log(f"✅ {log_adi} - Tarih değiştirildi: {previous_str} → {new_datetime}")
                        success_count += 1
                    else:
                        error_count += 1
                        self.log(f"❌ {log_adi} - Tarih doğrulanamadı!")
                else:
                    error_count += 1
                    self.log(f"❌ {log_adi} - Tarih güncellenemedi!")
            
            except Exception as e:
                error_count += 1
                self.log(f"❌ {os.path.basename(foto_path)} - Hata: {str(e)}")
            
            # İlerleme çubuğunu güncelle
            self.progress['value'] = i + 1
            self.status_var.set(f"İşleniyor... {i+1}/{len(self.secilen_fotograflar)}")
            self.update()
        
        # Sonuçları göster
        self.log("=" * 50)
        self.log(f"🎉 İŞLEM TAMAMLANDI!")
        self.log(f"✅ Başarılı: {success_count} fotoğraf")
        self.log(f"✅ Doğrulanan: {verified_count} fotoğraf") 
        self.log(f"❌ Hatalı: {error_count} fotoğraf")
        self.log(f"📁 İşlem yapılan: {tarama_klasoru}")
        self.log("=" * 50)
        
        messagebox.showinfo("Tamamlandı", 
                          f"İşlem tamamlandı!\n\n"
                          f"✅ Başarılı: {success_count} fotoğraf\n"
                          f"✅ Doğrulanan: {verified_count} fotoğraf\n"
                          f"❌ Hatalı: {error_count} fotoğraf\n\n"
                          f"Yeni tarih: {new_datetime}")
        
        self.status_var.set(f"Tamamlandı - Başarılı: {success_count}, Hatalı: {error_count}")

    def update_exif_date_advanced(self, file_path, new_datetime):
        """EXIF tarihini güncellemek için gelişmiş ve güvenilir yöntem"""
        try:
            # Dosya kontrolü
            if not os.path.exists(file_path):
                raise Exception(f"Dosya bulunamadı: {file_path}")
            
            if not os.access(file_path, os.W_OK):
                raise Exception(f"Dosya yazılabilir değil: {file_path}")
            
            # Yedek oluştur
            backup_path = file_path + '.backup'
            shutil.copy2(file_path, backup_path)
            
            try:
                # Tarih formatını kontrol et
                try:
                    datetime.datetime.strptime(new_datetime, '%Y:%m:%d %H:%M:%S')
                except ValueError:
                    raise Exception(f"Geçersiz tarih formatı: {new_datetime}")
                
                # Fotoğrafı PIL ile aç
                with Image.open(file_path) as img:
                    # Mevcut EXIF verisini al veya yeni oluştur
                    exif_dict = {}
                    try:
                        exif_dict = piexif.load(img.info.get('exif', b''))
                    except:
                        exif_dict = {"0th": {}, "Exif": {}, "GPS": {}, "1st": {}, "Interop": {}}
                    
                    # Tarih bilgilerini byte formatına çevir
                    new_datetime_bytes = new_datetime.encode('utf-8')
                    
                    # Tüm tarih tag'lerini güncelle
                    # DateTimeOriginal (EXIF tag 36867)
                    exif_dict['Exif'][piexif.ExifIFD.DateTimeOriginal] = new_datetime_bytes
                    # DateTimeDigitized (EXIF tag 36868)
                    exif_dict['Exif'][piexif.ExifIFD.DateTimeDigitized] = new_datetime_bytes
                    # DateTime (Image tag 306)
                    exif_dict['0th'][piexif.ImageIFD.DateTime] = new_datetime_bytes
                    
                    # EXIF verisini byte'a çevir
                    exif_bytes = piexif.dump(exif_dict)
                    
                    # Fotoğrafı kaydet (orijinal formatını koru)
                    img_format = img.format
                    if img_format == 'JPEG':
                        # JPEG için doğrudan exif parametresi ile kaydet
                        img.save(file_path, "JPEG", quality=95, exif=exif_bytes)
                        self.log(f"✓ JPEG EXIF güncellendi: {os.path.basename(file_path)}")
                    else:
                        # Diğer formatlar için piexif.insert kullan
                        img.save(file_path, img_format)
                        piexif.insert(exif_bytes, file_path)
                        self.log(f"✓ {img_format} EXIF güncellendi: {os.path.basename(file_path)}")
                    
                    return True
                    
            except Exception as e:
                # Hata durumunda yedeği geri yükle
                if os.path.exists(backup_path):
                    shutil.copy2(backup_path, file_path)
                raise Exception(f"EXIF güncellenemedi: {str(e)}")
            finally:
                # Yedek dosyayı temizle
                if os.path.exists(backup_path):
                    try:
                        os.remove(backup_path)
                    except:
                        pass
                        
        except Exception as e:
            self.log(f"❌ {os.path.basename(file_path)} - Hata: {str(e)}")
            return False

    def verify_exif_date(self, file_path, expected_datetime):
        """EXIF tarihinin doğru güncellendiğini doğrular"""
        try:
            current_date = self.fotograf_tarihini_al(file_path)
            if current_date:
                current_str = current_date.strftime('%Y:%m:%d %H:%M:%S')
                return current_str == expected_datetime
            return False
        except:
            return False    

    def log(self, message):
        """Log mesajını ekrana yazar"""
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        log_message = f"[{timestamp}] {message}\n"
        self.log_text.insert(tk.END, log_message)
        self.log_text.see(tk.END)
        self.update()


    def secili_fotograflara_tarih_ekle(self):
        """Sadece seçili fotoğraflara tarih ekler - YENİ EKLENDİ"""
        secili_indeksler = self.foto_listbox.curselection()
        
        if not secili_indeksler:
            messagebox.showwarning("Uyarı", "Lütfen tarih eklemek için fotoğraf seçin!")
            return
        
        self.secilen_fotograflar = [self.tum_fotograflar[i] for i in secili_indeksler]
        self.secili_fotograflari_isle()

    def secili_fotograflari_isle(self):
        """Seçili fotoğrafları işler - YENİ EKLENDİ"""
        if not self.secilen_fotograflar:
            messagebox.showwarning("Uyarı", "Lütfen tarih eklemek için fotoğraf seçin!")
            return
        
        if not self.date_var.get() or not self.time_var.get():
            messagebox.showwarning("Uyarı", "Lütfen tarih ve saat bilgisini girin!")
            return
        
        # Tarih formatı kontrolü
        try:
            datetime.datetime.strptime(self.date_var.get(), "%Y:%m:%d")
            datetime.datetime.strptime(self.time_var.get(), "%H:%M:%S")
        except ValueError:
            messagebox.showerror("Hata", "Tarih veya saat formatı hatalı!\nTarih: YYYY:AA:GG\nSaat: SS:DD:SS")
            return
        
        new_datetime = f"{self.date_var.get().strip()} {self.time_var.get().strip()}"
        
        result = messagebox.askyesno("Onay", 
                                   f"{len(self.secilen_fotograflar)} seçili fotoğrafa tarih eklenecek:\n"
                                   f"Yeni Tarih: {new_datetime}\n\n"
                                   f"Devam etmek istiyor musunuz?")
        if not result:
            return
        
        # Manuel klasör yolunu kullan
        tarama_klasoru = self.manuel_klasor_yolu if self.manuel_klasor_yolu else self.klasor_yolu
        
        self.log("🔄 Seçili fotoğraflara tarih ekleniyor...")
        self.log(f"📅 Yeni tarih: {new_datetime}")
        self.log(f"📁 İşlem yapılan klasör: {tarama_klasoru}")
        
        # İlerleme çubuğunu ayarla
        self.progress['maximum'] = len(self.secilen_fotograflar)
        self.progress['value'] = 0
        
        success_count = 0
        error_count = 0
        verified_count = 0
        
        for i, foto_path in enumerate(self.secilen_fotograflar):
            try:
                # Önceki tarihi kaydet (log için)
                previous_date = self.fotograf_tarihini_al(foto_path)
                previous_str = previous_date.strftime('%Y:%m:%d %H:%M:%S') if previous_date else "Tarihsiz"
                
                # Göreli yolu al (daha okunabilir log için)
                try:
                    goreli_yol = os.path.relpath(foto_path, tarama_klasoru)
                    log_adi = goreli_yol
                except:
                    log_adi = os.path.basename(foto_path)
                
                # Tarihi güncelle
                if self.update_exif_date_advanced(foto_path, new_datetime):
                    # Güncellemeyi doğrula
                    if self.verify_exif_date(foto_path, new_datetime):
                        verified_count += 1
                        self.log(f"✅ {log_adi} - Tarih değiştirildi: {previous_str} → {new_datetime}")
                        success_count += 1
                    else:
                        error_count += 1
                        self.log(f"❌ {log_adi} - Tarih doğrulanamadı!")
                else:
                    error_count += 1
                    self.log(f"❌ {log_adi} - Tarih güncellenemedi!")
            
            except Exception as e:
                error_count += 1
                self.log(f"❌ {os.path.basename(foto_path)} - Hata: {str(e)}")
            
            # İlerleme çubuğunu güncelle
            self.progress['value'] = i + 1
            self.status_var.set(f"İşleniyor... {i+1}/{len(self.secilen_fotograflar)}")
            self.update()
        
        # Sonuçları göster
        self.log("=" * 50)
        self.log(f"🎉 SEÇİLİ FOTOĞRAFLAR İŞLEMİ TAMAMLANDI!")
        self.log(f"✅ Başarılı: {success_count} fotoğraf")
        self.log(f"✅ Doğrulanan: {verified_count} fotoğraf") 
        self.log(f"❌ Hatalı: {error_count} fotoğraf")
        self.log("=" * 50)
        
        messagebox.showinfo("Tamamlandı", 
                          f"Seçili fotoğraflar işlemi tamamlandı!\n\n"
                          f"✅ Başarılı: {success_count} fotoğraf\n"
                          f"✅ Doğrulanan: {verified_count} fotoğraf\n"
                          f"❌ Hatalı: {error_count} fotoğraf\n\n"
                          f"Yeni tarih: {new_datetime}")
        
        self.status_var.set(f"Tamamlandı - Başarılı: {success_count}, Hatalı: {error_count}")

    def tarihsizleri_sec(self):
        """Sadece tarihsiz fotoğrafları seçer - YENİ EKLENDİ"""
        self.foto_listbox.selection_clear(0, tk.END)
        tarihsiz_sayisi = 0
        
        for i, foto_path in enumerate(self.tum_fotograflar):
            if not self.exif_tarihi_var_mi(foto_path):
                self.foto_listbox.selection_set(i)
                tarihsiz_sayisi += 1
        
        self.secilen_fotograflar = [self.tum_fotograflar[i] for i in self.foto_listbox.curselection()]
        self.selected_var.set(f"Seçilen: {tarihsiz_sayisi}")
        self.log(f"✓ {tarihsiz_sayisi} tarihsiz fotoğraf seçildi")

    def tarihlileri_sec(self):
        """Sadece tarihli fotoğrafları seçer - YENİ EKLENDİ"""
        self.foto_listbox.selection_clear(0, tk.END)
        tarihli_sayisi = 0
        
        for i, foto_path in enumerate(self.tum_fotograflar):
            if self.exif_tarihi_var_mi(foto_path):
                self.foto_listbox.selection_set(i)
                tarihli_sayisi += 1
        
        self.secilen_fotograflar = [self.tum_fotograflar[i] for i in self.foto_listbox.curselection()]
        self.selected_var.set(f"Seçilen: {tarihli_sayisi}")
        self.log(f"✓ {tarihli_sayisi} tarihli fotoğraf seçildi")

    def istatistikleri_goster(self):
        """Detaylı istatistikleri gösterir - YENİ EKLENDİ"""
        if not self.tum_fotograflar:
            messagebox.showinfo("Bilgi", "Henüz hiç fotoğraf taranmamış!")
            return
        
        # Manuel klasör yolunu kullan
        tarama_klasoru = self.manuel_klasor_yolu if self.manuel_klasor_yolu else self.klasor_yolu
        
        toplam_foto = len(self.tum_fotograflar)
        tarihli_foto = 0
        tarihsiz_foto = 0
        
        for foto_path in self.tum_fotograflar:
            if self.exif_tarihi_var_mi(foto_path):
                tarihli_foto += 1
            else:
                tarihsiz_foto += 1
        
        tarihli_yuzde = (tarihli_foto / toplam_foto) * 100 if toplam_foto > 0 else 0
        tarihsiz_yuzde = (tarihsiz_foto / toplam_foto) * 100 if toplam_foto > 0 else 0
        
        istatistik_metni = (
            f"📊 DETAYLI İSTATİSTİKLER\n"
            f"{'='*40}\n"
            f"📁 Klasör: {tarama_klasoru}\n"
            f"📅 Tarama Zamanı: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n"
            f"{'='*40}\n"
            f"📷 Toplam Fotoğraf: {toplam_foto}\n"
            f"✅ Tarihli Fotoğraf: {tarihli_foto} (%{tarihli_yuzde:.1f})\n"
            f"❌ Tarihsiz Fotoğraf: {tarihsiz_foto} (%{tarihsiz_yuzde:.1f})\n"
            f"{'='*40}\n"
            f"🎯 Mevcut Gösterim: {self.secim_tipi}\n"
            f"👆 Şu an Seçili: {len(self.secilen_fotograflar)} fotoğraf"
        )
        
        # İstatistikleri log'a da yaz
        self.log("📊 İSTATİSTİKLER GÖSTERİLDİ:")
        self.log(f"   Toplam: {toplam_foto} fotoğraf")
        self.log(f"   Tarihli: {tarihli_foto} (%{tarihli_yuzde:.1f})")
        self.log(f"   Tarihsiz: {tarihsiz_foto} (%{tarihsiz_yuzde:.1f})")
        
        messagebox.showinfo("Detaylı İstatistikler", istatistik_metni)


class RaporEkrani(tk.Toplevel):
    def __init__(self, parent, excel_dosyasi, hat_tarihsiz_istatistikleri=None):
        super().__init__(parent)
        self.parent = parent
        self.excel_dosyasi = excel_dosyasi
        self.hat_tarihsiz_istatistikleri = hat_tarihsiz_istatistikleri or {}
        
        self.title("Rapor Özeti - Direk İstatistikleri")
        self.geometry("1800x1000")
        self.configure(bg="#2c3e50")
        
        self.colors = {
            'dark_bg': '#2c3e50', 'light_bg': '#34495e', 'accent': '#3498db',
            'success': '#27ae60', 'danger': '#e74c3c', 'warning': '#f39c12',
            'text': '#ecf0f1', 'border': '#7f8c8d', 'info_bg': '#2c3e50'
        }

        # Log için Text widget'ı oluştur
        self.log_text_widget = None
        
        # Değişkenler
        self.veriler = None
        self.rapor_verileri = None
        self.siralama_durumu = {}
        self.mevcut_filtreli_veri = []
        self.oncelik_veriler = []  # Boş liste olarak başlat
        
        self.turkce_aylar = [
            'Ocak', 'Şubat', 'Mart', 'Nisan', 'Mayıs', 'Haziran',
            'Temmuz', 'Ağustos', 'Eylül', 'Ekim', 'Kasım', 'Aralık'
        ]
        
        self.turkce_gunler = ['Pzt', 'Sal', 'Çar', 'Per', 'Cum', 'Cmt', 'Paz']
        
        # NOTEBOOK OLUŞTUR - 3 SEKMELİ
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=5)
        
        # 1. SEKME: İstatistikler
        self.istatistik_frame = tk.Frame(self.notebook, bg=self.colors['dark_bg'])
        self.notebook.add(self.istatistik_frame, text='📊 İstatistikler')
        
        # 2. SEKME: Öncelik Analizi
        self.oncelik_frame = tk.Frame(self.notebook, bg=self.colors['dark_bg'])
        self.notebook.add(self.oncelik_frame, text='🚨 Öncelik Analizi')
        
        # 3. SEKME: Harita Analizi - BOŞ FRAME OLUŞTUR, LAZY LOADING YAP
        self.harita_frame = tk.Frame(self.notebook, bg=self.colors['dark_bg'])
        self.notebook.add(self.harita_frame, text='📍 Harita Analizi')
        
        # Harita sekmesi yüklendi mi flag'i
        self.harita_yuklendi = False
        
        # Notebook değişimini takip et
        self.notebook.bind("<<NotebookTabChanged>>", self.sekme_degisti)
        
        # Sekmeleri oluştur
        self.istatistik_sekmesi_olustur()
        self.oncelik_sekmesi_olustur()
        # HARİTA SEKMESİNİ OLUŞTURMA - sadece frame oluşturduk
        
        # ÖNCELİK VERİLERİNİ OTOMATİK YÜKLE
        self.after(100, self.oncelik_verileri_yukle)  # 100ms sonra otomatik yükle
        
        if not self.verileri_yukle():
            return




    def log_mesaj_ekle(self, mesaj):
        """Log mesajı ekler (AnaUygulama ile uyumlu)"""
        # Ana uygulamanın log'una yazmaya çalış
        try:
            if hasattr(self.parent, 'log_mesaj_ekle'):
                self.parent.log_mesaj_ekle(mesaj)
            else:
                # Kendi log'unu oluştur
                if not hasattr(self, 'txt_sonuc'):
                    # Text widget'ı oluştur
                    self.txt_sonuc = tk.Text(self, height=10, width=100, bg="#2c3e50", fg="white")
                    self.txt_sonuc.pack(fill='x', padx=10, pady=5)
                
                timestamp = datetime.datetime.now().strftime("[%H:%M:%S]")
                self.txt_sonuc.insert(tk.END, f"{timestamp} {mesaj}\n")
                self.txt_sonuc.see(tk.END)
        except Exception as e:
            print(f"Log hatası: {e}")

        
    def sekme_degisti(self, event):
        """Notebook sekmesi değiştiğinde çağrılır"""
        secili_sekme = self.notebook.index(self.notebook.select())
        
        # Eğer 3. sekme (Harita Analizi) seçildiyse ve henüz yüklenmediyse
        if secili_sekme == 2 and not self.harita_yuklendi:
            self.harita_sekmesi_olustur()
            self.harita_yuklendi = True
            
    def istatistik_sekmesi_olustur(self):
        """İstatistikler sekmesini oluşturur (mevcut içerik)"""
        # ANA PENCEREYİ 9 BİRİM OLARAK AYARLA
        self.istatistik_frame.grid_rowconfigure(0, weight=3)  # Üst bölüm: 3 birim
        self.istatistik_frame.grid_rowconfigure(1, weight=3)  # Grafik 1: 3 birim
        self.istatistik_frame.grid_rowconfigure(2, weight=3)  # Grafik 2: 3 birim
        self.istatistik_frame.grid_columnconfigure(0, weight=1)
        
        # 1. BÖLÜM: ÜST BAŞLIK VE İSTATİSTİKLER (3 BİRİM)
        top_frame = tk.Frame(self.istatistik_frame, bg=self.colors['dark_bg'])
        top_frame.grid(row=0, column=0, sticky='nsew', padx=20, pady=10)
        
        title_frame = tk.Frame(top_frame, bg=self.colors['dark_bg'])
        title_frame.pack(fill='x', pady=(0, 10))
        
        title_label = tk.Label(title_frame, text="📊 DİREK İSTATİSTİK RAPORU", 
                              font=('Segoe UI', 20, 'bold'),
                              fg=self.colors['text'], bg=self.colors['dark_bg'])
        title_label.pack()
        
        dosya_label = tk.Label(title_frame, text=f"Dosya: {os.path.basename(self.excel_dosyasi)}", 
                              font=('Segoe UI', 10),
                              fg=self.colors['accent'], bg=self.colors['dark_bg'])
        dosya_label.pack()
        
        stats_frame = tk.Frame(top_frame, bg=self.colors['dark_bg'])
        stats_frame.pack(fill='x', pady=10)
        
        self.kartlar_frame = tk.Frame(stats_frame, bg=self.colors['dark_bg'])
        self.kartlar_frame.pack(fill='x')
        
        # 2. VE 3. BÖLÜM İÇİN ANA İÇERİK FRAME
        content_frame = tk.Frame(self.istatistik_frame, bg=self.colors['dark_bg'])
        content_frame.grid(row=1, column=0, rowspan=2, sticky='nsew', padx=20, pady=(0, 10))
        
        # CONTENT FRAME'İ 2 SATIRA BÖL (GRAFİK 1 ve GRAFİK 2 için)
        content_frame.grid_rowconfigure(0, weight=1)  # Grafik 1: 3 birim
        content_frame.grid_rowconfigure(1, weight=1)  # Grafik 2: 3 birim
        content_frame.grid_columnconfigure(0, weight=1)
        content_frame.grid_columnconfigure(1, weight=1)
        
        # SOL TARAF: TABLO (GRAFİKLERLE AYNI YÜKSEKLİKTE)
        left_frame = tk.LabelFrame(content_frame, text="🛣️ Hat Detayları", 
                                  font=('Segoe UI', 14, 'bold'),
                                  fg=self.colors['text'], bg=self.colors['light_bg'],
                                  padx=15, pady=15)
        left_frame.grid(row=0, column=0, rowspan=2, sticky='nsew', padx=(0, 10))
        
        # SAĞ TARAF: GRAFİKLER (ÜST ÜSTE)
        right_frame = tk.Frame(content_frame, bg=self.colors['dark_bg'])
        right_frame.grid(row=0, column=1, rowspan=2, sticky='nsew', padx=(10, 0))
        
        # SAĞ FRAME'İ 2 EŞİT PARÇAYA BÖL
        right_frame.grid_rowconfigure(0, weight=1)  # Grafik 1: 3 birim
        right_frame.grid_rowconfigure(1, weight=1)  # Grafik 2: 3 birim
        right_frame.grid_columnconfigure(0, weight=1)
        
        # SOL FRAME İÇERİĞİ (TABLO)
        filter_frame = tk.Frame(left_frame, bg=self.colors['light_bg'])
        filter_frame.pack(fill='x', pady=(0, 10))
        
        tk.Label(filter_frame, text="İlçe:", font=('Segoe UI', 10, 'bold'),
                bg=self.colors['light_bg'], fg=self.colors['text']).pack(side='left', padx=(0, 5))
        
        self.ilce_filter_var = tk.StringVar(value="TÜMÜ")
        self.ilce_filter = ttk.Combobox(filter_frame, textvariable=self.ilce_filter_var, state="readonly", width=15)
        self.ilce_filter.pack(side='left', padx=(0, 15))
        
        tarih_filter_frame = tk.Frame(filter_frame, bg=self.colors['light_bg'])
        tarih_filter_frame.pack(side='left', padx=(0, 15))
        
        tk.Label(tarih_filter_frame, text="Tarih Aralığı:", font=('Segoe UI', 10, 'bold'),
                bg=self.colors['light_bg'], fg=self.colors['text']).pack(anchor='w')
        
        tarih_selection_frame = tk.Frame(tarih_filter_frame, bg=self.colors['light_bg'])
        tarih_selection_frame.pack(fill='x', pady=(5, 0))
        
        start_frame = tk.Frame(tarih_selection_frame, bg=self.colors['light_bg'])
        start_frame.pack(side='left', padx=(0, 10))
        
        tk.Label(start_frame, text="Başlangıç:", font=('Segoe UI', 9),
                bg=self.colors['light_bg'], fg=self.colors['text']).pack(anchor='w')
        
        self.baslangic_tarih_var = tk.StringVar()
        self.baslangic_entry = DateEntry(start_frame, 
                                       textvariable=self.baslangic_tarih_var,
                                       date_pattern='dd/mm/yyyy',
                                       width=12, 
                                       font=('Segoe UI', 9),
                                       background='white',
                                       foreground='black',
                                       borderwidth=1,
                                       locale='tr_TR')
        self.baslangic_entry.pack(pady=(2, 0))
        self.baslangic_entry.delete(0, tk.END)
        
        self.takvimi_turkcelestir(self.baslangic_entry)
        
        end_frame = tk.Frame(tarih_selection_frame, bg=self.colors['light_bg'])
        end_frame.pack(side='left', padx=(10, 0))
        
        tk.Label(end_frame, text="Bitiş:", font=('Segoe UI', 9),
                bg=self.colors['light_bg'], fg=self.colors['text']).pack(anchor='w')
        
        self.bitis_tarih_var = tk.StringVar()
        self.bitis_entry = DateEntry(end_frame, 
                                   textvariable=self.bitis_tarih_var,
                                   date_pattern='dd/mm/yyyy',
                                       width=12, 
                                       font=('Segoe UI', 9),
                                       background='white',
                                       foreground='black',
                                       borderwidth=1,
                                       locale='tr_TR')
        self.bitis_entry.pack(pady=(2, 0))
        self.bitis_entry.delete(0, tk.END)
        
        self.takvimi_turkcelestir(self.bitis_entry)
        
        button_frame = tk.Frame(filter_frame, bg=self.colors['light_bg'])
        button_frame.pack(side='left', padx=(15, 0))
        
        filter_btn = tk.Button(button_frame, text="Filtrele", font=('Segoe UI', 10, 'bold'),
                              bg=self.colors['accent'], fg='white', relief='raised',
                              command=self.filtrele, width=10)
        filter_btn.pack(side='left', padx=(0, 5))
        
        reset_btn = tk.Button(button_frame, text="Sıfırla", font=('Segoe UI', 10, 'bold'),
                             bg=self.colors['warning'], fg='white', relief='raised',
                             command=self.filtreleri_sifirla, width=10)
        reset_btn.pack(side='left')
        
        tree_frame = tk.Frame(left_frame, bg=self.colors['light_bg'])
        tree_frame.pack(fill='both', expand=True)
        

        # Tablo sütunlarını İSTENEN SIRAYA göre tanımla - TOPLAM MESAFE SÜTUNU EKLENDİ
        self.tree = ttk.Treeview(tree_frame, columns=(
            'Ilce', 'HatAdi', 'MusterkDirek', 'OgDirek', 'DirekSayisi', 
            'FotoSayisi', 'BulguSayisi', 'IlkTarih', 'SonTarih', 'ToplamGun', 'ToplamMesafe'
        ), show='headings')
        
        # Sütun başlıklarını İSTENEN SIRAYA göre ayarla - TOPLAM MESAFE SÜTUNU EKLENDİ
        self.tree.heading('Ilce', text='İlce', command=lambda: self.sutun_sirala('Ilce'))
        self.tree.heading('HatAdi', text='Hat Adı', command=lambda: self.sutun_sirala('HatAdi'))
        self.tree.heading('MusterkDirek', text='Müşterek Direk', command=lambda: self.sutun_sirala('MusterkDirek'))
        self.tree.heading('OgDirek', text='OG Direk', command=lambda: self.sutun_sirala('OgDirek'))
        self.tree.heading('DirekSayisi', text='Direk Sayısı', command=lambda: self.sutun_sirala('DirekSayisi'))
        self.tree.heading('FotoSayisi', text='Fotoğraf Sayısı', command=lambda: self.sutun_sirala('FotoSayisi'))
        self.tree.heading('BulguSayisi', text='Bulgu', command=lambda: self.sutun_sirala('BulguSayisi'))
        self.tree.heading('IlkTarih', text='İlk Çekim Tarihi', command=lambda: self.sutun_sirala('IlkTarih'))
        self.tree.heading('SonTarih', text='Son Çekim Tarihi', command=lambda: self.sutun_sirala('SonTarih'))
        self.tree.heading('ToplamGun', text='Toplam Gün', command=lambda: self.sutun_sirala('ToplamGun'))
        self.tree.heading('ToplamMesafe', text='Toplam Mesafe', command=lambda: self.sutun_sirala('ToplamMesafe'))
        
        # Sütun genişliklerini ayarla - TOPLAM MESAFE SÜTUNU EKLENDİ
        self.tree.column('Ilce', width=60, anchor='center')
        self.tree.column('HatAdi', width=210, anchor='w')
        self.tree.column('MusterkDirek', width=100, anchor='center')
        self.tree.column('OgDirek', width=80, anchor='center')
        self.tree.column('DirekSayisi', width=80, anchor='center')
        self.tree.column('FotoSayisi', width=100, anchor='center')
        self.tree.column('BulguSayisi', width=60, anchor='center')
        self.tree.column('IlkTarih', width=100, anchor='center')
        self.tree.column('SonTarih', width=100, anchor='center')
        self.tree.column('ToplamGun', width=80, anchor='center')
        self.tree.column('ToplamMesafe', width=100, anchor='center')
        
        # Çift tıklama olayını bağla
        self.tree.bind('<Double-1>', self.tabloya_cift_tikla)
        
        scrollbar = ttk.Scrollbar(tree_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # SAĞ TARAF: GRAFİKLER (EŞİT BOYUTTA)
        
        # GRAFİK 1: İlçelere göre direk dağılımı (3 BİRİM)
        self.graf1_frame = tk.LabelFrame(right_frame, text="🏙️ İlçelere Göre Direk Dağılımı", 
                                       font=('Segoe UI', 12, 'bold'),
                                       fg=self.colors['text'], bg=self.colors['light_bg'],
                                       padx=10, pady=10)
        self.graf1_frame.grid(row=0, column=0, sticky='nsew', pady=(0, 5))
        
        # GRAFİK 2: Aylara göre direk sayısı (3 BİRİM)
        self.graf2_frame = tk.LabelFrame(right_frame, text="📅 Aylara Göre Direk Sayısı", 
                                       font=('Segoe UI', 12, 'bold'),
                                       fg=self.colors['text'], bg=self.colors['light_bg'],
                                       padx=10, pady=10)
        self.graf2_frame.grid(row=1, column=0, sticky='nsew', pady=(5, 0))
        
        # BOŞ GRAFİKLERİ OLUŞTUR
        self.fig1 = plt.Figure(figsize=(6, 3.5), dpi=100)
        self.ax1 = self.fig1.add_subplot(111)
        self.canvas1 = FigureCanvasTkAgg(self.fig1, self.graf1_frame)
        self.canvas1.get_tk_widget().pack(fill='both', expand=True)
        
        self.fig2 = plt.Figure(figsize=(6, 3.5), dpi=100)
        self.ax2 = self.fig2.add_subplot(111)
        self.canvas2 = FigureCanvasTkAgg(self.fig2, self.graf2_frame)
        self.canvas2.get_tk_widget().pack(fill='both', expand=True)
        
    def oncelik_sekmesi_olustur(self):
        """Öncelik analizi sekmesini oluşturur - GÜNCELLENMİŞ"""
        # Başlık
        baslik_frame = tk.Frame(self.oncelik_frame, bg="#3498db", height=60)
        baslik_frame.pack(fill='x', padx=1, pady=1)
        baslik_frame.pack_propagate(False)
        
        tk.Label(baslik_frame, text="🚨 ÖNCELİK ANALİZİ - Hat Bazında Özet", 
                 font=('Segoe UI', 16, 'bold'), fg='white', bg="#3498db").pack(expand=True, pady=12)
        
        tk.Label(baslik_frame, text=f"Dosya: {os.path.basename(self.excel_dosyasi)}", 
                 font=('Segoe UI', 11), fg='white', bg="#3498db").pack(pady=(0, 12))
        
        # Filtreler
        filter_frame = tk.LabelFrame(self.oncelik_frame, text="🔍 Filtreler", 
                                     font=('Segoe UI', 11, 'bold'),
                                     fg='white', bg="#34495e", padx=15, pady=15)
        filter_frame.pack(fill='x', padx=20, pady=10)
        
        # Filtre grid
        filters_grid = tk.Frame(filter_frame, bg="#34495e")
        filters_grid.pack()
        
        # İlçe filtresi
        tk.Label(filters_grid, text="İlçe:", font=('Segoe UI', 10, 'bold'),
                 bg="#34495e", fg='white', width=10, anchor='w').grid(row=0, column=0, padx=8, pady=5, sticky='w')
        
        self.oncelik_ilce_var = tk.StringVar(value="TÜMÜ")
        self.oncelik_ilce_combo = ttk.Combobox(filters_grid, textvariable=self.oncelik_ilce_var, 
                                               state="readonly", width=24)
        self.oncelik_ilce_combo.grid(row=0, column=1, padx=8, pady=5)
        
        self.oncelik_ilce_combo.bind('<<ComboboxSelected>>', self.oncelik_ilce_degisti)
        
        # Hat filtresi
        tk.Label(filters_grid, text="Hat Adı:", font=('Segoe UI', 10, 'bold'),
                 bg="#34495e", fg='white', width=10, anchor='w').grid(row=0, column=2, padx=8, pady=5, sticky='w')
        
        self.oncelik_hat_var = tk.StringVar(value="TÜMÜ")
        self.oncelik_hat_combo = ttk.Combobox(filters_grid, textvariable=self.oncelik_hat_var,
                                              state="readonly", width=30)
        self.oncelik_hat_combo.grid(row=0, column=3, padx=8, pady=5)
        
        # Butonlar - "Verileri Yükle" butonu KALDIRILDI
        button_frame = tk.Frame(filter_frame, bg="#34495e")
        button_frame.pack(pady=10)
        
        tk.Button(button_frame, text="🔍 Filtrele", font=('Segoe UI', 10, 'bold'),
                  bg="#27ae60", fg='white', command=self.oncelik_filtrele,
                  width=14).pack(side='left', padx=8)
        
        tk.Button(button_frame, text="🔄 Sıfırla", font=('Segoe UI', 10, 'bold'),
                  bg="#e74c3c", fg='white', command=self.oncelik_filtreleri_sifirla,
                  width=14).pack(side='left', padx=8)
        
        tk.Button(button_frame, text="💾 Excel'e Aktar", font=('Segoe UI', 10, 'bold'),
                  bg="#9b59b6", fg='white', command=self.oncelik_detayli_excele_aktar,
                  width=20).pack(side='left', padx=8)
        
        # İstatistikler
        stats_frame = tk.LabelFrame(filter_frame, text="📊 İstatistikler", 
                                    font=('Segoe UI', 10, 'bold'),
                                    fg='white', bg="#34495e", padx=10, pady=10)
        stats_frame.pack(fill='x', pady=(10, 0))
        
        stats_inner = tk.Frame(stats_frame, bg="#34495e")
        stats_inner.pack()
        
        self.oncelik_toplam_var = tk.StringVar(value="Toplam: 0")
        self.oncelik_yapilan_var = tk.StringVar(value="Yapılan: 0")
        self.oncelik_yapilmayan_var = tk.StringVar(value="Yapılmayan: 0")
        
        tk.Label(stats_inner, textvariable=self.oncelik_toplam_var, 
                 font=('Segoe UI', 9, 'bold'), bg="#34495e", fg="#3498db").pack(side='left', padx=15)
        
        tk.Label(stats_inner, textvariable=self.oncelik_yapilan_var, 
                 font=('Segoe UI', 9, 'bold'), bg="#34495e", fg="#27ae60").pack(side='left', padx=15)
        
        tk.Label(stats_inner, textvariable=self.oncelik_yapilmayan_var, 
                 font=('Segoe UI', 9, 'bold'), bg="#34495e", fg="#e74c3c").pack(side='left', padx=15)
        
        # Tablo için ana frame
        table_main_frame = tk.LabelFrame(self.oncelik_frame, text="📋 Hat Bazında Öncelik Özeti", 
                                         font=('Segoe UI', 11, 'bold'),
                                         fg='white', bg="#34495e", padx=15, pady=15)
        table_main_frame.pack(fill='both', expand=True, padx=20, pady=(0, 10))
        
        # Tablo için canvas ve scrollbar
        canvas_frame = tk.Frame(table_main_frame, bg="#34495e")
        canvas_frame.pack(fill='both', expand=True)
        
        self.oncelik_canvas = tk.Canvas(canvas_frame, bg="#34495e", highlightthickness=0)
        scrollbar = ttk.Scrollbar(canvas_frame, orient="vertical", command=self.oncelik_canvas.yview)
        
        self.oncelik_tablo_frame = tk.Frame(self.oncelik_canvas, bg="#34495e")
        
        self.oncelik_canvas.create_window((0, 0), window=self.oncelik_tablo_frame, anchor="nw")
        self.oncelik_canvas.configure(yscrollcommand=scrollbar.set)
        
        self.oncelik_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        def configure_canvas(event):
            self.oncelik_canvas.configure(scrollregion=self.oncelik_canvas.bbox("all"))
        
        self.oncelik_tablo_frame.bind("<Configure>", configure_canvas)
        
        # İlk başta tablo başlıklarını oluştur
        self.oncelik_tablo_basliklari_olustur()
        
        # Yükleme mesajı göster
        self.oncelik_yukleniyor_label = tk.Label(self.oncelik_tablo_frame, 
                                                text="Öncelik verileri yükleniyor...\nLütfen bekleyin.",
                                                font=('Segoe UI', 12),
                                                bg="#34495e", fg="white")
        self.oncelik_yukleniyor_label.pack(expand=True, pady=50)



    def oncelik_tablo_basliklari_olustur(self):
        """Öncelik tablosunun başlıklarını oluşturur"""
        # Başlık satırı
        header_frame = tk.Frame(self.oncelik_tablo_frame, bg="#2c3e50")
        header_frame.pack(fill='x', pady=(0, 2))
        
        # Sütunlar
        column_widths = {
            'ilce': 100,
            'hat_adi': 200,
            'yapilan_bekleyebilir': 90,
            'yapilan_normal': 80,
            'yapilan_acil': 70,
            'yapilan_cok_acil': 90,
            'yapilmayan_bekleyebilir': 90,
            'yapilmayan_normal': 80,
            'yapilmayan_acil': 70,
            'yapilmayan_cok_acil': 90,
            'toplam_yapilan': 80,
            'toplam_yapilmayan': 90
        }
        
        # İlçe
        tk.Label(header_frame, text="İlçe", font=('Segoe UI', 10, 'bold'),
                 bg="#2c3e50", fg='white', width=20, anchor='center').pack(side='left', padx=2)
        
        # Hat Adı
        tk.Label(header_frame, text="Hat Adı", font=('Segoe UI', 10, 'bold'),
                 bg="#2c3e50", fg='white', width=39, anchor='w').pack(side='left', padx=28)
        
        # Yapılanlar frame
        yapilan_frame = tk.LabelFrame(header_frame, text="✅ BAKIM YAPILANLAR", 
                                      font=('Segoe UI', 9, 'bold'),
                                      fg='white', bg="#2c3e50", bd=1, relief='solid')
        yapilan_frame.pack(side='left', padx=5)
        
        yapilan_inner = tk.Frame(yapilan_frame, bg="#2c3e50")
        yapilan_inner.pack(pady=2)
        
        tk.Label(yapilan_inner, text="Bekleyebilir", font=('Segoe UI', 8, 'bold'),
                 bg="#27ae60", fg='white', width=14, anchor='center').pack(side='left', padx=1)
        tk.Label(yapilan_inner, text="Normal", font=('Segoe UI', 8, 'bold'),
                 bg="#3498db", fg='white', width=13, anchor='center').pack(side='left', padx=1)
        tk.Label(yapilan_inner, text="Acil", font=('Segoe UI', 8, 'bold'),
                 bg="#f39c12", fg='white', width=13, anchor='center').pack(side='left', padx=1)
        tk.Label(yapilan_inner, text="Çok Acil", font=('Segoe UI', 8, 'bold'),
                 bg="#e74c3c", fg='white', width=13, anchor='center').pack(side='left', padx=1)
        
        # Ayırıcı
        tk.Label(header_frame, text="|", font=('Segoe UI', 10, 'bold'),
                 bg="#2c3e50", fg='white').pack(side='left', padx=10)
        
        # Yapılmayanlar frame
        yapilmayan_frame = tk.LabelFrame(header_frame, text="❌ BAKIM YAPILMAYANLAR", 
                                         font=('Segoe UI', 9, 'bold'),
                                         fg='white', bg="#2c3e50", bd=1, relief='solid')
        yapilmayan_frame.pack(side='left', padx=5)
        
        yapilmayan_inner = tk.Frame(yapilmayan_frame, bg="#2c3e50")
        yapilmayan_inner.pack(pady=2)
        
        tk.Label(yapilmayan_inner, text="Bekleyebilir", font=('Segoe UI', 8, 'bold'),
                 bg="#27ae60", fg='white', width=14, anchor='center').pack(side='left', padx=1)
        tk.Label(yapilmayan_inner, text="Normal", font=('Segoe UI', 8, 'bold'),
                 bg="#3498db", fg='white', width=13, anchor='center').pack(side='left', padx=1)
        tk.Label(yapilmayan_inner, text="Acil", font=('Segoe UI', 8, 'bold'),
                 bg="#f39c12", fg='white', width=13, anchor='center').pack(side='left', padx=1)
        tk.Label(yapilmayan_inner, text="Çok Acil", font=('Segoe UI', 8, 'bold'),
                 bg="#e74c3c", fg='white', width=13, anchor='center').pack(side='left', padx=1)
        
        # Ayırıcı
        tk.Label(header_frame, text="|", font=('Segoe UI', 10, 'bold'),
                 bg="#2c3e50", fg='white').pack(side='left', padx=10)
        
        # Toplamlar
        toplam_frame = tk.LabelFrame(header_frame, text="📊 TOPLAMLAR", 
                                     font=('Segoe UI', 9, 'bold'),
                                     fg='white', bg="#2c3e50", bd=1, relief='solid')
        toplam_frame.pack(side='left', padx=5)
        
        toplam_inner = tk.Frame(toplam_frame, bg="#2c3e50")
        toplam_inner.pack(pady=2)
        
        tk.Label(toplam_inner, text="Yapılan", font=('Segoe UI', 8, 'bold'),
                 bg="#27ae60", fg='white', width=14, anchor='center').pack(side='left', padx=2)
        tk.Label(toplam_inner, text="Yapılmayan", font=('Segoe UI', 8, 'bold'),
                 bg="#e74c3c", fg='white', width=14, anchor='center').pack(side='left', padx=6)



    def oncelik_tabloyu_doldur(self):
        """Öncelik tablosunu doldurur"""
        # Önceki satırları temizle
        for widget in self.oncelik_tablo_frame.winfo_children():
            if widget != self.oncelik_tablo_frame.winfo_children()[0]:  # Başlık hariç
                widget.destroy()
        
        if not hasattr(self, 'oncelik_veriler') or not self.oncelik_veriler:
            return
        
        # Filtre değerlerini al
        secili_ilce = self.oncelik_ilce_var.get()
        secili_hat = self.oncelik_hat_var.get()
        
        # İstatistikleri sıfırla
        toplam_yapilan = 0
        toplam_yapilmayan = 0
        toplam_genel = 0
        
        # Verileri filtrele ve satırları oluştur
        for i, veri in enumerate(self.oncelik_veriler):
            # Filtrele
            if secili_ilce != 'TÜMÜ' and veri['ilce'] != secili_ilce:
                continue
            if secili_hat != 'TÜMÜ' and veri['hat_adi'] != secili_hat:
                continue
            
            # Satır frame'ini oluştur (zebra deseni)
            row_frame = tk.Frame(self.oncelik_tablo_frame, 
                                bg="#e8f4f8" if i % 2 == 0 else "#ffffff")
            row_frame.pack(fill='x', pady=1)
            
            # İlçe
            tk.Label(row_frame, text=veri['ilce'], font=('Segoe UI', 9),
                     bg=row_frame['bg'], fg='#000000', width=20,
                     anchor='center', bd=1, relief='solid').pack(side='left', padx=2, fill='y')
            
            # Hat Adı
            tk.Label(row_frame, text=veri['hat_adi'], font=('Segoe UI', 9),
                     bg=row_frame['bg'], fg='#000000', width=55,
                     anchor='w', bd=1, relief='solid').pack(side='left', padx=2, fill='y')
            
            # Yapılanlar
            yapilan_frame = tk.Frame(row_frame, bg=row_frame['bg'])
            yapilan_frame.pack(side='left', padx=5)
            
            # Bekleyebilir (Yapılan)
            tk.Label(yapilan_frame, text=str(veri['yapilan_bekleyebilir']), 
                     font=('Segoe UI', 9, 'bold'), bg="#bbf7d0", fg='#14532d', 
                     width=15, anchor='center', bd=1, relief='solid').pack(side='left', padx=1)
            
            # Normal (Yapılan)
            tk.Label(yapilan_frame, text=str(veri['yapilan_normal']), 
                     font=('Segoe UI', 9, 'bold'), bg="#86efac", fg='#14532d', 
                     width=13, anchor='center', bd=1, relief='solid').pack(side='left', padx=1)
            
            # Acil (Yapılan)
            tk.Label(yapilan_frame, text=str(veri['yapilan_acil']), 
                     font=('Segoe UI', 9, 'bold'), bg="#22c55e", fg='white', 
                     width=13, anchor='center', bd=1, relief='solid').pack(side='left', padx=1)
            
            # Çok Acil (Yapılan)
            tk.Label(yapilan_frame, text=str(veri['yapilan_cok_acil']), 
                     font=('Segoe UI', 9, 'bold'), bg="#15803d", fg='white', 
                     width=14, anchor='center', bd=1, relief='solid').pack(side='left', padx=1)
            
            # Ayırıcı
            tk.Label(row_frame, text="|", font=('Segoe UI', 10),
                     bg=row_frame['bg'], fg='#000000', width=2).pack(side='left', padx=5)
            
            # Yapılmayanlar
            yapilmayan_frame = tk.Frame(row_frame, bg=row_frame['bg'])
            yapilmayan_frame.pack(side='left', padx=5)
            
            # Bekleyebilir (Yapılmayan)
            tk.Label(yapilmayan_frame, text=str(veri['yapilmayan_bekleyebilir']), 
                     font=('Segoe UI', 9, 'bold'), bg="#fca5a5", fg='#7f1d1d', 
                     width=15, anchor='center', bd=1, relief='solid').pack(side='left', padx=1)
            
            # Normal (Yapılmayan)
            tk.Label(yapilmayan_frame, text=str(veri['yapilmayan_normal']), 
                     font=('Segoe UI', 9, 'bold'), bg="#f87171", fg='#7f1d1d', 
                     width=13, anchor='center', bd=1, relief='solid').pack(side='left', padx=1)
            
            # Acil (Yapılmayan)
            tk.Label(yapilmayan_frame, text=str(veri['yapilmayan_acil']), 
                     font=('Segoe UI', 9, 'bold'), bg="#ef4444", fg='white', 
                     width=13, anchor='center', bd=1, relief='solid').pack(side='left', padx=1)
            
            # Çok Acil (Yapılmayan)
            tk.Label(yapilmayan_frame, text=str(veri['yapilmayan_cok_acil']), 
                     font=('Segoe UI', 9, 'bold'), bg="#dc2626", fg='white', 
                     width=14, anchor='center', bd=1, relief='solid').pack(side='left', padx=1)
            
            # Ayırıcı
            tk.Label(row_frame, text="|", font=('Segoe UI', 10),
                     bg=row_frame['bg'], fg='#000000', width=2).pack(side='left', padx=5)
            
            # Toplamlar
            toplam_frame = tk.Frame(row_frame, bg=row_frame['bg'])
            toplam_frame.pack(side='left', padx=5)
            
            # Toplam Yapılan
            tk.Label(toplam_frame, text=str(veri['toplam_yapilan']), 
                     font=('Segoe UI', 9, 'bold'), bg="#27ae60", fg='white', 
                     width=14, anchor='center', bd=1, relief='solid').pack(side='left', padx=1)
            
            # Toplam Yapılmayan
            tk.Label(toplam_frame, text=str(veri['toplam_yapilmayan']), 
                     font=('Segoe UI', 9, 'bold'), bg="#e74c3c", fg='white', 
                     width=15, anchor='center', bd=1, relief='solid').pack(side='left', padx=1)
            
            # İstatistikleri güncelle
            toplam_yapilan += veri['toplam_yapilan']
            toplam_yapilmayan += veri['toplam_yapilmayan']
            toplam_genel += veri['toplam_genel']
        
        # İstatistikleri güncelle
        self.oncelik_toplam_var.set(f"Toplam: {toplam_genel}")
        self.oncelik_yapilan_var.set(f"Yapılan: {toplam_yapilan}")
        self.oncelik_yapilmayan_var.set(f"Yapılmayan: {toplam_yapilmayan}")
        
        # Tabloyu yeniden boyutlandır
        self.oncelik_canvas.configure(scrollregion=self.oncelik_canvas.bbox("all"))




    def oncelik_verileri_yukle(self):
        """Rapor Excel'inden öncelik verilerini yükler"""
        try:
            self.log_mesaj_ekle("📊 Öncelik verileri yükleniyor...")
            
            if not os.path.exists(self.excel_dosyasi):
                messagebox.showerror("Hata", f"Excel dosyası bulunamadı: {self.excel_dosyasi}")
                return
            
            # Fotoğraf Klasörleri sayfasını oku
            df_fotograf = pd.read_excel(self.excel_dosyasi, sheet_name='Fotoğraf Klasörleri')
            
            # Gerekli sütunları kontrol et
            gerekli_sutunlar = ['Aob', 'Hat adı', 'Direk No', 'Öncelik', 'Yapıldı mı?', 'Tespit Notu']
            for sutun in gerekli_sutunlar:
                if sutun not in df_fotograf.columns:
                    messagebox.showerror("Hata", f"Excel'de '{sutun}' sütunu bulunamadı!")
                    return
            
            # Öncelik ve Yapıldı mı? sütunlarını temizle
            df_fotograf['Öncelik'] = df_fotograf['Öncelik'].fillna('').astype(str).str.strip()
            df_fotograf['Yapıldı mı?'] = df_fotograf['Yapıldı mı?'].fillna('').astype(str).str.strip()
            
            # Hat bazında öncelik istatistiklerini hesapla
            hat_oncelik_istatistikleri = {}
            
            for index, row in df_fotograf.iterrows():
                try:
                    ilce = str(row['Aob']).strip()
                    hat_adi = str(row['Hat adı']).strip()
                    oncelik = row['Öncelik'].lower()
                    yapildi_mi = str(row['Yapıldı mı?']).lower()
                    
                    # Direk numarasını temizle
                    direk_no = str(row['Direk No']).strip()
                    
                    # Tespit notunu kontrol et
                    tespit_notu = str(row['Tespit Notu']).strip() if pd.notna(row['Tespit Notu']) else ""
                    has_tespit = tespit_notu != ""
                    
                    if not ilce or not hat_adi or not direk_no:
                        continue
                    
                    # YAPILDI MI? FİLTRESİNİ UYGULA - HARİTA FİLTRESİNE GÖRE
                    # DÜZELTME: self.yapildi_var'ı kontrol et, eğer yoksa 'TÜMÜ' kullan
                    if hasattr(self, 'yapildi_var'):
                        secili_yapildi = self.yapildi_var.get()
                    else:
                        secili_yapildi = 'TÜMÜ'
                    
                    # Sadece filtre uygulanmışsa kontrol et
                    if secili_yapildi != 'TÜMÜ':
                        if secili_yapildi == 'YAPILDI':
                            # Yapıldı olanları filtrele
                            yapildi_durum = any(keyword in yapildi_mi for keyword in ['evet', 'yapıldı', 'x', '✓', '✅'])
                            if not yapildi_durum:
                                continue
                        elif secili_yapildi == 'YAPILMADI':
                            # Yapılmadı olanları filtrele
                            yapildi_durum = any(keyword in yapildi_mi for keyword in ['evet', 'yapıldı', 'x', '✓', '✅'])
                            if yapildi_durum:
                                continue
                    
                    hat_anahtari = (ilce, hat_adi)
                    
                    if hat_anahtari not in hat_oncelik_istatistikleri:
                        hat_oncelik_istatistikleri[hat_anahtari] = {
                            'ilce': ilce,
                            'hat_adi': hat_adi,
                            'yapilan_bekleyebilir': 0,
                            'yapilan_normal': 0,
                            'yapilan_acil': 0,
                            'yapilan_cok_acil': 0,
                            'yapilmayan_bekleyebilir': 0,
                            'yapilmayan_normal': 0,
                            'yapilmayan_acil': 0,
                            'yapilmayan_cok_acil': 0,
                            'toplam_yapilan': 0,
                            'toplam_yapilmayan': 0,
                            'toplam_genel': 0,
                            'tespit_sayisi': 0
                        }
                    
                    # Yapıldı mı? kontrolü
                    yapildi = False
                    if any(keyword in yapildi_mi for keyword in ['evet', 'yapıldı', 'x', '✓', '✅']):
                        yapildi = True
                    
                    # Öncelik kategorisine göre say
                    if 'bekleyebilir' in oncelik:
                        if yapildi:
                            hat_oncelik_istatistikleri[hat_anahtari]['yapilan_bekleyebilir'] += 1
                            hat_oncelik_istatistikleri[hat_anahtari]['toplam_yapilan'] += 1
                        else:
                            hat_oncelik_istatistikleri[hat_anahtari]['yapilmayan_bekleyebilir'] += 1
                            hat_oncelik_istatistikleri[hat_anahtari]['toplam_yapilmayan'] += 1
                        
                    elif 'normal' in oncelik:
                        if yapildi:
                            hat_oncelik_istatistikleri[hat_anahtari]['yapilan_normal'] += 1
                            hat_oncelik_istatistikleri[hat_anahtari]['toplam_yapilan'] += 1
                        else:
                            hat_oncelik_istatistikleri[hat_anahtari]['yapilmayan_normal'] += 1
                            hat_oncelik_istatistikleri[hat_anahtari]['toplam_yapilmayan'] += 1
                        
                    elif 'acil' in oncelik and 'çok acil' not in oncelik and 'cok acil' not in oncelik:
                        if yapildi:
                            hat_oncelik_istatistikleri[hat_anahtari]['yapilan_acil'] += 1
                            hat_oncelik_istatistikleri[hat_anahtari]['toplam_yapilan'] += 1
                        else:
                            hat_oncelik_istatistikleri[hat_anahtari]['yapilmayan_acil'] += 1
                            hat_oncelik_istatistikleri[hat_anahtari]['toplam_yapilmayan'] += 1
                        
                    elif 'çok acil' in oncelik or 'cok acil' in oncelik:
                        if yapildi:
                            hat_oncelik_istatistikleri[hat_anahtari]['yapilan_cok_acil'] += 1
                            hat_oncelik_istatistikleri[hat_anahtari]['toplam_yapilan'] += 1
                        else:
                            hat_oncelik_istatistikleri[hat_anahtari]['yapilmayan_cok_acil'] += 1
                            hat_oncelik_istatistikleri[hat_anahtari]['toplam_yapilmayan'] += 1
                    
                    # Tespit sayısını güncelle
                    if has_tespit:
                        hat_oncelik_istatistikleri[hat_anahtari]['tespit_sayisi'] += 1
                    
                    # Toplam genel sayıyı güncelle
                    hat_oncelik_istatistikleri[hat_anahtari]['toplam_genel'] += 1
                    
                except Exception as e:
                    continue
            
            # Verileri listeye çevir
            self.oncelik_veriler = list(hat_oncelik_istatistikleri.values())
            self.oncelik_veriler.sort(key=lambda x: (x['ilce'], x['hat_adi']))
            
            # İstatistikleri hesapla
            toplam = sum(v['toplam_genel'] for v in self.oncelik_veriler)
            yapilan = sum(v['toplam_yapilan'] for v in self.oncelik_veriler)
            yapilmayan = sum(v['toplam_yapilmayan'] for v in self.oncelik_veriler)
            
            self.oncelik_toplam_var.set(f"Toplam: {toplam}")
            self.oncelik_yapilan_var.set(f"Yapılan: {yapilan}")
            self.oncelik_yapilmayan_var.set(f"Yapılmayan: {yapilmayan}")
            
            # Combobox'ları doldur
            ilce_list = ['TÜMÜ'] + sorted(set(v['ilce'] for v in self.oncelik_veriler))
            self.oncelik_ilce_combo['values'] = ilce_list
            
            hat_list = ['TÜMÜ'] + sorted(set(v['hat_adi'] for v in self.oncelik_veriler))
            self.oncelik_hat_combo['values'] = hat_list
            
            # Tabloyu doldur
            self.oncelik_tabloyu_doldur()
            
            # Yükleme mesajını kaldır (güvenli şekilde)
            try:
                if hasattr(self, 'oncelik_yukleniyor_label') and self.oncelik_yukleniyor_label.winfo_exists():
                    self.oncelik_yukleniyor_label.destroy()
            except:
                pass
            
            self.log_mesaj_ekle(f"✅ Öncelik verileri yüklendi: {len(self.oncelik_veriler)} hat, {toplam} kayıt")
            
            # DÜZELTME: Filtre mesajını doğru şekilde göster
            if hasattr(self, 'yapildi_var'):
                yapildi_filtresi = self.yapildi_var.get()
                if yapildi_filtresi != 'TÜMÜ':
                    self.log_mesaj_ekle(f"🔍 Yapıldı mı? filtresi: {yapildi_filtresi}")
            
        except Exception as e:
            messagebox.showerror("Hata", f"Öncelik verileri yüklenirken hata: {str(e)}")
            import traceback
            print(f"Hata detayı: {traceback.format_exc()}")
            
            # Hata durumunda mesaj göster (güvenli şekilde)
            try:
                if hasattr(self, 'oncelik_yukleniyor_label') and self.oncelik_yukleniyor_label.winfo_exists():
                    self.oncelik_yukleniyor_label.config(text=f"Hata oluştu:\n{str(e)[:50]}...", 
                                                       fg="#e74c3c")
            except:
                pass

        
    def oncelik_filtrele(self, ilk_yukleme=False):
        """Öncelik analizinde filtreleme yapar - Exception handling eklenmiş"""
        try:
            print(f"\n🔍 ÖNCELİK FİLTRELEME BAŞLATILIYOR...")
            
            # 1. TABLOYU TAMAMEN TEMİZLE
            for item in self.oncelik_tree.get_children():
                self.oncelik_tree.delete(item)
            
            # 2. oncelik_veriler değişkenini kontrol et
            if not hasattr(self, 'oncelik_veriler') or not self.oncelik_veriler:
                print("❌ HATA: oncelik_veriler yüklenmemiş!")
                self.oncelik_tree.insert('', 'end', values=(
                    "HATA", "Veriler yüklenemedi", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"
                ))
                return
            
            print(f"   Toplam {len(self.oncelik_veriler)} hat verisi var")
            
            # 3. Filtre değerlerini al
            secili_ilce = self.oncelik_ilce_var.get()
            secili_hat = self.oncelik_hat_var.get()
            
            filtered_count = 0
            
            # 4. VERİLERİ FİLTRELE VE TABLOYA EKLE
            for i, veri in enumerate(self.oncelik_veriler):
                try:
                    # İlçe filtresi
                    if secili_ilce != 'TÜMÜ' and veri['ilce'] != secili_ilce:
                        continue
                    
                    # Hat filtresi
                    if secili_hat != 'TÜMÜ' and veri['hat_adi'] != secili_hat:
                        continue
                    
                    # TABLOYA EKLE - YENİ FORMAT (12 sütun)
                    item = self.oncelik_tree.insert('', 'end', values=(
                        veri['ilce'],                               # İlçe
                        veri['hat_adi'],                           # Hat Adı
                        # BAKIM YAPILANLAR
                        veri['yapilan_bekleyebilir'],              # Yapılan - Bekleyebilir
                        veri['yapilan_normal'],                    # Yapılan - Normal
                        veri['yapilan_acil'],                      # Yapılan - Acil
                        veri['yapilan_cok_acil'],                  # Yapılan - Çok Acil
                        # BAKIM YAPILMAYANLAR
                        veri['yapilmayan_bekleyebilir'],           # Yapılmayan - Bekleyebilir
                        veri['yapilmayan_normal'],                 # Yapılmayan - Normal
                        veri['yapilmayan_acil'],                   # Yapılmayan - Acil
                        veri['yapilmayan_cok_acil'],               # Yapılmayan - Çok Acil
                        # TOPLAMLAR
                        veri['toplam_bakim_yapilan'],              # Toplam Bakım Yapılan
                        veri['toplam_genel']                       # Toplam Genel
                    ))
                    
                    # Zebra çizgi efekti
                    if filtered_count % 2 == 0:
                        self.oncelik_tree.item(item, tags=('row_even',))
                    else:
                        self.oncelik_tree.item(item, tags=('row_odd',))
                    
                    filtered_count += 1
                    
                except Exception as e:
                    print(f"⚠️ Hat {i} eklenirken hata: {e}")
                    continue
            
            print(f"✅ {filtered_count} hat filtrelenip tabloya eklendi")
            
            # 5. EĞER FİLTRE SONUCU BOŞSA MESAJ GÖSTER
            if filtered_count == 0:
                print("ℹ️ Filtre sonucu boş")
                self.oncelik_tree.insert('', 'end', values=(
                    "BİLGİ", "Filtreye uygun veri bulunamadı", "-", "-", "-", "-", "-", "-", "-", "-", "-", "-"
                ))
            
        except Exception as e:
            print(f"❌ Filtreleme hatası: {str(e)}")
            import traceback
            print(f"🔍 Hata detayı:\n{traceback.format_exc()}")
            
            # Hata durumunda tabloya mesaj ekle
            try:
                for item in self.oncelik_tree.get_children():
                    self.oncelik_tree.delete(item)
                
                self.oncelik_tree.insert('', 'end', values=(
                    "HATA", f"Filtreleme hatası: {str(e)[:30]}", "0", "0", "0", "0", "0", "0", "0", "0", "0", "0"
                ))
            except:
                pass

            
    def oncelik_filtrele(self, event=None):
        """Öncelik tablosunu filtreler"""
        self.oncelik_tabloyu_doldur()

    def oncelik_ilce_degisti(self, event=None):
        """İlçe değiştiğinde hat listesini günceller"""
        secili_ilce = self.oncelik_ilce_var.get()
        
        if not hasattr(self, 'oncelik_veriler') or not self.oncelik_veriler:
            return
        
        if secili_ilce == "TÜMÜ":
            hat_list = ['TÜMÜ'] + sorted(set(v['hat_adi'] for v in self.oncelik_veriler))
        else:
            hat_list = ['TÜMÜ'] + sorted(set(v['hat_adi'] for v in self.oncelik_veriler if v['ilce'] == secili_ilce))
        
        self.oncelik_hat_combo['values'] = hat_list
        self.oncelik_hat_var.set("TÜMÜ")

    def oncelik_filtreleri_sifirla(self):
        """Öncelik filtrelerini sıfırlar"""
        self.oncelik_ilce_var.set("TÜMÜ")
        self.oncelik_hat_var.set("TÜMÜ")
        
        # Tüm ilçeler için hat listesini yenile
        if hasattr(self, 'oncelik_veriler') and self.oncelik_veriler:
            hat_list = ['TÜMÜ'] + sorted(set(v['hat_adi'] for v in self.oncelik_veriler))
            self.oncelik_hat_combo['values'] = hat_list
        
        self.oncelik_tabloyu_doldur()

    def oncelik_detayli_excele_aktar(self):
        """Filtrelenmiş öncelik detaylarını Excel'e aktarır - SADECE DETAYLAR SAYFASI"""
        try:
            
            # Orijinal Excel dosyasını oku
            df_fotograf = pd.read_excel(self.excel_dosyasi, sheet_name='Fotoğraf Klasörleri')
            
            # Sadece öncelik bilgisi olan satırları filtrele (TEMEL FİLTRE)
            df_oncelik = df_fotograf[
                df_fotograf['Öncelik'].notna() & 
                (df_fotograf['Öncelik'] != '') & 
                (df_fotograf['Öncelik'] != 'nan')
            ].copy()
            
            if df_oncelik.empty:
                messagebox.showinfo("Bilgi", "Öncelik bilgisi olan veri bulunamadı.")
                return
            
            # TABLODAKİ FİLTRELERİ UYGULA
            secili_ilce = self.oncelik_ilce_var.get()
            secili_hat = self.oncelik_hat_var.get()
            
            # İlçe filtresi
            if secili_ilce != "TÜMÜ":
                df_oncelik = df_oncelik[df_oncelik['Aob'] == secili_ilce]
            
            # Hat filtresi
            if secili_hat != "TÜMÜ":
                df_oncelik = df_oncelik[df_oncelik['Hat adı'] == secili_hat]
            
            if df_oncelik.empty:
                messagebox.showinfo("Bilgi", 
                                  f"Seçilen filtreler için veri bulunamadı:\n"
                                  f"• İlçe: {secili_ilce}\n"
                                  f"• Hat: {secili_hat}")
                return
            
            # Kaydetme yeri seç
            kaydetme_yolu = filedialog.asksaveasfilename(
                title="Filtrelenmiş Öncelik Listesini Kaydet",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                initialfile=f"oncelik_filtreli_{dt.now().strftime('%Y%m%d_%H%M')}.xlsx"
            )
            
            if not kaydetme_yolu:
                return
            
            # İstenen sütunları seç - LAT/LON EKLENDİ
            detayli_sutunlar = ['Aob', 'Hat adı', 'Direk No', 'Öncelik', 'Tespit Notu', 
                               'Fotoğraf Sayısı', 'En Yeni Fotoğraf Tarihi', 'LAT', 'LON']
            
            # Mevcut sütunları kontrol et
            mevcut_sutunlar = []
            for col in detayli_sutunlar:
                if col in df_oncelik.columns:
                    mevcut_sutunlar.append(col)
                else:
                    print(f"⚠️ {col} sütunu bulunamadı")
            
            if not mevcut_sutunlar:
                messagebox.showinfo("Bilgi", "Gerekli sütunlar bulunamadı.")
                return
            
            # Filtrelenmiş DataFrame
            df_detay = df_oncelik[mevcut_sutunlar].copy()
            
            # LAT ve LON sütunlarını temizle (NaN değerleri boş string yap)
            if 'LAT' in df_detay.columns:
                df_detay['LAT'] = df_detay['LAT'].apply(lambda x: '' if pd.isna(x) else x)
            if 'LON' in df_detay.columns:
                df_detay['LON'] = df_detay['LON'].apply(lambda x: '' if pd.isna(x) else x)
            
            # Excel'e kaydet - SADECE 1 SAYFA
            with pd.ExcelWriter(kaydetme_yolu, engine='openpyxl') as writer:
                # SADECE DETAYLI LİSTE SAYFASI
                df_detay.to_excel(writer, sheet_name='Filtrelenmiş Detaylar', index=False)
                
                # Formatlama
                workbook = writer.book
                
                # SAYFA: Filtrelenmiş Detaylar formatlama
                ws_detay = writer.sheets['Filtrelenmiş Detaylar']
                
                # Başlık stilini uygula (ORTAYA HİZALI)
                header_fill = PatternFill(start_color="27ae60", end_color="27ae60", fill_type="solid")
                header_font = Font(bold=True, color="FFFFFF")
                
                for col in range(1, len(df_detay.columns) + 1):
                    cell = ws_detay.cell(row=1, column=col)
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal='center', vertical='center')  # BAŞLIK ORTADA
                
                # Öncelik renk kodlama ve zebra deseni
                for row in range(2, ws_detay.max_row + 1):
                    # Zebra deseni
                    if row % 2 == 0:
                        row_fill = PatternFill(start_color="f2f9f2", end_color="f2f9f2", fill_type="solid")
                    else:
                        row_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
                    
                    # Tüm satıra zebra deseni ve SOLA HİZALAMA uygula
                    for col in range(1, len(df_detay.columns) + 1):
                        cell = ws_detay.cell(row=row, column=col)
                        cell.fill = row_fill
                        cell.alignment = Alignment(horizontal='left', vertical='center')  # VERİLER SOLA HİZALI
                    
                    # Öncelik renk kodlama (D sütunu - 4. sütun)
                    oncelik_cell = ws_detay.cell(row=row, column=4)  # Öncelik sütunu
                    oncelik = str(oncelik_cell.value).lower() if oncelik_cell.value else ""
                    
                    if 'bekleyebilir' in oncelik:
                        oncelik_cell.font = Font(color="27ae60", bold=True)
                    elif 'normal' in oncelik:
                        oncelik_cell.font = Font(color="3498db", bold=True)
                    elif 'acil' in oncelik:
                        oncelik_cell.font = Font(color="f39c12", bold=True)
                    elif 'çok acil' in oncelik or 'cok acil' in oncelik:
                        oncelik_cell.font = Font(color="e74c3c", bold=True)
                    
                    # LAT ve LON sütunları için sayı formatı
                    if 'LAT' in mevcut_sutunlar:
                        lat_index = mevcut_sutunlar.index('LAT') + 1
                        lat_cell = ws_detay.cell(row=row, column=lat_index)
                        if lat_cell.value and lat_cell.value != '':
                            try:
                                lat_value = float(lat_cell.value)
                                lat_cell.value = lat_value
                                lat_cell.number_format = '0.000000'
                            except:
                                pass
                    
                    if 'LON' in mevcut_sutunlar:
                        lon_index = mevcut_sutunlar.index('LON') + 1
                        lon_cell = ws_detay.cell(row=row, column=lon_index)
                        if lon_cell.value and lon_cell.value != '':
                            try:
                                lon_value = float(lon_cell.value)
                                lon_cell.value = lon_value
                                lon_cell.number_format = '0.000000'
                            except:
                                pass
                    
                    # Koordinat bilgisi olmayan satırları işaretle
                    if 'LAT' in mevcut_sutunlar and 'LON' in mevcut_sutunlar:
                        lat_index = mevcut_sutunlar.index('LAT') + 1
                        lon_index = mevcut_sutunlar.index('LON') + 1
                        
                        lat_cell = ws_detay.cell(row=row, column=lat_index)
                        lon_cell = ws_detay.cell(row=row, column=lon_index)
                        
                        lat_value = lat_cell.value
                        lon_value = lon_cell.value
                        
                        # Koordinat yoksa sarı renkle işaretle
                        if not lat_value or not lon_value or lat_value == '' or lon_value == '':
                            lat_cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                            lon_cell.fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
                            lat_cell.value = "KOORDİNAT YOK" if not lat_value or lat_value == '' else lat_value
                            lon_cell.value = "KOORDİNAT YOK" if not lon_value or lon_value == '' else lon_value
                
                # Sütun genişliklerini otomatik ayarla
                for column in ws_detay.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    ws_detay.column_dimensions[column_letter].width = adjusted_width
                
            # Başarı mesajı
            filter_msg = ""
            if secili_ilce != "TÜMÜ":
                filter_msg += f"• İlçe: {secili_ilce}\n"
            if secili_hat != "TÜMÜ":
                filter_msg += f"• Hat: {secili_hat}\n"
            
            if not filter_msg:
                filter_msg = "• Filtre: TÜMÜ (filtresiz)\n"
            
            messagebox.showinfo("Başarılı", 
                              f"✅ FİLTRELENMİŞ VERİLER EXCEL'E AKTARILDI\n\n"
                              f"📁 Dosya: {os.path.basename(kaydetme_yolu)}\n\n"
                              f"🔍 UYGULANAN FİLTRELER:\n{filter_msg}\n"
                              f"📊 AKTARILAN VERİ:\n"
                              f"• {len(df_detay)} kayıt\n\n"
                              f"📌 ÖZELLİKLER:\n"
                              f"• Zebra desenli satırlar\n"
                              f"• Öncelik renk kodlaması\n"
                              f"• Koordinat formatı (6 ondalık)\n"
                              f"• Koordinat eksikleri sarı renkte\n"
                              f"• Başlık orta, veriler sola hizalı\n"  # Bu satır eklendi
                              f"• Otomatik sütun genişliği")
            
        except Exception as e:
            messagebox.showerror("Hata", f"Excel kaydedilirken hata: {str(e)}")
            print(f"Excel aktarım hatası: {e}")
            import traceback
            traceback.print_exc()


    def _oncelik_seviyesi_hesapla(self, toplam):
        """Toplam değere göre öncelik seviyesini hesaplar"""
        if toplam <= 5:
            return 'BEKLEYEBİLİR'
        elif toplam <= 15:
            return 'NORMAL'
        elif toplam <= 30:
            return 'ACİL'
        else:
            return 'ÇOK ACİL'

    def show_error_in_table(self, error_message):
            """Tabloya hata mesajı ekler"""
            try:
                # Tabloyu temizle
                for item in self.oncelik_tree.get_children():
                    self.oncelik_tree.delete(item)
                
                # Hata mesajını ekle
                self.oncelik_tree.insert('', 'end', values=(
                    "HATA", error_message, "0", "0", "0", "0", "0"
                ))
            except Exception as e:
                print(f"❌ Tabloya hata eklenirken hata: {e}")



                


    def harita_sekmesi_olustur(self):
        """Harita analizi sekmesini oluşturur - SADECE İLK KEZ TIKLANDIĞINDA"""
        try:
            # Harita sekmesini tamamen boşalt
            for widget in self.harita_frame.winfo_children():
                widget.destroy()
            
            # Yükleme mesajı göster
            loading_label = tk.Label(self.harita_frame, 
                                   text="Harita yükleniyor...\nLütfen bekleyin.",
                                   font=('Segoe UI', 14),
                                   bg="#2c3e50", fg="white")
            loading_label.pack(expand=True)
            self.update()
            
            # Başlık frame
            baslik_frame = tk.Frame(self.harita_frame, bg="#3498db", height=60)
            baslik_frame.pack(fill='x', padx=1, pady=1)
            baslik_frame.pack_propagate(False)
            
            tk.Label(baslik_frame, text="📍 HARİTA ANALİZİ - BULGU GÖRSELLEŞTİRME", 
                    font=('Segoe UI', 16, 'bold'), fg='white', bg="#3498db").pack(expand=True, pady=15)
            
            tk.Label(baslik_frame, text=f"Dosya: {os.path.basename(self.excel_dosyasi)}", 
                    font=('Segoe UI', 11), fg='white', bg="#3498db").pack(pady=(0, 15))
            
            # Yükleme mesajını kaldır
            loading_label.destroy()
            
            # GömülüHaritaAnalizEkrani'ni ekle
            self.gomulu_harita = GömülüHaritaAnalizEkrani(self.harita_frame, self.excel_dosyasi)
            self.gomulu_harita.pack(fill='both', expand=True, padx=10, pady=(0, 10))
            
            print(f"✅ Harita analizi sekmesi LAZY LOADING ile yüklendi")
            
        except Exception as e:
            messagebox.showerror("Hata", f"Harita sekmesi oluşturulamadı: {str(e)}")
            import traceback
            print(f"Harita sekmesi hatası: {traceback.format_exc()}")
            
            # Hata durumunda bilgi mesajı göster
            error_frame = tk.Frame(self.harita_frame, bg="#2c3e50")
            error_frame.pack(expand=True)
            
            tk.Label(error_frame, text="Harita yüklenemedi!", 
                    font=('Segoe UI', 16, 'bold'), fg='red', bg="#2c3e50").pack(pady=10)
            
            tk.Label(error_frame, text=f"Hata: {str(e)}", 
                    font=('Segoe UI', 11), fg='white', bg="#2c3e50", wraplength=600).pack(pady=10)
            
            tk.Button(error_frame, text="Tekrar Dene", 
                     command=self.harita_sekmesi_olustur,
                     bg="#3498db", fg='white', font=('Segoe UI', 11)).pack(pady=20)


    
    def pyqt_kontrol_et(self):
        """PyQt6'nın kurulu olup olmadığını kontrol eder"""
        try:
            import sys
            from PyQt6.QtCore import Qt, QUrl
            from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout
            from PyQt6.QtWebEngineWidgets import QWebEngineView
            from PyQt6.QtWebEngineCore import QWebEngineSettings
            
            self.pyqt_available = True
            self.harita_uyari_var.set("✅ PyQt6 kurulu - Harita TKinter içinde gösterilebilir")
        except ImportError:
            self.pyqt_available = False
            self.harita_uyari_var.set("⚠️ PyQt6 kurulu değil - Harita tarayıcıda açılacak")
    
    def harita_analiz_ekrani_ac(self):
        """HaritaAnalizEkrani'nı RaporEkrani içine gömülü olarak açar"""
        try:
            # Harita sekmesindeki mevcut içeriği temizle
            for widget in self.harita_frame.winfo_children():
                widget.destroy()
            
            # HaritaAnalizEkrani'ni DOĞRUDAN harita_frame içine göm
            self.harita_ekrani = GömülüHaritaAnalizEkrani(self.harita_frame, self.excel_dosyasi)
            self.harita_ekrani.pack(fill='both', expand=True)
            
        except Exception as e:
            messagebox.showerror("Hata", f"Harita ekranı açılırken hata: {str(e)}")
            import traceback
            print(f"Harita ekranı açma hatası: {traceback.format_exc()}")
            



    def tum_hat_fotograflarini_ac(self, ilce, hat_adi):
        """Bir hata ait TÜM fotoğrafları EXIF düzenleyicide açar - GÜNCELLENMİŞ"""
        try:
            # İlgili hat için TÜM klasör yollarını bul
            klasor_yollari = self.hat_klasor_yollari_bul(ilce, hat_adi)
            
            if not klasor_yollari:
                messagebox.showwarning("Uyarı", f"{ilce} - {hat_adi} için klasör bulunamadı!")
                return
            
            # Tüm klasörlerdeki fotoğrafları ORİJİNAL YOLLARIYLA topla
            tum_fotograflar = []
            for klasor_yolu in klasor_yollari:
                klasor_fotograflari = self.klasordeki_tum_fotograflari_bul(klasor_yolu)
                tum_fotograflar.extend(klasor_fotograflari)
            
            if not tum_fotograflar:
                messagebox.showinfo("Bilgi", f"{ilce} - {hat_adi} için fotoğraf bulunamadı!")
                return
            
            # İlk klasörü ana klasör olarak kullan
            if klasor_yollari:
                ana_klasor_yolu = os.path.dirname(klasor_yollari[0])
                
                # EXIF düzenleyiciyi aç - TÜM fotoğraflar seçili olarak
                exif_editor = ExifDateEditor(self, ana_klasor_yolu, secim_tipi="tumunu")
                exif_editor.log(f"🎯 TÜM HAT FOTOĞRAFLARI: {ilce} - {hat_adi}")
                exif_editor.log(f"📊 Toplam {len(tum_fotograflar)} fotoğraf")
                exif_editor.log(f"📍 Orijinal klasör kullanılıyor: {ana_klasor_yolu}")
                exif_editor.log(f"🎯 Tüm fotoğraflar otomatik seçildi")
                exif_editor.log(f"📅 'Sadece Tarihsizleri Göster' butonuna tıklayarak filtreleyebilirsiniz")
                    
            else:
                messagebox.showerror("Hata", "Klasör bulunamadı!")
                
        except Exception as e:
            messagebox.showerror("Hata", f"Tüm hat fotoğrafları açılırken hata: {str(e)}")

    def klasordeki_tum_fotograflari_bul(self, klasor_yolu):
        """Bir klasördeki tüm fotoğrafları bulur"""
        foto_uzantilari = ('.jpg', '.jpeg', '.png', '.tiff', '.tif', '.bmp')
        tum_fotograflar = []
        
        try:
            for root, dirs, files in os.walk(klasor_yolu):
                for dosya in files:
                    if any(dosya.lower().endswith(ext) for ext in foto_uzantilari):
                        full_path = os.path.join(root, dosya)
                        tum_fotograflar.append(full_path)
        except Exception as e:
            print(f"Klasör tarama hatası {klasor_yolu}: {e}")
        
        return tum_fotograflar





    def verileri_yukle(self):
        """Excel dosyasındaki Hat Özeti sayfasından verileri yükler ve Bulgu sayılarını hesaplar"""
        try:
            # Önce dosyanın var olup olmadığını kontrol et
            if not os.path.exists(self.excel_dosyasi):
                messagebox.showerror("Hata", f"Excel dosyası bulunamadı:\n{self.excel_dosyasi}")
                self.destroy()
                return False

            # 1. AŞAMA: HAT ÖZETİ SAYFASINI OKU (direk sayıları buradan alınacak)
            try:
                df_hat_ozet = pd.read_excel(self.excel_dosyasi, sheet_name='Hat Özeti')
                print(f"✅ DEBUG: Hat Özeti sayfası okundu - {len(df_hat_ozet)} satır")
                
                # Gerekli sütunları kontrol et
                required_columns = ['İlçe', 'Hat Adı', 'Müşterek Direk', 'OG Direk', 'Direk Sayısı']
                for col in required_columns:
                    if col not in df_hat_ozet.columns:
                        print(f"⚠️ UYARI: '{col}' sütunu Hat Özeti sayfasında bulunamadı")
                        # Eksik sütunları boş olarak oluştur
                        df_hat_ozet[col] = 0
                
            except Exception as e:
                messagebox.showerror("Hata", 
                                   f"Hat Özeti sayfası okunamadı:\n{str(e)}\n\n"
                                   f"Dosya: {self.excel_dosyasi}")
                self.destroy()
                return False
            
            # 2. AŞAMA: FOTOĞRAF KLASÖRLERİ SAYFASINI OKU (diğer bilgiler için)
            try:
                df_fotograf = pd.read_excel(self.excel_dosyasi, sheet_name='Fotoğraf Klasörleri')
                print(f"✅ DEBUG: Fotoğraf Klasörleri sayfası okundu - {len(df_fotograf)} satır")
                
                # Mesafe bilgisi kontrolü
                if 'Mesafe (m)' in df_fotograf.columns:
                    mesafe_dolu = df_fotograf['Mesafe (m)'].notna().sum()
                    print(f"🔢 DEBUG: Mesafe sütununda {mesafe_dolu} dolu satır")
                    
                # Bulgu sayıları için Tespit Notu kontrolü
                if 'Tespit Notu' in df_fotograf.columns:
                    print(f"✅ DEBUG: Tespit Notu sütunu mevcut")
                    
            except Exception as e:
                print(f"⚠️ UYARI: Fotoğraf Klasörleri sayfası okunamadı: {e}")
                df_fotograf = pd.DataFrame()

            # 3. AŞAMA: MESAFE BİLGİLERİNİ HESAPLA (Fotoğraf Klasörleri'nden)
            hat_toplam_mesafe = self.hat_toplam_mesafe_hesapla(df_fotograf) if not df_fotograf.empty else {}
            print(f"📊 DEBUG: Toplam {len(hat_toplam_mesafe)} hat için mesafe hesaplandı")
            
            # 4. AŞAMA: BULGU SAYILARINI HESAPLA
            hat_bulgu_sayilari = self.hat_bulgu_sayilari_hesapla(df_fotograf) if not df_fotograf.empty else {}
            
            # 5. AŞAMA: TOPLAM İSTATİSTİKLERİ HESAPLA
            toplam_ilce = df_hat_ozet['İlçe'].nunique() if len(df_hat_ozet) > 0 else 0
            toplam_hat = df_hat_ozet['Hat Adı'].nunique() if len(df_hat_ozet) > 0 else 0
            toplam_direk = df_hat_ozet['Direk Sayısı'].sum() if len(df_hat_ozet) > 0 and 'Direk Sayısı' in df_hat_ozet.columns else 0
            toplam_mustertek = df_hat_ozet['Müşterek Direk'].sum() if len(df_hat_ozet) > 0 and 'Müşterek Direk' in df_hat_ozet.columns else 0
            toplam_og = df_hat_ozet['OG Direk'].sum() if len(df_hat_ozet) > 0 and 'OG Direk' in df_hat_ozet.columns else 0
            
            # Fotoğraf sayısı için Fotoğraf Klasörleri'nden hesaplama
            if not df_fotograf.empty and 'Fotoğraf Sayısı' in df_fotograf.columns:
                toplam_foto = df_fotograf['Fotoğraf Sayısı'].sum()
            else:
                # Yaklaşık hesapla: her direk için ortalama 2 fotoğraf
                toplam_foto = toplam_direk * 2
            
            # Toplam bulgu
            toplam_bulgu = sum(hat_bulgu_sayilari.values()) if hat_bulgu_sayilari else 0
            
            # Toplam mesafe
            toplam_genel_mesafe = sum(hat_toplam_mesafe.values()) if hat_toplam_mesafe else 0

            # 6. AŞAMA: YENİ: Hat bazında ay verilerini sakla - FİLTRELEME İÇİN GEREKLİ
            self.hat_ay_detaylari = {}
            ay_detay = {}  # Tüm veriler için ay detayı
            
            if not df_fotograf.empty and 'Ay' in df_fotograf.columns:
                for index, row in df_fotograf.iterrows():
                    ay = row.get('Ay')
                    aob = row.get('Aob')
                    hat_adi = row.get('Hat adı')
                    hat_anahtari = (aob, hat_adi)
                    
                    if pd.notna(ay):
                        try:
                            # Ay'ı standart formata getir (2 haneli)
                            if isinstance(ay, (int, float)):
                                ay_str = f"{int(ay):02d}"
                            else:
                                ay_str = str(ay).strip().zfill(2)
                            
                            # Geçerli ay mı kontrol et
                            if ay_str in ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12']:
                                # HAT BAZINDA AY DETAYI (filtreleme için)
                                if hat_anahtari not in self.hat_ay_detaylari:
                                    self.hat_ay_detaylari[hat_anahtari] = {}
                                
                                if ay_str not in self.hat_ay_detaylari[hat_anahtari]:
                                    self.hat_ay_detaylari[hat_anahtari][ay_str] = 0
                                
                                self.hat_ay_detaylari[hat_anahtari][ay_str] += 1
                                
                                # TOPLAM AY DETAYI (tüm veriler için)
                                if ay_str not in ay_detay:
                                    ay_detay[ay_str] = 0
                                ay_detay[ay_str] += 1
                        except Exception as e:
                            print(f"Ay işleme hatası: {e}")
                            continue
            
            # TEST: Eğer ay verisi yoksa test verisi ekle
            if not ay_detay:
                print("UYARI: Ay verisi bulunamadı, test verisi ekleniyor...")
                ay_detay = {
                    '01': 150, '02': 200, '03': 180, '04': 220, 
                    '05': 250, '06': 300, '07': 280, '08': 320,
                    '09': 190, '10': 210, '11': 230, '12': 260
                }
            
            # 7. AŞAMA: HAT DETAY VERİLERİNİ HAZIRLA (HAT ÖZETİ'NDEN AL)
            hat_detay = []
            for index, row in df_hat_ozet.iterrows():
                # Hat Özeti'ndeki temel bilgileri al
                ilce_adi = row.get('İlçe', '')
                hat_adi = row.get('Hat Adı', '')
                
                hat_anahtari = (ilce_adi, hat_adi)
                
                # Bu hat için bulgu sayısını al
                bulgu_sayisi = hat_bulgu_sayilari.get(hat_anahtari, 0)
                
                # Bu hat için toplam mesafeyi al
                toplam_mesafe = hat_toplam_mesafe.get(hat_anahtari, 0)
                
                # HAT ÖZETİ'NDEN DİREK SAYILARINI AL
                mustertek_direk = row.get('Müşterek Direk', 0)
                og_direk = row.get('OG Direk', 0)
                direk_sayisi = row.get('Direk Sayısı', 0)
                
                # Fotoğraf sayısını hesapla (Fotoğraf Klasörleri'nden)
                foto_sayisi = 0
                if not df_fotograf.empty and 'Fotoğraf Sayısı' in df_fotograf.columns:
                    try:
                        # Hat için kayıtları filtrele
                        mask = (df_fotograf['Aob'] == ilce_adi) & (df_fotograf['Hat adı'] == hat_adi)
                        hat_fotograflari = df_fotograf[mask]
                        
                        if not hat_fotograflari.empty:
                            foto_sayisi = int(hat_fotograflari['Fotoğraf Sayısı'].sum())
                    except:
                        foto_sayisi = 0

                # Eğer fotoğraf sayısı 0 ise, direk sayısından tahmin et
                if foto_sayisi == 0 and direk_sayisi > 0:
                    foto_sayisi = direk_sayisi * 2  # Her direk için ortalama 2 fotoğraf

                print(f"📊 HAT ÖZETİ: {ilce_adi} - {hat_adi}")
                print(f"   Müşterek: {mustertek_direk}, OG: {og_direk}, Toplam: {direk_sayisi}")
                print(f"   Fotoğraf: {foto_sayisi}, Bulgu: {bulgu_sayisi}")
                
                hat_detay.append({
                    'ilce': ilce_adi,
                    'hat_adi': hat_adi,
                    'musterk_direk': mustertek_direk,
                    'og_direk': og_direk,
                    'direk_sayisi': direk_sayisi,
                    'foto_sayisi': foto_sayisi,
                    'ilk_tarih': row.get('İlk Çekim Tarihi', ''),
                    'son_tarih': row.get('Son Çekim Tarihi', ''),
                    'toplam_gun': row.get('Toplam Gün', 0),
                    'bulgu_sayisi': bulgu_sayisi,
                    'toplam_mesafe': toplam_mesafe,
                    'ilk_tarih_datetime': self.tarih_metninden_datetime(row.get('İlk Çekim Tarihi', '')),
                    'son_tarih_datetime': self.tarih_metninden_datetime(row.get('Son Çekim Tarihi', '')),
                    'tarihsiz_yuzde': self.hat_tarihsiz_istatistikleri.get(hat_anahtari, 0),
                    'trafo_direk': 0,  # Hat Özeti'nde trafo bilgisi yok
                    'og_direk_saf': 0   # Hat Özeti'nde OG Direk (saf) bilgisi yok
                })
            
            # 8. AŞAMA: RAPOR VERİLERİNİ SAKLA
            self.rapor_verileri = {
                'toplam_direk': toplam_direk,
                'toplam_ilce': toplam_ilce,
                'toplam_hat': toplam_hat,
                'toplam_foto': toplam_foto,
                'hat_detay': hat_detay,
                'ay_detay': ay_detay,
                'ilce_list': df_fotograf['Aob'].unique().tolist() if not df_fotograf.empty and 'Aob' in df_fotograf.columns else [],
                'toplam_bulgu': toplam_bulgu,
                'toplam_mesafe': toplam_genel_mesafe,
                'toplam_mustertek': toplam_mustertek,
                'toplam_og': toplam_og,
                'toplam_trafo': 0,  # Hat Özeti'nde trafo bilgisi yok
                'toplam_og_saf': 0   # Hat Özeti'nde OG Direk (saf) bilgisi yok
            }
            
            print(f"🎯 DEBUG: Rapor verileri Hat Özeti'nden hazırlandı")
            print(f"📊 DEBUG: Toplam Müşterek Direk: {toplam_mustertek}")
            print(f"📊 DEBUG: Toplam OG Direk: {toplam_og}")
            print(f"📊 DEBUG: Toplam Direk: {toplam_direk}")
            
            # 9. AŞAMA: GÖRSELLEŞTİRMELERİ GÜNCELLE
            self.istatistik_kartlari_olustur()
            self.grafikleri_ciz()
            self.tablolari_doldur()
            self.filtre_combobox_doldur()
            
            # Öncelik verilerini yükle
            self.after(1000, self.oncelik_verileri_yukle)
            
            return True
        
        except Exception as e:
            messagebox.showerror("Hata", f"Excel dosyası işlenirken hata oluştu: {str(e)}\n\n"
                                      f"Dosya: {self.excel_dosyasi}")
            self.destroy()
            return False

    

    def hat_toplam_mesafe_hesapla(self, df_fotograf):
        """Her hat için toplam mesafeyi hesaplar - DÜZELTİLMİŞ VERSİYON"""
        hat_toplam_mesafe = {}
        
        try:
            # DataFrame boş mu kontrol et
            if df_fotograf.empty:
                print("❌ DEBUG: DataFrame boş")
                return hat_toplam_mesafe
            
            # Gerekli sütunları kontrol et
            required_columns = ['Aob', 'Hat adı', 'Mesafe (m)']
            missing_columns = [col for col in required_columns if col not in df_fotograf.columns]
            
            if missing_columns:
                print(f"❌ DEBUG: Eksik sütunlar: {missing_columns}")
                print(f"🔍 DEBUG: Mevcut sütunlar: {df_fotograf.columns.tolist()}")
                return hat_toplam_mesafe
            
            print(f"✅ DEBUG: {len(df_fotograf)} satır veri işlenecek")
            
            # Her hat için grupla ve toplam mesafeyi hesapla
            for hat_anahtari, group in df_fotograf.groupby(['Aob', 'Hat adı']):
                aob, hat_adi = hat_anahtari
                hat_mesafe_toplam = 0
                gecerli_mesafe_sayisi = 0
                
                for index, row in group.iterrows():
                    mesafe = row.get('Mesafe (m)', '')
                    
                    # Boş değerleri atla
                    if pd.isna(mesafe) or mesafe == '':
                        continue
                    
                    # Mesafe değerini sayıya çevirmeye çalış
                    try:
                        mesafe_str = str(mesafe).strip()
                        
                        # Boş veya geçersiz değerleri atla
                        if mesafe_str in ['', 'Koordinat yok', 'Hesaplanamadı', 'nan', 'None', '-']:
                            continue
                        
                        # Eğer zaten sayıysa doğrudan kullan
                        if isinstance(mesafe, (int, float)):
                            mesafe_float = float(mesafe)
                        else:
                            # String ise temizle ve çevir
                            # Nokta ve virgülleri temizle - binlik ayraç olabilir
                            mesafe_clean = mesafe_str.replace('.', '').replace(',', '.')
                            # Harf ve boşlukları temizle
                            mesafe_clean = ''.join(c for c in mesafe_clean if c.isdigit() or c == '.')
                            
                            if mesafe_clean and mesafe_clean != '.':
                                mesafe_float = float(mesafe_clean)
                            else:
                                continue
                        
                        # ÇOK BÜYÜK sayıları kontrol et (muhtemelen format hatası)
                        if mesafe_float > 1000000:  # 1000 km'den büyükse şüpheli
                            print(f"⚠️  ŞÜPHELİ MESAFE: {hat_anahtari} - {mesafe_float} (çok büyük)")
                            # Büyük sayıları atla veya böl
                            continue
                        
                        # Mantıklı bir mesafe mi kontrol et (0-1000 km arası)
                        if 0 <= mesafe_float <= 1000000:
                            hat_mesafe_toplam += mesafe_float
                            gecerli_mesafe_sayisi += 1
                        else:
                            print(f"⚠️  GEÇERSİZ MESAFE: {hat_anahtari} - {mesafe_float}")
                            
                    except (ValueError, TypeError) as e:
                        print(f"⚠️  Mesafe değeri çevrilemedi: '{mesafe_str}' - Hata: {e}")
                        continue
                
                # Sadece geçerli mesafesi olan hatları ekle
                if gecerli_mesafe_sayisi > 0:
                    hat_toplam_mesafe[hat_anahtari] = hat_mesafe_toplam
                    print(f"✅ {hat_anahtari}: {gecerli_mesafe_sayisi} direk, toplam {hat_mesafe_toplam:.2f} m = {hat_mesafe_toplam/1000:.2f} km")
            
            print(f"✅ DEBUG: {len(hat_toplam_mesafe)} hat için toplam mesafe hesaplandı")
            
            # Toplamları kontrol et
            toplam_mesafe = sum(hat_toplam_mesafe.values())
            print(f"📊 GENEL TOPLAM: {toplam_mesafe:.2f} m = {toplam_mesafe/1000:.2f} km")
            
        except Exception as e:
            print(f"❌ DEBUG: Toplam mesafe hesaplama hatası: {str(e)}")
            import traceback
            print(f"🔍 DEBUG: Hata detayı: {traceback.format_exc()}")
        
        return hat_toplam_mesafe


	

    def filtre_combobox_doldur(self):  # Metod ismini düzeltin
        if self.rapor_verileri:
            ilce_list = self.rapor_verileri['ilce_list']
            self.ilce_filter['values'] = ['TÜMÜ'] + ilce_list

    def tarih_metninden_datetime(self, tarih_metni):
        """Tarih metnini datetime objesine çevirir - DÜZELTİLDİ (isim hatası giderildi)"""
        if not tarih_metni or tarih_metni == 'Belirsiz' or tarih_metni == 'Tarih yok' or tarih_metni == 'Tarihi Yok':
            return None
        
        try:
            # Sadece tarih formatı (gg/aa/yyyy)
            if len(tarih_metni) == 10 and '/' in tarih_metni:
                return datetime.datetime.strptime(tarih_metni, '%d/%m/%Y')
            # Tarih ve saat formatı (gg/aa/yyyy ss:dd:ss)
            elif ' ' in tarih_metni:
                return datetime.datetime.strptime(tarih_metni, '%d/%m/%Y %H:%M:%S')
            else:
                return None
        except:
            return None

    def hat_bulgu_sayilari_hesapla(self, df_fotograf):
        """Her hat için Tespit Notu sayısını hesaplar"""
        hat_bulgu_sayilari = {}
        
        try:
            # Tespit Notu sütunu var mı kontrol et
            if 'Tespit Notu' not in df_fotograf.columns:
                print("UYARI: 'Tespit Notu' sütunu bulunamadı")
                return hat_bulgu_sayilari
            
            # Her satırı işle
            for index, row in df_fotograf.iterrows():
                aob = row.get('Aob', '')
                hat_adi = row.get('Hat adı', '')
                tespit_notu = row.get('Tespit Notu', '')
                
                # Tespit Notu boş değilse say
                if pd.notna(tespit_notu) and str(tespit_notu).strip() != '':
                    hat_anahtari = (aob, hat_adi)
                    
                    if hat_anahtari not in hat_bulgu_sayilari:
                        hat_bulgu_sayilari[hat_anahtari] = 0
                    
                    hat_bulgu_sayilari[hat_anahtari] += 1
            
            print(f"✅ Bulgu sayıları hesaplandı: {len(hat_bulgu_sayilari)} hat için toplam {sum(hat_bulgu_sayilari.values())} bulgu")
            
        except Exception as e:
            print(f"❌ Bulgu sayısı hesaplama hatası: {str(e)}")
        
        return hat_bulgu_sayilari
    
    def hat_bulgu_sayilari_hesapla(self, df_fotograf):
            """Her hat için Tespit Notu sayısını hesaplar"""
            hat_bulgu_sayilari = {}
            
            try:
                # DataFrame boşsa boş döndür
                if df_fotograf.empty:
                    return hat_bulgu_sayilari
                
                # Tespit Notu sütunu var mı kontrol et
                if 'Tespit Notu' not in df_fotograf.columns:
                    print("UYARI: 'Tespit Notu' sütunu bulunamadı")
                    return hat_bulgu_sayilari
                
                # Aob ve Hat adı sütunları var mı kontrol et
                if 'Aob' not in df_fotograf.columns or 'Hat adı' not in df_fotograf.columns:
                    print("UYARI: 'Aob' veya 'Hat adı' sütunları bulunamadı")
                    return hat_bulgu_sayilari
                
                # Her satırı işle
                for index, row in df_fotograf.iterrows():
                    aob = row.get('Aob', '')
                    hat_adi = row.get('Hat adı', '')
                    tespit_notu = row.get('Tespit Notu', '')
                    
                    # Tespit Notu boş değilse say
                    if pd.notna(tespit_notu) and str(tespit_notu).strip() != '':
                        hat_anahtari = (aob, hat_adi)
                        
                        if hat_anahtari not in hat_bulgu_sayilari:
                            hat_bulgu_sayilari[hat_anahtari] = 0
                        
                        hat_bulgu_sayilari[hat_anahtari] += 1
                
                print(f"✅ Bulgu sayıları hesaplandı: {len(hat_bulgu_sayilari)} hat için toplam {sum(hat_bulgu_sayilari.values())} bulgu")
                
            except Exception as e:
                print(f"❌ Bulgu sayısı hesaplama hatası: {str(e)}")
            
            return hat_bulgu_sayilari

    
    def istatistik_kartlari_olustur(self):
        """İstatistik kartlarını oluşturur - DÜZELTİLMİŞ VERSİYON"""
        for widget in self.kartlar_frame.winfo_children():
            widget.destroy()
            
        # YENİ: Toplam Mesafe kartını ekle - DÜZELTİLDİ
        kart_verileri = [
            {"baslik": "Toplam İlçe", "deger": self.rapor_verileri['toplam_ilce'], "birim": "adet", "renk": "#27ae60", "icon": "🏙️"},
            {"baslik": "Toplam Hat", "deger": self.rapor_verileri['toplam_hat'], "birim": "adet", "renk": "#e74c3c", "icon": "🛣️"},
            {"baslik": "Toplam Direk", "deger": self.rapor_verileri['toplam_direk'], "birim": "adet", "renk": "#3498db", "icon": "🏗️"},
            {"baslik": "Toplam Fotoğraf", "deger": self.rapor_verileri['toplam_foto'], "birim": "adet", "renk": "#f39c12", "icon": "📷"},
            {"baslik": "Toplam Bulgu", "deger": self.rapor_verileri['toplam_bulgu'], "birim": "adet", "renk": "#9b59b6", "icon": "🔍"},
            {"baslik": "Toplam Mesafe", "deger": f"{self.rapor_verileri['toplam_mesafe']/1000:.2f}", "birim": "km", "renk": "#1abc9c", "icon": "📏"},  # DÜZELTİLDİ: metre -> km
        ]
        
        # Grid layout için satır ve sütun ayarı
        rows = 1
        cols = len(kart_verileri)
        
        for i, kart in enumerate(kart_verileri):
            kart_frame = tk.Frame(self.kartlar_frame, bg=kart['renk'], relief='raised', bd=2)
            kart_frame.grid(row=0, column=i, sticky='nsew', padx=3, pady=2)
            
            # Grid hücrelerinin eşit genişlemesi için
            self.kartlar_frame.grid_columnconfigure(i, weight=1)
            self.kartlar_frame.grid_rowconfigure(0, weight=1)
            
            # Kart frame'inin içeriği için grid
            kart_frame.grid_columnconfigure(0, weight=1)
            kart_frame.grid_rowconfigure(0, weight=1)  # İkon ve başlık
            kart_frame.grid_rowconfigure(1, weight=1)  # Değer ve birim
            
            # İkon ve başlık - ÜST SATIR
            top_frame = tk.Frame(kart_frame, bg=kart['renk'])
            top_frame.grid(row=0, column=0, sticky='ew', padx=8, pady=(8, 2))
            
            # İkon SOL - BÜYÜK
            icon_label = tk.Label(top_frame, text=kart['icon'], font=('Segoe UI', 16),
                                bg=kart['renk'], fg='white')
            icon_label.pack(side='left', padx=(0, 8))
            
            # Başlık - BÜYÜK
            baslik_label = tk.Label(top_frame, text=kart['baslik'], font=('Segoe UI', 11, 'bold'),
                                  bg=kart['renk'], fg='white', wraplength=80, justify='left')
            baslik_label.pack(side='left', fill='x', expand=True)
            
            # Değer ve birim - ALT SATIR
            bottom_frame = tk.Frame(kart_frame, bg=kart['renk'])
            bottom_frame.grid(row=1, column=0, sticky='ew', padx=8, pady=(2, 8))
            
            # Değer SOL - BÜYÜK
            deger_text = kart['deger']
            
            # Sayısal değerleri formatla, string değerleri olduğu gibi bırak
            if isinstance(deger_text, (int, float)):
                if kart['baslik'] == "Toplam Mesafe":
                    # Mesafe için özel format (2 ondalık)
                    deger_text = f"{deger_text:.2f}"
                else:
                    # Diğer sayısal değerler için binlik ayracı
                    deger_text = f"{deger_text:,}"
            
            deger_label = tk.Label(bottom_frame, text=deger_text, font=('Segoe UI', 18, 'bold'),
                                 bg=kart['renk'], fg='white')
            deger_label.pack(side='left', padx=(0, 8))
            
            # Birim SAĞ - BÜYÜK
            birim_label = tk.Label(bottom_frame, text=kart['birim'], font=('Segoe UI', 12, 'bold'),
                                 bg=kart['renk'], fg='white')
            birim_label.pack(side='left')
            
        # Toplam Bulgu kartına tıklanabilirlik ekle
        if len(self.kartlar_frame.winfo_children()) >= 5:  # 5. kart Toplam Bulgu
            bulgu_kart = self.kartlar_frame.winfo_children()[4]
            bulgu_kart.bind('<Button-1>', lambda e: self.toplam_bulgu_tikla())
            bulgu_kart.bind('<Enter>', lambda e: bulgu_kart.config(cursor="hand2"))
            bulgu_kart.bind('<Leave>', lambda e: bulgu_kart.config(cursor=""))




    def tablolari_doldur(self):
        if self.rapor_verileri:
            self.mevcut_filtreli_veri = sorted(self.rapor_verileri['hat_detay'], 
                                             key=lambda x: x['direk_sayisi'], reverse=True)
            self.tabloyu_guncelle(self.mevcut_filtreli_veri)

    def tabloyu_guncelle(self, hat_verileri):
        """Tabloyu günceller - DÜZELTİLMİŞ VERSİYON"""
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        for detay in hat_verileri:
            # Tarih değerlerini kontrol et ve "Tarihi Yok" olarak göster
            ilk_tarih = detay['ilk_tarih'] if detay['ilk_tarih'] not in ['Belirsiz', 'Tarih yok'] else 'Tarihi Yok'
            son_tarih = detay['son_tarih'] if detay['son_tarih'] not in ['Belirsiz', 'Tarih yok'] else 'Tarihi Yok'
            
            # SARI RENK KOŞULU - SADECE TOPLAM GÜN SÜTUNU İÇİN
            hat_anahtari = (detay['ilce'], detay['hat_adi'])
            tarihsiz_yuzde = detay.get('tarihsiz_yuzde', 0)
            if tarihsiz_yuzde > 0:
                toplam_gun_deger = f"⚠️ {detay['toplam_gun']} gün"  # Uyarı işareti + koyu yazı
            else:
                toplam_gun_deger = f"{detay['toplam_gun']} gün"
            
            # YENİ: Toplam mesafe bilgisini formatla - DÜZELTİLDİ
            toplam_mesafe = detay.get('toplam_mesafe', 0)
            if toplam_mesafe > 0:
                mesafe_deger = f"{toplam_mesafe/1000:.2f} km"  # DÜZELTİLDİ: metre -> km
            else:
                mesafe_deger = "-"
            
            # Tabloya satır ekle - TOPLAM MESAFE SÜTUNU EKLENDİ
            item = self.tree.insert('', 'end', values=(
                detay['ilce'],
                detay['hat_adi'],
                f"{detay['musterk_direk']:,}",
                f"{detay['og_direk']:,}", 
                f"{detay['direk_sayisi']:,}",
                f"{detay['foto_sayisi']:,}",
                f"{detay['bulgu_sayisi']:,}",
                ilk_tarih,
                son_tarih,
                toplam_gun_deger,
                mesafe_deger  # YENİ: Toplam Mesafe sütunu
            ))
            
            # Tarihsiz fotoğraf yüzdesi 0'dan büyükse özel tag ekle
            if tarihsiz_yuzde > 0:
                self.tree.item(item, tags=('uyari_gun',))
        
        # Tag stillerini ayarla - koyu yazı tipi
        self.tree.tag_configure('uyari_gun', font=('Segoe UI', 10, 'bold'), foreground='#d35400')


    def arayuz_olustur(self):
        # ANA PENCEREYİ 9 BİRİM OLARAK AYARLA
        self.grid_rowconfigure(0, weight=3)  # Üst bölüm: 3 birim
        self.grid_rowconfigure(1, weight=3)  # Grafik 1: 3 birim
        self.grid_rowconfigure(2, weight=3)  # Grafik 2: 3 birim
        self.grid_columnconfigure(0, weight=1)
        
        # 1. BÖLÜM: ÜST BAŞLIK VE İSTATİSTİKLER (3 BİRİM)
        top_frame = tk.Frame(self, bg=self.colors['dark_bg'])
        top_frame.grid(row=0, column=0, sticky='nsew', padx=20, pady=10)
        
        title_frame = tk.Frame(top_frame, bg=self.colors['dark_bg'])
        title_frame.pack(fill='x', pady=(0, 10))
        
        title_label = tk.Label(title_frame, text="📊 DİREK İSTATİSTİK RAPORU", 
                              font=('Segoe UI', 20, 'bold'),
                              fg=self.colors['text'], bg=self.colors['dark_bg'])
        title_label.pack()
        
        dosya_label = tk.Label(title_frame, text=f"Dosya: {os.path.basename(self.excel_dosyasi)}", 
                              font=('Segoe UI', 10),
                              fg=self.colors['accent'], bg=self.colors['dark_bg'])
        dosya_label.pack()
        
        stats_frame = tk.Frame(top_frame, bg=self.colors['dark_bg'])
        stats_frame.pack(fill='x', pady=10)
        
        self.kartlar_frame = tk.Frame(stats_frame, bg=self.colors['dark_bg'])
        self.kartlar_frame.pack(fill='x')
        
        # 2. VE 3. BÖLÜM İÇİN ANA İÇERİK FRAME
        content_frame = tk.Frame(self, bg=self.colors['dark_bg'])
        content_frame.grid(row=1, column=0, rowspan=2, sticky='nsew', padx=20, pady=(0, 10))
        
        # CONTENT FRAME'İ 2 SATIRA BÖL (GRAFİK 1 ve GRAFİK 2 için)
        content_frame.grid_rowconfigure(0, weight=1)  # Grafik 1: 3 birim
        content_frame.grid_rowconfigure(1, weight=1)  # Grafik 2: 3 birim
        content_frame.grid_columnconfigure(0, weight=1)
        content_frame.grid_columnconfigure(1, weight=1)
        
        # SOL TARAF: TABLO (GRAFİKLERLE AYNI YÜKSEKLİKTE)
        left_frame = tk.LabelFrame(content_frame, text="🛣️ Hat Detayları", 
                                  font=('Segoe UI', 14, 'bold'),
                                  fg=self.colors['text'], bg=self.colors['light_bg'],
                                  padx=15, pady=15)
        left_frame.grid(row=0, column=0, rowspan=2, sticky='nsew', padx=(0, 10))
        
        # SAĞ TARAF: GRAFİKLER (ÜST ÜSTE)
        right_frame = tk.Frame(content_frame, bg=self.colors['dark_bg'])
        right_frame.grid(row=0, column=1, rowspan=2, sticky='nsew', padx=(10, 0))
        
        # SAĞ FRAME'İ 2 EŞİT PARÇAYA BÖL
        right_frame.grid_rowconfigure(0, weight=1)  # Grafik 1: 3 birim
        right_frame.grid_rowconfigure(1, weight=1)  # Grafik 2: 3 birim
        right_frame.grid_columnconfigure(0, weight=1)
        
        # SOL FRAME İÇERİĞİ (TABLO)
        filter_frame = tk.Frame(left_frame, bg=self.colors['light_bg'])
        filter_frame.pack(fill='x', pady=(0, 10))
        
        tk.Label(filter_frame, text="İlçe:", font=('Segoe UI', 10, 'bold'),
                bg=self.colors['light_bg'], fg=self.colors['text']).pack(side='left', padx=(0, 5))
        
        self.ilce_filter_var = tk.StringVar(value="TÜMÜ")
        self.ilce_filter = ttk.Combobox(filter_frame, textvariable=self.ilce_filter_var, state="readonly", width=15)
        self.ilce_filter.pack(side='left', padx=(0, 15))
        
        tarih_filter_frame = tk.Frame(filter_frame, bg=self.colors['light_bg'])
        tarih_filter_frame.pack(side='left', padx=(0, 15))
        
        tk.Label(tarih_filter_frame, text="Tarih Aralığı:", font=('Segoe UI', 10, 'bold'),
                bg=self.colors['light_bg'], fg=self.colors['text']).pack(anchor='w')
        
        tarih_selection_frame = tk.Frame(tarih_filter_frame, bg=self.colors['light_bg'])
        tarih_selection_frame.pack(fill='x', pady=(5, 0))
        
        start_frame = tk.Frame(tarih_selection_frame, bg=self.colors['light_bg'])
        start_frame.pack(side='left', padx=(0, 10))
        
        tk.Label(start_frame, text="Başlangıç:", font=('Segoe UI', 9),
                bg=self.colors['light_bg'], fg=self.colors['text']).pack(anchor='w')
        
        self.baslangic_tarih_var = tk.StringVar()
        self.baslangic_entry = DateEntry(start_frame, 
                                       textvariable=self.baslangic_tarih_var,
                                       date_pattern='dd/mm/yyyy',
                                       width=12, 
                                       font=('Segoe UI', 9),
                                       background='white',
                                       foreground='black',
                                       borderwidth=1,
                                       locale='tr_TR')
        self.baslangic_entry.pack(pady=(2, 0))
        self.baslangic_entry.delete(0, tk.END)
        
        self.takvimi_turkcelestir(self.baslangic_entry)
        
        end_frame = tk.Frame(tarih_selection_frame, bg=self.colors['light_bg'])
        end_frame.pack(side='left', padx=(10, 0))
        
        tk.Label(end_frame, text="Bitiş:", font=('Segoe UI', 9),
                bg=self.colors['light_bg'], fg=self.colors['text']).pack(anchor='w')
        
        self.bitis_tarih_var = tk.StringVar()
        self.bitis_entry = DateEntry(end_frame, 
                                   textvariable=self.bitis_tarih_var,
                                   date_pattern='dd/mm/yyyy',
                                   width=12, 
                                   font=('Segoe UI', 9),
                                       background='white',
                                       foreground='black',
                                       borderwidth=1,
                                       locale='tr_TR')
        self.bitis_entry.pack(pady=(2, 0))
        self.bitis_entry.delete(0, tk.END)
        
        self.takvimi_turkcelestir(self.bitis_entry)
        
        button_frame = tk.Frame(filter_frame, bg=self.colors['light_bg'])
        button_frame.pack(side='left', padx=(15, 0))
        
        filter_btn = tk.Button(button_frame, text="Filtrele", font=('Segoe UI', 10, 'bold'),
                              bg=self.colors['accent'], fg='white', relief='raised',
                              command=self.filtrele, width=10)
        filter_btn.pack(side='left', padx=(0, 5))
        
        reset_btn = tk.Button(button_frame, text="Sıfırla", font=('Segoe UI', 10, 'bold'),
                             bg=self.colors['warning'], fg='white', relief='raised',
                             command=self.filtreleri_sifirla, width=10)
        reset_btn.pack(side='left')
        
        tree_frame = tk.Frame(left_frame, bg=self.colors['light_bg'])
        tree_frame.pack(fill='both', expand=True)
        

        # Tablo sütunlarını İSTENEN SIRAYA göre tanımla - TOPLAM MESAFE SÜTUNU EKLENDİ
        self.tree = ttk.Treeview(tree_frame, columns=(
            'Ilce', 'HatAdi', 'MusterkDirek', 'OgDirek', 'DirekSayisi', 
            'FotoSayisi', 'BulguSayisi', 'IlkTarih', 'SonTarih', 'ToplamGun', 'ToplamMesafe'  # YENİ: ToplamMesafe eklendi
        ), show='headings')
        
        # Sütun başlıklarını İSTENEN SIRAYA göre ayarla - TOPLAM MESAFE SÜTUNU EKLENDİ
        self.tree.heading('Ilce', text='İlce', command=lambda: self.sutun_sirala('Ilce'))
        self.tree.heading('HatAdi', text='Hat Adı', command=lambda: self.sutun_sirala('HatAdi'))
        self.tree.heading('MusterkDirek', text='Müşterek Direk', command=lambda: self.sutun_sirala('MusterkDirek'))
        self.tree.heading('OgDirek', text='OG Direk', command=lambda: self.sutun_sirala('OgDirek'))
        self.tree.heading('DirekSayisi', text='Direk Sayısı', command=lambda: self.sutun_sirala('DirekSayisi'))
        self.tree.heading('FotoSayisi', text='Fotoğraf Sayısı', command=lambda: self.sutun_sirala('FotoSayisi'))
        self.tree.heading('BulguSayisi', text='Bulgu', command=lambda: self.sutun_sirala('BulguSayisi'))
        self.tree.heading('IlkTarih', text='İlk Çekim Tarihi', command=lambda: self.sutun_sirala('IlkTarih'))
        self.tree.heading('SonTarih', text='Son Çekim Tarihi', command=lambda: self.sutun_sirala('SonTarih'))
        self.tree.heading('ToplamGun', text='Toplam Gün', command=lambda: self.sutun_sirala('ToplamGun'))
        self.tree.heading('ToplamMesafe', text='Toplam Mesafe', command=lambda: self.sutun_sirala('ToplamMesafe'))  # YENİ
        
        # Sütun genişliklerini ayarla - TOPLAM MESAFE SÜTUNU EKLENDİ
        self.tree.column('Ilce', width=120, anchor='center')
        self.tree.column('HatAdi', width=200, anchor='w')
        self.tree.column('MusterkDirek', width=120, anchor='center')
        self.tree.column('OgDirek', width=100, anchor='center')
        self.tree.column('DirekSayisi', width=100, anchor='center')
        self.tree.column('FotoSayisi', width=120, anchor='center')
        self.tree.column('BulguSayisi', width=80, anchor='center')
        self.tree.column('IlkTarih', width=120, anchor='center')
        self.tree.column('SonTarih', width=120, anchor='center')
        self.tree.column('ToplamGun', width=100, anchor='center')
        self.tree.column('ToplamMesafe', width=120, anchor='center')  # YENİ        
        
        # Çift tıklama olayını bağla
        self.tree.bind('<Double-1>', self.tabloya_cift_tikla)
        
        scrollbar = ttk.Scrollbar(tree_frame, orient='vertical', command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # SAĞ TARAF: GRAFİKLER (EŞİT BOYUTTA)
        
        # GRAFİK 1: İlçelere göre direk dağılımı (3 BİRİM)
        self.graf1_frame = tk.LabelFrame(right_frame, text="🏙️ İlçelere Göre Direk Dağılımı", 
                                       font=('Segoe UI', 12, 'bold'),
                                       fg=self.colors['text'], bg=self.colors['light_bg'],
                                       padx=10, pady=10)
        self.graf1_frame.grid(row=0, column=0, sticky='nsew', pady=(0, 5))
        
        # GRAFİK 2: Aylara göre direk sayısı (3 BİRİM)
        self.graf2_frame = tk.LabelFrame(right_frame, text="📅 Aylara Göre Direk Sayısı", 
                                       font=('Segoe UI', 12, 'bold'),
                                       fg=self.colors['text'], bg=self.colors['light_bg'],
                                       padx=10, pady=10)
        self.graf2_frame.grid(row=1, column=0, sticky='nsew', pady=(5, 0))
        
        # BOŞ GRAFİKLERİ OLUŞTUR
        self.fig1 = plt.Figure(figsize=(6, 3.5), dpi=100)
        self.ax1 = self.fig1.add_subplot(111)
        self.canvas1 = FigureCanvasTkAgg(self.fig1, self.graf1_frame)
        self.canvas1.get_tk_widget().pack(fill='both', expand=True)
        
        self.fig2 = plt.Figure(figsize=(6, 3.5), dpi=100)
        self.ax2 = self.fig2.add_subplot(111)
        self.canvas2 = FigureCanvasTkAgg(self.fig2, self.graf2_frame)
        self.canvas2.get_tk_widget().pack(fill='both', expand=True)
    
    def tabloya_cift_tikla(self, event):
        """Tablo hücresine çift tıklandığında hangi sütuna tıklandığını tespit eder ve ilgili işlemi yapar"""
        item = self.tree.selection()[0] if self.tree.selection() else None
        if not item:
            return
        
        # Tıklanan sütunu tespit et
        x = event.x
        column = self.tree.identify_column(x)
        column_index = int(column.replace('#', '')) - 1  # Sütun indeksi (0'dan başlar)

        values = self.tree.item(item, 'values')
        if not values:
            return
        
        ilce = values[0]
        hat_adi = values[1]
        direk_no = values[3]  # Direk No sütunu (4. sütun, indeks 3)

        # DİREK NO SÜTUNUNA TIKLANDIĞINDA TESPİT EXCEL'İNİ FİLTRELİ AÇ - YENİ EKLENDİ
        if column_index == 3:  # Direk No sütunu (4. sütun, indeks 3)
            self.direk_no_icin_tespit_excelini_ac(ilce, hat_adi, direk_no)
        
        # BULGU SÜTUNUNA TIKLANDIĞINDA EXCEL DOSYASINI AÇ
        elif column_index == 6:  # Bulgu sütunu (7. sütun, indeks 6)
            bulgu_sayisi = values[6]
            if bulgu_sayisi != '0' and bulgu_sayisi != '':
                self.hat_excel_dosyasini_ac(ilce, hat_adi)
            else:
                messagebox.showinfo("Bilgi", "Bu hat için bulgu bulunmuyor")
        
        # Tarih sütunlarına tıklandığında EXIF düzenleyiciyi aç
        elif column_index == 7:  # İlk Tarih sütunu (8. sütun, indeks 7)
            tarih_degeri = values[7]
            if tarih_degeri != 'Tarih yok' and tarih_degeri != 'Belirsiz' and tarih_degeri != 'Tarihi Yok':
                self.tarih_klasorunu_ac(ilce, hat_adi, tarih_degeri, "ilk")
            else:
                messagebox.showinfo("Bilgi", "İlk tarih bilgisi bulunmuyor")
            
        elif column_index == 8:  # Son Tarih sütunu (9. sütun, indeks 8)
            tarih_degeri = values[8]
            if tarih_degeri != 'Tarih yok' and tarih_degeri != 'Belirsiz' and tarih_degeri != 'Tarihi Yok':
                self.tarih_klasorunu_ac(ilce, hat_adi, tarih_degeri, "son")
            else:
                messagebox.showinfo("Bilgi", "Son tarih bilgisi bulunmuyor")
        
        # TOPLAM GÜN sütununa tıklandığında DİREKT EXIF düzenleyiciyi aç (soru sormadan)
        elif column_index == 9:  # Toplam Gün sütunu (10. sütun, indeks 9)
            toplam_gun = values[9]
            if toplam_gun != '0 gün' and toplam_gun != 'Belirsiz':
                # DİREKT EXIF düzenleyiciyi aç - TÜM fotoğraflar seçili olarak
                self.tum_hat_fotograflarini_ac(ilce, hat_adi)
            else:
                messagebox.showinfo("Bilgi", "Bu hat için tarih bilgisi bulunmuyor")
        
        # HAT ADI sütununa tıklandığında klasörü aç
        elif column_index == 1:  # Hat Adı sütunu (2. sütun, indeks 1)
            self.hat_klasoru_ac(ilce, hat_adi)
        
        else:
            # Diğer sütunlara tıklandığında hiçbir şey yapma
            return

    def hat_excel_dosyasini_ac(self, ilce, hat_adi):
        """Rapor Excel'ini BULGULAR FİLTRELİ açar ve HAT TESPİT Excel'ini de açar - DİREK ID TEMİZLİ"""
        try:
            # Rapor dosyasını bulmak için farklı yollar deneyelim
            rapor_dosyasi = None
            
            # 1. Önce AnaUygulama'daki kaydetme_yeri'ni kontrol et
            if hasattr(self, 'parent') and hasattr(self.parent, 'kaydetme_yeri'):
                ana_rapor = self.parent.kaydetme_yeri.get()
                if ana_rapor and os.path.exists(ana_rapor):
                    rapor_dosyasi = ana_rapor
                    self.log(f"📁 Ana uygulamadan rapor bulundu: {os.path.basename(rapor_dosyasi)}")
            
            # 2. RaporEkrani oluşturulurken verilen excel_dosyasi'ni kontrol et
            if not rapor_dosyasi and hasattr(self, 'excel_dosyasi'):
                if self.excel_dosyasi and os.path.exists(self.excel_dosyasi):
                    rapor_dosyasi = self.excel_dosyasi
                    self.log(f"📁 RaporEkrani'nden rapor bulundu: {os.path.basename(rapor_dosyasi)}")
            
            # 3. Son çare: kullanıcıya sor
            if not rapor_dosyasi:
                rapor_dosyasi = filedialog.askopenfilename(
                    title="Rapor Excel Dosyasını Seçin",
                    filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
                )
                if not rapor_dosyasi:
                    return
            
            if not os.path.exists(rapor_dosyasi):
                messagebox.showerror("Hata", f"Rapor dosyası bulunamadı: {rapor_dosyasi}")
                return
            
            self.log(f"🔍 Rapor analiz ediliyor: {os.path.basename(rapor_dosyasi)}")
            
            # Rapor Excel'ini oku
            df_rapor = pd.read_excel(rapor_dosyasi, sheet_name='Fotoğraf Klasörleri')
            
            # Sütunları logla
            self.log(f"📊 Excel sütunları: {df_rapor.columns.tolist()}")
            
            # 1. İlçe ve hat adına göre filtrele
            if 'Aob' not in df_rapor.columns or 'Hat adı' not in df_rapor.columns:
                messagebox.showerror("Hata", "Excel'de 'Aob' veya 'Hat adı' sütunu bulunamadı!")
                return
            
            filtre_ilce_hat = (df_rapor['Aob'] == ilce) & (df_rapor['Hat adı'] == hat_adi)
            hat_satirlari = df_rapor[filtre_ilce_hat]
            
            self.log(f"🔍 {ilce}-{hat_adi} için {len(hat_satirlari)} satır bulundu")
            
            if hat_satirlari.empty:
                messagebox.showinfo("Bilgi", f"{ilce} - {hat_adi} için kayıt bulunamadı!")
                
                # Mevcut ilçe ve hatları göster
                ilceler = df_rapor['Aob'].dropna().unique()
                hatlar = df_rapor['Hat adı'].dropna().unique()
                self.log(f"📋 Mevcut ilçeler: {ilceler}")
                self.log(f"📋 Mevcut hatlar: {hatlar}")
                return
            
            # 2. Tespit Notu olan satırları bul (BULGULAR)
            tespit_sutunu = 'Tespit Notu'
            if tespit_sutunu not in hat_satirlari.columns:
                messagebox.showwarning("Uyarı", f"'{tespit_sutunu}' sütunu bulunamadı!")
                self.log(f"📋 Mevcut sütunlar: {hat_satirlari.columns.tolist()}")
                return
            
            # Tespit Notu dolu olan satırları filtrele
            filtre_tespit = (hat_satirlari[tespit_sutunu].notna()) & (hat_satirlari[tespit_sutunu] != '')
            bulgu_satirlari = hat_satirlari[filtre_tespit]
            
            self.log(f"🎯 {ilce}-{hat_adi} için {len(bulgu_satirlari)} bulgu satırı bulundu")
            
            if bulgu_satirlari.empty:
                messagebox.showinfo("Bilgi", f"{ilce} - {hat_adi} için bulgu bulunamadı!")
                # Bulgu yoksa sadece hat Excel'ini aç
                self.hat_tespit_excelini_ac_temizli(ilce, hat_adi)
                return
            
            # Sadece HAT TESPİT Excel'ini aç - DİREK ID TEMİZLİ
            self.hat_tespit_excelini_ac_temizli(ilce, hat_adi)
            
        except Exception as e:
            messagebox.showerror("Hata", f"Excel açılırken hata: {str(e)}")
            import traceback
            self.log(f"🔍 Hata detayı: {traceback.format_exc()}")

    def direk_id_isle_excel(self, direk_id_serisi):
        """Excel için Direk ID temizleme - pandas Series uyumlu"""
        try:
            def temizle_deger(deger):
                if pd.isna(deger) or deger == '' or deger is None:
                    return ''
                
                deger_str = str(deger).strip()
                deger_str = deger_str.replace('-', '')  # Tüm - işaretlerini kaldır
                
                if not deger_str:
                    return ''
                
                if deger_str.isdigit():
                    try:
                        return int(deger_str)
                    except:
                        return deger_str
                else:
                    return deger_str
            
            return direk_id_serisi.apply(temizle_deger)
            
        except Exception as e:
            self.log(f"⚠️ Excel Direk ID temizleme hatası: {str(e)}")
            return direk_id_serisi

    def hat_tespit_excelini_ac_temizli(self, ilce, hat_adi):
        """Orjinal Excel'i aç, temizle, hizala, filtre ekle ve kaydet - SESSİZ AÇMA"""
        try:
            # 1. İlgili hat için klasör yollarını bul
            klasor_yollari = self.hat_klasor_yollari_bul(ilce, hat_adi)
            
            if not klasor_yollari:
                self.log(f"⚠️ {ilce}-{hat_adi} için klasör bulunamadı!")
                return
            
            # 2. Hat klasörünü bul
            ilk_klasor = klasor_yollari[0]
            hat_klasoru = os.path.dirname(ilk_klasor)
            
            if not os.path.exists(hat_klasoru):
                self.log(f"⚠️ Hat klasörü bulunamadı: {hat_klasoru}")
                return
            
            # 3. Excel dosyalarını ara
            excel_dosyalari = []
            for dosya in os.listdir(hat_klasoru):
                if dosya.lower().endswith(('.xlsx', '.xls')):
                    excel_dosyalari.append(os.path.join(hat_klasoru, dosya))
            
            if not excel_dosyalari:
                self.log(f"⚠️ {hat_klasoru} klasöründe Excel dosyası bulunamadı!")
                return
            
            # 4. ORİJİNAL Excel dosyasını seç
            orijinal_excel = excel_dosyalari[0]
            
            self.log(f"📂 Orjinal Excel: {os.path.basename(orijinal_excel)}")
            self.log(f"📁 Klasör: {hat_klasoru}")
            
            # 5. Excel dosyası açık mı kontrol et
            try:
                # Dosyayı okuma modunda açmaya çalış
                with open(orijinal_excel, 'rb'):
                    pass
            except PermissionError:
                self.log("⚠️ Excel dosyası açık! Lütfen Excel'i kapatıp tekrar deneyin.")
                messagebox.showwarning("Excel Açık", 
                                     f"{os.path.basename(orijinal_excel)} dosyası açık!\n\n"
                                     f"Lütfen Excel'i kapatıp tekrar deneyin.")
                return
            
            # 6. Excel'i AÇ ve İŞLE
            workbook = load_workbook(orijinal_excel)
            
            # İlk worksheet'i al
            worksheet = workbook.active
            
            self.log(f"📊 Sayfa: {worksheet.title}, Satır: {worksheet.max_row}, Sütun: {worksheet.max_column}")
            
            # 7. Direk ID sütununu bul ve temizle
            direk_id_col_index = None
            for col in range(1, worksheet.max_column + 1):
                cell = worksheet.cell(row=1, column=col)
                if cell.value and isinstance(cell.value, str):
                    if "direk" in cell.value.lower() and "id" in cell.value.lower():
                        direk_id_col_index = col
                        self.log(f"✅ Direk ID sütunu bulundu: {col}. sütun")
                        break
            
            # 8. Direk ID'leri temizle
            if direk_id_col_index:
                for row in range(2, worksheet.max_row + 1):
                    cell = worksheet.cell(row=row, column=direk_id_col_index)
                    if cell.value:
                        # Direk ID'yi temizle
                        temiz_deger = self.direk_id_degerini_temizle_yerel(str(cell.value))
                        cell.value = temiz_deger
                self.log(f"✅ {worksheet.max_row - 1} Direk ID temizlendi")
            
            # 9. TÜM HÜCRELERİ SOLA HİZALA
            for row in worksheet.iter_rows(min_row=1, max_row=worksheet.max_row, 
                                          min_col=1, max_col=worksheet.max_column):
                for cell in row:
                    cell.alignment = Alignment(horizontal='left', vertical='center')
            
            self.log("✅ Tüm hücreler sola hizalandı")
            
            # 10. BAŞLIK SATIRINI BİÇİMLENDİR (mavi arkaplan, beyaz yazı)
            header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
            header_font = Font(bold=True, color="FFFFFF")
            
            for col in range(1, worksheet.max_column + 1):
                header_cell = worksheet.cell(row=1, column=col)
                header_cell.fill = header_fill
                header_cell.font = header_font
                header_cell.alignment = Alignment(horizontal='center', vertical='center')
            
            self.log("✅ Başlık satırı biçimlendirildi")
            
            # 11. FİLTRE EKLE
            if worksheet.max_row > 1:
                filter_range = f"A1:{get_column_letter(worksheet.max_column)}{worksheet.max_row}"
                worksheet.auto_filter.ref = filter_range
                self.log(f"✅ Filtre eklendi: {filter_range}")
            
            # 12. ORİJİNAL DOSYAYI KAYDET
            try:
                workbook.save(orijinal_excel)
                self.log(f"💾 Orjinal dosya kaydedildi: {os.path.basename(orijinal_excel)}")
                
                # Workbook'u kapat
                workbook.close()
                
            except PermissionError as e:
                self.log(f"❌ Dosya kaydedilemedi: {e}")
                messagebox.showerror("Kayıt Hatası", 
                                   f"Dosya kaydedilemedi:\n\n{str(e)}\n\n"
                                   "Excel dosyası açık olabilir. Lütfen kapatıp tekrar deneyin.")
                return
            
            # 13. EXCEL DOSYASINI SESSİZCE AÇ (etkileşim kutusu olmadan)
            self.log(f"🚀 Excel sessizce açılıyor...")
            
            try:
                # Windows için sessiz açma
                if sys.platform == "win32":
                    # Method 1: os.startfile (genellikle sorunsuz çalışır)
                    os.startfile(orijinal_excel)
                    
                    # Alternatif method 2: subprocess ile sessiz açma
                    # subprocess.Popen([orijinal_excel], shell=True)
                    
                elif sys.platform == "darwin":
                    subprocess.run(["open", orijinal_excel])
                else:
                    subprocess.run(["xdg-open", orijinal_excel])
                    
                self.log(f"✅ {ilce}-{hat_adi} TESPİT Excel'i başarıyla açıldı!")
                
            except Exception as e:
                self.log(f"⚠️ Excel açılırken uyarı: {e}")
                # Son çare: kullanıcıya dosya yolunu göster
                messagebox.showinfo("Excel Açıldı", 
                                  f"Excel dosyası işlendi ve kaydedildi.\n\n"
                                  f"Dosya: {orijinal_excel}\n\n"
                                  f"Dosyayı manuel olarak açabilirsiniz.")
            
            self.log("=" * 50)
            self.log("🎯 İŞLEMLER TAMAMLANDI:")
            self.log("   1. Direk ID'ler temizlendi (- işaretleri kaldırıldı)")
            self.log("   2. Tüm hücreler sola hizalandı")
            self.log("   3. Başlık satırı biçimlendirildi")
            self.log("   4. Filtre eklendi")
            self.log("   5. Orjinal dosya kaydedildi")
            self.log("=" * 50)
            
        except Exception as e:
            self.log(f"❌ HATA: {str(e)}")
            import traceback
            self.log(f"🔍 Traceback: {traceback.format_exc()[:500]}")  # İlk 500 karakter

    
    def direk_id_degerini_temizle_yerel(self, direk_id):
        """Yerel versiyon - AnaUygulama'daki metodun aynısı"""
        if not direk_id or direk_id.lower() == 'nan' or direk_id == 'None':
            return ''
        
        # String'e çevir
        direk_str = str(direk_id).strip()
        
        # "-" işaretlerini kaldır (sondaki ve baştaki)
        direk_str = direk_str.rstrip('-').lstrip('-')
        
        if not direk_str:
            return ''
        
        # Rakam mı kontrol et
        if direk_str[0].isdigit():
            try:
                # Nokta veya virgül içeriyor mu kontrol et
                if '.' in direk_str or ',' in direk_str:
                    # Float'a çevir, sonra integer'a çevirmeye çalış
                    float_deger = float(direk_str.replace(',', '.'))
                    # Eğer tam sayı ise integer'a çevir
                    if float_deger.is_integer():
                        return int(float_deger)
                    else:
                        return float_deger
                else:
                    # Doğrudan integer'a çevir
                    return int(direk_str)
            except (ValueError, TypeError):
                # Sayıya çevrilemezse orijinal string değeri kullan
                return direk_str
        else:
            # Metin ile başlıyorsa string olarak sakla
            return direk_str
            



    def log(self, message):
        """Log mesajı için basit metod"""
        # Eğer log_text varsa kullan, yoksa konsola yaz
        if hasattr(self, 'txt_sonuc') and self.txt_sonuc:
            self.txt_sonuc.insert(tk.END, f"{message}\n")
            self.txt_sonuc.see(tk.END)
        else:
            print(f"[Rapor] {message}")

        
    def hat_klasoru_ac(self, ilce, hat_adi):
        """Hat klasörünü Windows Explorer'da açar"""
        try:
            # İlgili hat için TÜM klasör yollarını bul
            klasor_yollari = self.hat_klasor_yollari_bul(ilce, hat_adi)
            
            if not klasor_yollari:
                messagebox.showwarning("Uyarı", f"{ilce} - {hat_adi} için klasör bulunamadı!")
                return
            
            # İlk klasörün BİR ÜST KLASÖRÜNÜ al (hat klasörünü bul)
            ilk_klasor = klasor_yollari[0]
            hat_klasoru = os.path.dirname(ilk_klasor)  # Bir üst klasör (hat klasörü)
            
            if os.path.exists(hat_klasoru):
                # Windows Explorer'da HAT klasörünü aç
                if sys.platform == "win32":
                    os.startfile(hat_klasoru)
                elif sys.platform == "darwin":
                    subprocess.run(["open", hat_klasoru])
                else:
                    subprocess.run(["xdg-open", hat_klasoru])
                
                print(f"📁 Hat klasörü açıldı: {hat_klasoru}")
            else:
                messagebox.showwarning("Uyarı", f"Hat klasörü bulunamadı: {hat_klasoru}")
                
        except Exception as e:
            messagebox.showerror("Hata", f"Klasör açılırken hata: {str(e)}")

    def tarih_klasorunu_ac(self, ilce, hat_adi, tarih_degeri, tarih_tipi):
        """Belirli bir tarihe ait klasörü bul ve EXIF düzenleyiciyi aç"""
        try:
            # İlgili hat için TÜM klasör yollarını bul
            klasor_yollari = self.hat_klasor_yollari_bul(ilce, hat_adi)
        
            if not klasor_yollari:
                messagebox.showwarning("Uyarı", f"{ilce} - {hat_adi} için klasör bulunamadı!")
                return
        
            # Tarihi datetime objesine çevir
            if ' ' in tarih_degeri:  # Tarih ve saat birlikte
                tarih_obj = datetime.datetime.strptime(tarih_degeri, '%d/%m/%Y %H:%M:%S')
            else:  # Sadece tarih
                tarih_obj = datetime.datetime.strptime(tarih_degeri, '%d/%m/%Y')
        
            # Bu tarihe ait klasörü bul
            hedef_klasor = self.tarihe_ait_klasoru_bul(klasor_yollari, tarih_obj, tarih_tipi)
        
            if hedef_klasor:
                # EXIF düzenleyiciyi aç - TÜM fotoğraflar için
                exif_editor = ExifDateEditor(self, hedef_klasor)
                exif_editor.log(f"🕒 {tarih_tipi.upper()} TARİH KLASÖRÜ: {os.path.basename(hedef_klasor)}")
                exif_editor.log(f"📅 Hedef Tarih: {tarih_degeri}")
                exif_editor.log(f"📍 Konum: {ilce} - {hat_adi}")
            else:
                messagebox.showinfo("Bilgi", 
                                  f"{ilce} - {hat_adi} için {tarih_degeri} tarihine ait klasör bulunamadı.\n\n"
                                  f"Tüm klasörleri kontrol etmek için hat adına çift tıklayın.")
            
        except ValueError as e:
            messagebox.showerror("Hata", f"Tarih formatı hatası: {tarih_degeri}\n{e}")
    
    def tarihe_ait_klasoru_bul(self, klasor_yollari, hedef_tarih, tarih_tipi):
        """Belirli bir tarihe ait klasörü bulur"""
        for klasor_yolu in klasor_yollari:
            # Klasördeki tüm fotoğrafları kontrol et
            foto_tarihleri = self.klasor_fotograf_tarihlerini_bul(klasor_yolu)
            
            if not foto_tarihleri:
                continue
                
            if tarih_tipi == "ilk":
                # İlk tarih için: en eski fotoğrafın tarihini kontrol et
                klasor_ilk_tarih = min(foto_tarihleri)
                if self.tarihler_esit_mi(klasor_ilk_tarih, hedef_tarih):
                    return klasor_yolu
            else:  # "son"
                # Son tarih için: en yeni fotoğrafın tarihini kontrol et
                klasor_son_tarih = max(foto_tarihleri)
                if self.tarihler_esit_mi(klasor_son_tarih, hedef_tarih):
                    return klasor_yolu
        
        return None
    
    def tarihler_esit_mi(self, tarih1, tarih2):
        """İki tarihin aynı gün olup olmadığını kontrol et (saat farkını gözardı et)"""
        return (tarih1.year == tarih2.year and 
                tarih1.month == tarih2.month and 
                tarih1.day == tarih2.day)
    
    def hat_klasorunu_ac(self, ilce, hat_adi, klasor_yollari):
        """Tüm hat klasörlerini EXIF düzenleyicide aç"""
        if klasor_yollari:
            # İlk klasörü kullan
            secilen_klasor = klasor_yollari[0]
            
            # EXIF düzenleyiciyi aç - TÜM fotoğraflar için
            exif_editor = ExifDateEditor(self, secilen_klasor)
            exif_editor.log(f"📍 Tüm hat klasörleri: {ilce} - {hat_adi}")
            exif_editor.log(f"📁 Seçilen klasör: {os.path.basename(secilen_klasor)}")
            if len(klasor_yollari) > 1:
                exif_editor.log(f"📂 Toplam {len(klasor_yollari)} klasör bulundu")
        else:
            messagebox.showwarning("Uyarı", f"{ilce} - {hat_adi} için klasör bulunamadı!")
    
    def hat_klasor_yollari_bul(self, ilce, hat_adi):
        """İlçe ve hat adına göre TÜM klasör yollarını bulur"""
        try:
            df = pd.read_excel(self.excel_dosyasi, sheet_name='Fotoğraf Klasörleri')
            
            # İlçe ve hat adına göre filtrele
            filtrelenmis = df[(df['Aob'] == ilce) & (df['Hat adı'] == hat_adi)]
            
            klasor_yollari = []
            
            for index, row in filtrelenmis.iterrows():
                klasor_yolu = row['Klasör Yolu']
                if isinstance(klasor_yolu, str) and os.path.exists(klasor_yolu):
                    klasor_yollari.append(klasor_yolu)
                else:
                    # Klasör yolunu bulamazsa, Excel'deki diğer bilgileri kullanarak yol oluşturmayı dene
                    alternatif_yol = self.alternatif_klasor_yolu_olustur(row, ilce, hat_adi)
                    if alternatif_yol and os.path.exists(alternatif_yol):
                        klasor_yollari.append(alternatif_yol)
            
            return list(set(klasor_yollari))  # Tekilleri döndür
            
        except Exception as e:
            print(f"Klasör yolu bulma hatası: {e}")
            return []

    def alternatif_klasor_yolu_olustur(self, row, ilce, hat_adi):
        """Alternatif klasör yolu oluşturur"""
        try:
            # Ana klasör yolunu Excel dosyasının olduğu dizinden tahmin et
            excel_dizin = os.path.dirname(self.excel_dosyasi)
            
            # Olası klasör yapılarını dene
            olası_yollar = [
                os.path.join(excel_dizin, ilce, hat_adi, str(row['Direk No'])),
                os.path.join(excel_dizin, "Fotoğraflar", ilce, hat_adi, str(row['Direk No'])),
                os.path.join(excel_dizin, "Images", ilce, hat_adi, str(row['Direk No'])),
                os.path.join(os.path.expanduser("~"), "Desktop", ilce, hat_adi, str(row['Direk No'])),
            ]
            
            for yol in olası_yollar:
                if os.path.exists(yol):
                    return yol
            
            return None
            
        except Exception as e:
            print(f"Alternatif yol oluşturma hatası: {e}")
            return None
    
    def klasor_fotograf_tarihlerini_bul(self, klasor_yolu):
        """Bir klasördeki tüm fotoğrafların EXIF tarihlerini bulur"""
        tarihler = []
        foto_uzantilari = ('.jpg', '.jpeg', '.png', '.tiff', '.tif', '.bmp')
        
        try:
            for root, dirs, files in os.walk(klasor_yolu):
                for dosya in files:
                    if any(dosya.lower().endswith(ext) for ext in foto_uzantilari):
                        full_path = os.path.join(root, dosya)
                        tarih = self.fotograf_tarihini_al(full_path)
                        if tarih:
                            tarihler.append(tarih)
        except Exception as e:
            print(f"Klasör tarama hatası {klasor_yolu}: {e}")
        
        return tarihler
    
    def fotograf_tarihini_al(self, dosya_yolu):
        """Fotoğrafın EXIF tarihini alır - GÜNCELLENMİŞ"""
        try:
            with Image.open(dosya_yolu) as img:
                exif_data = img.getexif()
                if exif_data:
                    # Öncelikle çekilme tarihini ara
                    datetime_original = exif_data.get(36867)  # DateTimeOriginal
                    if not datetime_original:
                        datetime_original = exif_data.get(306)  # DateTime
                    
                    if datetime_original:
                        try:
                            return datetime.datetime.strptime(datetime_original, '%Y:%m:%d %H:%M:%S')
                        except ValueError:
                            try:
                                return datetime.datetime.strptime(datetime_original, '%Y-%m:%d %H:%M:%S')
                            except:
                                return None
                
                # Piexif ile de dene
                try:
                    exif_dict = piexif.load(dosya_yolu)
                    if piexif.ExifIFD.DateTimeOriginal in exif_dict['Exif']:
                        tarih_bytes = exif_dict['Exif'][piexif.ExifIFD.DateTimeOriginal]
                        tarih_str = tarih_bytes.decode('utf-8')
                        return datetime.datetime.strptime(tarih_str, '%Y:%m:%d %H:%M:%S')
                except:
                    pass
                
                return None
                
        except Exception as e:
            print(f"EXIF okuma hatası {dosya_yolu}: {e}")
            return None

    def takvimi_turkcelestir(self, date_entry):
        try:
            calendar = date_entry._top_cal
            if calendar:
                for i, month in enumerate(self.turkce_aylar):
                    calendar._month_names[i] = month
                
                for i, day in enumerate(self.turkce_gunler):
                    calendar._week_days[i] = day
        except Exception:
            # Hata olursa hiçbir şey yapma, sessizce devam et
            pass
    
    def sutun_sirala(self, sutun):
        if not self.mevcut_filtreli_veri:
            return
        
        if sutun not in self.siralama_durumu:
            self.siralama_durumu[sutun] = False
        
        ters = self.siralama_durumu[sutun]
        
        if sutun == 'Ilce':
            self.mevcut_filtreli_veri.sort(key=lambda x: x['ilce'], reverse=ters)
        elif sutun == 'HatAdi':
            self.mevcut_filtreli_veri.sort(key=lambda x: x['hat_adi'], reverse=ters)
        elif sutun == 'MusterkDirek':
            self.mevcut_filtreli_veri.sort(key=lambda x: x['musterk_direk'], reverse=ters)
        elif sutun == 'OgDirek':
            self.mevcut_filtreli_veri.sort(key=lambda x: x['og_direk'], reverse=ters)
        elif sutun == 'DirekSayisi':
            self.mevcut_filtreli_veri.sort(key=lambda x: x['direk_sayisi'], reverse=ters)
        elif sutun == 'FotoSayisi':
            self.mevcut_filtreli_veri.sort(key=lambda x: x['foto_sayisi'], reverse=ters)
        elif sutun == 'BulguSayisi':
            self.mevcut_filtreli_veri.sort(key=lambda x: x['bulgu_sayisi'], reverse=ters)
        elif sutun == 'IlkTarih':
            self.mevcut_filtreli_veri.sort(key=lambda x: x['ilk_tarih_datetime'] or datetime.datetime.min, reverse=ters)
        elif sutun == 'SonTarih':
            self.mevcut_filtreli_veri.sort(key=lambda x: x['son_tarih_datetime'] or datetime.datetime.min, reverse=ters)
        elif sutun == 'ToplamGun':
            self.mevcut_filtreli_veri.sort(key=lambda x: x['toplam_gun'], reverse=ters)
        elif sutun == 'ToplamMesafe':  # YENİ: Toplam Mesafe sıralama
            self.mevcut_filtreli_veri.sort(key=lambda x: x['toplam_mesafe'], reverse=ters)
        
        self.tabloyu_guncelle(self.mevcut_filtreli_veri)
        self.sutun_basliklarini_guncelle(sutun, ters)
        self.siralama_durumu[sutun] = not ters
    
    def sutun_basliklarini_guncelle(self, aktif_sutun, ters):
        sutun_bilgileri = {
            'Ilce': 'İlce',
            'HatAdi': 'Hat Adı', 
            'MusterkDirek': 'Müşterek Direk',
            'OgDirek': 'OG Direk',
            'DirekSayisi': 'Direk Sayısı',
            'FotoSayisi': 'Fotoğraf Sayısı',
            'BulguSayisi': 'Bulgu',  # YENİ: Bulgu sütunu
            'IlkTarih': 'İlk Çekim Tarihi',
            'SonTarih': 'Son Çekim Tarihi',
            'ToplamGun': 'Toplam Gün'
        }
        
        for sutun, orijinal_text in sutun_bilgileri.items():
            if sutun == aktif_sutun:
                if ters:
                    self.tree.heading(sutun, text=orijinal_text + ' ▼')
                else:
                    self.tree.heading(sutun, text=orijinal_text + ' ▲')
            else:
                self.tree.heading(sutun, text=orijinal_text)
    
    def mevcut_filtreli_veriyi_al(self):
        if not self.rapor_verileri:
            return []
            
        secili_ilce = self.ilce_filter_var.get()
        baslangic_tarih = self.tarih_metninden_al(self.baslangic_tarih_var.get())
        bitis_tarih = self.tarih_metninden_al(self.bitis_tarih_var.get())
        
        filtrelenmis_veri = self.rapor_verileri['hat_detay']
        
        if secili_ilce != 'TÜMÜ':
            filtrelenmis_veri = [hat for hat in filtrelenmis_veri if hat['ilce'] == secili_ilce]
        
        if baslangic_tarih or bitis_tarih:
            yeni_filtre = []
            for hat in filtrelenmis_veri:
                if hat['son_tarih_datetime']:
                    hat_tarihi = hat['son_tarih_datetime']
                    hat_tarihi_date = hat_tarihi.date()
                    
                    baslangic_kosul = True
                    bitis_kosul = True
                    
                    if baslangic_tarih:
                        baslangic_kosul = hat_tarihi_date >= baslangic_tarih
                    if bitis_tarih:
                        bitis_kosul = hat_tarihi_date <= bitis_tarih
                    
                    if baslangic_kosul and bitis_kosul:
                        yeni_filtre.append(hat)
                else:
                    yeni_filtre.append(hat)
            
            filtrelenmis_veri = yeni_filtre
        
        return filtrelenmis_veri

    def filtrele(self):
        """Filtreleme yapar ve filtreli verilere göre istatistikleri günceller"""
        # Mevcut filtreli veriyi al
        self.mevcut_filtreli_veri = self.mevcut_filtreli_veriyi_al()
        self.mevcut_filtreli_veri.sort(key=lambda x: x['direk_sayisi'], reverse=True)
        self.siralama_durumu = {}
        self.sutun_basliklarini_guncelle(None, False)
        self.tabloyu_guncelle(self.mevcut_filtreli_veri)
        
        # FİLTRELİ VERİYE GÖRE İSTATİSTİKLERİ YENİDEN HESAPLA
        self.filtreli_istatistikleri_guncelle()

    def filtreli_istatistikleri_guncelle(self):
        """Filtreli verilere göre istatistikleri günceller"""
        if not self.mevcut_filtreli_veri:
            return
        
        # Filtreli verilere göre yeni istatistikleri hesapla
        filtrelenmis_istatistikler = self.filtreli_istatistikleri_hesapla(self.mevcut_filtreli_veri)
        
        # İstatistik kartlarını güncelle
        self.istatistik_kartlari_guncelle(filtrelenmis_istatistikler)
        
        # Grafikleri güncelle
        self.grafikleri_guncelle_filtreli(self.mevcut_filtreli_veri, filtrelenmis_istatistikler)
        
    def filtreli_istatistikleri_hesapla(self, filtreli_veri):
        """Filtreli verilere göre istatistikleri hesaplar"""
        try:
            if not filtreli_veri:
                return {}
            
            # Toplam istatistikleri hesapla
            toplam_ilce = len(set(hat['ilce'] for hat in filtreli_veri))
            toplam_hat = len(filtreli_veri)  # Filtrelenmiş hat sayısı
            toplam_direk = sum(hat['direk_sayisi'] for hat in filtreli_veri)
            toplam_foto = sum(hat['foto_sayisi'] for hat in filtreli_veri)
            toplam_bulgu = sum(hat['bulgu_sayisi'] for hat in filtreli_veri)
            toplam_mesafe = sum(hat.get('toplam_mesafe', 0) for hat in filtreli_veri)
            
            # Ay dağılımı (orijinal Excel'den alınan veriler kullanılamaz, bu yüzden boş bırakıyoruz)
            # Eğer ay verilerini de filtrelemek isterseniz, self.hat_ay_detaylari sözlüğünü kullanmalısınız
            ay_detay = {}
            
            return {
                'toplam_ilce': toplam_ilce,
                'toplam_hat': toplam_hat,
                'toplam_direk': toplam_direk,
                'toplam_foto': toplam_foto,
                'toplam_bulgu': toplam_bulgu,
                'toplam_mesafe': toplam_mesafe,
                'ay_detay': ay_detay,
                'hat_detay': filtreli_veri
            }
            
        except Exception as e:
            print(f"Filtreli istatistik hesaplama hatası: {e}")
            return {}

    def filtreli_istatistikleri_hesapla(self, filtreli_veri):
        """Filtreli verilere göre istatistikleri hesaplar"""
        try:
            if not filtreli_veri:
                return {}
            
            # Toplam istatistikleri hesapla
            toplam_ilce = len(set(hat['ilce'] for hat in filtreli_veri))
            toplam_hat = len(set((hat['ilce'], hat['hat_adi']) for hat in filtreli_veri))
            toplam_direk = sum(hat['direk_sayisi'] for hat in filtreli_veri)
            toplam_foto = sum(hat['foto_sayisi'] for hat in filtreli_veri)
            toplam_bulgu = sum(hat['bulgu_sayisi'] for hat in filtreli_veri)
            toplam_mesafe = sum(hat.get('toplam_mesafe', 0) for hat in filtreli_veri)
            
            # Aylara göre dağılımı hesapla (orijinal Excel'den alınamıyor, bu nedenle mevcut verilerden yola çıkıyoruz)
            # Eğer Excel'den ay bilgisi almak isterseniz, bu kısmı değiştirmeniz gerekir
            ay_detay = {}
            
            # Filtrelenmiş istatistikleri döndür
            return {
                'toplam_ilce': toplam_ilce,
                'toplam_hat': toplam_hat,
                'toplam_direk': toplam_direk,
                'toplam_foto': toplam_foto,
                'toplam_bulgu': toplam_bulgu,
                'toplam_mesafe': toplam_mesafe,
                'ay_detay': ay_detay,
                'hat_detay': filtreli_veri
            }
            
        except Exception as e:
            print(f"Filtreli istatistik hesaplama hatası: {e}")
            return {}
            
    def istatistik_kartlari_guncelle(self, filtrelenmis_istatistikler):
        """İstatistik kartlarını filtreli verilere göre günceller"""
        for widget in self.kartlar_frame.winfo_children():
            widget.destroy()
        
        # Kart verilerini filtreli istatistiklerden al
        toplam_mesafe_km = filtrelenmis_istatistikler.get('toplam_mesafe', 0) / 1000 if filtrelenmis_istatistikler.get('toplam_mesafe', 0) > 0 else 0
        
        kart_verileri = [
            {"baslik": "Toplam İlçe", "deger": filtrelenmis_istatistikler.get('toplam_ilce', 0), "birim": "adet", "renk": "#27ae60", "icon": "🏙️"},
            {"baslik": "Toplam Hat", "deger": filtrelenmis_istatistikler.get('toplam_hat', 0), "birim": "adet", "renk": "#e74c3c", "icon": "🛣️"},
            {"baslik": "Toplam Direk", "deger": filtrelenmis_istatistikler.get('toplam_direk', 0), "birim": "adet", "renk": "#3498db", "icon": "🏗️"},
            {"baslik": "Toplam Fotoğraf", "deger": filtrelenmis_istatistikler.get('toplam_foto', 0), "birim": "adet", "renk": "#f39c12", "icon": "📷"},
            {"baslik": "Toplam Bulgu", "deger": filtrelenmis_istatistikler.get('toplam_bulgu', 0), "birim": "adet", "renk": "#9b59b6", "icon": "🔍"},
            {"baslik": "Toplam Mesafe", "deger": f"{toplam_mesafe_km:.2f}", "birim": "km", "renk": "#1abc9c", "icon": "📏"},
        ]
        
        # Grid layout için satır ve sütun ayarı
        rows = 1
        cols = len(kart_verileri)
        
        for i, kart in enumerate(kart_verileri):
            kart_frame = tk.Frame(self.kartlar_frame, bg=kart['renk'], relief='raised', bd=2)
            kart_frame.grid(row=0, column=i, sticky='nsew', padx=3, pady=2)
            
            # Grid hücrelerinin eşit genişlemesi için
            self.kartlar_frame.grid_columnconfigure(i, weight=1)
            self.kartlar_frame.grid_rowconfigure(0, weight=1)
            
            # Kart frame'inin içeriği için grid
            kart_frame.grid_columnconfigure(0, weight=1)
            kart_frame.grid_rowconfigure(0, weight=1)
            kart_frame.grid_rowconfigure(1, weight=1)
            
            # İkon ve başlık - ÜST SATIR
            top_frame = tk.Frame(kart_frame, bg=kart['renk'])
            top_frame.grid(row=0, column=0, sticky='ew', padx=8, pady=(8, 2))
            
            icon_label = tk.Label(top_frame, text=kart['icon'], font=('Segoe UI', 16),
                                bg=kart['renk'], fg='white')
            icon_label.pack(side='left', padx=(0, 8))
            
            baslik_label = tk.Label(top_frame, text=kart['baslik'], font=('Segoe UI', 11, 'bold'),
                                  bg=kart['renk'], fg='white', wraplength=80, justify='left')
            baslik_label.pack(side='left', fill='x', expand=True)
            
            # Değer ve birim - ALT SATIR
            bottom_frame = tk.Frame(kart_frame, bg=kart['renk'])
            bottom_frame.grid(row=1, column=0, sticky='ew', padx=8, pady=(2, 8))
            
            deger_text = kart['deger']
            
            # Sayısal değerleri formatla
            if isinstance(deger_text, (int, float)):
                if kart['baslik'] == "Toplam Mesafe":
                    deger_text = f"{deger_text:.2f}"
                else:
                    deger_text = f"{deger_text:,}"
            
            deger_label = tk.Label(bottom_frame, text=deger_text, font=('Segoe UI', 18, 'bold'),
                                 bg=kart['renk'], fg='white')
            deger_label.pack(side='left', padx=(0, 8))
            
            birim_label = tk.Label(bottom_frame, text=kart['birim'], font=('Segoe UI', 12, 'bold'),
                                 bg=kart['renk'], fg='white')
            birim_label.pack(side='left')
            
    

        
    
    def tarih_metninden_al(self, tarih_metni):
        if not tarih_metni:
            return None
        
        try:
            return datetime.datetime.strptime(tarih_metni, '%d/%m/%Y').date()
        except:
            return None
    
    def filtreleri_sifirla(self):
        """Filtreleri sıfırlar ve orijinal verilere döner"""
        self.ilce_filter_var.set("TÜMÜ")
        self.baslangic_tarih_var.set("")
        self.bitis_tarih_var.set("")
        self.baslangic_entry.delete(0, tk.END)
        self.bitis_entry.delete(0, tk.END)
        self.siralama_durumu = {}
        self.sutun_basliklarini_guncelle(None, False)
        
        # Orijinal verilere dön
        self.mevcut_filtreli_veri = sorted(self.rapor_verileri['hat_detay'], 
                                         key=lambda x: x['direk_sayisi'], reverse=True)
        self.tabloyu_guncelle(self.mevcut_filtreli_veri)
        
        # Orijinal istatistikleri geri yükle
        self.istatistik_kartlari_olustur()
        self.grafikleri_ciz()

    def grafikleri_ciz(self):
        """Grafikleri çizer - TÜM veriler için (filtreleme olmadan)"""
        # Grafik 1: İlçelere göre direk dağılımı - HARF SIRASINA GÖRE
        self.ax1.clear()
        
        ilce_verileri = {}
        for detay in self.rapor_verileri['hat_detay']:
            ilce = detay['ilce']
            if ilce not in ilce_verileri:
                ilce_verileri[ilce] = 0
            ilce_verileri[ilce] += detay['direk_sayisi']
        
        # İlçeleri harf sırasına göre sırala
        ilce_adlari = sorted(ilce_verileri.keys())
        direk_sayilari = [ilce_verileri[ilce] for ilce in ilce_adlari]
        
        colors = ['#3498db', '#27ae60', '#e74c3c', '#f39c12', '#9b59b6', '#1abc9c', '#34495e', '#d35400']
        
        if ilce_adlari and direk_sayilari:
            bars = self.ax1.bar(ilce_adlari, direk_sayilari, color=colors[:len(ilce_adlari)])
            self.ax1.set_ylabel('Direk Sayısı', fontweight='bold')
            self.ax1.set_title('İlçelere Göre Direk Dağılımı (Tüm Veriler)', fontsize=12, fontweight='bold')
            
            for bar in bars:
                height = bar.get_height()
                self.ax1.text(bar.get_x() + bar.get_width()/2., height,
                             f'{int(height):,}', ha='center', va='bottom', fontweight='bold')
            
            self.ax1.tick_params(axis='x', rotation=45)
            self.fig1.subplots_adjust(top=0.90, bottom=0.20, left=0.12, right=0.95)
        else:
            self.ax1.text(0.5, 0.5, 'İlçe verisi bulunamadı', 
                         ha='center', va='center', transform=self.ax1.transAxes,
                         fontsize=12, color='gray')
            self.fig1.subplots_adjust(top=0.90, bottom=0.20, left=0.12, right=0.95)
        
        self.canvas1.draw()
        
        # Grafik 2: Aylara göre direk sayısı - TÜM veriler için
        self.ax2.clear()
        
        aylar = list(self.rapor_verileri['ay_detay'].keys())
        ay_direkleri = list(self.rapor_verileri['ay_detay'].values())
        
        if aylar and ay_direkleri:
            ay_sirasi = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12']
            sirali_aylar = [ay for ay in ay_sirasi if ay in aylar]
            sirali_degerler = [self.rapor_verileri['ay_detay'][ay] for ay in sirali_aylar]
            
            if sirali_aylar and sirali_degerler:
                ay_isimleri = ['Oca', 'Şub', 'Mar', 'Nis', 'May', 'Haz', 'Tem', 'Ağu', 'Eyl', 'Eki', 'Kas', 'Ara']
                ay_etiketi = []
                for ay in sirali_aylar:
                    try:
                        ay_index = int(ay) - 1
                        if 0 <= ay_index < len(ay_isimleri):
                            ay_etiketi.append(ay_isimleri[ay_index])
                        else:
                            ay_etiketi.append(ay)
                    except:
                        ay_etiketi.append(ay)
                
                bars = self.ax2.bar(ay_etiketi, sirali_degerler, color='#e74c3c', alpha=0.7)
                self.ax2.set_ylabel('Direk Sayısı', fontweight='bold')
                self.ax2.set_title('Aylara Göre Direk Sayısı (Tüm Veriler)', fontsize=12, fontweight='bold')
                
                for i, v in enumerate(sirali_degerler):
                    self.ax2.text(i, v + 0.1, f'{v:,}', ha='center', va='bottom', fontweight='bold')
                
                self.ax2.set_xticks(range(len(ay_etiketi)))
                self.ax2.set_xticklabels(ay_etiketi)
                self.fig2.subplots_adjust(top=0.90, bottom=0.20, left=0.12, right=0.95)
            else:
                self.ax2.text(0.5, 0.5, 'Sıralanmış ay verisi bulunamadı', 
                             ha='center', va='center', transform=self.ax2.transAxes,
                             fontsize=12, color='gray')
                self.fig2.subplots_adjust(top=0.90, bottom=0.20, left=0.12, right=0.95)
        else:
            self.ax2.text(0.5, 0.5, 'Ay verisi bulunamadı', 
                         ha='center', va='center', transform=self.ax2.transAxes,
                         fontsize=12, color='gray')
            self.fig2.subplots_adjust(top=0.90, bottom=0.20, left=0.12, right=0.95)
        
        self.canvas2.draw()

    def grafikleri_guncelle_filtreli(self, filtreli_veri, istatistikler):
        """Grafikleri filtreli verilere göre günceller"""
        # Grafik 1: Filtrelenmiş ilçelere göre direk dağılımı
        self.ax1.clear()
        
        ilce_verileri = {}
        for hat in filtreli_veri:
            ilce = hat['ilce']
            if ilce not in ilce_verileri:
                ilce_verileri[ilce] = 0
            ilce_verileri[ilce] += hat['direk_sayisi']
        
        # İlçeleri harf sırasına göre sırala
        ilce_adlari = sorted(ilce_verileri.keys())
        direk_sayilari = [ilce_verileri[ilce] for ilce in ilce_adlari]
        
        colors = ['#3498db', '#27ae60', '#e74c3c', '#f39c12', '#9b59b6', '#1abc9c', '#34495e', '#d35400']
        
        if ilce_adlari and direk_sayilari:
            bars = self.ax1.bar(ilce_adlari, direk_sayilari, color=colors[:len(ilce_adlari)])
            self.ax1.set_ylabel('Direk Sayısı', fontweight='bold')
            self.ax1.set_title('İlçelere Göre Direk Dağılımı (Filtrelenmiş)', fontsize=12, fontweight='bold')
            
            for bar in bars:
                height = bar.get_height()
                self.ax1.text(bar.get_x() + bar.get_width()/2., height,
                             f'{int(height):,}', ha='center', va='bottom', fontweight='bold')
            
            self.ax1.tick_params(axis='x', rotation=45)
            self.fig1.subplots_adjust(top=0.90, bottom=0.20, left=0.12, right=0.95)
        else:
            self.ax1.text(0.5, 0.5, 'Filtrelenmiş ilçe verisi bulunamadı', 
                         ha='center', va='center', transform=self.ax1.transAxes,
                         fontsize=12, color='gray')
            self.fig1.subplots_adjust(top=0.90, bottom=0.20, left=0.12, right=0.95)
        
        self.canvas1.draw()
        
        # Grafik 2: Filtrelenmiş hatların ay dağılımı
        self.ax2.clear()
        
        # Filtrelenmiş hatların ay dağılımını hesapla
        filtrelenmis_ay_detay = {}
        
        for hat in filtreli_veri:
            hat_anahtari = (hat['ilce'], hat['hat_adi'])
            
            # Bu hat için ay verilerini al
            if hasattr(self, 'hat_ay_detaylari') and hat_anahtari in self.hat_ay_detaylari:
                for ay, sayi in self.hat_ay_detaylari[hat_anahtari].items():
                    if ay not in filtrelenmis_ay_detay:
                        filtrelenmis_ay_detay[ay] = 0
                    filtrelenmis_ay_detay[ay] += sayi
        
        # Eğer filtreli ay verisi yoksa, mesaj göster
        if not filtrelenmis_ay_detay:
            self.ax2.text(0.5, 0.5, 'Filtrelenmiş veriler için\nay bilgisi bulunamadı', 
                         ha='center', va='center', transform=self.ax2.transAxes,
                         fontsize=12, color='gray')
            self.fig2.subplots_adjust(top=0.90, bottom=0.20, left=0.12, right=0.95)
            self.canvas2.draw()
            return
        
        # Ay sıralaması
        ay_sirasi = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12']
        sirali_aylar = [ay for ay in ay_sirasi if ay in filtrelenmis_ay_detay]
        sirali_degerler = [filtrelenmis_ay_detay[ay] for ay in sirali_aylar]
        
        if sirali_aylar and sirali_degerler:
            # Ay isimleri (Türkçe kısaltma)
            ay_isimleri = ['Oca', 'Şub', 'Mar', 'Nis', 'May', 'Haz', 'Tem', 'Ağu', 'Eyl', 'Eki', 'Kas', 'Ara']
            
            # Ay numaralarını isimlere çevir
            ay_etiketi = []
            for ay in sirali_aylar:
                try:
                    ay_index = int(ay) - 1
                    if 0 <= ay_index < len(ay_isimleri):
                        ay_etiketi.append(ay_isimleri[ay_index])
                    else:
                        ay_etiketi.append(ay)
                except:
                    ay_etiketi.append(ay)
            
            # Grafiği çiz
            bars = self.ax2.bar(ay_etiketi, sirali_degerler, color='#e74c3c', alpha=0.7)
            self.ax2.set_ylabel('Direk Sayısı', fontweight='bold')
            self.ax2.set_title('Aylara Göre Direk Sayısı (Filtrelenmiş)', fontsize=12, fontweight='bold')
            
            # Değerleri üste yaz
            for i, v in enumerate(sirali_degerler):
                self.ax2.text(i, v + 0.1, f'{v:,}', ha='center', va='bottom', fontweight='bold')
            
            self.ax2.set_xticks(range(len(ay_etiketi)))
            self.ax2.set_xticklabels(ay_etiketi)
            self.fig2.subplots_adjust(top=0.90, bottom=0.20, left=0.12, right=0.95)
        else:
            self.ax2.text(0.5, 0.5, 'Filtrelenmiş ay verisi bulunamadı', 
                         ha='center', va='center', transform=self.ax2.transAxes,
                         fontsize=12, color='gray')
            self.fig2.subplots_adjust(top=0.90, bottom=0.20, left=0.12, right=0.95)
        
        self.canvas2.draw()


    

    


    def harita_verileri_yukle(self):
        """Harita için verileri yükler"""
        try:
            # Fotoğraf Klasörleri sayfasını oku
            self.harita_df_veri = pd.read_excel(self.excel_dosyasi, sheet_name='Fotoğraf Klasörleri')
            
            # Sütunları kontrol et
            required_cols = ['LAT', 'LON', 'Aob', 'Hat adı', 'Direk No', 'Öncelik', 'Tespit Notu']
            for col in required_cols:
                if col not in self.harita_df_veri.columns:
                    self.harita_bilgi_var.set(f"❌ Eksik sütun: {col}")
                    return
            
            # Koordinatları temizle
            self.harita_df_veri['LAT'] = pd.to_numeric(self.harita_df_veri['LAT'], errors='coerce')
            self.harita_df_veri['LON'] = pd.to_numeric(self.harita_df_veri['LON'], errors='coerce')
            
            # Filtrelenmiş veriyi başlangıçta tüm veri olarak ayarla
            self.harita_df_filtreli = self.harita_df_veri.copy()
            
            # Listeleri doldur
            self.harita_listeleri_doldur()
            
            # İstatistikleri güncelle
            self.harita_istatistikleri_guncelle()
            
            self.harita_bilgi_var.set("✅ Veriler başarıyla yüklendi")
            
        except Exception as e:
            self.harita_bilgi_var.set(f"❌ Veri yükleme hatası: {str(e)}")

    def harita_listeleri_doldur(self):
        """Harita filtresi listelerini doldurur"""
        try:
            # İlçe listesi
            if 'Aob' in self.harita_df_veri.columns:
                ilce_list = ['TÜMÜ'] + sorted(self.harita_df_veri['Aob'].dropna().unique().tolist())
                self.harita_ilce_combo['values'] = ilce_list
            
            # Hat listesi
            if 'Hat adı' in self.harita_df_veri.columns:
                hat_list = ['TÜMÜ'] + sorted(self.harita_df_veri['Hat adı'].dropna().unique().tolist())
                self.harita_hat_combo['values'] = hat_list
                
        except Exception as e:
            print(f"Harita liste doldurma hatası: {e}")

    def harita_istatistikleri_guncelle(self):
        """Harita istatistiklerini günceller"""
        try:
            if self.harita_df_filtreli is not None:
                # Toplam kayıt
                total_points = len(self.harita_df_filtreli)
                self.harita_toplam_kayit_var.set(f"{total_points}")
                
                # Geçerli koordinat
                valid_points = len(self.harita_df_filtreli[self.harita_df_filtreli['LAT'].notna() & self.harita_df_filtreli['LON'].notna()])
                self.harita_gecerli_koordinat_var.set(f"{valid_points}")
                
                # Haritadaki nokta
                self.harita_nokta_var.set(f"{valid_points}")
                
                # Öncelik dağılımı
                if 'Öncelik' in self.harita_df_filtreli.columns:
                    # Çok acil
                    cok_acil_count = self.harita_df_filtreli['Öncelik'].astype(str).str.lower().str.contains('çok acil|cok acil', na=False).sum()
                    self.harita_cok_acil_var.set(f"{cok_acil_count}")
                    
                    # Acil
                    acil_count = self.harita_df_filtreli['Öncelik'].astype(str).str.lower().str.contains('acil', na=False).sum()
                    self.harita_acil_var.set(f"{acil_count}")
                    
                    # Normal
                    normal_count = self.harita_df_filtreli['Öncelik'].astype(str).str.lower().str.contains('normal', na=False).sum()
                    self.harita_normal_var.set(f"{normal_count}")
                    
                    # Bekleyebilir
                    bekleyebilir_count = self.harita_df_filtreli['Öncelik'].astype(str).str.lower().str.contains('bekleyebilir', na=False).sum()
                    self.harita_bekleyebilir_var.set(f"{bekleyebilir_count}")
                
                self.harita_bilgi_var.set(f"📊 Toplam {total_points} kayıt, {valid_points} geçerli koordinat")
                
        except Exception as e:
            print(f"Harita istatistik güncelleme hatası: {e}")

    def harita_filtrele(self):
        """Harita filtrelerini uygular"""
        try:
            if self.harita_df_veri is None:
                self.harita_bilgi_var.set("❌ Veriler yüklenmemiş")
                return
            
            self.harita_bilgi_var.set("🔍 Filtreleme yapılıyor...")
            
            # Filtrelemeyi uygula
            filtered_df = self.harita_df_veri.copy()
            
            # 1. İlçe filtresi
            secili_ilce = self.harita_ilce_var.get()
            if secili_ilce != "TÜMÜ" and secili_ilce:
                filtered_df = filtered_df[filtered_df['Aob'] == secili_ilce]
            
            # 2. Hat filtresi
            secili_hat = self.harita_hat_var.get()
            if secili_hat != "TÜMÜ" and secili_hat:
                filtered_df = filtered_df[filtered_df['Hat adı'] == secili_hat]
            
            # Filtrelenmiş veriyi kaydet
            self.harita_df_filtreli = filtered_df
            
            # İstatistikleri güncelle
            self.harita_istatistikleri_guncelle()
            
            # Haritayı oluştur ve göster
            self.harita_olustur_ve_goster()
            
            self.harita_bilgi_var.set(f"✅ {len(filtered_df)} kayıt filtrelendi")
            
        except Exception as e:
            self.harita_bilgi_var.set(f"❌ Filtreleme hatası: {str(e)}")

    def harita_filtreleri_sifirla(self):
        """Harita filtresini sıfırlar"""
        try:
            self.harita_ilce_var.set("TÜMÜ")
            self.harita_hat_var.set("TÜMÜ")
            
            # Veriyi sıfırla
            self.harita_df_filtreli = self.harita_df_veri.copy()
            self.harita_istatistikleri_guncelle()
            self.harita_olustur_ve_goster()
            
            self.harita_bilgi_var.set("✅ Filtreler sıfırlandı")
            
        except Exception as e:
            self.harita_bilgi_var.set(f"❌ Filtre sıfırlama hatası: {str(e)}")

    def harita_olustur_ve_goster(self):
        """Harita oluşturur ve embed olarak gösterir - MASAÜSTÜ/RAFOR/RaporHarita.html dosyasına kaydeder"""
        try:
            if not hasattr(self, 'harita_df_filtreli') or self.harita_df_filtreli is None or self.harita_df_filtreli.empty:
                self.harita_bilgi_var.set("❌ Harita için veri bulunamadı!")
                return
            
            self.harita_bilgi_var.set("🔄 Harita oluşturuluyor...")
            self.update()
            
            # Bilgi label'ını gizle
            if hasattr(self, 'harita_baslangic_label'):
                self.harita_baslangic_label.pack_forget()
            
            # Koordinat geçerliliği kontrolü
            valid = self.harita_df_filtreli['LAT'].notna() & self.harita_df_filtreli['LON'].notna()
            df_map = self.harita_df_filtreli[valid].copy()
            
            if df_map.empty:
                self.harita_bilgi_var.set("❌ Geçerli koordinatlı veri bulunamadı!")
                return
            
            # Merkez hesapla
            try:
                center_lat = df_map['LAT'].astype(float).mean()
                center_lon = df_map['LON'].astype(float).mean()
            except:
                # Varsayılan merkez (Türkiye)
                center_lat = 39.9334
                center_lon = 32.8597
            
            # Folium harita oluştur
            harita = folium.Map(location=[center_lat, center_lon], 
                              zoom_start=12, 
                              control_scale=True,
                              tiles=None)  # Hiç tile ekleme
            
            # TÜM TILE LAYER'LARI EKLE
            # 1. Normal Harita (OpenStreetMap)
            folium.TileLayer(
                tiles='https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',
                name='🗺️ Normal Harita',
                attr='© OpenStreetMap contributors',
                max_zoom=19,
                control=True
            ).add_to(harita)
            
            # 2. Uydu + Etiketler (Google Hybrid)
            folium.TileLayer(
                tiles='https://mt1.google.com/vt/lyrs=y&x={x}&y={y}&z={z}',
                attr='Google Hybrid',
                name='🛰️🔤 Uydu + Etiketler',
                max_zoom=22,
                control=True
            ).add_to(harita)
            
            # Marker'ları ekle
            eklenen = 0
            
            for idx, row in df_map.iterrows():
                try:
                    lat = float(row['LAT'])
                    lon = float(row['LON'])
                except:
                    continue
                
                # Öncelik rengini belirle
                oncelik = str(row.get('Öncelik', '')).lower()
                renk = 'gray'
                if 'çok acil' in oncelik or 'cok acil' in oncelik:
                    renk = 'red'
                elif 'acil' in oncelik:
                    renk = 'orange'
                elif 'normal' in oncelik:
                    renk = 'blue'
                elif 'bekleyebilir' in oncelik:
                    renk = 'green'
                
                # Popup içeriği
                popup_html = f"""
                <div style='font-family: Arial; width: 260px;'>
                    <h4 style='color: {renk}; margin:0 0 6px 0'>{row.get('Öncelik', 'Belirsiz')}</h4>
                    <b>İlçe:</b> {row.get('Aob', '')}<br>
                    <b>Hat:</b> {row.get('Hat adı', '')}<br>
                    <b>Direk No:</b> {row.get('Direk No', '')}<br>
                    <b>Tespit:</b> {str(row.get('Tespit Notu', ''))[:150]}<br>
                    <b>Koordinat:</b> {lat:.6f}, {lon:.6f}
                </div>
                """
                
                marker = folium.Marker(
                    location=[lat, lon],
                    popup=folium.Popup(popup_html, max_width=320),
                    tooltip=f"{row.get('Hat adı','')} - {row.get('Direk No','')}",
                    icon=folium.Icon(color=renk, icon='info-sign')
                ).add_to(harita)
                eklenen += 1
            
            # Layer kontrol ekle
            folium.LayerControl(collapsed=False, position='topright').add_to(harita)
            
            # HARİTAYI GEÇİCİ DOSYAYA KAYDET
            import tempfile
            temp_file = tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False, encoding='utf-8')
            harita.save(temp_file.name)
            temp_file.close()
            
            self.map_temp_html = temp_file.name
            
            # İstatistikleri güncelle
            self.harita_nokta_var.set(f"{eklenen}")
            self.harita_gecerli_koordinat_var.set(f"{eklenen}")
            
            # Bilgi güncelle
            self.harita_bilgi_var.set(f"✅ Harita hazır: {eklenen} nokta")
            
            # Haritayı gömülü olarak göster
            self.harita_goster_embed(self.map_temp_html)
            
        except Exception as e:
            self.harita_bilgi_var.set(f"❌ Harita oluşturma hatası: {str(e)}")
            import traceback
            traceback.print_exc()

    def harita_html_goster(self, html_string):
        """HTML haritayı gömülü olarak gösterir"""
        try:
            # Önceki içeriği temizle
            if hasattr(self, 'harita_map_frame'):
                for widget in self.harita_map_frame.winfo_children():
                    widget.destroy()
            
            # Basit bir WebView oluşturmak için tkinter Text widget'ını kullan
            html_viewer = tk.Text(self.harita_map_frame, wrap='none')
            html_viewer.insert('1.0', html_string)
            html_viewer.config(state='disabled', bg='white')
            
            scrollbar_y = ttk.Scrollbar(self.harita_map_frame, orient='vertical', command=html_viewer.yview)
            scrollbar_x = ttk.Scrollbar(self.harita_map_frame, orient='horizontal', command=html_viewer.xview)
            html_viewer.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
            
            html_viewer.pack(side='left', fill='both', expand=True)
            scrollbar_y.pack(side='right', fill='y')
            scrollbar_x.pack(side='bottom', fill='x')
            
            # Alternatif olarak, web tarayıcıda açma seçeneği ekleyelim
            button_frame = tk.Frame(self.harita_map_frame, bg='white')
            button_frame.pack(side='bottom', fill='x')
            
            tk.Button(button_frame, text="🌐 Haritayı Tarayıcıda Aç", 
                     command=self.harita_tarayiciya_ac,
                     bg="#3498db", fg='white', font=('Segoe UI', 9)).pack(pady=5)
            
        except Exception as e:
            # Hata durumunda basit mesaj
            error_label = tk.Label(self.harita_map_frame,
                                 text=f"Harita gösterilemedi: {str(e)}\n\n"
                                      "Haritayı tarayıcıda açmak için butona tıklayın.",
                                 font=('Segoe UI', 11),
                                 bg="black", fg="red",
                                 justify='center')
            error_label.pack(expand=True)
            
            button_frame = tk.Frame(self.harita_map_frame, bg='black')
            button_frame.pack(side='bottom', fill='x')
            
            tk.Button(button_frame, text="🌐 Haritayı Tarayıcıda Aç", 
                     command=self.harita_tarayiciya_ac,
                     bg="#3498db", fg='white', font=('Segoe UI', 9)).pack(pady=5)

    def harita_tarayiciya_ac(self):
        """Haritayı tarayıcıda açar"""
        try:
            if not hasattr(self, 'harita_df_filtreli') or self.harita_df_filtreli is None:
                messagebox.showwarning("Uyarı", "Önce harita verilerini yükleyin!")
                return
            
            # Geçici HTML dosyası oluştur
            import tempfile
            import webbrowser
            
            temp_file = tempfile.NamedTemporaryFile(mode='w', suffix='.html', delete=False, encoding='utf-8')
            
            # Koordinat geçerliliği kontrolü
            valid = self.harita_df_filtreli['LAT'].notna() & self.harita_df_filtreli['LON'].notna()
            df_map = self.harita_df_filtreli[valid].copy()
            
            if df_map.empty:
                messagebox.showwarning("Uyarı", "Geçerli koordinatlı veri bulunamadı!")
                return
            
            # Merkez hesapla
            try:
                center_lat = df_map['LAT'].astype(float).mean()
                center_lon = df_map['LON'].astype(float).mean()
            except:
                center_lat = 39.9334
                center_lon = 32.8597
            
            # Folium harita oluştur
            harita = folium.Map(location=[center_lat, center_lon], 
                              zoom_start=12, 
                              control_scale=True)
            
            # Marker'ları ekle
            for idx, row in df_map.iterrows():
                try:
                    lat = float(row['LAT'])
                    lon = float(row['LON'])
                except:
                    continue
                
                # Öncelik rengini belirle
                oncelik = str(row.get('Öncelik', '')).lower()
                renk = 'gray'
                if 'çok acil' in oncelik or 'cok acil' in oncelik:
                    renk = 'red'
                elif 'acil' in oncelik:
                    renk = 'orange'
                elif 'normal' in oncelik:
                    renk = 'blue'
                elif 'bekleyebilir' in oncelik:
                    renk = 'green'
                
                # Popup içeriği
                popup_html = f"""
                <div style='font-family: Arial; width: 260px;'>
                    <h4 style='color: {renk}; margin:0 0 6px 0'>{row.get('Öncelik', 'Belirsiz')}</h4>
                    <b>İlçe:</b> {row.get('Aob', '')}<br>
                    <b>Hat:</b> {row.get('Hat adı', '')}<br>
                    <b>Direk No:</b> {row.get('Direk No', '')}<br>
                    <b>Tespit:</b> {str(row.get('Tespit Notu', ''))[:150]}<br>
                    <b>Koordinat:</b> {lat:.6f}, {lon:.6f}
                </div>
                """
                
                marker = folium.Marker(
                    location=[lat, lon],
                    popup=folium.Popup(popup_html, max_width=320),
                    tooltip=f"{row.get('Hat adı','')} - {row.get('Direk No','')}",
                    icon=folium.Icon(color=renk, icon='info-sign')
                ).add_to(harita)
            
            harita.save(temp_file.name)
            temp_file.close()
            
            # Tarayıcıda aç
            webbrowser.open(f'file://{os.path.abspath(temp_file.name)}')
            self.harita_bilgi_var.set(f"✅ Harita tarayıcıda açıldı: {len(df_map)} nokta")
            
        except Exception as e:
            messagebox.showerror("Hata", f"Harita açılırken hata: {str(e)}")
        

    def harita_listeleri_doldur(self):
        """Harita sekmesindeki combo box'ları doldurur"""
        try:
            if os.path.exists(self.excel_dosyasi):
                df = pd.read_excel(self.excel_dosyasi, sheet_name='Fotoğraf Klasörleri')
                
                # İlçe listesi
                ilce_list = ['TÜMÜ'] + sorted(df['Aob'].dropna().unique().tolist())
                self.harita_ilce_combo['values'] = ilce_list
                
                # Hat listesi
                hat_list = ['TÜMÜ'] + sorted(df['Hat adı'].dropna().unique().tolist())
                self.harita_hat_combo['values'] = hat_list
                
        except Exception as e:
            print(f"Harita listeleri doldurma hatası: {e}")
    
    

    def embed_harita_ekrani(self, harita_ekrani, sol_frame, sag_frame):
        """HaritaAnalizEkrani'nın içeriğini mevcut frame'lere göm"""
        try:
            # HaritaAnalizEkrani'nın tüm child'larını bul
            harita_children = harita_ekrani.winfo_children()
            
            # Her child'ı kontrol et ve uygun frame'e yerleştir
            for child in harita_children:
                child_type = str(type(child))
                
                # Sol frame'e yerleştirilecekler (filtreler, istatistikler)
                if 'Frame' in child_type:
                    # İçindeki widget'lara göre karar ver
                    child_children = child.winfo_children()
                    for subchild in child_children:
                        subchild_type = str(type(subchild))
                        
                        # Filtreler frame'ini bul
                        if 'LabelFrame' in subchild_type and 'Filtreler' in subchild.cget('text'):
                            # Bu frame'i sol_frame'e taşı
                            subchild.pack_forget()
                            subchild.pack(in_=sol_frame, fill='both', expand=True, padx=10, pady=10)
                        
                        # İstatistikler frame'ini bul
                        elif 'LabelFrame' in subchild_type and ('İstatistikler' in subchild.cget('text') or 
                                                               'Öncelik Dağılımı' in subchild.cget('text')):
                            subchild.pack_forget()
                            subchild.pack(in_=sol_frame, fill='both', expand=True, padx=10, pady=10)
                
                # Sağ frame'e yerleştirilecekler (harita)
                elif 'LabelFrame' in child_type and 'HARİTA GÖRÜNÜMÜ' in child.cget('text'):
                    child.pack_forget()
                    child.pack(in_=sag_frame, fill='both', expand=True)
            
            # HaritaAnalizEkrani penceresini gizle (artık gerek yok)
            harita_ekrani.withdraw()
            
        except Exception as e:
            print(f"Harita gömme hatası: {e}")

    def apply_harita_filters(self, harita_ekrani):
        """Harita filtresini uygula"""
        try:
            secili_ilce = self.harita_ilce_var.get()
            secili_hat = self.harita_hat_var.get()
            sadece_oncelikli = self.harita_oncelik_var.get()
            kumelenme = self.harita_kumelenme_var.get()
            
            # Filtreleri uygula
            if hasattr(harita_ekrani, 'ilce_var'):
                harita_ekrani.ilce_var.set(secili_ilce)
            
            if hasattr(harita_ekrani, 'hat_var'):
                harita_ekrani.hat_var.set(secili_hat)
            
            if hasattr(harita_ekrani, 'oncelik_dagilimi_var'):
                harita_ekrani.oncelik_dagilimi_var.set(sadece_oncelikli)
            
            if hasattr(harita_ekrani, 'kumelenme_var'):
                harita_ekrani.kumelenme_var.set(kumelenme)
            
            # Filtreleme yap
            if hasattr(harita_ekrani, 'filtrele'):
                harita_ekrani.filtrele()
                
        except Exception as e:
            print(f"Filtre uygulama hatası: {e}")
    
    def harita_istatistikleri_guncelle(self):
        """Harita istatistiklerini günceller"""
        try:
            if hasattr(self, 'harita_ekrani') and self.harita_ekrani:
                if hasattr(self.harita_ekrani, 'df_filtreli'):
                    df = self.harita_ekrani.df_filtreli
                    if df is not None:
                        toplam = len(df)
                        gecerli = len(df[df['LAT'].notna() & df['LON'].notna()])
                        
                        # Öncelikli sayısı
                        oncelikli = 0
                        if 'Öncelik' in df.columns:
                            oncelikli = df['Öncelik'].astype(str).str.lower().str.contains(
                                'çok acil|cok acil|acil', na=False).sum()
                        
                        self.harita_toplam_var.set(f"Toplam: {toplam}")
                        self.harita_gecerli_var.set(f"Koordinatlı: {gecerli}")
                        self.harita_oncelikli_var.set(f"Öncelikli: {oncelikli}")
        except Exception as e:
            print(f"İstatistik güncelleme hatası: {e}")
    
    def harita_filtreleri_sifirla(self):
        """Harita filtresini sıfırlar"""
        self.harita_ilce_var.set("TÜMÜ")
        self.harita_hat_var.set("TÜMÜ")
        self.harita_oncelik_var.set(False)
        self.harita_kumelenme_var.set(True)        

class AnaUygulama:
    def __init__(self, root):
        self.root = root
        self.root.title("Halil İbrahim Başkaya")
        self.root.geometry("1200x1000")
        self.root.minsize(900, 850)
        self.root.resizable(True, True)
        self.root.configure(bg="#2c3e50")
        
        # Program seçimi
        self.secili_program = tk.StringVar(value="fotograf")
        
        # Diğer değişkenler
        self.calisma_devam_ediyor = False
        self.eklenen_klasor_sayisi = tk.IntVar(value=0)
        self.atlanan_klasor_sayisi = tk.IntVar(value=0)
        self.toplam_klasor_sayisi = tk.IntVar(value=0)
        self.ana_klasor = tk.StringVar()
        self.kaydetme_yeri = tk.StringVar()
        self.ilerleme_yuzde = tk.StringVar(value="%0")
        
        # Threading için
        self.thread_lock = threading.Lock()
        self._sehir_oncesi = {}  # Önbellek
        self._last_update = 0    # UI güncelleme kontrolü
        
        # Drone için değişkenler
        self.toplam_foto_sayisi = tk.IntVar(value=0)

        
        # Excel dosya yolları
        self.direk_musterek_yolu = None
        self.direk_og_yolu = None
        
        # Stil ayarları
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        self.colors = {
            'dark_bg': '#2c3e50', 'light_bg': '#34495e', 'accent': '#3498db',
            'success': '#27ae60', 'danger': '#e74c3c', 'warning': '#f39c12',
            'text': '#ecf0f1', 'border': '#7f8c8d', 'info_bg': '#2c3e50',
            'selected': '#1e8449', 'frame_bg': '#3d566e', 'frame_border': '#1a252f'
        }
        
        self.sehirler = ['DEMRE', 'ELMALI', 'FİNİKE', 'KALKAN', 'KAŞ', 'KEMER', 'KORKUTELİ', 'KUMLUCA']
        
        self.configure_styles()
        self.arayuz_olustur()
        
    def configure_styles(self):
        self.style.configure('Dark.TFrame', background=self.colors['dark_bg'])
        self.style.configure('Light.TFrame', background=self.colors['light_bg'])
        self.style.configure('Info.TFrame', background=self.colors['info_bg'])
        self.style.configure('Frame.TFrame', background=self.colors['frame_bg'])
        
        self.style.configure('Title.TLabel', 
                            font=('Segoe UI', 18, 'bold'),
                            foreground=self.colors['text'],
                            background=self.colors['dark_bg'])
        
        self.style.configure('Subtitle.TLabel', 
                            font=('Segoe UI', 11, 'bold'),
                            foreground=self.colors['accent'],
                            background=self.colors['dark_bg'])
        
        self.style.configure('Light.TLabel', 
                            font=('Segoe UI', 10),
                            foreground=self.colors['text'],
                            background=self.colors['light_bg'])
        
        self.style.configure('Stats.TLabel', 
                            font=('Segoe UI', 11, 'bold'),
                            foreground=self.colors['text'],
                            background=self.colors['info_bg'])
        
        self.style.configure('StatsValue.TLabel', 
                            font=('Segoe UI', 12, 'bold'),
                            foreground=self.colors['accent'],
                            background=self.colors['info_bg'])
        
        self.style.configure('Accent.TButton', 
                            font=('Segoe UI', 10, 'bold'),
                            foreground='white',
                            background=self.colors['accent'],
                            borderwidth=1,
                            focusthickness=3,
                            focuscolor=self.colors['accent'])
        
        self.style.map('Accent.TButton', background=[('active', '#2980b9')])
        
        self.style.configure('Success.TButton', 
                            font=('Segoe UI', 10, 'bold'),
                            foreground='white',
                            background=self.colors['success'],
                            borderwidth=1,
                            focusthickness=3,
                            focuscolor=self.colors['success'])
        
        self.style.map('Success.TButton', background=[('active', '#219a52')])
        
        self.style.configure('Danger.TButton', 
                            font=('Segoe UI', 10, 'bold'),
                            foreground='white',
                            background=self.colors['danger'],
                            borderwidth=1,
                            focusthickness=3,
                            focuscolor=self.colors['danger'])
        
        self.style.map('Danger.TButton', background=[('active', '#c0392b')])
        
        self.style.configure('Custom.TEntry',
                            fieldbackground=self.colors['light_bg'],
                            foreground=self.colors['text'],
                            borderwidth=1,
                            focusthickness=3,
                            focuscolor=self.colors['accent'])
        
        self.style.configure('Custom.Horizontal.TProgressbar',
                            thickness=20,
                            troughcolor=self.colors['light_bg'],
                            background=self.colors['accent'],
                            lightcolor=self.colors['accent'],
                            darkcolor=self.colors['accent'])
        
        self.style.configure('Light.TCheckbutton',
                            background=self.colors['light_bg'],
                            foreground=self.colors['text'],
                            font=('Segoe UI', 10))
        
    def create_3d_frame(self, parent, **kwargs):
        frame = tk.Frame(parent, 
                        bg=self.colors['frame_bg'],
                        relief='ridge', 
                        bd=2,
                        highlightbackground=self.colors['frame_border'],
                        highlightthickness=1,
                        **kwargs)
        return frame
        
    def create_input_frame(self, parent, text, variable, browse_command=None, width=50):
        frame = self.create_3d_frame(parent)
        
        label = tk.Label(frame, text=text, font=('Segoe UI', 10, 'bold'),
                        bg=self.colors['frame_bg'], fg=self.colors['text'],
                        width=15, anchor='w')
        label.pack(side='left', padx=8, pady=6)
        
        entry = ttk.Entry(frame, textvariable=variable, 
                         style='Custom.TEntry', width=width, 
                         font=('Segoe UI', 10))
        entry.pack(side='left', padx=5, fill='x', expand=True)
        
        if browse_command:
            btn = ttk.Button(frame, text="Gözat", 
                           command=browse_command, 
                           style='Accent.TButton', width=10)
            btn.pack(side='left', padx=5)
            
        return frame
        
    def create_progress_frame(self, parent, text):
        frame = self.create_3d_frame(parent)
        
        label = tk.Label(frame, text=text, font=('Segoe UI', 10, 'bold'),
                        bg=self.colors['frame_bg'], fg=self.colors['text'],
                        width=15, anchor='w')
        label.pack(side='left', padx=8, pady=8)
        
        progress = ttk.Progressbar(frame, style='Custom.Horizontal.TProgressbar', 
                                 mode='determinate', length=300)
        progress.pack(side='left', padx=(0, 10), fill='x', expand=True)
        
        percent_label = tk.Label(frame, textvariable=self.ilerleme_yuzde,
                               font=('Segoe UI', 10, 'bold'),
                               bg=self.colors['frame_bg'], fg=self.colors['accent'],
                               width=5)
        percent_label.pack(side='left', padx=(0, 5))
        
        return frame, progress
        
    def arayuz_olustur(self):
        main_container = ttk.Frame(self.root, style='Dark.TFrame', padding=15)
        main_container.pack(fill='both', expand=True)
        
        title_frame = ttk.Frame(main_container, style='Dark.TFrame')
        title_frame.pack(fill='x', pady=(0, 10))
        
        title_label = ttk.Label(title_frame, 
                               text="RAPOR", 
                               style='Title.TLabel')
        title_label.pack()
        
        program_frame = ttk.Frame(title_frame, style='Dark.TFrame')
        program_frame.pack(pady=8)
        


        
        
        self.lbl_aciklama = ttk.Label(title_frame, style='Subtitle.TLabel')
        self.lbl_aciklama.pack(pady=(3, 0))
        
        stats_frame = self.create_3d_frame(main_container)
        stats_frame.pack(fill='x', pady=(0, 10), padx=5)
        
        stats_grid = ttk.Frame(stats_frame, style='Dark.TFrame')
        stats_grid.pack(pady=8, padx=10)
        
        ttk.Label(stats_grid, text="Eklenen Direk Sayısı:", style='Stats.TLabel').grid(row=0, column=0, padx=15)
        ttk.Label(stats_grid, textvariable=self.eklenen_klasor_sayisi, style='StatsValue.TLabel').grid(row=0, column=1, padx=15)
        
        ttk.Label(stats_grid, text="Atlanan Direk Seyısı:", style='Stats.TLabel').grid(row=0, column=2, padx=15)
        ttk.Label(stats_grid, textvariable=self.atlanan_klasor_sayisi, style='StatsValue.TLabel').grid(row=0, column=3, padx=15)
        
        ttk.Label(stats_grid, text="Toplam Direk Sayısı:", style='Stats.TLabel').grid(row=0, column=4, padx=15)
        ttk.Label(stats_grid, textvariable=self.toplam_klasor_sayisi, style='StatsValue.TLabel').grid(row=0, column=5, padx=15)

        
        
        
        self.content_frame = ttk.Frame(main_container, style='Light.TFrame')
        self.content_frame.pack(fill='both', expand=True, padx=8, pady=3)
        
        self.guncelle_icerik_alani()
        
        self.lbl_durum = tk.Label(main_container, text="Hazır", font=('Segoe UI', 12, 'bold'), 
                                 fg=self.colors['success'], bg=self.colors['dark_bg'])
        self.lbl_durum.pack(pady=3)
        
        result_frame = self.create_3d_frame(main_container)
        result_frame.pack(fill='both', expand=True, padx=5, pady=3)
        
        lbl_sonuc = tk.Label(result_frame, text="İşlem Sonuçları:", 
                           font=('Segoe UI', 11, 'bold'),
                           bg=self.colors['frame_bg'], fg=self.colors['text'])
        lbl_sonuc.pack(anchor='w', pady=(8, 4), padx=10)
        
        text_frame = ttk.Frame(result_frame, style='Light.TFrame')
        text_frame.pack(fill='both', expand=True, padx=8, pady=(0, 8))
        
        self.txt_sonuc = tk.Text(text_frame, height=5, width=80,
                                bg=self.colors['light_bg'], fg=self.colors['text'],
                                font=('Consolas', 10), relief='sunken', 
                                borderwidth=2, highlightthickness=1,
                                highlightcolor=self.colors['accent'],
                                highlightbackground=self.colors['border'])
        
        scrollbar = ttk.Scrollbar(text_frame, orient='vertical', command=self.txt_sonuc.yview)
        self.txt_sonuc.configure(yscrollcommand=scrollbar.set)
        
        self.txt_sonuc.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        self.program_degisti()
        
    def program_degisti(self):
        try:
            self.istatistikleri_sifirla()
        except Exception as e:
            # Hata durumunda sadece log yaz, uygulamayı çökertme
            print(f"Program değişimi sırasında hata: {e}")
            self.txt_sonuc.insert(tk.END, f"⚠️ Program değişimi sırasında hata: {e}\n")
        
        # Program değişimi işlemlerine devam et...
        if self.secili_program.get() == "fotograf":
            self.lbl_aciklama.config(text="Klasörlerinizdeki fotoğrafları profesyonel şekilde listeleyin")


        elif self.secili_program.get() == "drone":
            self.lbl_aciklama.config(text="Drone GPS verilerinizi Excel'e aktarın")
            self.foto_radio.config(fg=self.colors['text'])

        elif self.secili_program.get() == "boyutlandirici":
            self.lbl_aciklama.config(text="Fotoğraflarınızı toplu olarak boyutlandırın")
            self.foto_radio.config(fg=self.colors['text'])

        elif self.secili_program.get() == "excel_merge":
            self.lbl_aciklama.config(text="Excel dosyalarınızı birleştirin ve filtreleyin")
            self.foto_radio.config(fg=self.colors['text'])

        else:
            self.lbl_aciklama.config(text="Farklı klasörlerdeki fotoğrafları birleştirin")
            self.foto_radio.config(fg=self.colors['text'])

        
        self.guncelle_icerik_alani()
        
    def guncelle_icerik_alani(self):
        for widget in self.content_frame.winfo_children():
            widget.destroy()
            
        if self.secili_program.get() == "fotograf":
            self.Klasor_Listele_arayuz_olustur()

    
    def Klasor_Listele_arayuz_olustur(self):
        frame_ana_klasor = self.create_input_frame(
            self.content_frame, 
            "Ana Klasör:", 
            self.ana_klasor, 
            self.ana_klasor_sec
        )
        frame_ana_klasor.pack(fill='x', padx=15, pady=8)
        
        frame_kaydetme = self.create_input_frame(
            self.content_frame,
            "Kaydetme Yeri:",
            self.kaydetme_yeri,
            self.kaydetme_yeri_sec
        )
        frame_kaydetme.pack(fill='x', padx=15, pady=8)
        
        # YENİ: Tarihsiz kontrol seçeneği - "Orijinalleri Taşı" gibi tasarım
        tarihsiz_frame = self.create_3d_frame(self.content_frame)
        tarihsiz_frame.pack(fill='x', padx=15, pady=8)
        
        self.tarihsiz_kontrol_var = tk.BooleanVar(value=False)
        
        # "Orijinalleri Taşı" tasarımına uygun checkbox
        self.tarihsiz_check = tk.Checkbutton(tarihsiz_frame, 
                                           text="📊 Tarihsiz Fotoğrafları Rapora Ekle",
                                           variable=self.tarihsiz_kontrol_var,
                                           font=('Segoe UI', 10, 'bold'),
                                           bg=self.colors['frame_bg'],
                                           fg=self.colors['text'],
                                           selectcolor='white',
                                           activebackground=self.colors['frame_bg'],
                                           activeforeground='black',
                                           command=self.tarihsiz_kontrol_degisti)
        self.tarihsiz_check.pack(side='left', padx=15, pady=8)
        
        frame_ilerleme, self.ilerleme = self.create_progress_frame(
            self.content_frame,
            "İlerleme Durumu:"
        )
        frame_ilerleme.pack(fill='x', padx=15, pady=12)
        
        frame_butonlar = self.create_3d_frame(self.content_frame)
        frame_butonlar.pack(pady=15, padx=15)
        
        ttk.Button(frame_butonlar, text="İşlemi Başlat", 
                  command=self.fotograf_klasorleri_listele, 
                  style='Success.TButton', width=15).pack(side='left', padx=12, pady=10)
        
        ttk.Button(frame_butonlar, text="Rapor Göster", 
                  command=self.rapor_goster, 
                  style='Accent.TButton', width=15).pack(side='left', padx=12, pady=10)
        
        ttk.Button(frame_butonlar, text="Çıkış", 
                  command=self.root.destroy, 
                  style='Danger.TButton', width=15).pack(side='left', padx=12, pady=10)
    


    def tarihsiz_kontrol_degisti(self):
        """Tarihsiz kontrol checkbox'ı değiştiğinde renk ve stil güncellemesi"""
        if self.tarihsiz_kontrol_var.get():
            self.tarihsiz_check.config(fg='black')  # Seçiliyse siyah renk
            self.txt_sonuc.insert(tk.END, "ℹ️ Tarihsiz fotoğraf kontrolü AKTİF (Yavaş mod)\n")
        else:
            self.tarihsiz_check.config(fg=self.colors['text'])  # Seçili değilse normal renk
            self.txt_sonuc.insert(tk.END, "ℹ️ Tarihsiz fotoğraf kontrolü PASİF (Hızlı mod)\n")
        self.txt_sonuc.see(tk.END)


    
    def excel_secim_ve_islem_baslat(self):
        """Excel dosyalarını seçmek için dialog aç ve işlemi başlat"""
        if not self.ana_klasor.get():
            messagebox.showerror("Hata", "Lütfen bir ana klasör seçin!")
            return
            
        if not self.kaydetme_yeri.get():
            messagebox.showerror("Hata", "Lütfen kaydetme yeri seçin!")
            return
        
        # Excel seçim dialogunu aç
        excel_dialog = ExcelSecimDialog(self.root)
        self.root.wait_window(excel_dialog)
        
        # Seçilen dosya yollarını al
        self.direk_musterek_yolu = excel_dialog.direk_musterek_yolu
        self.direk_og_yolu = excel_dialog.direk_og_yolu
        
        # Eğer hiç dosya seçilmediyse uyarı ver
        if not self.direk_musterek_yolu and not self.direk_og_yolu:
            messagebox.showwarning("Uyarı", "Hiç Excel dosyası seçilmedi! Direk tipi bilgisi eklenmeyecek.")
        
        # İşlemi başlat
        self.fotograf_klasorleri_listele()
    

    
    def alt_klasor_check_degisti(self):
        if self.alt_klasor_dahil.get():
            self.alt_klasor_check.config(fg='black')
        else:
            self.alt_klasor_check.config(fg=self.colors['text'])
    
    def boyut_check_degisti(self):
        if self.boyutlandir_birlesik.get():
            self.boyut_check.config(fg='black')
            self.entry_boyut.config(state='normal')
        else:
            self.boyut_check.config(fg=self.colors['text'])
            self.entry_boyut.config(state='disabled')
    
    def orijinal_check_degisti(self):
        if self.orijinalleri_tasi.get():
            self.orijinal_check.config(fg='black')
        else:
            self.orijinal_check.config(fg=self.colors['text'])
    
    def boyutlandir_birlesik_toggle(self):
        if self.boyutlandir_birlesik.get():
            self.entry_boyut.config(state='normal')
        else:
            self.entry_boyut.config(state='disabled')
    
    def kaynak_klasor_temizle(self):
        self.kaynak_klasorler.clear()
        self.kaynak_listbox.delete(0, tk.END)

    
    def istatistikleri_sifirla(self):
        self.eklenen_klasor_sayisi.set(0)
        self.atlanan_klasor_sayisi.set(0)
        self.toplam_klasor_sayisi.set(0)
        self.toplam_foto_sayisi.set(0)

        
        # İlerleme çubuğunu güvenli şekilde sıfırla
        try:
            if hasattr(self, 'ilerleme') and self.ilerleme:
                self.ilerleme['value'] = 0
        except tk.TclError:
            # İlerleme çubuğu mevcut değilse, hata verme
            pass
        
        try:
            if hasattr(self, 'excel_ilerleme') and self.excel_ilerleme:
                self.excel_ilerleme['value'] = 0
        except tk.TclError:
            # Excel ilerleme çubuğu mevcut değilse, hata verme
            pass
        
        self.ilerleme_yuzde.set("%0")
        self.txt_sonuc.delete(1.0, tk.END)
        self.lbl_durum.config(text="Hazır", fg=self.colors['success'])
        self._sehir_oncesi.clear()

        
    def log_mesaj_ekle(self, mesaj):
        """İşlem sonuçlarına zaman damgalı mesaj ekler"""
        timestamp = datetime.datetime.now().strftime("[%H:%M:%S]")
        log_mesaj = f"{timestamp} {mesaj}\n"
        self.txt_sonuc.insert(tk.END, log_mesaj)
        self.txt_sonuc.see(tk.END)
        self.root.update()  # self.update() yerine self.root.update()


    
    def ana_klasor_sec(self):
        klasor = filedialog.askdirectory(title="Ana Klasörü Seçin")
        if klasor:
            self.ana_klasor.set(klasor)
            
    def klasor_sec(self, degisken):
        klasor = filedialog.askdirectory(title="Klasör Seçin")
        if klasor:
            degisken.set(klasor)
            
    def kaydetme_yeri_sec(self):
        baslangic_yolu = os.path.join(os.path.expanduser("~"), "Documents")
        if not os.path.exists(baslangic_yolu):
            baslangic_yolu = os.path.expanduser("~")
            
        dosya = filedialog.asksaveasfilename(
            initialdir=baslangic_yolu,
            title="Excel Dosyasını Kaydet",
            defaultextension=".xlsx",
            filetypes=[("Excel dosyaları", "*.xlsx"), ("Tüm dosyalar", "*.*")]
        )
        if dosya:
            self.kaydetme_yeri.set(dosya)

    
    def kaynak_klasor_sil(self):
        secili_indeksler = self.kaynak_listbox.curselection()
        for indeks in reversed(secili_indeksler):
            self.kaynak_klasorler.pop(indeks)
            self.kaynak_listbox.delete(indeks)

    


    def fotograf_klasorleri_listele(self):
        if not self.ana_klasor.get():
            messagebox.showerror("Hata", "Lütfen bir ana klasör seçin!")
            return
            
        if not self.kaydetme_yeri.get():
            messagebox.showerror("Hata", "Lütfen kaydetme yeri seçin!")
            return
        
        # 1. ÖNCE MASADÜSTÜNDEKİ DRONE KLASÖRÜNDEN TÜM EXCEL DOSYALARINI OTOMATİK BUL
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        drone_klasoru = os.path.join(desktop_path, "Drone")
        
        # Excel dosya adlarını tanımla - NİHAİ_DİREK İLK SIRADA
        nihai_dosya_adlari = ["Nihai_Direk.xlsx", "Nihai_Direk.xls", "NihaiDirek.xlsx", 
                             "nihai_direk.xlsx", "NİHAİ_DİREK.xlsx", "nihai.xlsx"]
        
        musterek_dosya_adlari = ["direk_musterek.xlsx", "direk_musterek.xls", 
                                "musterek_direk.xlsx", "musterek_direk.xls",
                                "müşterek_direk.xlsx", "müşterek_direk.xls"]
        
        og_dosya_adlari = ["direk_og.xlsx", "direk_og.xls", "og_direk.xlsx", 
                          "og_direk.xls", "OG_Direk.xlsx", "OG_Direk.xls"]
        
        # Dosyaları bul - ÖNCE NİHAİ_DİREK
        self.nihai_direk_yolu = None
        self.direk_musterek_yolu = None
        self.direk_og_yolu = None
        
        if os.path.exists(drone_klasoru):
            self.log_mesaj_ekle(f"🔍 Masaüstü Drone klasörü taranıyor: {drone_klasoru}")
            
            # 1. NİHAİ DİREK EXCEL'İNİ ARA
            for dosya_adi in nihai_dosya_adlari:
                potansiyel_yol = os.path.join(drone_klasoru, dosya_adi)
                if os.path.exists(potansiyel_yol):
                    self.nihai_direk_yolu = potansiyel_yol
                    self.log_mesaj_ekle(f"🏆 Nihai Direk Excel'i bulundu: {dosya_adi}")
                    break
            
            # 2. MÜŞTEREK DİREK EXCEL'İNİ ARA
            for dosya_adi in musterek_dosya_adlari:
                potansiyel_yol = os.path.join(drone_klasoru, dosya_adi)
                if os.path.exists(potansiyel_yol):
                    self.direk_musterek_yolu = potansiyel_yol
                    self.log_mesaj_ekle(f"✅ Müşterek direk Excel'i bulundu: {dosya_adi}")
                    break
            
            # 3. OG DİREK EXCEL'İNİ ARA
            for dosya_adi in og_dosya_adlari:
                potansiyel_yol = os.path.join(drone_klasoru, dosya_adi)
                if os.path.exists(potansiyel_yol):
                    self.direk_og_yolu = potansiyel_yol
                    self.log_mesaj_ekle(f"✅ OG direk Excel'i bulundu: {dosya_adi}")
                    break
            
            # Eğer dosyalar bulunamadıysa, klasördeki tüm Excel dosyalarını ara
            excel_dosyalari = []
            for dosya in os.listdir(drone_klasoru):
                if dosya.lower().endswith(('.xlsx', '.xls')):
                    tam_yol = os.path.join(drone_klasoru, dosya)
                    excel_dosyalari.append((dosya, tam_yol))
            
            # Bulunan Excel dosyalarını analiz et
            for dosya_adi, dosya_yolu in excel_dosyalari:
                dosya_adi_lower = dosya_adi.lower()
                
                # Nihai_Direk bulunamadıysa ara
                if not self.nihai_direk_yolu:
                    if any(keyword in dosya_adi_lower for keyword in ['nihai', 'nihai_direk', 'nihaidirek', 'final', 'son']):
                        self.nihai_direk_yolu = dosya_yolu
                        self.log_mesaj_ekle(f"🏆 Nihai Direk Excel'i bulundu: {dosya_adi}")
                
                # Müşterek bulunamadıysa ara
                if not self.direk_musterek_yolu:
                    if any(keyword in dosya_adi_lower for keyword in ['musterek', 'müşterek', 'ortak']):
                        self.direk_musterek_yolu = dosya_yolu
                        self.log_mesaj_ekle(f"✅ Müşterek direk Excel'i bulundu: {dosya_adi}")
                
                # OG bulunamadıysa ara
                if not self.direk_og_yolu:
                    if any(keyword in dosya_adi_lower for keyword in ['og', 'og_direk', 'ogdirek', 'trafo']):
                        self.direk_og_yolu = dosya_yolu
                        self.log_mesaj_ekle(f"✅ OG direk Excel'i bulundu: {dosya_adi}")
        
        # 2. EĞER OTOMATİK BULUNAMADIYSA, KULLANICIYA SOR
        bulunan_dosyalar = []
        if self.nihai_direk_yolu:
            bulunan_dosyalar.append(f"Nihai: {os.path.basename(self.nihai_direk_yolu)}")
        if self.direk_musterek_yolu:
            bulunan_dosyalar.append(f"Müşterek: {os.path.basename(self.direk_musterek_yolu)}")
        if self.direk_og_yolu:
            bulunan_dosyalar.append(f"OG: {os.path.basename(self.direk_og_yolu)}")
        
        if bulunan_dosyalar:
            self.log_mesaj_ekle(f"✅ Excel dosyaları otomatik bulundu: {', '.join(bulunan_dosyalar)}")
        else:
            self.log_mesaj_ekle("⚠️ Excel dosyaları otomatik bulunamadı, manuel seçime geçiliyor...")
            
            # Excel seçim dialogunu aç (Nihai_Direk dahil)
            excel_dialog = ExcelSecimDialog(self.root, nihai_direk_eklendi=True)
            self.root.wait_window(excel_dialog)
            
            # Seçilen dosya yollarını al
            self.nihai_direk_yolu = excel_dialog.nihai_direk_yolu
            self.direk_musterek_yolu = excel_dialog.direk_musterek_yolu
            self.direk_og_yolu = excel_dialog.direk_og_yolu
        
        # 3. RAPOR İŞLEMİNİ BAŞLAT
        self.calisma_devam_ediyor = True
        self.lbl_durum.config(text="Rapor oluşturma işlemi başlatılıyor...", fg=self.colors['accent'])
        self.txt_sonuc.delete(1.0, tk.END)
        self.txt_sonuc.insert(tk.END, "Rapor oluşturma işlemi başlatılıyor...\n")
        self.ilerleme['value'] = 0
        self.ilerleme_yuzde.set("%0")
        self.root.update()
        
        # Direkt rapor işlemini başlat
        threading.Thread(target=self.fotograf_klasorleri_listele_thread_optimized, daemon=True).start()

    
    def tum_yol_ayiricilari_duzelt(self, yol):
        if yol:
            yol = yol.replace('/', '\\')
            yol = os.path.normpath(yol)
        return yol
    
    def fotograf_klasorleri_listele_thread_optimized(self):
        try:
            # EXCEL DOSYA DURUMUNU LOGLA - BURAYA EKLİYORUZ
            self.log_mesaj_ekle("📊 Excel dosya durumu kontrol ediliyor...")
            
            if hasattr(self, 'direk_musterek_yolu') and self.direk_musterek_yolu:
                self.log_mesaj_ekle(f"✅ Müşterek Excel: {os.path.basename(self.direk_musterek_yolu)}")
            else:
                self.log_mesaj_ekle("⚠️ Müşterek Excel dosyası bulunamadı - Direk tipi bilgisi eklenmeyecek")
                
            if hasattr(self, 'direk_og_yolu') and self.direk_og_yolu:
                self.log_mesaj_ekle(f"✅ OG Excel: {os.path.basename(self.direk_og_yolu)}")
            else:
                self.log_mesaj_ekle("⚠️ OG Excel dosyası bulunamadı - Direk tipi bilgisi eklenmeyecek")
            
            start_time = time.time()
            self.log_mesaj_ekle("🔄 Hızlı tarama başlatıldı...")
            self.throttled_update()
            
            # KULLANICI SEÇENEĞİ - YENİ EKLENDİ
            tarihsiz_tara = self.tarihsiz_kontrol_var.get()
            
            # 1. AŞAMA: Klasör tarama
            self.log_mesaj_ekle("📁 Klasörler taranıyor...")
            self.throttled_update()
            tum_klasorler = self.hizli_klasor_tarama(self.ana_klasor.get())
            tarama_suresi = time.time() - start_time
            
            self.log_mesaj_ekle(f"✅ {len(tum_klasorler)} klasör {tarama_suresi:.2f}s'de bulundu")
            self.throttled_update()
            
            if not tum_klasorler:
                self.log_mesaj_ekle("❌ Hiç fotoğraf klasörü bulunamadı!")
                return
            
            # 2. AŞAMA: Klasör işleme
            self.log_mesaj_ekle("🔍 Klasör bilgileri işleniyor...")
            
            if tarihsiz_tara:
                self.log_mesaj_ekle("⚡ Tüm fotoğraflar tarih kontrolü yapılacak (Yavaş mod)")
            else:
                self.log_mesaj_ekle("⚡ Sadece çekim tarihleri alınacak (Hızlı mod)")
            
            self.log_mesaj_ekle(f"⚡ {self.optimal_thread_sayisi()} thread ile paralel işlem")
            self.throttled_update()
            
            foto_klasorleri = []
            atlanan_klasorler = []
            hat_tarihleri = {}
            
            max_workers = self.optimal_thread_sayisi()
            
            with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
                future_to_klasor = {
                    executor.submit(self.klasor_isle, klasor, tarihsiz_tara): klasor 
                    for klasor in tum_klasorler
                }
                
                completed = 0
                for future in concurrent.futures.as_completed(future_to_klasor):
                    completed += 1
                    try:
                        result = future.result(timeout=15)
                        if result:
                            klasor_veri, atlama_veri = result
                            if klasor_veri:
                                foto_klasorleri.append(klasor_veri)
                                
                                # Hat bazlı tarih analizi
                                hat_anahtari = (klasor_veri['Aob'], klasor_veri['Hat adı'])
                                if hat_anahtari not in hat_tarihleri:
                                    hat_tarihleri[hat_anahtari] = {
                                        'ilk_tarih': None,
                                        'son_tarih': None,
                                        'tarihler': []
                                    }
                                
                                # Tarih bilgisini datetime'a çevir
                                try:
                                    tarih_str = klasor_veri['En Yeni Fotoğraf Tarihi']
                                    if tarih_str != 'Tarih alınamadı' and tarih_str != 'Tarihi Yok':
                                        tarih_obj = datetime.datetime.strptime(tarih_str, '%d/%m/%Y %H:%M:%S')
                                        hat_tarihleri[hat_anahtari]['tarihler'].append(tarih_obj)
                                except:
                                    pass
                                
                                with self.thread_lock:
                                    self.eklenen_klasor_sayisi.set(len(foto_klasorleri))
                                    
                            elif atlama_veri:
                                atlanan_klasorler.append(atlama_veri)
                                with self.thread_lock:
                                    self.atlanan_klasor_sayisi.set(len(atlanan_klasorler))
                    
                    except concurrent.futures.TimeoutError:
                        klasor_yolu = future_to_klasor[future]
                        self.log_mesaj_ekle(f"⏰ Timeout: {os.path.basename(klasor_yolu)}")
                    except Exception as e:
                        klasor_yolu = future_to_klasor[future]
                        self.log_mesaj_ekle(f"⚠️ Hata: {os.path.basename(klasor_yolu)} - {str(e)}")
                    
                    ilerleme = (completed / len(tum_klasorler)) * 100
                    with self.thread_lock:
                        self.ilerleme['value'] = ilerleme
                        self.ilerleme_yuzde.set(f"%{int(ilerleme)}")
                        self.toplam_klasor_sayisi.set(completed)
                    
                    if completed % 25 == 0:
                        self.throttled_update()
            
            # YENİ: KLASÖRLERİ SIRALA - Hat bazında ve sıra numarasına göre
            self.log_mesaj_ekle("🔢 Klasörler sıralanıyor...")
            self.throttled_update()
            foto_klasorleri = self.klasorleri_sirala(foto_klasorleri)
            
            # 3. AŞAMA: Hat sürelerini hesaplama
            self.log_mesaj_ekle("📅 Hat süreleri hesaplanıyor...")
            self.throttled_update()
            hat_sureleri = self.hat_surelerini_hesapla(hat_tarihleri)
            
            # 4. AŞAMA: Direk tipi bilgilerini alma
            self.log_mesaj_ekle("🏷️ Direk tipi bilgileri alınıyor...")
            self.throttled_update()
            direk_tipleri = self.direk_tipi_bilgilerini_al(foto_klasorleri)
            
            # 5. AŞAMA: Tarihsiz istatistikleri hesaplama (SADECE SEÇİLİRSE)
            hat_tarihsiz_istatistikleri = {}
            if tarihsiz_tara:
                self.log_mesaj_ekle("📊 Tarihsiz fotoğraf istatistikleri hesaplanıyor...")
                self.throttled_update()
                hat_tarihsiz_istatistikleri = self.hat_bazi_tarihsiz_hesapla(foto_klasorleri)
            else:
                self.log_mesaj_ekle("⏩ Tarihsiz kontrol atlanıyor (hızlı mod)")
                self.throttled_update()
                # Boş istatistik gönder
                hat_tarihsiz_istatistikleri = {}
            
            self.throttled_update()
            
            # 6. AŞAMA: Excel oluşturma
            self.log_mesaj_ekle("📊 Excel dosyası oluşturuluyor...")
            self.throttled_update()
            # Tarihsiz istatistiklerini parametre olarak gönderiyoruz
            self.excel_islemi_tamamla(foto_klasorleri, atlanan_klasorler, hat_sureleri, hat_tarihsiz_istatistikleri)
            
        except Exception as e:
            self.lbl_durum.config(text="Hata oluştu", fg=self.colors['danger'])
            self.log_mesaj_ekle(f"❌ Beklenmeyen hata: {str(e)}")
            import traceback
            self.log_mesaj_ekle(f"🔍 Hata detayı: {traceback.format_exc()}")



    def klasorleri_sirala(self, foto_klasorleri):
        """Klasörleri hat bazında ve sıra numarasına göre sıralar"""
        try:
            # Önce hat bazında grupla
            hat_gruplari = {}
            
            for klasor in foto_klasorleri:
                hat_anahtari = (klasor['Aob'], klasor['Hat adı'])
                
                if hat_anahtari not in hat_gruplari:
                    hat_gruplari[hat_anahtari] = []
                
                hat_gruplari[hat_anahtari].append(klasor)
            
            # Her hat grubunu sıra numarasına göre sırala
            sirali_klasorler = []
            
            for hat_anahtari, klasor_listesi in hat_gruplari.items():
                # Sıra numarasına göre sırala
                sirali_hat_klasorleri = sorted(klasor_listesi, 
                                             key=lambda x: self.sira_no_ile_sirala(x['Orijinal Klasör Adı']))
                sirali_klasorler.extend(sirali_hat_klasorleri)
                
                # Debug info
                self.log_mesaj_ekle(f"🔢 {hat_anahtari[0]} - {hat_anahtari[1]}: {len(sirali_hat_klasorleri)} klasör sıralandı")
            
            self.log_mesaj_ekle(f"✅ Toplam {len(sirali_klasorler)} klasör sıralandı")
            return sirali_klasorler
            
        except Exception as e:
            self.log_mesaj_ekle(f"⚠️ Sıralama hatası: {str(e)}")
            # Hata durumunda orijinal listeyi döndür
            return foto_klasorleri

    def sira_no_ile_sirala(self, orijinal_klasor_adi):
        """Sıralama için sıra numarasını çıkarır (sayısal karşılaştırma için)"""
        try:
            sira_no_str = self.sira_no_ayikla(orijinal_klasor_adi)
            if sira_no_str and sira_no_str.isdigit():
                return int(sira_no_str)
            else:
                # Sıra numarası yoksa veya geçersizse, çok büyük bir sayı döndürerek sona at
                return 999999
        except:
            return 999999


    def hat_surelerini_hesapla(self, hat_tarihleri):
        """Hat bazında sadece çekim yapılan günleri hesaplar"""
        hat_sureleri = {}
        
        for hat_anahtari, tarih_bilgileri in hat_tarihleri.items():
            if tarih_bilgileri['tarihler']:
                # Tüm tarihleri al ve sadece tarih kısmını (gün bazında) al
                tum_tarihler = tarih_bilgileri['tarihler']
                
                # Benzersiz çekim günlerini bul (sadece tarih, saat yok)
                cekim_gunleri = set()
                for tarih in tum_tarihler:
                    gun_bazinda = tarih.date()  # Sadece tarih kısmını al, saati at
                    cekim_gunleri.add(gun_bazinda)
                
                # Benzersiz çekim günlerini sırala
                sirali_gunler = sorted(cekim_gunleri)
                
                if sirali_gunler:
                    ilk_gun = sirali_gunler[0]
                    son_gun = sirali_gunler[-1]
                    
                    # TOPLAM İŞ GÜNÜ = Benzersiz çekim günü sayısı
                    toplam_is_gunu = len(cekim_gunleri)
                    
                    hat_sureleri[hat_anahtari] = {
                        'ilk_tarih': ilk_gun.strftime('%d/%m/%Y'),
                        'son_tarih': son_gun.strftime('%d/%m/%Y'),
                        'toplam_gun': toplam_is_gunu,
                        'benzersiz_cekim_gunleri': len(cekim_gunleri),
                        'tum_tarih_araligi': f"{ilk_gun.strftime('%d/%m/%Y')} - {son_gun.strftime('%d/%m/%Y')}",
                        'tarih_araligi_gun_sayisi': (son_gun - ilk_gun).days + 1
                    }
                else:
                    hat_sureleri[hat_anahtari] = {
                        'ilk_tarih': 'Tarihi Yok',
                        'son_tarih': 'Tarihi Yok', 
                        'toplam_gun': 0,
                        'benzersiz_cekim_gunleri': 0,
                        'tarih_araligi': 'Tarihi Yok'
                    }
            else:
                hat_sureleri[hat_anahtari] = {
                    'ilk_tarih': 'Tarihi Yok',
                    'son_tarih': 'Tarihi Yok', 
                    'toplam_gun': 0,
                    'benzersiz_cekim_gunleri': 0,
                    'tarih_araligi': 'Tarihi Yok'
                }
        
        return hat_sureleri
    
    def hizli_klasor_tarama(self, ana_klasor):
        foto_uzantilari = ('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.webp')
        foto_klasorleri = []
        
        for root, dirs, files in os.walk(ana_klasor):
            for dosya in files:
                if any(dosya.lower().endswith(ext) for ext in foto_uzantilari):
                    foto_klasorleri.append(root)
                    break
        
        return foto_klasorleri
    
    def optimal_thread_sayisi(self):
        try:
            cpu_sayisi = os.cpu_count() or 1
            return min(cpu_sayisi, 4)
        except:
            return 2
    
    def throttled_update(self):
        current_time = time.time()
        if current_time - self._last_update > 0.1:
            self.root.update()  # Burada da self.root.update() kullan
            self._last_update = current_time
            # Text widget'ını her güncellemede en sona kaydır
            self.txt_sonuc.see(tk.END)
    
    def exif_tarihini_al(self, dosya_yolu):
        try:
            with Image.open(dosya_yolu) as img:
                exif_data = img.getexif()
                if exif_data:
                    # Öncelikle çekilme tarihini ara (EXIF tag 36867 veya 306)
                    datetime_original = exif_data.get(36867)  # DateTimeOriginal
                    if not datetime_original:
                        datetime_original = exif_data.get(306)  # DateTime
                    
                    if datetime_original:
                        # EXIF tarih formatı: '2025:07:20 12:30:45'
                        try:
                            return datetime.datetime.strptime(datetime_original, '%Y:%m:%d %H:%M:%S')
                        except ValueError:
                            # Bazı format farklılıkları olabilir
                            try:
                                return datetime.datetime.strptime(datetime_original, '%Y-%m-%d %H:%M:%S')
                            except:
                                pass
                
                # EXIF yoksa veya okunamazsa None döndür
                return None
                
        except Exception as e:
            print(f"EXIF okuma hatası {dosya_yolu}: {e}")
            return None

    def klasor_isle(self, klasor_yolu, tarihsiz_tara=False):
        klasor_adi = os.path.basename(klasor_yolu)
        foto_uzantilari = ('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.webp')
        
        try:
            try:
                dosyalar = os.listdir(klasor_yolu)
            except (PermissionError, OSError) as e:
                return None, {
                    'Klasör Yolu': self.tum_yol_ayiricilari_duzelt(klasor_yolu),
                    'Direk No': klasor_adi,
                    'Sebep': f'Erişim hatası: {str(e)}'
                }
            
            rakam_kismi = self.klasor_adi_kontrol(klasor_adi)
            
            if rakam_kismi is None or rakam_kismi == "0":
                return None, {
                    'Klasör Yolu': self.tum_yol_ayiricilari_duzelt(klasor_yolu),
                    'Direk No': klasor_adi,
                    'Sebep': 'Uygun format değil' if rakam_kismi is None else 'Direk No 0 olamaz'
                }
            
            foto_sayisi = 0
            son_foto_tarihi = None
            
            # Klasördeki fotoğrafları sırala ve son fotoğrafı bul
            foto_dosyalari = []
            for dosya in dosyalar:
                if any(dosya.lower().endswith(ext) for ext in foto_uzantilari):
                    foto_sayisi += 1
                    foto_dosyalari.append(dosya)
            
            # Fotoğrafları sırala (alfabetik olarak sonuncuyu al)
            if foto_dosyalari:
                foto_dosyalari.sort()
                son_foto = foto_dosyalari[-1]  # En son sıradaki fotoğraf
                dosya_yolu = os.path.join(klasor_yolu, son_foto)
                
                # Önce EXIF'ten çekilme tarihini al
                son_foto_tarihi = self.exif_tarihini_al(dosya_yolu)
                
                # Eğer EXIF tarihi yoksa, "Tarihi Yok" olarak işaretle
                if son_foto_tarihi is None:
                    son_foto_tarihi = None  # "Tarihi Yok" olarak işaretle
            
            # Tarih bilgisini işle
            if son_foto_tarihi is not None:
                tarih_bilgileri = self.tarihi_ayristir(son_foto_tarihi)
            else:
                # Tarih yoksa "Tarihi Yok" olarak işaretle
                tarih_bilgileri = {
                    'tam_tarih': 'Tarihi Yok',
                    'yil': 'Tarihi Yok',
                    'ay': 'Tarihi Yok', 
                    'gun': 'Tarihi Yok'
                }
            
            aob_adi = self.aob_adi_bul(klasor_yolu)
            hat_adi = self.bir_ust_klasor_adi_al(klasor_yolu)
            
            try:
                klasor_adi_sayi = int(rakam_kismi)
            except ValueError:
                klasor_adi_sayi = rakam_kismi
            
            mesaj = f"✓ {aob_adi} -> {hat_adi} -> {klasor_adi_sayi} ({foto_sayisi} fotoğraf, Tarih: {tarih_bilgileri['tam_tarih']})"
            with self.thread_lock:
                self.log_mesaj_ekle(mesaj)
            
            return {
                'Klasör Yolu': self.tum_yol_ayiricilari_duzelt(klasor_yolu),
                'Aob': aob_adi,
                'Hat adı': hat_adi,
                'Direk No': klasor_adi_sayi,
                'Fotoğraf Sayısı': foto_sayisi,
                'Orijinal Klasör Adı': klasor_adi,
                'Gün': tarih_bilgileri['gun'],
                'Ay': tarih_bilgileri['ay'],
                'Yıl': tarih_bilgileri['yil'],
                'En Yeni Fotoğraf Tarihi': tarih_bilgileri['tam_tarih']
            }, None
            
        except Exception as e:
            return None, {
                'Klasör Yolu': self.tum_yol_ayiricilari_duzelt(klasor_yolu),
                'Direk No': klasor_adi,
                'Sebep': f'İşlem hatası: {str(e)}'
            }
    
    def tarihi_ayristir(self, tarih_objesi):
        if tarih_objesi is None:
            return {
                'tam_tarih': 'Tarihi Yok',
                'yil': 'Tarihi Yok',
                'ay': 'Tarihi Yok', 
                'gun': 'Tarihi Yok'
            }
        
        try:
            return {
                'tam_tarih': tarih_objesi.strftime('%d/%m/%Y %H:%M:%S'),
                'yil': str(tarih_objesi.year),
                'ay': str(tarih_objesi.month).zfill(2),
                'gun': str(tarih_objesi.day).zfill(2)
            }
        except Exception as e:
            return {
                'tam_tarih': 'Tarihi Yok',
                'yil': 'Tarihi Yok',
                'ay': 'Tarihi Yok',
                'gun': 'Tarihi Yok'
            }
    
    def aob_adi_bul(self, klasor_yolu):
        if klasor_yolu in self._sehir_oncesi:
            return self._sehir_oncesi[klasor_yolu]
            
        parts = klasor_yolu.split(os.sep)
        for part in parts:
            temiz_part = part.strip().upper()
            for sehir in self.sehirler:
                if sehir in temiz_part:
                    self._sehir_oncesi[klasor_yolu] = sehir
                    return sehir
        
        result = parts[-1] if parts else "Bilinmiyor"
        self._sehir_oncesi[klasor_yolu] = result
        return result
    
    def bir_ust_klasor_adi_al(self, klasor_yolu):
        if klasor_yolu:
            ust_klasor_yolu = os.path.dirname(klasor_yolu)
            return os.path.basename(ust_klasor_yolu)
        return ""
    
    def klasor_adi_kontrol(self, klasor_adi):
        """
        Revize edilmiş klasör adı kontrol fonksiyonu
        Yeni format: "74 - a123456" → "123456" desteği eklendi
        """
        
        # ÖNCE YENİ FORMAT: "74 - a123456", "74 - b123457" (SADECE TEK HARF)
        try:
            if ' - ' in klasor_adi:
                parts = klasor_adi.split(' - ', 1)
                if len(parts) == 2:
                    sıra_kısmı = parts[0].strip()      # "74"
                    direk_kısmı = parts[1].strip()     # "a123456", "b123457", vb.
                    
                    # Sıra kısmı sayı mı kontrol et
                    if sıra_kısmı and sıra_kısmı.isdigit():
                        # Direk kısmı: TAM OLARAK 1 HARF + RAKAMLAR olmalı
                        if (len(direk_kısmı) >= 2 and 
                            direk_kısmı[0].isalpha() and 
                            len(direk_kısmı[0]) == 1 and  # SADECE 1 HARF
                            direk_kısmı[1:].isdigit()):   # GERİ KALANI SADECE RAKAM
                            
                            rakam_kismi = direk_kısmı[1:]  # "123456"
                            # Baştaki sıfırları temizle, "0" değerini koru
                            rakam_kismi = rakam_kismi.lstrip('0') or '0'
                            if rakam_kismi != "0":
                                return rakam_kismi
        except:
            pass
        
        # SONRA ORİJİNAL FORMATLARI KONTROL ET
        if klasor_adi.count('-') >= 2:
            parts = klasor_adi.split('-', 2)
            if len(parts) == 3:
                rakam_kismi = parts[1].strip()
                rakam_kismi = ''.join(filter(str.isdigit, rakam_kismi))
                if rakam_kismi and rakam_kismi != "0":
                    return rakam_kismi
        
        elif ' - ' in klasor_adi:
            parts = klasor_adi.split(' - ', 1)
            if len(parts) == 2:
                rakam_kismi = parts[1].strip()
                if not rakam_kismi or not rakam_kismi[0].isdigit():
                    return None
                rakamlar = []
                for char in rakam_kismi:
                    if char.isdigit():
                        rakamlar.append(char)
                    elif rakamlar:
                        break
                rakam_kismi = ''.join(rakamlar)
                if rakam_kismi and rakam_kismi != "0":
                    return rakam_kismi
        
        elif '-' in klasor_adi:
            parts = klasor_adi.split('-', 1)
            if len(parts) == 2:
                rakam_kismi = parts[1].strip()
                if not rakam_kismi or not rakam_kismi[0].isdigit():
                    return None
                rakamlar = []
                for char in rakam_kismi:
                    if char.isdigit():
                        rakamlar.append(char)
                    elif rakamlar:
                        break
                rakam_kismi = ''.join(rakamlar)
                if rakam_kismi and rakam_kismi != "0":
                    return rakam_kismi
        
        return None
    

    def excel_islemi_tamamla(self, foto_klasorleri, atlanan_klasorler, hat_sureleri=None, hat_tarihsiz_istatistikleri=None):
        start_time = time.time()
        self.log_mesaj_ekle("💾 Excel dosyası oluşturuluyor...")
        self.throttled_update()
        
        try:
            # 1. AŞAMA: Workbook oluşturma
            self.log_mesaj_ekle("📗 Workbook oluşturuluyor...")
            workbook = Workbook()
            workbook.remove(workbook.active)
            
            # 2. AŞAMA: Koordinat bilgilerini al
            self.log_mesaj_ekle("📍 Koordinat bilgileri alınıyor...")
            koordinat_bilgileri = self.koordinat_bilgilerini_al(foto_klasorleri)
            
            # 3. AŞAMA: TESPİT BİLGİLERİNİ BİRLEŞTİR (ÜÇLÜ ANAHTAR İLE) - YAPILDI MI? BİLGİSİ DE ALINACAK
            self.log_mesaj_ekle("🔍 Tespit bilgileri birleştiriliyor...")
            tespit_verileri = self.tespit_bilgilerini_birlestir(foto_klasorleri)
            
            # 4. AŞAMA: Mesafe bilgilerini hesapla (YENİ: hat uzunluklarını da al)
            self.log_mesaj_ekle("📏 Mesafe hesaplamaları yapılıyor...")
            mesafe_sonuclari_dict = self.mesafe_hesapla(foto_klasorleri, koordinat_bilgileri)
            mesafe_sonuclari = mesafe_sonuclari_dict.get('mesafe_sonuclari', {})
            hat_uzunluklari = mesafe_sonuclari_dict.get('hat_uzunluklari', {})  # YENİ: Hat uzunluklarını al
            
            # 5. AŞAMA: Diğer bilgileri al (TESPİT VERİLERİNİ PARAMETRE OLARAK GÖNDER)
            self.log_mesaj_ekle("🏷️ Direk tipi ve tespit bilgileri alınıyor...")
            bilgiler = self.direk_tipi_bilgilerini_al(foto_klasorleri, tespit_verileri)
            direk_tipleri = bilgiler.get('direk_tipleri', {})
            trafo_direkleri = bilgiler.get('trafo_direkleri', {})
            tespit_bilgileri = bilgiler.get('tespit_bilgileri', {})
            
            # YENİ: TESPİT VERİLERİNİ SÖZLÜĞE DÖNÜŞTÜR (ÜÇLÜ ANAHTAR İLE)
            tespit_dict = {}
            for tespit in tespit_verileri:
                # İlçe bilgisini al (hem İlçe hem Aob için)
                ilce = tespit.get('İlçe', tespit.get('Aob', '')).strip()
                
                # Hat adı bilgisini al (tüm olası sütun adları)
                hat_adi = tespit.get('Hat Adı', 
                           tespit.get('Hat adı', 
                           tespit.get('Hat Adı', 
                           tespit.get('Hat', '')))).strip()
                
                # Direk numarasını al (hem Direk No hem Direk ID için)
                direk_no = tespit.get('Direk No', tespit.get('Direk ID', ''))
                temiz_direk_no = self.direk_no_temizle(direk_no)
                
                if ilce and hat_adi and temiz_direk_no:
                    # ÜÇLÜ ANAHTAR: (İlçe, Hat Adı, Direk No)
                    anahtar = (ilce, hat_adi, temiz_direk_no)
                    
                    tespit_dict[anahtar] = {
                        'oncelik': tespit.get('Öncelik', ''),
                        'tespit_notu': tespit.get('Tespit Notu', ''),
                        'yapildi_mi': tespit.get('Yapıldı mı?', tespit.get('Yapıldı mı', ''))
                    }
                        
            if foto_klasorleri:
                # 6. AŞAMA: Fotoğraf Klasörleri sayfası
                self.log_mesaj_ekle("📊 Fotoğraf Klasörleri sayfası hazırlanıyor...")
                ws_veri = workbook.create_sheet('Fotoğraf Klasörleri')
                
                basliklar = [
                    'Sıra No', 'Klasör Yolu', 'Aob', 'Hat adı', 'Direk No', 'Fotoğraf Sayısı',
                    'Orijinal Klasör Adı', 'Gün', 'Ay', 'Yıl', 'En Yeni Fotoğraf Tarihi',
                    'Toplam Gün', 'Direk Tipi', 'LAT', 'LON', 'Mesafe (m)', 
                    'Öncelik', 'Tespit Notu', 'Yapıldı mı?'  # YAPILDI MI? SÜTUNU
                ]
                
                for col_num, column_title in enumerate(basliklar, 1):
                    cell = ws_veri.cell(row=1, column=col_num)
                    cell.value = column_title
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
                    cell.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                       top=Side(style='thin'), bottom=Side(style='thin'))
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                
                # 7. AŞAMA: Veri satırlarını ekleme (ÜÇLÜ ANAHTAR İLE TESPİT VERİSİ EŞLEŞTİRME)
                self.log_mesaj_ekle(f"📝 {len(foto_klasorleri)} satır veri ekleniyor...")

                for row_num, row_data in enumerate(foto_klasorleri, 2):
                    # Temel bilgileri al - Fotoğraf Klasörleri sayfasından
                    aob = row_data.get('Aob', row_data.get('İlçe', '')).strip()
                    hat_adi = row_data.get('Hat adı', row_data.get('Hat Adı', '')).strip()
                    
                    # Direk numarasını al ve temizle
                    klasor_adi = str(row_data.get('Direk No', row_data.get('Direk ID', '')))
                    temiz_direk_no = self.direk_no_temizle(klasor_adi)
                    
                    # Sıra No bilgisini al - BU SATIR DAHA ERKEN OLMALI!
                    orijinal_klasor_adi = row_data.get('Orijinal Klasör Adı', '')
                    sira_no = self.sira_no_ayikla(orijinal_klasor_adi)  # BU SATIRI ERKEN TANIMLA!
                    
                    hat_anahtari = (aob, hat_adi)
                    hat_sure = hat_sureleri.get(hat_anahtari, {}) if hat_sureleri else {}
                    
                    # ÜÇLÜ ANAHTAR İLE TESPİT BİLGİLERİNİ ARA
                    uc_anahtar_1 = (aob, hat_adi, temiz_direk_no)
                    uc_anahtar_2 = (aob.upper() if aob else '', hat_adi.upper() if hat_adi else '', temiz_direk_no)
                    uc_anahtar_3 = (aob.lower() if aob else '', hat_adi.lower() if hat_adi else '', temiz_direk_no)
                    
                    # Tespit bilgilerini al
                    oncelik = ""
                    tespit_notu = ""
                    yapildi_mi = ""
                    
                    if uc_anahtar_1 in tespit_dict:
                        tespit_info = tespit_dict[uc_anahtar_1]
                    elif uc_anahtar_2 in tespit_dict:
                        tespit_info = tespit_dict[uc_anahtar_2]
                    elif uc_anahtar_3 in tespit_dict:
                        tespit_info = tespit_dict[uc_anahtar_3]
                    else:
                        tespit_info = None
                    
                    if tespit_info:
                        oncelik = tespit_info.get('oncelik', '')
                        tespit_notu = tespit_info.get('tespit_notu', '')
                        yapildi_mi = tespit_info.get('yapildi_mi', '')
                    
                    # Direk tipi bilgisini al
                    direk_tipi = direk_tipleri.get(temiz_direk_no, '') if temiz_direk_no in direk_tipleri else ''
                    
                    # Trafo direği bilgisini al
                    trafo_diregi = trafo_direkleri.get(temiz_direk_no, '') if temiz_direk_no in trafo_direkleri else ''
                    
                    # Koordinat bilgilerini al
                    koordinat_info = koordinat_bilgileri.get(temiz_direk_no, {})
                    lat = koordinat_info.get('lat', '')
                    lon = koordinat_info.get('lon', '')
                    
                    # Mesafe bilgisini al
                    uc_mesafe_anahtar = (aob, hat_adi, sira_no, temiz_direk_no)
                    mesafe = mesafe_sonuclari.get(uc_mesafe_anahtar, '')
                    
                    # HÜCRE DEĞERLERİ - FOTOĞRAF KLASÖRLERİ SAYFASI
                    cell_values = [
                        sira_no,  # ŞİMDİ TANIMLI!
                        row_data.get('Klasör Yolu', ''),
                        aob,
                        hat_adi,
                        row_data.get('Direk No', row_data.get('Direk ID', '')),
                        row_data.get('Fotoğraf Sayısı', ''),
                        orijinal_klasor_adi,
                        row_data.get('Gün', ''),
                        row_data.get('Ay', ''),
                        row_data.get('Yıl', ''),
                        row_data.get('En Yeni Fotoğraf Tarihi', ''),    
                        hat_sure.get('toplam_gun', ''),
                        direk_tipi,
                        lat,
                        lon,
                        mesafe,
                        oncelik,
                        tespit_notu,
                        yapildi_mi
                    ]
                    
                    for col_num, cell_value in enumerate(cell_values, 1):
                        cell = ws_veri.cell(row=row_num, column=col_num)
                        cell.value = cell_value
                        
                        # Mesafe sütunu için özel sayı formatı
                        if col_num == 17:
                            if isinstance(cell_value, (int, float)):
                                cell.number_format = '#,##0.0'
                        
                        cell.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                           top=Side(style='thin'), bottom=Side(style='thin'))
                        cell.alignment = Alignment(horizontal='left', vertical='center')
                
                # 8. AŞAMA: Sütun genişliklerini ayarla
                self.log_mesaj_ekle("📏 Sütun genişlikleri ayarlanıyor...")
                self.auto_adjust_column_widths(ws_veri)
                ws_veri.auto_filter.ref = ws_veri.dimensions
                
                # 9. AŞAMA: Hat Özeti sayfası (YENİ: Hat Uzunluğu sütunu eklendi)
                if hat_sureleri:
                    self.log_mesaj_ekle("📈 Hat Özeti sayfası oluşturuluyor...")
                    ws_hat_ozet = workbook.create_sheet('Hat Özeti')
                    
                    # YENİ: 'Hat Uzunluğu (km)' sütunu eklendi (Toplam Gün'den sonra)
                    hat_basliklar = [
                        'İlçe', 'Hat Adı', 'Müşterek Direk', 'OG Direk', 
                        'Direk Sayısı', 'Fotoğraf Sayısı', 'İlk Çekim Tarihi', 
                        'Son Çekim Tarihi', 'Toplam Gün', 'Hat Uzunluğu (km)'  # YENİ SÜTUN
                    ]
                    
                    for col_num, column_title in enumerate(hat_basliklar, 1):
                        cell = ws_hat_ozet.cell(row=1, column=col_num)
                        cell.value = column_title
                        cell.font = Font(bold=True, color="FFFFFF")
                        cell.fill = PatternFill(start_color="27ae60", end_color="27ae60", fill_type="solid")
                        cell.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                           top=Side(style='thin'), bottom=Side(style='thin'))
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                    
                    # Direk listelerini dış dosyalardan yükle
                    self.log_mesaj_ekle("📂 Müşterek ve OG direk listeleri yükleniyor...")
                    musterek_direk_listesi = self.musterek_direk_listesi_yukle()
                    og_direk_listesi = self.og_direk_listesi_yukle()
                    
                    # Hat detaylarını topla
                    hat_detaylari = {}
                    for row_data in foto_klasorleri:
                        aob = row_data.get('Aob', '')
                        hat_adi = row_data.get('Hat adı', '')
                        hat_anahtari = (aob, hat_adi)
                        
                        if hat_anahtari not in hat_detaylari:
                            hat_detaylari[hat_anahtari] = {
                                'ilce': aob,
                                'hat_adi': hat_adi,
                                'musterk_direk': 0,
                                'og_direk': 0,
                                'direk_sayisi': 0,
                                'foto_sayisi': 0
                            }
                        
                        klasor_adi = str(row_data.get('Direk No', ''))
                        temiz_direk_no = self.direk_no_temizle(klasor_adi)
                        
                        # Müşterek direk kontrolü - dış dosyadan
                        if temiz_direk_no in musterek_direk_listesi:
                            hat_detaylari[hat_anahtari]['musterk_direk'] += 1
                        
                        # OG direk kontrolü - dış dosyadan
                        if temiz_direk_no in og_direk_listesi:
                            hat_detaylari[hat_anahtari]['og_direk'] += 1
                        
                        hat_detaylari[hat_anahtari]['direk_sayisi'] += 1
                        hat_detaylari[hat_anahtari]['foto_sayisi'] += row_data.get('Fotoğraf Sayısı', 0)
                    
                    # Hat süre bilgilerini ekle
                    for hat_anahtari, detay in hat_detaylari.items():
                        if hat_anahtari in hat_sureleri:
                            detay['ilk_tarih'] = hat_sureleri[hat_anahtari]['ilk_tarih']
                            detay['son_tarih'] = hat_sureleri[hat_anahtari]['son_tarih']
                            detay['toplam_gun'] = hat_sureleri[hat_anahtari]['toplam_gun']
                        
                        # YENİ: Hat uzunluğunu ekle (km cinsinden)
                        if hat_anahtari in hat_uzunluklari:
                            hat_uzunluk_km = hat_uzunluklari[hat_anahtari] / 1000.0
                            detay['hat_uzunlugu'] = round(hat_uzunluk_km, 3)  # 3 ondalık basamak
                        else:
                            detay['hat_uzunlugu'] = 0.0
                    
                    # Hatları sırala ve ekle
                    sirali_hatlar = sorted(hat_detaylari.values(), 
                                         key=lambda x: x['direk_sayisi'], 
                                         reverse=True)
                    
                    for row_num, hat_detay in enumerate(sirali_hatlar, 2):
                        cell_values = [
                            hat_detay['ilce'],
                            hat_detay['hat_adi'],
                            hat_detay['musterk_direk'],
                            hat_detay['og_direk'],
                            hat_detay['direk_sayisi'],
                            hat_detay['foto_sayisi'],
                            hat_detay['ilk_tarih'],
                            hat_detay['son_tarih'],
                            hat_detay['toplam_gun'],
                            hat_detay['hat_uzunlugu']  # YENİ: Hat uzunluğu
                        ]
                        
                        for col_num, cell_value in enumerate(cell_values, 1):
                            cell = ws_hat_ozet.cell(row=row_num, column=col_num)
                            cell.value = cell_value
                            
                            # YENİ: Hat uzunluğu sütunu için özel format (km cinsinden, 3 ondalık)
                            if col_num == 10 and isinstance(cell_value, (int, float)):
                                cell.number_format = '#,##0.000'
                            
                            cell.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                               top=Side(style='thin'), bottom=Side(style='thin'))
                            cell.alignment = Alignment(horizontal='left', vertical='center')
                    
                    self.auto_adjust_column_widths(ws_hat_ozet)
                    ws_hat_ozet.auto_filter.ref = ws_hat_ozet.dimensions
            
            # 10. AŞAMA: BİRLEŞTİRİLMİŞ_TESPİT sayfasını oluştur - TAMAMEN YENİ FORMAT
            if tespit_verileri:
                self.log_mesaj_ekle("📊 Birleştirilmiş_Tespit sayfası oluşturuluyor...")
                ws_tespit = workbook.create_sheet('Birleştirilmiş_Tespit')
                
                # DEĞİŞTİRİLDİ: 'Direk ID' yerine 'Direk No'
                tespit_basliklar = [
                    'İlçe', 'Hat Adı', 'Direk No', 'Tespit Notu', 'Tespit Kategorisi', 
                    'Öncelik', 'Fotoğraf Yolu', 'Enlem', 'Boylam', 'Birim', 'Yapıldı mı?'
                ]
                
                # BAŞLIK SATIRI FORMATLAMA - MAVİ ARKA PLAN, KALIN BEYAZ YAZI
                for col_num, column_title in enumerate(tespit_basliklar, 1):
                    cell = ws_tespit.cell(row=1, column=col_num)
                    cell.value = column_title
                    cell.font = Font(bold=True, color="FFFFFF")  # Kalın beyaz yazı
                    cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")  # Mavi arkaplan
                    cell.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                       top=Side(style='thin'), bottom=Side(style='thin'))
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                
                # VERİ SATIRLARINI EKLE - TÜM SATIRLARI SOLA HİZALI
                for row_num, tespit_data in enumerate(tespit_verileri, 2):
                    # DEĞİŞTİRİLDİ: 'Direk ID' yerine 'Direk No'
                    direk_no = tespit_data.get('Direk No', '')
                    direk_no_temiz = self.direk_no_temizle(direk_no)
                    
                    # Direk No'yu sayıya çevirmeye çalış
                    try:
                        if direk_no_temiz and direk_no_temiz != '':
                            if '.' in direk_no_temiz or ',' in direk_no_temiz:
                                direk_no_temiz = float(direk_no_temiz.replace(',', '.'))
                            else:
                                direk_no_temiz = int(direk_no_temiz)
                    except (ValueError, TypeError):
                        # Sayıya çevrilemezse string olarak bırak
                        pass
                    
                    # Fotoğraf Yolu için özel işlem - KÖPRÜ OLUŞTUR
                    foto_yolu = tespit_data.get('Fotoğraf Yolu', '')
                    
                    # Tüm sütunları ekle - Fotoğraf Yolu için hyperlink
                    for col_num, sutun_adi in enumerate(tespit_basliklar, 1):
                        cell = ws_tespit.cell(row=row_num, column=col_num)
                        
                        if sutun_adi == 'Fotoğraf Yolu' and foto_yolu:
                            # Fotoğraf yolunu kontrol et
                            if os.path.exists(foto_yolu):
                                # TAM YOL VAR VE GEÇERLİ - KÖPRÜ OLUŞTUR
                                cell.value = os.path.basename(foto_yolu)  # Görünen metin (sadece dosya adı)
                                cell.hyperlink = foto_yolu  # Köprü linki (tam yol)
                                cell.font = Font(color="0000FF", underline="single")  # Mavi ve altı çizili
                            else:
                                # Yol geçerli değilse sadece metin göster
                                cell.value = foto_yolu if foto_yolu else ""
                                cell.font = Font(color="000000")  # Siyah
                            
                            cell.alignment = Alignment(horizontal='left', vertical='center')
                        elif sutun_adi == 'Direk No':
                            cell.value = direk_no_temiz
                            cell.alignment = Alignment(horizontal='left', vertical='center')
                        else:
                            cell.value = tespit_data.get(sutun_adi, '')
                            cell.alignment = Alignment(horizontal='left', vertical='center')
                    
                    # Hücre sınırlarını ekle
                    for col_num in range(1, len(tespit_basliklar) + 1):
                        cell = ws_tespit.cell(row=row_num, column=col_num)
                        cell.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                           top=Side(style='thin'), bottom=Side(style='thin'))
                
                # SÜTUN GENİŞLİKLERİNİ OTOMATİK AYARLA
                for column in ws_tespit.columns:
                    max_length = 0
                    column_letter = get_column_letter(column[0].column)
                    
                    for cell in column:
                        try:
                            if cell.value:
                                cell_length = len(str(cell.value))
                                if cell_length > max_length:
                                    max_length = cell_length
                        except:
                            pass
                    
                    adjusted_width = min(max_length + 2, 50)
                    ws_tespit.column_dimensions[column_letter].width = adjusted_width
                
                # FİLTRE EKLE
                if ws_tespit.max_row > 1:
                    filter_range = f"A1:{get_column_letter(ws_tespit.max_column)}{ws_tespit.max_row}"
                    ws_tespit.auto_filter.ref = filter_range
                    self.log_mesaj_ekle(f"✅ Tespit sayfasına filtre eklendi: {filter_range}")
                
                self.log_mesaj_ekle(f"✅ {len(tespit_verileri)} tespit verisi Birleştirilmiş_Tespit sayfasına eklendi (Fotoğraf Yolu köprülü)")
            
            # 11. AŞAMA: Atlanan klasörler sayfası
            if atlanan_klasorler:
                self.log_mesaj_ekle("📝 Atlanan klasörler sayfası oluşturuluyor...")
                ws_atlanan = workbook.create_sheet('Klasör Uyumsuz')
                
                basliklar = ['Klasör Yolu', 'Direk No', 'Sebep']
                for col_num, column_title in enumerate(basliklar, 1):
                    cell = ws_atlanan.cell(row=1, column=col_num)
                    cell.value = column_title
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.fill = PatternFill(start_color="e74c3c", end_color="e74c3c", fill_type="solid")
                    cell.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                                       top=Side(style='thin'), bottom=Side(style='thin'))
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                
                for row_num, atlama_veri in enumerate(atlanan_klasorler, 2):
                    ws_atlanan.cell(row=row_num, column=1, value=atlama_veri.get('Klasör Yolu', ''))
                    ws_atlanan.cell(row=row_num, column=2, value=atlama_veri.get('Direk No', ''))
                    ws_atlanan.cell(row=row_num, column=3, value=atlama_veri.get('Sebep', ''))
                
                self.auto_adjust_column_widths(ws_atlanan)
                ws_atlanan.auto_filter.ref = ws_atlanan.dimensions
            
            # 12. AŞAMA: Dosyayı kaydet
            self.log_mesaj_ekle("💾 Excel dosyası kaydediliyor...")
            workbook.save(self.kaydetme_yeri.get())
            
            # 13. AŞAMA: Sonuçları göster
            toplam_sure = time.time() - start_time
            self.lbl_durum.config(text="İşlem tamamlandı ✓", fg=self.colors['success'])
            self.log_mesaj_ekle(f"\n🎉 İŞLEM TAMAMLANDI!")
            self.log_mesaj_ekle(f"⏱️ Toplam süre: {toplam_sure:.2f}s")
            self.log_mesaj_ekle(f"📁 Excel dosyası: {self.kaydetme_yeri.get()}")
            self.log_mesaj_ekle(f"✅ {len(foto_klasorleri)} klasör eklendi")
            self.log_mesaj_ekle(f"❌ {len(atlanan_klasorler)} klasör atlandı")
            self.log_mesaj_ekle(f"🔍 {len(tespit_verileri)} tespit verisi eklendi")
            
            # Rapor ekranını aç
            self.rapor_ekrani_ac(self.kaydetme_yeri.get(), hat_tarihsiz_istatistikleri)
            
        except Exception as e:
            self.lbl_durum.config(text="Excel oluşturma hatası", fg=self.colors['danger'])
            self.log_mesaj_ekle(f"❌ Excel oluşturma hatası: {str(e)}")
            import traceback
            self.log_mesaj_ekle(f"🔍 Hata detayı: {traceback.format_exc()}")





    def tespit_bilgilerini_birlestir(self, foto_klasorleri):
        """Tespit Excel dosyalarını birleştirir ve döndürür (İlçe bilgisi ve TAM FOTOĞRAF YOLU eklenmiş şekilde)"""
        tespit_verileri = []
        
        try:
            # ÖNCE: Fotoğraf klasörleri verisini kullanarak direk-fotoğraf eşleştirmesi oluştur
            direk_foto_eslestirme = {}
            
            for klasor_veri in foto_klasorleri:
                klasor_yolu = klasor_veri.get('Klasör Yolu', '')
                direk_no = klasor_veri.get('Direk No', '')
                temiz_direk_no = self.direk_no_temizle(str(direk_no))
                aob = klasor_veri.get('Aob', '')
                hat_adi = klasor_veri.get('Hat adı', '')
                
                if temiz_direk_no and klasor_yolu:
                    # Anahtar: (İlçe, Hat Adı, Direk No)
                    anahtar = (aob, hat_adi, temiz_direk_no)
                    
                    # Bu klasördeki fotoğrafları bul
                    foto_dosyalari = []
                    try:
                        for dosya in os.listdir(klasor_yolu):
                            if dosya.lower().endswith(('.jpg', '.jpeg', '.png', '.tiff', '.tif', '.bmp')):
                                tam_yol = os.path.join(klasor_yolu, dosya)
                                foto_dosyalari.append((dosya, tam_yol))
                    except:
                        foto_dosyalari = []
                    
                    direk_foto_eslestirme[anahtar] = foto_dosyalari
            
            self.log_mesaj_ekle(f"📍 {len(direk_foto_eslestirme)} direk için fotoğraf bilgisi hazırlandı")
            
            # Ana klasördeki tüm Excel dosyalarını bul
            excel_dosyalari = []
            for root, dirs, files in os.walk(self.ana_klasor.get()):
                for dosya in files:
                    if dosya.lower().endswith(('.xlsx', '.xls')):
                        tam_yol = os.path.join(root, dosya)
                        excel_dosyalari.append(tam_yol)
            
            if not excel_dosyalari:
                self.log_mesaj_ekle("ℹ️ Tespit Excel dosyası bulunamadı")
                return tespit_verileri
            
            self.log_mesaj_ekle(f"🔍 {len(excel_dosyalari)} Excel dosyası taranıyor...")
            
            # Hem 'Direk No' hem de 'Direk ID' için arama yap
            istenen_sutunlar = [
                'İlçe', 'Hat Adı', 'Tespit Notu', 'Tespit Kategorisi', 
                'Öncelik', 'Fotoğraf Yolu', 'Enlem', 'Boylam', 'Birim', 'Yapıldı mı?'
            ]
            
            # DEĞİŞKENLER: Direk No ve Direk ID - "Sıra No" hariç
            direk_no_degiskenleri = ['Direk No', 'Direk No.', 'Direk_No', 'DirekNo', 'Direk no', 'direk no']
            direk_id_degiskenleri = ['Direk ID', 'Direk Id', 'Direk_ID', 'DirekID', 'Direk', 'No']
            
            for dosya in excel_dosyalari:
                try:
                    # İlçeyi klasör yolundan bul
                    dosya_klasoru = os.path.dirname(dosya)
                    ilce_adi = self.aob_adi_bul(dosya_klasoru)
                    
                    self.log_mesaj_ekle(f"🔍 {os.path.basename(dosya)} - Klasör: {dosya_klasoru}")
                    self.log_mesaj_ekle(f"   🏙️ İlçe: {ilce_adi}")
                    
                    df = pd.read_excel(dosya, header=None)
                    
                    # Gerçek başlık satırını bul
                    baslik_satiri_idx = None
                    for idx in range(min(10, len(df))):
                        satir = df.iloc[idx]
                        satir_metin = ' '.join(str(x).strip() for x in satir.values if pd.notna(x) and str(x).strip())
                        
                        # "İlçe", "Hat Adı" ve "Yapıldı mı?" hariç diğer sütunları kontrol et
                        kontrol_sutunlari = [s for s in istenen_sutunlar if s not in ['İlçe', 'Hat Adı', 'Yapıldı mı?']]
                        
                        # 'Direk No' veya 'Direk ID' kontrolü - SADECE bu değişkenlerle
                        direk_kontrol = False
                        for direk_degisken in direk_no_degiskenleri + direk_id_degiskenleri:
                            if any(direk_degisken.lower() in str(x).lower() for x in satir.values if pd.notna(x)):
                                direk_kontrol = True
                                break
                        
                        # Diğer sütunların kontrolü
                        diger_kontrol = all(any(sutun_adi.lower() in str(x).lower() for x in satir.values if pd.notna(x)) 
                                           for sutun_adi in kontrol_sutunlari if sutun_adi not in ['Direk No', 'Direk ID'])
                        
                        if direk_kontrol and diger_kontrol:
                            baslik_satiri_idx = idx
                            self.log_mesaj_ekle(f"   ✓ Başlık satırı {idx+1}. satırda bulundu")
                            break
                    
                    if baslik_satiri_idx is not None:
                        df = pd.read_excel(dosya, header=baslik_satiri_idx)
                    else:
                        df = pd.read_excel(dosya, header=0)
                        self.log_mesaj_ekle(f"ℹ️ {os.path.basename(dosya)} - İlk satır başlık olarak kullanıldı")
                    
                    # Sütun eşleştirme - ÖNCE 'Direk No', SONRA 'Direk ID' ara
                    mevcut_sutunlar = df.columns.tolist()
                    sutun_eslesmeleri = {}
                    
                    # 'Direk No' veya 'Direk ID' sütununu bul - "Sıra No" HARİÇ
                    direk_sutunu = None
                    for mevcut_sutun in mevcut_sutunlar:
                        mevcut_sutun_clean = str(mevcut_sutun).strip().lower()
                        
                        # ÖNCE 'Direk No' arasın - "Sıra No" kontrolü YOK
                        for direk_no_degisken in direk_no_degiskenleri:
                            if direk_no_degisken.lower() == mevcut_sutun_clean:
                                direk_sutunu = mevcut_sutun
                                sutun_eslesmeleri[mevcut_sutun] = 'Direk No'
                                self.log_mesaj_ekle(f"   ✓ 'Direk No' sütunu bulundu: {mevcut_sutun}")
                                break
                        
                        if direk_sutunu:
                            break
                    
                    # Eğer 'Direk No' bulunamadıysa, 'Direk ID' arasın - "Sıra No" HARİÇ
                    if not direk_sutunu:
                        for mevcut_sutun in mevcut_sutunlar:
                            mevcut_sutun_clean = str(mevcut_sutun).strip().lower()
                            
                            for direk_id_degisken in direk_id_degiskenleri:
                                # SADECE tam eşleşme veya kısmi eşleşme, "Sıra No" DEĞİL
                                if direk_id_degisken.lower() == mevcut_sutun_clean or \
                                   (direk_id_degisken.lower() in mevcut_sutun_clean and 'sıra' not in mevcut_sutun_clean):
                                    direk_sutunu = mevcut_sutun
                                    sutun_eslesmeleri[mevcut_sutun] = 'Direk ID'
                                    self.log_mesaj_ekle(f"   ✓ 'Direk ID' sütunu bulundu: {mevcut_sutun}")
                                    break
                            
                            if direk_sutunu:
                                break
                    
                    # Diğer sütunları bul
                    for istenen_sutun in istenen_sutunlar:
                        if istenen_sutun in ['İlçe', 'Hat Adı', 'Direk No', 'Direk ID']:
                            continue
                            
                        eslesen_sutun = None
                        for mevcut_sutun in mevcut_sutunlar:
                            mevcut_sutun_clean = str(mevcut_sutun).strip().lower()
                            istenen_sutun_clean = str(istenen_sutun).strip().lower()
                            
                            if (mevcut_sutun_clean == istenen_sutun_clean or 
                                istenen_sutun_clean in mevcut_sutun_clean):
                                eslesen_sutun = mevcut_sutun
                                sutun_eslesmeleri[mevcut_sutun] = istenen_sutun
                                break
                        
                        if eslesen_sutun:
                            self.log_mesaj_ekle(f"   ✓ '{istenen_sutun}' sütunu bulundu: {eslesen_sutun}")
                    
                    # Verileri işle
                    dosya_klasor_baslik = os.path.basename(os.path.dirname(dosya))
                    dosya_veri_sayisi = 0
                    
                    for index, row in df.iterrows():
                        try:
                            tespit_data = {
                                'İlçe': ilce_adi,
                                'Hat Adı': dosya_klasor_baslik
                            }
                            
                            # Direk No/ID bilgisini al
                            direk_degeri = ""
                            if direk_sutunu and direk_sutunu in row:
                                direk_degeri = str(row[direk_sutunu]) if pd.notna(row[direk_sutunu]) else ""
                                if direk_degeri:
                                    temiz_direk_no = self.direk_no_temizle(direk_degeri)
                                    tespit_data['Direk No'] = temiz_direk_no
                                    
                                    # DEBUG: Direk No değerini logla
                                    if dosya_veri_sayisi < 3:  # İlk 3 kayıt için
                                        self.log_mesaj_ekle(f"   🔍 Örnek Direk No: '{direk_degeri}' -> '{temiz_direk_no}'")
                                else:
                                    tespit_data['Direk No'] = ""
                            else:
                                tespit_data['Direk No'] = ""
                            
                            # Diğer sütunları al
                            for istenen_sutun in istenen_sutunlar:
                                if istenen_sutun in ['İlçe', 'Hat Adı', 'Direk No', 'Direk ID', 'Fotoğraf Yolu']:
                                    continue
                                    
                                eslesen_mevcut_sutun = None
                                for mevcut_sutun, standart_sutun in sutun_eslesmeleri.items():
                                    if standart_sutun == istenen_sutun:
                                        eslesen_mevcut_sutun = mevcut_sutun
                                        break
                                
                                if eslesen_mevcut_sutun and eslesen_mevcut_sutun in row:
                                    deger = row[eslesen_mevcut_sutun]
                                    tespit_data[istenen_sutun] = deger if pd.notna(deger) else ""
                                else:
                                    tespit_data[istenen_sutun] = ""
                            
                            # "Fotoğraf Yolu" için özel işlem - TAM YOLU BUL
                            foto_yolu_degeri = ""
                            if 'Fotoğraf Yolu' in sutun_eslesmeleri.values():
                                # Eşleşen sütunu bul
                                foto_sutunu = None
                                for mevcut_sutun, standart_sutun in sutun_eslesmeleri.items():
                                    if standart_sutun == 'Fotoğraf Yolu':
                                        foto_sutunu = mevcut_sutun
                                        break
                                
                                if foto_sutunu and foto_sutunu in row:
                                    foto_yolu_degeri = str(row[foto_sutunu]) if pd.notna(row[foto_sutunu]) else ""
                            
                            # ŞİMDİ FOTOĞRAFIN TAM YOLUNU BUL
                            tam_foto_yolu = ""
                            
                            if tespit_data['Direk No'] and ilce_adi and dosya_klasor_baslik:
                                # Anahtar: (İlçe, Hat Adı, Direk No)
                                anahtar = (ilce_adi, dosya_klasor_baslik, tespit_data['Direk No'])
                                
                                if anahtar in direk_foto_eslestirme:
                                    foto_listesi = direk_foto_eslestirme[anahtar]
                                    
                                    if foto_yolu_degeri:
                                        # Tespit dosyasında belirtilen fotoğraf adını ara
                                        for foto_adi, foto_tam_yol in foto_listesi:
                                            if foto_adi == foto_yolu_degeri or foto_yolu_degeri in foto_adi:
                                                tam_foto_yolu = foto_tam_yol
                                                break
                                    
                                    # Eğer tam yol bulunamadıysa, ilk fotoğrafı kullan
                                    if not tam_foto_yolu and foto_listesi:
                                        tam_foto_yolu = foto_listesi[0][1]  # İlk fotoğrafın tam yolu
                            
                            # Fotoğraf yolunu kaydet (tam yol veya orijinal değer)
                            tespit_data['Fotoğraf Yolu'] = tam_foto_yolu if tam_foto_yolu else foto_yolu_degeri
                            
                            # "Yapıldı mı?" sütunu için özel kontrol
                            if 'Yapıldı mı?' not in tespit_data:
                                tespit_data['Yapıldı mı?'] = ""
                            
                            # Birim filtresi
                            if tespit_data.get('Birim', ''):
                                birim_str = str(tespit_data['Birim']).strip().lower()
                                if birim_str in ['aob', ''] or pd.isna(tespit_data['Birim']):
                                    tespit_verileri.append(tespit_data)
                                    dosya_veri_sayisi += 1
                        
                        except Exception as e:
                            continue
                    
                    self.log_mesaj_ekle(f"✅ {os.path.basename(dosya)} - {dosya_veri_sayisi} tespit verisi eklendi (İlçe: {ilce_adi})")
                    
                except Exception as e:
                    self.log_mesaj_ekle(f"⚠️ {os.path.basename(dosya)} - Hata: {str(e)}")
            
            # İlçe bazında istatistik
            ilce_bazli_sayilar = {}
            for veri in tespit_verileri:
                ilce = veri.get('İlçe', 'Bilinmeyen')
                if ilce not in ilce_bazli_sayilar:
                    ilce_bazli_sayilar[ilce] = 0
                ilce_bazli_sayilar[ilce] += 1
            
            self.log_mesaj_ekle(f"📊 Toplam {len(tespit_verileri)} tespit verisi birleştirildi")
            for ilce, sayi in ilce_bazli_sayilar.items():
                self.log_mesaj_ekle(f"   📍 {ilce}: {sayi} tespit")
            
            # DEBUG: İlk 5 tespit verisini göster
            if tespit_verileri:
                self.log_mesaj_ekle("🔍 İlk 5 tespit verisi örneği:")
                for i, tespit in enumerate(tespit_verileri[:5]):
                    self.log_mesaj_ekle(f"   {i+1}. İlçe: {tespit.get('İlçe', '')}, "
                                       f"Direk No: {tespit.get('Direk No', '')}, "
                                       f"Fotoğraf Yolu: {os.path.basename(tespit.get('Fotoğraf Yolu', '')) if tespit.get('Fotoğraf Yolu') else 'Yok'}")
            
        except Exception as e:
            self.log_mesaj_ekle(f"❌ Tespit birleştirme hatası: {str(e)}")
        
        return tespit_verileri


            

    def mesafe_hesapla(self, foto_klasorleri, koordinat_bilgileri):
        """Her direk için bir önceki direkle arasındaki mesafeyi METRE cinsinden hesaplar ve toplam hat uzunluklarını hesaplar"""
        mesafe_sonuclari = {}
        hat_uzunluklari = {}  # YENİ: Hat bazında toplam uzunluklar
        
        try:
            mesafe_sonuclari.clear()
            hat_uzunluklari.clear()  # YENİ: Hat uzunluklarını temizle
            self.log_mesaj_ekle("🔄 Mesafe hesaplamaları yapılıyor (Nihai Direk özel durumu ile)...")
            
            # Direk tipi bilgilerini al
            direk_tipi_bilgileri = self.direk_tipi_bilgilerini_al(foto_klasorleri)
            
            # direk_tipleri sözlüğünden Nihai Direk'leri filtrele
            direk_tipleri = direk_tipi_bilgileri.get('direk_tipleri', {})
            nihai_direkler = {k: True for k, v in direk_tipleri.items() if v == 'Nihai Direk'}
            
            # DEBUG log
            self.log_mesaj_ekle(f"🔍 Mesafe hesaplama için {len(nihai_direkler)} Nihai Direk bulundu")
            if nihai_direkler:
                nihai_list = list(nihai_direkler.keys())[:10]
                self.log_mesaj_ekle(f"🔍 Nihai direk no'ları (ilk 10): {nihai_list}")
                if '7' in nihai_direkler:
                    self.log_mesaj_ekle(f"✅ Direk 7 Nihai Direk olarak tanındı!")
                else:
                    self.log_mesaj_ekle(f"❌ Direk 7 Nihai Direk olarak tanınMADI!")
            
            hat_gruplari = {}
            
            for klasor in foto_klasorleri:
                hat_anahtari = (klasor['Aob'], klasor['Hat adı'])
                if hat_anahtari not in hat_gruplari:
                    hat_gruplari[hat_anahtari] = []
                    hat_uzunluklari[hat_anahtari] = 0.0  # YENİ: Her hat için uzunluk sıfırla
                hat_gruplari[hat_anahtari].append(klasor)
            
            for hat_anahtari, klasor_listesi in hat_gruplari.items():
                aob, hat_adi = hat_anahtari
                sirali_klasorler = sorted(klasor_listesi, 
                                        key=lambda x: self.sira_no_ile_sirala(x['Orijinal Klasör Adı']))
                
                onceki_lat = None
                onceki_lon = None
                onceki_direk_no = None
                onceki_sira_no = None
                onceki_direk_tipi = None  # YENİ: Önceki direğin tipini tut
                
                # Geriye doğru arama için tüm direkleri kaydedelim
                hat_direkleri = []
                
                for i, klasor in enumerate(sirali_klasorler):
                    klasor_adi = str(klasor.get('Direk No', ''))
                    temiz_direk_no = self.direk_no_temizle(klasor_adi)
                    sira_no = self.sira_no_ayikla(klasor.get('Orijinal Klasör Adı', ''))
                    
                    # Bu direğin tipini belirle
                    direk_tipi = "Normal"
                    if temiz_direk_no in nihai_direkler:
                        direk_tipi = "Nihai Direk"
                    
                    # ÜÇLÜ ANAHTAR: (Aob, Hat-Adı, SıraNo, DirekNo)
                    uc_anahtar = (aob, hat_adi, sira_no, temiz_direk_no)
                    
                    # Koordinat bilgilerini al
                    koordinat_info = koordinat_bilgileri.get(temiz_direk_no, {})
                    lat = koordinat_info.get('lat')
                    lon = koordinat_info.get('lon')
                    
                    # Hat direklerini kaydet (geriye doğru arama için)
                    hat_direkleri.append({
                        'index': i,
                        'anahtar': uc_anahtar,
                        'direk_no': temiz_direk_no,
                        'sira_no': sira_no,
                        'tip': direk_tipi,
                        'lat': lat,
                        'lon': lon,
                        'mesafe': 0  # Varsayılan mesafe
                    })
                
                for i, klasor in enumerate(sirali_klasorler):
                    klasor_adi = str(klasor.get('Direk No', ''))
                    temiz_direk_no = self.direk_no_temizle(klasor_adi)
                    sira_no = self.sira_no_ayikla(klasor.get('Orijinal Klasör Adı', ''))
                    
                    # ÜÇLÜ ANAHTAR: (Aob, Hat-Adı, SıraNo, DirekNo)
                    uc_anahtar = (aob, hat_adi, sira_no, temiz_direk_no)
                    
                    koordinat_info = koordinat_bilgileri.get(temiz_direk_no, {})
                    simdiki_lat = koordinat_info.get('lat')
                    simdiki_lon = koordinat_info.get('lon')
                    
                    # Bu direğin tipini belirle
                    direk_tipi = "Normal"
                    if temiz_direk_no in nihai_direkler:
                        direk_tipi = "Nihai Direk"
                    
                    if i == 0:
                        # 1- İlk direk için mesafe yok
                        mesafe_deger = ""
                        self.log_mesaj_ekle(f"📍 {aob}-{hat_adi}-{sira_no}-{temiz_direk_no}: İlk direk (mesafe yok)")
                    
                    elif onceki_direk_tipi == "Nihai Direk":
                        # 2- ÖNCEKİ DİREK NİHAİ DİREK İSE: Geriye doğru EN KISA mesafeyi bul (NİHAİ DİREK HARİÇ - SADECE NORMAL DİREKLER)
                        if (simdiki_lat is not None and simdiki_lon is not None):
                            try:
                                # Geriye doğru TÜM NORMAL direklerde EN KISA mesafeyi bul (NİHAİ DİREKLER HARİÇ)
                                en_kisa_mesafe = float('inf')
                                en_kisa_direk = None
                                
                                # Geriye doğru ara (i-2'den 0'a kadar) - NİHAİ DİREK'i ATLA
                                for j in range(i-2, -1, -1):  # i-2'den başla çünkü i-1 Nihai Direk
                                    geri_direk = hat_direkleri[j]
                                    geri_lat = geri_direk['lat']
                                    geri_lon = geri_direk['lon']
                                    geri_tipi = geri_direk['tip']
                                    
                                    # SADECE NORMAL DİREKLERİ KONTROL ET (NİHAİ DİREKLERİ ATLA)
                                    if geri_tipi == "Nihai Direk":
                                        continue  # Nihai Direk'leri atla
                                    
                                    # Koordinatı olan NORMAL direkleri kontrol et
                                    if geri_lat is not None and geri_lon is not None:
                                        # Mevcut direk ile gerideki NORMAL direk arasındaki mesafeyi hesapla
                                        mesafe = self.koordinat_mesafe_hesapla(geri_lat, geri_lon, simdiki_lat, simdiki_lon)
                                        
                                        # EN KISA MESAFEYİ KAYDET
                                        if mesafe < en_kisa_mesafe:
                                            en_kisa_mesafe = mesafe
                                            en_kisa_direk = geri_direk
                                
                                if en_kisa_direk is not None and en_kisa_mesafe != float('inf'):
                                    # En kısa mesafeyi bulduk
                                    mesafe_deger = float(en_kisa_mesafe)
                                    # YENİ: Mesafeyi hat uzunluğuna ekle
                                    if mesafe_deger > 0:
                                        hat_uzunluklari[hat_anahtari] += mesafe_deger
                                    
                                    self.log_mesaj_ekle(f"🔍 {aob}-{hat_adi}-{sira_no}-{temiz_direk_no}: Nihai direkten sonra - {en_kisa_direk['direk_no']} (Normal) ile EN KISA mesafe {mesafe_deger:.2f} m")
                                    
                                    # Debug: Tüm kontrol edilen NORMAL direkleri göster
                                    debug_mesaj = f"    Kontrol edilen NORMAL direkler: "
                                    mesafe_listesi = []
                                    for j in range(i-2, -1, -1):
                                        geri_direk = hat_direkleri[j]
                                        geri_tipi = geri_direk['tip']
                                        if geri_tipi != "Nihai Direk":  # Sadece Normal direkler
                                            geri_lat = geri_direk['lat']
                                            geri_lon = geri_direk['lon']
                                            if geri_lat is not None and geri_lon is not None:
                                                m = self.koordinat_mesafe_hesapla(geri_lat, geri_lon, simdiki_lat, simdiki_lon)
                                                mesafe_listesi.append(f"{geri_direk['direk_no']}:{m:.1f}m")
                                    
                                    if mesafe_listesi:
                                        debug_mesaj += ", ".join(mesafe_listesi)
                                        self.log_mesaj_ekle(debug_mesaj)
                                else:
                                    # Geriye doğru NORMAL direk bulunamadı
                                    mesafe_deger = ""
                                    self.log_mesaj_ekle(f"⏭️ {aob}-{hat_adi}-{sira_no}-{temiz_direk_no}: Nihai direkten sonra geriye doğru NORMAL direk bulunamadı")
                            
                            except Exception as e:
                                mesafe_deger = "Hesaplanamadı"
                                self.log_mesaj_ekle(f"⚠️ Mesafe hesaplama hatası {uc_anahtar}: {str(e)}")
                        else:
                            mesafe_deger = "Koordinat yok"
                            self.log_mesaj_ekle(f"⚠️ {uc_anahtar}: Koordinat eksik")
                    
                    elif (onceki_lat is not None and onceki_lon is not None and 
                          simdiki_lat is not None and simdiki_lon is not None):
                        # 3- NORMAL DURUM: Önceki direkle (Normal ise) arasındaki mesafeyi hesapla
                        try:
                            mesafe_metre = self.koordinat_mesafe_hesapla(onceki_lat, onceki_lon, simdiki_lat, simdiki_lon)
                            mesafe_deger = float(mesafe_metre)
                            # YENİ: Mesafeyi hat uzunluğuna ekle
                            if mesafe_deger > 0:
                                hat_uzunluklari[hat_anahtari] += mesafe_deger
                            
                            tip_str = " (Nihai)" if onceki_direk_tipi == "Nihai Direk" else " (Normal)"
                            self.log_mesaj_ekle(f"📏 {aob}-{hat_adi}-{sira_no}-{temiz_direk_no}: {onceki_direk_no}{tip_str} ile mesafe {mesafe_deger:.2f} m")
                        except Exception as e:
                            mesafe_deger = "Hesaplanamadı"
                            self.log_mesaj_ekle(f"⚠️ Mesafe hesaplama hatası {uc_anahtar}: {str(e)}")
                    else:
                        mesafe_deger = "Koordinat yok"
                        self.log_mesaj_ekle(f"⚠️ {uc_anahtar}: Koordinat eksik")
                    
                    # ÜÇLÜ ANAHTAR ile kaydet
                    mesafe_sonuclari[uc_anahtar] = mesafe_deger
                    
                    # Bir sonraki iterasyon için değerleri güncelle
                    onceki_lat = simdiki_lat
                    onceki_lon = simdiki_lon
                    onceki_direk_no = temiz_direk_no
                    onceki_sira_no = sira_no
                    onceki_direk_tipi = direk_tipi  # Önceki direğin tipini güncelle
                    
                    # Hat direklerini güncelle (mesafe bilgisini kaydet)
                    if i < len(hat_direkleri):
                        hat_direkleri[i]['mesafe'] = mesafe_deger if isinstance(mesafe_deger, (int, float)) else 0
                
                # YENİ: Hat uzunluğunu km'ye çevir ve logla
                hat_uzunluk_km = hat_uzunluklari[hat_anahtari] / 1000.0
                self.log_mesaj_ekle(f"📏 {hat_anahtari} toplam uzunluk: {hat_uzunluk_km:.3f} km")
                self.log_mesaj_ekle(f"✅ {hat_anahtari} için {len([v for v in mesafe_sonuclari.values() if v and v != 'Koordinat yok' and v != 'Hesaplanamadı'])} mesafe hesaplandı")
            
            # YENİ: Toplam hat uzunlukları istatistiği
            toplam_hat_uzunluk_km = sum(hat_uzunluklari.values()) / 1000.0
            self.log_mesaj_ekle(f"📊 Toplam {len(hat_uzunluklari)} hat için uzunluk hesaplandı")
            self.log_mesaj_ekle(f"📏 Toplam hat uzunluğu: {toplam_hat_uzunluk_km:.2f} km")
            
            # YENİ: Nihai direk istatistikleri
            nihai_direk_sayisi = len(nihai_direkler)
            self.log_mesaj_ekle(f"🏁 Toplam {nihai_direk_sayisi} nihai direk bulundu")
            
            # YENİ: Hat uzunluklarını döndür
            sonuc = {
                'mesafe_sonuclari': mesafe_sonuclari,
                'hat_uzunluklari': hat_uzunluklari  # YENİ: Hat uzunluklarını da döndür
            }
            
            return sonuc
            
        except Exception as e:
            self.log_mesaj_ekle(f"❌ Mesafe hesaplama hatası: {str(e)}")
        
        return {'mesafe_sonuclari': {}, 'hat_uzunluklari': {}}
    



    def koordinat_mesafe_hesapla(self, lat1, lon1, lat2, lon2):
        """DOĞRU HAVERSINE FORMÜLÜ - Metre cinsinden"""
        try:
            # Koordinatları float'a çevir
            lat1 = float(lat1)
            lon1 = float(lon1)
            lat2 = float(lat2)
            lon2 = float(lon2)
            
            # Dünya yarıçapı (metre cinsinden)
            R = 6371000  # 6371 km * 1000 = 6371000 metre
            
            # Dereceleri radyana çevir
            lat1_rad = math.radians(lat1)
            lon1_rad = math.radians(lon1)
            lat2_rad = math.radians(lat2)
            lon2_rad = math.radians(lon2)
            
            # Koordinat farkları
            dlat = lat2_rad - lat1_rad
            dlon = lon2_rad - lon1_rad
            
            # Haversine formülü
            a = math.sin(dlat/2)**2 + math.cos(lat1_rad) * math.cos(lat2_rad) * math.sin(dlon/2)**2
            c = 2 * math.atan2(math.sqrt(a), math.sqrt(1-a))
            
            # Mesafe (metre cinsinden)
            distance = R * c
            
            return distance  # metre cinsinden döndür
            
        except Exception as e:
            self.log_mesaj_ekle(f"❌ Mesafe hesaplama hatası: {str(e)}")
            return 0.0

    


    def sira_no_ayikla(self, orijinal_klasor_adi):
        """Orijinal Klasör Adı'ndan sıra numarasını ayıklar"""
        if not orijinal_klasor_adi:
            return ""
        
        try:
            # String'e çevir
            klasor_adi = str(orijinal_klasor_adi).strip()
            
            # " - " veya "-" işaretlerini kontrol et
            if " - " in klasor_adi:
                parts = klasor_adi.split(" - ", 1)
            elif "-" in klasor_adi:
                parts = klasor_adi.split("-", 1)
            else:
                return ""  # Ayırıcı yoksa boş döndür
            
            if len(parts) >= 1:
                # İlk kısmı al ve sadece rakamları çıkar
                ilk_kisim = parts[0].strip()
                rakamlar = []
                
                for char in ilk_kisim:
                    if char.isdigit():
                        rakamlar.append(char)
                    elif rakamlar:  # Rakam başladıktan sonra rakam dışı karakter gelirse dur
                        break
                
                if rakamlar:
                    return ''.join(rakamlar)
            
            return ""
            
        except Exception as e:
            print(f"Sıra no ayıklama hatası: {e}")
            return ""

    
    def koordinat_bilgilerini_al(self, veri):
        """Direk numaralarına göre koordinat bilgilerini alır"""
        koordinat_bilgileri = {}
        
        try:
            # 1. Ana veriden direk numaralarını hazırla
            ana_direk_set = set()
            for row_data in veri:
                direk_no = str(row_data.get('Direk No', ''))
                temiz_no = self.direk_no_temizle(direk_no)
                if temiz_no and temiz_no != 'nan':
                    ana_direk_set.add(temiz_no)

            self.txt_sonuc.insert(tk.END, f"📍 {len(ana_direk_set)} direk için koordinat bilgisi aranıyor\n")
            
            if not ana_direk_set:
                return {}

            # 2. Müşterek direk Excel'inden koordinatları al
            if self.direk_musterek_yolu and os.path.exists(self.direk_musterek_yolu):
                try:
                    df_musterek = pd.read_excel(self.direk_musterek_yolu)
                    
                    # Sütun isimlerini bul
                    direkno_col = None
                    lat_col = None
                    lon_col = None
                    
                    for col in df_musterek.columns:
                        col_upper = col.upper()
                        if 'DIREK' in col_upper and ('NO' in col_upper or 'NUMARA' in col_upper):
                            direkno_col = col
                        elif 'LAT' in col_upper or 'ENLEM' in col_upper:
                            lat_col = col
                        elif 'LON' in col_upper or 'BOYLAM' in col_upper:
                            lon_col = col
                    
                    if direkno_col and lat_col and lon_col:
                        for index, row in df_musterek.iterrows():
                            direk_no = str(row[direkno_col])
                            temiz_no = self.direk_no_temizle(direk_no)
                            
                            if temiz_no and temiz_no in ana_direk_set:
                                lat = row[lat_col]
                                lon = row[lon_col]
                                
                                # Koordinat değerlerini kontrol et
                                if pd.notna(lat) and pd.notna(lon):
                                    koordinat_bilgileri[temiz_no] = {
                                        'lat': lat,
                                        'lon': lon
                                    }
                        
                        self.txt_sonuc.insert(tk.END, f"✓ Müşterek Excel'inden {len(koordinat_bilgileri)} koordinat bulundu\n")
                        
                except Exception as e:
                    self.txt_sonuc.insert(tk.END, f"⚠️ Müşterek Excel koordinat okuma hatası: {str(e)}\n")

            # 3. OG direk Excel'inden koordinatları al (müşterekte bulunmayanlar için)
            if self.direk_og_yolu and os.path.exists(self.direk_og_yolu):
                try:
                    df_og = pd.read_excel(self.direk_og_yolu)
                    
                    # Sütun isimlerini bul
                    direkno_col = None
                    lat_col = None
                    lon_col = None
                    
                    for col in df_og.columns:
                        col_upper = col.upper()
                        if 'DIREK' in col_upper and ('NO' in col_upper or 'NUMARA' in col_upper):
                            direkno_col = col
                        elif 'LAT' in col_upper or 'ENLEM' in col_upper:
                            lat_col = col
                        elif 'LON' in col_upper or 'BOYLAM' in col_upper:
                            lon_col = col
                    
                    if direkno_col and lat_col and lon_col:
                        og_koordinat_sayisi = 0
                        for index, row in df_og.iterrows():
                            direk_no = str(row[direkno_col])
                            temiz_no = self.direk_no_temizle(direk_no)
                            
                            # Sadece müşterekte bulunmayan direkler için
                            if temiz_no and temiz_no in ana_direk_set and temiz_no not in koordinat_bilgileri:
                                lat = row[lat_col]
                                lon = row[lon_col]
                                
                                # Koordinat değerlerini kontrol et
                                if pd.notna(lat) and pd.notna(lon):
                                    koordinat_bilgileri[temiz_no] = {
                                        'lat': lat,
                                        'lon': lon
                                    }
                                    og_koordinat_sayisi += 1
                        
                        if og_koordinat_sayisi > 0:
                            self.txt_sonuc.insert(tk.END, f"✓ OG Excel'inden {og_koordinat_sayisi} koordinat bulundu\n")
                        
                except Exception as e:
                    self.txt_sonuc.insert(tk.END, f"⚠️ OG Excel koordinat okuma hatası: {str(e)}\n")

            # İstatistikler
            toplam_koordinat = len(koordinat_bilgileri)
            koordinatsiz = len(ana_direk_set) - toplam_koordinat
            
            self.txt_sonuc.insert(tk.END, f"📍 Toplam {toplam_koordinat} direk için koordinat bulundu\n")
            if koordinatsiz > 0:
                self.txt_sonuc.insert(tk.END, f"📍 {koordinatsiz} direk için koordinat bulunamadı\n")
                
        except Exception as e:
            self.txt_sonuc.insert(tk.END, f"⚠️ Koordinat bilgisi alma hatası: {str(e)}\n")
        
        return koordinat_bilgileri



    def rapor_ekrani_ac(self, excel_dosyasi, hat_tarihsiz_istatistikleri=None):
        try:
            # Mevcut rapor ekranını kapat
            if hasattr(self, 'rapor_ekrani'):
                try:
                    self.rapor_ekrani.destroy()
                    time.sleep(0.5)  # Kapanması için bekle
                except:
                    pass
            
            # Yeni rapor ekranı oluştur - main thread'de
            def create_report():
                try:
                    self.rapor_ekrani = RaporEkrani(self.root, excel_dosyasi, hat_tarihsiz_istatistikleri)
                except Exception as e:
                    messagebox.showerror("Hata", f"Rapor ekranı oluşturulamadı: {str(e)}")
            
            # 1 saniye bekle ve oluştur
            self.root.after(1000, create_report)
            
        except Exception as e:
            messagebox.showwarning("Uyarı", f"Rapor ekranı açılamadı: {str(e)}")
    
    def rapor_goster(self):
        if not self.kaydetme_yeri.get() or not os.path.exists(self.kaydetme_yeri.get()):
            messagebox.showerror("Hata", "Önce bir Excel dosyası oluşturun!")
            return
        
        # Excel dosyasından hat istatistiklerini oku
        try:
            df_hat_ozet = pd.read_excel(self.kaydetme_yeri.get(), sheet_name='Hat Özeti')
            hat_tarihsiz_istatistikleri = {}
            
            for index, row in df_hat_ozet.iterrows():
                hat_anahtari = (row['İlçe'], row['Hat Adı'])
                # Excel'de tarihsiz yüzde sütunu yoksa 0 olarak kabul et
                hat_tarihsiz_istatistikleri[hat_anahtari] = 0
                
        except:
            hat_tarihsiz_istatistikleri = {}
        
        self.rapor_ekrani_ac(self.kaydetme_yeri.get(), hat_tarihsiz_istatistikleri)
    
    def auto_adjust_column_widths(self, worksheet):
        for column in worksheet.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)
            
            for cell in column:
                try:
                    if cell.value:
                        cell_length = len(str(cell.value))
                        if cell_length > max_length:
                            max_length = cell_length
                except:
                    pass
            
            header_cell = worksheet[f"{column_letter}1"]
            if header_cell.value:
                header_length = len(str(header_cell.value))
                if header_length > max_length:
                    max_length = header_length
            
            adjusted_width = min(max_length + 3, 100)
            adjusted_width = max(adjusted_width, 12)
            
            worksheet.column_dimensions[column_letter].width = adjusted_width
    
    def direk_tipi_bilgilerini_al(self, veri, tespit_verileri=None):
        """GÜNCELLENDİ: ÖNCE Nihai_Direk, SONRA TRAFO, SONRA diğer Excel'ler"""
        try:
            # 1. Ana veriden direk numaralarını hazırla
            ana_direk_set = set()
            for row_data in veri:
                direk_no = str(row_data.get('Direk No', ''))
                temiz_no = self.direk_no_temizle(direk_no)
                if temiz_no and temiz_no != 'nan':
                    ana_direk_set.add(temiz_no)

            self.txt_sonuc.insert(tk.END, f"📊 Ana Excel'de {len(ana_direk_set)} direk numarası bulundu\n")
            
            if not ana_direk_set:
                return {}

            direk_tipleri = {}
            trafo_direkleri = {}
            tespit_bilgileri = {}
            nihai_direkler = {}  # YENİ: Nihai direkleri ayrıca kaydet
            
            # AŞAMA 1: ÖNCE NİHAİ_DİREK EXCEL'İNİ KONTROL ET
            if hasattr(self, 'nihai_direk_yolu') and self.nihai_direk_yolu and os.path.exists(self.nihai_direk_yolu):
                try:
                    self.txt_sonuc.insert(tk.END, "🔍 Nihai_Direk Excel'i kontrol ediliyor...\n")
                    df_nihai = pd.read_excel(self.nihai_direk_yolu)
                    
                    # Direk numarası sütununu bul
                    direkno_col = None
                    for col in df_nihai.columns:
                        col_upper = col.upper()
                        if 'DIREK' in col_upper and ('NO' in col_upper or 'NUMARA' in col_upper or 'ID' in col_upper or 'NUM' in col_upper):
                            direkno_col = col
                            break
                    
                    if direkno_col:
                        nihai_direk_sayisi = 0
                        for index, row in df_nihai.iterrows():
                            direk_no = str(row[direkno_col]) if pd.notna(row[direkno_col]) else ""
                            temiz_no = self.direk_no_temizle(direk_no)
                            
                            if temiz_no and temiz_no in ana_direk_set:
                                # NİHAİ DİREK BULUNDU - EN YÜKSEK ÖNCELİK
                                direk_tipleri[temiz_no] = 'Nihai Direk'
                                nihai_direkler[temiz_no] = True  # YENİ: Nihai direkleri kaydet
                                nihai_direk_sayisi += 1
                        
                        self.txt_sonuc.insert(tk.END, f"🏆 Nihai_Direk Excel'inden {nihai_direk_sayisi} direk bulundu\n")
                        
                        # DEBUG: Bulunan Nihai direkleri göster
                        if nihai_direkler:
                            nihai_list = list(nihai_direkler.keys())
                            self.txt_sonuc.insert(tk.END, f"🔍 Bulunan Nihai direkleri: {nihai_list[:10]}\n")
                            if '7' in nihai_list:
                                self.txt_sonuc.insert(tk.END, f"✅ Direk 7 Nihai Direk olarak işaretlendi!\n")
                    else:
                        self.txt_sonuc.insert(tk.END, "⚠️ Nihai_Direk Excel'inde direk no sütunu bulunamadı\n")
                        
                except Exception as e:
                    self.txt_sonuc.insert(tk.END, f"⚠️ Nihai_Direk Excel okuma hatası: {str(e)}\n")
            else:
                self.txt_sonuc.insert(tk.END, "ℹ️ Nihai_Direk Excel'i bulunamadı veya seçilmedi\n")
            
            # AŞAMA 2: NİHAİ_DİREK'TE OLMAYAN DİREKLER İÇİN TRAFO KONTROLÜ (OG Excel'inden)
            nihai_olmayan_direkler = ana_direk_set - set(direk_tipleri.keys())
            
            if nihai_olmayan_direkler and self.direk_og_yolu and os.path.exists(self.direk_og_yolu):
                try:
                    self.txt_sonuc.insert(tk.END, "🔍 TRAFO DİREĞİ kontrol ediliyor...\n")
                    df_og = pd.read_excel(self.direk_og_yolu)
                    
                    direkno_col = None
                    tip_col = None
                    
                    for col in df_og.columns:
                        col_upper = col.upper()
                        if 'DIREK' in col_upper and ('NO' in col_upper or 'NUMARA' in col_upper):
                            direkno_col = col
                        elif 'TIP' in col_upper or 'TİP' in col_upper or 'TYPE' in col_upper:
                            tip_col = col
                    
                    if direkno_col and tip_col:
                        trafo_sayisi = 0
                        for index, row in df_og.iterrows():
                            direk_no = str(row[direkno_col])
                            temiz_no = self.direk_no_temizle(direk_no.strip())
                            tip_deger = str(row[tip_col]).strip().upper() if pd.notna(row[tip_col]) else ""
                            
                            # SADECE nihai olmayan direkleri kontrol et
                            if temiz_no and temiz_no in nihai_olmayan_direkler:
                                if "TRAFO" in tip_deger or "TRAFO DİREĞİ" in tip_deger.upper():
                                    # TRAFO DİREĞİ BULUNDU - HEM M HEM N sütunları için kaydet
                                    
                                    direk_tipleri[temiz_no] = "Trafo Direği"  # M sütunu için
                                    trafo_sayisi += 1
                        
                        self.txt_sonuc.insert(tk.END, f"🏭 OG Excel'inden {trafo_sayisi} trafo direği bulundu\n")
                        
                except Exception as e:
                    self.txt_sonuc.insert(tk.END, f"⚠️ TRAFO kontrolü okuma hatası: {str(e)}\n")
            
            # AŞAMA 3: MÜŞTEREK DİREK KONTROLÜ (nihai ve trafo olmayanlar için)
            trafo_olmayan_direkler = nihai_olmayan_direkler - set(trafo_direkleri.keys())
            
            if trafo_olmayan_direkler and self.direk_musterek_yolu and os.path.exists(self.direk_musterek_yolu):
                try:
                    self.txt_sonuc.insert(tk.END, "🔍 Müşterek Direk kontrol ediliyor...\n")
                    df_musterek = pd.read_excel(self.direk_musterek_yolu)
                    direkno_col = next((col for col in df_musterek.columns 
                                      if 'DIREKNO' in col.upper() or 'NO' in col.upper() or 'NUMARA' in col.upper()), None)
                    
                    if direkno_col:
                        musterek_direkler = set()
                        for direk_no in df_musterek[direkno_col].dropna().astype(str):
                            temiz_no = self.direk_no_temizle(direk_no.strip())
                            if temiz_no:
                                musterek_direkler.add(temiz_no)
                        
                        # SADECE nihai ve trafo olmayan direkleri kontrol et
                        eslesenler = trafo_olmayan_direkler.intersection(musterek_direkler)
                        for direk_no in eslesenler:
                            direk_tipleri[direk_no] = 'Müşterek Direk'
                        
                        self.txt_sonuc.insert(tk.END, f"🤝 Müşterek Excel'inde eşleşen: {len(eslesenler)} direk\n")
                        
                except Exception as e:
                    self.txt_sonuc.insert(tk.END, f"⚠️ Müşterek Excel okuma hatası: {str(e)}\n")
            
            # AŞAMA 4: OG DİREK KONTROLÜ (nihai, trafo ve müşterek olmayanlar için)
            kalan_direkler = trafo_olmayan_direkler - set(direk_tipleri.keys())
            
            if kalan_direkler and self.direk_og_yolu and os.path.exists(self.direk_og_yolu):
                try:
                    self.txt_sonuc.insert(tk.END, "🔍 OG Direk kontrol ediliyor...\n")
                    df_og = pd.read_excel(self.direk_og_yolu)
                    direkno_col = next((col for col in df_og.columns 
                                      if 'DIREKNO' in col.upper() or 'NO' in col.upper() or 'NUMARA' in col.upper()), None)
                    
                    if direkno_col:
                        og_direkler = set()
                        for direk_no in df_og[direkno_col].dropna().astype(str):
                            temiz_no = self.direk_no_temizle(direk_no.strip())
                            if temiz_no:
                                og_direkler.add(temiz_no)
                        
                        # Kalan direkleri kontrol et
                        eslesenler = kalan_direkler.intersection(og_direkler)
                        for direk_no in eslesenler:
                            direk_tipleri[direk_no] = 'Og Direk'
                        
                        self.txt_sonuc.insert(tk.END, f"⚡ OG Excel'inde eşleşen: {len(eslesenler)} direk\n")
                        
                except Exception as e:
                    self.txt_sonuc.insert(tk.END, f"⚠️ OG Excel okuma hatası: {str(e)}\n")
            
            # AŞAMA 5: TESPİT VERİLERİNİ İŞLE
            if tespit_verileri:
                for tespit_data in tespit_verileri:
                    ilce = tespit_data.get('İlçe', '').strip()
                    kaynak_klasor = tespit_data.get('Hat Adı', '').strip()
                    direk_no = tespit_data.get('Direk No', '')
                    temiz_no = self.direk_no_temizle(direk_no)
                    
                    if ilce and kaynak_klasor and temiz_no:
                        anahtar = (ilce, kaynak_klasor, temiz_no)
                        tespit_bilgileri[anahtar] = {
                            'oncelik': tespit_data.get('Öncelik', ''),
                            'tespit_notu': tespit_data.get('Tespit Notu', ''),
                            'yapildi_mi': tespit_data.get('Yapıldı mı?', '')
                        }
            
            # İSTATİSTİKLER
            eslesmeyenler = ana_direk_set - set(direk_tipleri.keys()) - set(trafo_direkleri.keys())
            
            self.txt_sonuc.insert(tk.END, f"\n📊 DIREK TİPİ DAĞILIMI:\n")
            self.txt_sonuc.insert(tk.END, f"   🏆 Nihai Direk: {len([k for k, v in direk_tipleri.items() if v == 'Nihai Direk'])}\n")
            self.txt_sonuc.insert(tk.END, f"   🤝 Müşterek Direk: {len([k for k, v in direk_tipleri.items() if v == 'Müşterek Direk'])}\n")
            self.txt_sonuc.insert(tk.END, f"   ⚡ Og Direk: {len([k for k, v in direk_tipleri.items() if v == 'Og Direk'])}\n")
            self.txt_sonuc.insert(tk.END, f"   🏭 Trafo Direği: {len(trafo_direkleri)}\n")
            self.txt_sonuc.insert(tk.END, f"   ❓ Tipi Belirlenemeyen: {len(eslesmeyenler)}\n")
            self.txt_sonuc.insert(tk.END, f"   📋 Tespit Bilgisi Bulunan: {len(tespit_bilgileri)} direk\n")
            
            # SONUÇ - nihai_direkler de eklenmeli
            sonuc = {
                'direk_tipleri': direk_tipleri,
                'trafo_direkleri': trafo_direkleri,
                'tespit_bilgileri': tespit_bilgileri,
                'nihai_direkler': nihai_direkler  # YENİ: Nihai direkleri de döndür
            }
            
            return sonuc
            
        except Exception as e:
            self.txt_sonuc.insert(tk.END, f"⚠️ Direk tipi hesaplama hatası: {str(e)}\n")
            return {'direk_tipleri': {}, 'trafo_direkleri': {}, 'tespit_bilgileri': {}, 'nihai_direkler': {}}


    def musterek_direk_listesi_yukle(self):
        """direk_musterek.xlsx dosyasından direk listesini yükler"""
        try:
            # Önce dosya yolunu kontrol et
            if not hasattr(self, 'direk_musterek_yolu') or not self.direk_musterek_yolu:
                self.log_mesaj_ekle("⚠️ Müşterek Excel dosya yolu tanımlanmamış")
                return set()
            
            dosya_yolu = self.direk_musterek_yolu
            if os.path.exists(dosya_yolu):
                self.log_mesaj_ekle(f"📄 Müşterek direk listesi yükleniyor: {dosya_yolu}")
                df = pd.read_excel(dosya_yolu)
                direk_listesi = set()
                
                # Direk No sütununu bul (çeşitli olasılıkları deneyelim)
                direkno_col = None
                for col in df.columns:
                    col_upper = str(col).upper().strip()
                    if 'DIREK' in col_upper and ('NO' in col_upper or 'NUMARA' in col_upper or 'ID' in col_upper or 'NUM' in col_upper):
                        direkno_col = col
                        break
                
                # Eğer yukarıdaki kriterlerle bulunamazsa, sütun adında 'NO' geçiyor mu diye bak
                if not direkno_col:
                    for col in df.columns:
                        col_upper = str(col).upper().strip()
                        if 'NO' in col_upper or 'ID' in col_upper or 'NUM' in col_upper:
                            direkno_col = col
                            break
                
                if direkno_col:
                    for direk_no in df[direkno_col]:
                        if pd.notna(direk_no):  # NaN değerleri atla
                            temiz_no = self.direk_no_temizle(str(direk_no))
                            if temiz_no and temiz_no != '' and temiz_no != 'NAN':
                                direk_listesi.add(temiz_no)
                    
                    self.log_mesaj_ekle(f"✅ {len(direk_listesi)} müşterek direk yüklendi")
                    return direk_listesi
                else:
                    self.log_mesaj_ekle(f"⚠️ Müşterek dosyasında direk no sütunu bulunamadı. Sütunlar: {list(df.columns)}")
                    return set()
            else:
                self.log_mesaj_ekle(f"⚠️ Müşterek Excel dosyası bulunamadı: {dosya_yolu}")
                return set()
        except Exception as e:
            self.log_mesaj_ekle(f"❌ Müşterek direk listesi yükleme hatası: {str(e)}")
            import traceback
            self.log_mesaj_ekle(f"🔍 Hata detayı: {traceback.format_exc()}")
            return set()

    def og_direk_listesi_yukle(self):
        """direk_og.xlsx dosyasından direk listesini yükler"""
        try:
            # Önce dosya yolunu kontrol et
            if not hasattr(self, 'direk_og_yolu') or not self.direk_og_yolu:
                self.log_mesaj_ekle("⚠️ OG Excel dosya yolu tanımlanmamış")
                return set()
            
            dosya_yolu = self.direk_og_yolu
            if os.path.exists(dosya_yolu):
                self.log_mesaj_ekle(f"📄 OG direk listesi yükleniyor: {dosya_yolu}")
                df = pd.read_excel(dosya_yolu)
                direk_listesi = set()
                
                # Direk No sütununu bul (çeşitli olasılıkları deneyelim)
                direkno_col = None
                for col in df.columns:
                    col_upper = str(col).upper().strip()
                    if 'DIREK' in col_upper and ('NO' in col_upper or 'NUMARA' in col_upper or 'ID' in col_upper or 'NUM' in col_upper):
                        direkno_col = col
                        break
                
                # Eğer yukarıdaki kriterlerle bulunamazsa, sütun adında 'NO' geçiyor mu diye bak
                if not direkno_col:
                    for col in df.columns:
                        col_upper = str(col).upper().strip()
                        if 'NO' in col_upper or 'ID' in col_upper or 'NUM' in col_upper:
                            direkno_col = col
                            break
                
                if direkno_col:
                    for direk_no in df[direkno_col]:
                        if pd.notna(direk_no):  # NaN değerleri atla
                            temiz_no = self.direk_no_temizle(str(direk_no))
                            if temiz_no and temiz_no != '' and temiz_no != 'NAN':
                                direk_listesi.add(temiz_no)
                    
                    self.log_mesaj_ekle(f"✅ {len(direk_listesi)} OG direk yüklendi")
                    return direk_listesi
                else:
                    self.log_mesaj_ekle(f"⚠️ OG dosyasında direk no sütunu bulunamadı. Sütunlar: {list(df.columns)}")
                    return set()
            else:
                self.log_mesaj_ekle(f"⚠️ OG Excel dosyası bulunamadı: {dosya_yolu}")
                return set()
        except Exception as e:
            self.log_mesaj_ekle(f"❌ OG direk listesi yükleme hatası: {str(e)}")
            import traceback
            self.log_mesaj_ekle(f"🔍 Hata detayı: {traceback.format_exc()}")
            return set()


    
    def tespit_excelini_oku(self):
        """Rapor Excel'indeki Birleştirilmiş_Tespit sayfasını okur (İlçe bilgisi ile)"""
        tespit_verileri = {}
        
        try:
            # Rapor Excel dosyasının yolunu bul
            rapor_dosya_yolu = self.kaydetme_yeri.get()
            
            if not rapor_dosya_yolu or not os.path.exists(rapor_dosya_yolu):
                self.txt_sonuc.insert(tk.END, "⚠️ Rapor Excel dosyası bulunamadı\n")
                return tespit_verileri
            
            # Birleştirilmiş_Tespit sayfasını oku
            try:
                df_tespit = pd.read_excel(rapor_dosya_yolu, sheet_name='Birleştirilmiş_Tespit')
            except:
                self.txt_sonuc.insert(tk.END, "⚠️ Birleştirilmiş_Tespit sayfası bulunamadı\n")
                return tespit_verileri
            
            # DEĞİŞTİRİLDİ: 'Direk ID' yerine 'Direk No'
            gerekli_sutunlar = ['İlçe', 'Direk No', 'Öncelik', 'Tespit Notu', 'Yapıldı mı?']  # DEĞİŞTİRİLDİ
            mevcut_sutunlar = df_tespit.columns.tolist()
            
            # Sütun eşleştirme
            sutun_eslesmeleri = {}
            for gerekli_sutun in gerekli_sutunlar:
                for mevcut_sutun in mevcut_sutunlar:
                    if gerekli_sutun.lower() in mevcut_sutun.lower():
                        sutun_eslesmeleri[gerekli_sutun] = mevcut_sutun
                        break
            
            # Tespit verilerini işle
            for index, row in df_tespit.iterrows():
                try:
                    # DEĞİŞTİRİLDİ: 'Direk ID' yerine 'Direk No'
                    direk_no_sutunu = sutun_eslesmeleri.get('Direk No')  # DEĞİŞTİRİLDİ
                    if not direk_no_sutunu:
                        continue
                        
                    direk_no = str(row[direk_no_sutunu])
                    temiz_no = self.direk_no_temizle(direk_no)
                    
                    if not temiz_no or temiz_no == 'nan':
                        continue
                    
                    # İlçe bilgisini al
                    ilce_sutunu = sutun_eslesmeleri.get('İlçe')
                    ilce = ""
                    if ilce_sutunu and pd.notna(row[ilce_sutunu]):
                        ilce = str(row[ilce_sutunu])
                    
                    # Öncelik, Tespit Notu ve Yapıldı mı?'yı al
                    oncelik = ""
                    tespit_notu = ""
                    yapildi_mi = ""
                    
                    oncelik_sutunu = sutun_eslesmeleri.get('Öncelik')
                    if oncelik_sutunu and pd.notna(row[oncelik_sutunu]):
                        oncelik = str(row[oncelik_sutunu])
                    
                    tespit_notu_sutunu = sutun_eslesmeleri.get('Tespit Notu')
                    if tespit_notu_sutunu and pd.notna(row[tespit_notu_sutunu]):
                        tespit_notu = str(row[tespit_notu_sutunu])
                    
                    yapildi_mi_sutunu = sutun_eslesmeleri.get('Yapıldı mı?')
                    if yapildi_mi_sutunu and pd.notna(row[yapildi_mi_sutunu]):
                        yapildi_mi = str(row[yapildi_mi_sutunu])
                    
                    # Tespit bilgilerini kaydet (ilçe bilgisi ile birlikte)
                    tespit_verileri[temiz_no] = {
                        'ilce': ilce,
                        'oncelik': oncelik,
                        'tespit_notu': tespit_notu,
                        'yapildi_mi': yapildi_mi
                    }
                    
                except Exception as e:
                    continue
            
            self.txt_sonuc.insert(tk.END, f"✅ Birleştirilmiş_Tespit sayfasından {len(tespit_verileri)} direk için bilgi alındı\n")
            
        except Exception as e:
            self.txt_sonuc.insert(tk.END, f"⚠️ Tespit Excel okuma hatası: {str(e)}\n")
        
        return tespit_verileri



    
    def direk_no_temizle(self, direk_no):
        """Direk numarasını temizleyerek standart formata getirir"""
        if not direk_no or str(direk_no).lower() in ['nan', 'none', 'null', '']:
            return ""  # TÜM boş değerler için "" döndür
        
        # String'e çevir ve temizle
        direk_no = str(direk_no).strip()
        
        # Eğer boş string ise None döndür
        if not direk_no:
            return None
        
        # Özel durum: "0" veya "0000" gibi değerleri kontrol et
        if direk_no.replace('0', '') == '':
            return '0'
        
        # Sayısal değerleri kontrol et (örn: 123.0 -> 123)
        try:
            # Float ise integer'a çevir, baştaki sıfırları kaldır
            if '.' in direk_no:
                cleaned = str(int(float(direk_no)))
            else:
                # Doğrudan integer'a çevirmeye çalış
                cleaned = str(int(direk_no))
            
            # Baştaki sıfırları kaldır, ama "0" değerini koru
            cleaned = cleaned.lstrip('0') or '0'
            return cleaned
            
        except (ValueError, TypeError):
            # Sayıya çevrilemezse, sadece rakamları al
            sadece_rakamlar = ''.join(filter(str.isdigit, direk_no))
            if sadece_rakamlar:
                # Baştaki sıfırları kaldır, ama "0" değerini koru
                sadece_rakamlar = sadece_rakamlar.lstrip('0') or '0'
                return sadece_rakamlar
            
            # Hiç rakam yoksa orijinal değeri döndür (temizlenmiş)
            return direk_no

    
    
    def excel_ac(self, dosya_yolu):
        try:
            if sys.platform == "win32":
                os.startfile(dosya_yolu)
            elif sys.platform == "darwin":
                subprocess.run(["open", dosya_yolu])
            else:
                subprocess.run(["xdg-open", dosya_yolu])
            return True
        except Exception as e:
            print(f"Excel açılırken hata: {e}")
            return False


    
    

    

    def isim_sablonu_olustur(self, foto_info, sira_no, dosya_sayaclari):
        orijinal_adi = os.path.splitext(foto_info['orijinal_adi'])[0]
        uzanti = os.path.splitext(foto_info['orijinal_adi'])[1]
        klasor_adi = foto_info['klasor_adi']
        kaynak_klasor_adi = os.path.basename(foto_info['kaynak_klasor'])
    
        # İsim şablonunu al
        sablon = self.isim_sablonu.get().strip()
    
        # Eğer şablon sadece "DJI" ise, DJI_1'den başlayacak şekilde değiştir
        if sablon == "DJI":
            sablon = "DJI_{sira_no}"
    
        # Şablonu işle
        yeni_isim = sablon
        yeni_isim = yeni_isim.replace('{orijinal_adi}', orijinal_adi)
        yeni_isim = yeni_isim.replace('{klasor_adi}', klasor_adi)
        yeni_isim = yeni_isim.replace('{kaynak_klasr}', kaynak_klasor_adi)
        yeni_isim = yeni_isim.replace('{sira_no}', str(sira_no))
    
        # Geçersiz karakterleri temizle
        yeni_isim = re.sub(r'[<>:"/\\|?*]', '_', yeni_isim)
    
        return yeni_isim + uzanti.lower()


    def hat_bazi_tarihsiz_hesapla(self, veri):
        """Her hat için tarihsiz fotoğraf yüzdesini hesaplar"""
        hat_istatistikleri = {}
        
        # 1. Önce hatları grupla ve doğrudan sayım yap
        for row_data in veri:
            aob = row_data.get('Aob', '')
            hat_adi = row_data.get('Hat adı', '')
            hat_anahtari = (aob, hat_adi)
            
            if hat_anahtari not in hat_istatistikleri:
                hat_istatistikleri[hat_anahtari] = {
                    'toplam_foto': 0,
                    'tarihsiz_foto': 0
                }
            
            # Doğrudan bu klasörün istatistiklerini hesapla ve ekle
            klasor_yolu = row_data.get('Klasör Yolu', '')
            if klasor_yolu and os.path.exists(klasor_yolu):
                klasor_istatistik = self.klasor_tarihsiz_say(klasor_yolu)
                hat_istatistikleri[hat_anahtari]['toplam_foto'] += klasor_istatistik['toplam']
                hat_istatistikleri[hat_anahtari]['tarihsiz_foto'] += klasor_istatistik['tarihsiz']
        
        # 2. Her hat için tarihsiz fotoğraf yüzdesini hesapla
        hat_tarihsiz_oranlari = {}
        for hat_anahtari, istatistik in hat_istatistikleri.items():
            toplam_foto = istatistik['toplam_foto']
            tarihsiz_foto = istatistik['tarihsiz_foto']
            
            # Yüzde hesapla
            if toplam_foto > 0:
                tarihsiz_yuzde = (tarihsiz_foto / toplam_foto) * 100
            else:
                tarihsiz_yuzde = 0
                
            hat_tarihsiz_oranlari[hat_anahtari] = round(tarihsiz_yuzde, 1)
            
            # Debug info
            print(f"🔍 {hat_anahtari}: {toplam_foto} fotoğraf, {tarihsiz_foto} tarihsiz (%{tarihsiz_yuzde:.1f})")
        
        return hat_tarihsiz_oranlari

    def klasor_tarihsiz_say(self, klasor_yolu):
        """Bir klasördeki tarihsiz fotoğrafları sayar - HIZLI versiyon"""
        foto_uzantilari = ('.jpg', '.jpeg', '.png', '.tiff', '.tif', '.bmp')
        toplam = 0
        tarihsiz = 0
        
        try:
            for root, dirs, files in os.walk(klasor_yolu):
                for dosya in files:
                    if any(dosya.lower().endswith(ext) for ext in foto_uzantilari):
                        toplam += 1
                        dosya_yolu = os.path.join(root, dosya)
                        if not self.exif_tarihi_var_mi_hizli(dosya_yolu):
                            tarihsiz += 1
        except Exception as e:
            print(f"Klasör tarama hatası {klasor_yolu}: {e}")
        
        return {'toplam': toplam, 'tarihsiz': tarihsiz}

    def exif_tarihi_var_mi_hizli(self, dosya_yolu):
        """EXIF tarihi kontrolü - HIZLI versiyon"""
        try:
            with Image.open(dosya_yolu) as img:
                exif_data = img.getexif()
                if exif_data:
                    datetime_original = exif_data.get(36867)  # DateTimeOriginal
                    datetime_digitized = exif_data.get(36868)  # DateTimeDigitized  
                    datetime_normal = exif_data.get(306)       # DateTime
                    
                    if datetime_original or datetime_digitized or datetime_normal:
                        return True
                return False
        except:
            return False



    def excel_klasor_sec(self):
        """Klasör seçme dialogunu açar"""
        try:
            klasor = filedialog.askdirectory(
                title="Excel Dosyalarının Bulunduğu Klasörü Seçin"
            )
            if klasor:
                self.excel_birlestirme_klasor = klasor
                
                # SADECE klasör seçimi alanını güncelle, kaydetme yerini değil
                for widget in self.content_frame.winfo_children():
                    if isinstance(widget, tk.Frame) and hasattr(widget, 'winfo_children'):
                        # İlk frame'i bul (klasör seçimi frame'i)
                        frame_children = widget.winfo_children()
                        if len(frame_children) >= 2:  # Label ve Entry var
                            entry_widget = frame_children[1]  # Entry widget'ı
                            if isinstance(entry_widget, tk.Entry):
                                # Sadece klasör seçimi entry'sini güncelle
                                if "Klasör Seç:" in str(frame_children[0].cget('text')):
                                    entry_widget.config(state='normal')
                                    entry_widget.delete(0, tk.END)
                                    entry_widget.insert(0, klasor)
                                    entry_widget.config(state='readonly')
                                    break
                
                self.log_mesaj_ekle(f"✅ Klasör seçildi: {klasor}")
                # Klasörü otomatik tara
                self.excel_klasoru_tara()
                
        except Exception as e:
            messagebox.showerror("Hata", f"Klasör seçilirken hata oluştu: {str(e)}")

    def excel_klasoru_tara(self):
        """Seçili klasördeki Excel dosyalarını tarar"""
        if not hasattr(self, 'excel_birlestirme_klasor') or not self.excel_birlestirme_klasor:
            messagebox.showwarning("Uyarı", "Lütfen önce bir klasör seçin!")
            return
        
        try:
            self.excel_birlestirme_dosyalari = []
            
            if self.alt_klasor_tara_var.get():
                # Alt klasörler dahil
                for root, dirs, files in os.walk(self.excel_birlestirme_klasor):
                    for dosya in files:
                        if dosya.lower().endswith(('.xlsx', '.xls')):
                            tam_yol = os.path.join(root, dosya)
                            self.excel_birlestirme_dosyalari.append(tam_yol)
            else:
                # Sadece seçili klasör
                for dosya in os.listdir(self.excel_birlestirme_klasor):
                    if dosya.lower().endswith(('.xlsx', '.xls')):
                        tam_yol = os.path.join(self.excel_birlestirme_klasor, dosya)
                        self.excel_birlestirme_dosyalari.append(tam_yol)
            
            self.excel_birlestirme_toplam_dosya.set(len(self.excel_birlestirme_dosyalari))
            self.excel_birlestirme_islenen_dosya.set(0)
            
            self.log_mesaj_ekle(f"✅ {len(self.excel_birlestirme_dosyalari)} Excel dosyası bulundu")
            
            if self.excel_birlestirme_dosyalari:
                self.log_mesaj_ekle("📋 Bulunan dosyalar:")
                for dosya in self.excel_birlestirme_dosyalari[:10]:  # İlk 10'u göster
                    goreli_yol = os.path.relpath(dosya, self.excel_birlestirme_klasor)
                    self.log_mesaj_ekle(f"   📄 {goreli_yol}")
                
                if len(self.excel_birlestirme_dosyalari) > 10:
                    self.log_mesaj_ekle(f"   ... ve {len(self.excel_birlestirme_dosyalari) - 10} dosya daha")
            
        except Exception as e:
            self.log_mesaj_ekle(f"❌ Klasör taranırken hata: {str(e)}")

    def excel_kaydetme_yeri_sec(self):
        """Kaydetme yeri seçme dialogunu açar"""
        try:
            baslangic_dizin = os.path.expanduser("~")
            if hasattr(self, 'excel_birlestirme_klasor') and self.excel_birlestirme_klasor:
                baslangic_dizin = self.excel_birlestirme_klasor
                
            dosya = filedialog.asksaveasfilename(
                title="Birleştirilmiş Excel Dosyasını Kaydet",
                initialdir=baslangic_dizin,
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
            )
            if dosya:
                self.excel_birlestirme_kaydetme_yeri.set(dosya)
                self.log_mesaj_ekle(f"💾 Kaydetme yeri: {dosya}")
        except Exception as e:
            messagebox.showerror("Hata", f"Kaydetme yeri seçilirken hata oluştu: {str(e)}")

    def excel_birlestirmeyi_baslat(self):
        """Excel birleştirme işlemini başlatır"""
        if not hasattr(self, 'excel_birlestirme_klasor') or not self.excel_birlestirme_klasor:
            messagebox.showerror("Hata", "Lütfen önce bir klasör seçin!")
            return
            
        if not self.excel_birlestirme_dosyalari:
            messagebox.showwarning("Uyarı", "Seçilen klasörde Excel dosyası bulunamadı!")
            return
            
        if not self.excel_birlestirme_kaydetme_yeri.get():
            messagebox.showerror("Hata", "Lütfen kaydetme yerini seçin!")
            return
        
        self.calisma_devam_ediyor = True
        self.lbl_durum.config(text="Excel dosyaları birleştiriliyor...", fg=self.colors['accent'])
        self.txt_sonuc.delete(1.0, tk.END)
        self.txt_sonuc.insert(tk.END, "Excel dosyaları birleştiriliyor...\n")
        self.excel_ilerleme['value'] = 0
        self.ilerleme_yuzde.set("%0")
        self.root.update()
        
        threading.Thread(target=self.excel_birlestir_thread, daemon=True).start()

    def excel_birlestir_thread(self):
        """Excel dosyalarını birleştirme işlemini thread'de çalıştırır - Hat Adı BİLGİLİ VERSİYON"""
        try:
            tum_veriler = []
            basarili_dosyalar = 0
            hatali_dosyalar = 0
            
            # İSTATİSTİK DEĞİŞKENLERİ
            toplam_satir = 0
            toplam_cbs_satir = 0
            toplam_aob_satir = 0
            toplam_bos_satir = 0
            toplam_eklenen_satir = 0
            toplam_eklenmeyen_satir = 0
            
            # İSTENEN SÜTUN LİSTESİ - "SIRA NO" KALDIRILDI, Hat Adı SÜTUNU EKLENDİ
            istenen_sutunlar = [
                'Hat Adı', 'Direk No', 'Tespit Notu', 'Tespit Kategorisi', 
                'Öncelik', 'Fotoğraf Yolu', 'Enlem', 'Boylam', 'Birim'
            ]
            
            self.excel_ilerleme['maximum'] = len(self.excel_birlestirme_dosyalari)
            self.excel_ilerleme['value'] = 0
            
            for i, dosya in enumerate(self.excel_birlestirme_dosyalari):
                try:
                    self.lbl_durum.config(text=f"İşleniyor: {os.path.basename(dosya)}")
                    
                    # Excel dosyasını header=None ile oku (başlık satırı olmadan)
                    df = pd.read_excel(dosya, header=None)
                    
                    # GERÇEK BAŞLIK SATIRINI BUL
                    baslik_satiri_idx = None
                    for idx in range(min(15, len(df))):  # İlk 15 satırı kontrol et
                        satir = df.iloc[idx]
                        satir_metin = ' '.join(str(x).strip() for x in satir.values if pd.notna(x) and str(x).strip())
                        
                        # İstenen sütun isimlerini ara (büyük/küçük harf duyarsız)
                        if all(any(sutun_adi.lower() in str(x).lower() for x in satir.values if pd.notna(x)) 
                               for sutun_adi in istenen_sutunlar if sutun_adi and sutun_adi != 'Hat Adı'):
                            baslik_satiri_idx = idx
                            break
                    
                    if baslik_satiri_idx is not None:
                        # Başlık satırını bulduk, onu kullanarak tekrar oku
                        df = pd.read_excel(dosya, header=baslik_satiri_idx)
                        self.log_mesaj_ekle(f"ℹ️ {os.path.basename(dosya)} - Başlık {baslik_satiri_idx + 1}. satırda bulundu")
                    else:
                        # Başlık bulunamazsa, ilk satırı başlık olarak kullan
                        df = pd.read_excel(dosya, header=0)
                        self.log_mesaj_ekle(f"ℹ️ {os.path.basename(dosya)} - İlk satır başlık olarak kullanıldı")
                    
                    # SADECE İSTENEN 9 SÜTUNU SEÇ - "SIRA NO" KALDIRILDI, Hat Adı EKLENDİ
                    mevcut_sutunlar = df.columns.tolist()
                    secilecek_sutunlar = []
                    sutun_eslesmeleri = {}
                    
                    # İstenen sütunları bul (büyük/küçük harf duyarsız ve kısmi eşleşme)
                    for istenen_sutun in istenen_sutunlar:
                        # "Hat Adı" sütunu özel işlem - mevcut Excel'de yok, biz ekleyeceğiz
                        if istenen_sutun == 'Hat Adı':
                            continue
                            
                        eslesen_sutun = None
                        for mevcut_sutun in mevcut_sutunlar:
                            mevcut_sutun_clean = str(mevcut_sutun).strip().lower()
                            istenen_sutun_clean = str(istenen_sutun).strip().lower()
                            
                            # Tam eşleşme
                            if mevcut_sutun_clean == istenen_sutun_clean:
                                eslesen_sutun = mevcut_sutun
                                break
                            # Kısmi eşleşme
                            elif istenen_sutun_clean in mevcut_sutun_clean:
                                eslesen_sutun = mevcut_sutun
                                break
                        
                        if eslesen_sutun:
                            secilecek_sutunlar.append(eslesen_sutun)
                            sutun_eslesmeleri[eslesen_sutun] = istenen_sutun
                        else:
                            # Eşleşen sütun yoksa, boş sütun ekleyeceğiz
                            pass
                    
                    # SADECE BULUNAN SÜTUNLARI SEÇ - DİĞERLERİNİ TAMAMEN SİL
                    df_filtreli = pd.DataFrame()
                    
                    # ÖNCE TÜM VERİ SATIRLARI İÇİN Hat Adı BİLGİSİNİ HAZIRLA
                    # Dosyanın bulunduğu klasör adını al
                    dosya_klasoru = os.path.basename(os.path.dirname(dosya))
                    
                    for istenen_sutun in istenen_sutunlar:
                        # "Hat Adı" sütunu özel işlem - TÜM SATIRLARA AYNI DEĞERİ VER
                        if istenen_sutun == 'Hat Adı':
                            continue
                            
                        # Bu sütun için eşleşme var mı kontrol et
                        eslesen_mevcut_sutun = None
                        for mevcut_sutun, standart_sutun in sutun_eslesmeleri.items():
                            if standart_sutun == istenen_sutun:
                                eslesen_mevcut_sutun = mevcut_sutun
                                break
                        
                        if eslesen_mevcut_sutun and eslesen_mevcut_sutun in df.columns:
                            # DİREK ID SÜTUNU ÖZEL İŞLEM
                            if istenen_sutun == 'Direk ID':
                                # Direk ID sütununu özel işle
                                df_filtreli[istenen_sutun] = self.direk_id_isle(df[eslesen_mevcut_sutun])
                            else:
                                df_filtreli[istenen_sutun] = df[eslesen_mevcut_sutun]
                        else:
                            # Sütun bulunamadı, boş sütun oluştur
                            df_filtreli[istenen_sutun] = [""] * len(df)
                    
                    # ŞİMDİ Hat Adı SÜTUNUNU EKLE - TÜM SATIRLARA AYNI DEĞER
                    if not df_filtreli.empty:
                        df_filtreli['Hat Adı'] = dosya_klasoru
                        # Sütun sırasını düzelt: Hat Adı ilk sütun olmalı
                        sutun_sirasi = ['Hat Adı'] + [sutun for sutun in istenen_sutunlar if sutun != 'Hat Adı']
                        df_filtreli = df_filtreli[sutun_sirasi]
                    
                    # İSTATİSTİK: Toplam satır sayısı (FİLTRE ÖNCESİ)
                    dosya_toplam_satir = len(df_filtreli)
                    toplam_satir += dosya_toplam_satir
                    
                    # BİRİM SÜTUNU İSTATİSTİKLERİ (FİLTRE ÖNCESİ)
                    if 'Birim' in df_filtreli.columns:
                        # "Birim" sütununu temizle
                        birim_serisi = df_filtreli['Birim'].astype(str).str.strip()
                        
                        # İSTATİSTİK: Filtreleme öncesi birim dağılımı
                        dosya_cbs_satir = birim_serisi.str.lower().str.contains('cbs', na=False).sum()
                        dosya_aob_satir = (birim_serisi.str.lower() == 'aob').sum()
                        dosya_bos_satir = ((birim_serisi == '') | (birim_serisi == 'nan') | 
                                         (birim_serisi == 'None') | df_filtreli['Birim'].isna()).sum()
                        
                        toplam_cbs_satir += dosya_cbs_satir
                        toplam_aob_satir += dosya_aob_satir
                        toplam_bos_satir += dosya_bos_satir
                        
                        # FİLTRELEME: Sadece "Aob" varyasyonları ve boş hücreler
                        def is_valid_birim(birim_deger):
                            if pd.isna(birim_deger) or birim_deger in ['', 'nan', 'None']:
                                return True
                            birim_str = str(birim_deger).strip().lower()
                            return birim_str == 'aob'
                        
                        birim_mask = birim_serisi.apply(is_valid_birim)
                        birim_mask.index = df_filtreli.index
                        
                        # Filtreyi uygula
                        df_filtreli = df_filtreli[birim_mask]
                        
                        # İSTATİSTİK: Eklenen ve eklenmeyen satırlar
                        dosya_eklenen_satir = len(df_filtreli)
                        dosya_eklenmeyen_satir = dosya_toplam_satir - dosya_eklenen_satir
                        
                        toplam_eklenen_satir += dosya_eklenen_satir
                        toplam_eklenmeyen_satir += dosya_eklenmeyen_satir
                        
                        # Filtreleme istatistikleri
                        self.log_mesaj_ekle(f"🔍 {os.path.basename(dosya)} - Birim filtresi:")
                        self.log_mesaj_ekle(f"   📊 Toplam: {dosya_toplam_satir} satır")
                        self.log_mesaj_ekle(f"   ✅ Eklenecek: {dosya_eklenen_satir} satır")
                        self.log_mesaj_ekle(f"   ❌ Eklenmeyecek: {dosya_eklenmeyen_satir} satır")
                        self.log_mesaj_ekle(f"   📋 Birim dağılımı - CBS: {dosya_cbs_satir}, AOB: {dosya_aob_satir}, Boş: {dosya_bos_satir}")
                    
                    # BAŞLIK SATIRLARINI VE BOŞ SATIRLARI FİLTRELE
                    if not df_filtreli.empty:
                        mask = pd.Series([True] * len(df_filtreli), index=df_filtreli.index)
                        
                        for sutun in df_filtreli.columns:
                            if df_filtreli[sutun].dtype == 'object' and sutun != 'Hat Adı':
                                # Başlık benzeri metinler içeren satırları bul
                                try:
                                    baslik_mask = ~df_filtreli[sutun].astype(str).str.contains(
                                        '|'.join([s for s in istenen_sutunlar if s != 'Hat Adı'] + ['Direk Numarası', 'YAPILACAK MI', 'K2']), 
                                        case=False, na=False
                                    )
                                    # Mask'ı hizala
                                    baslik_mask.index = df_filtreli.index
                                    mask = mask & baslik_mask
                                except Exception as e:
                                    self.log_mesaj_ekle(f"⚠️ Başlık filtresi hatası: {str(e)}")
                                    continue
                        
                        # Filtreyi uygula - mask'ı DataFrame index'i ile hizala
                        mask.index = df_filtreli.index
                        df_filtreli = df_filtreli[mask]
                        
                        # Tümü boş satırları sil (Hat Adı hariç)
                        bos_sutunlar = [col for col in df_filtreli.columns if col != 'Hat Adı']
                        df_filtreli = df_filtreli.dropna(subset=bos_sutunlar, how='all')
                        
                        # Hat Adı SÜTUNUNU KONTROL ET VE GÜNCELLE
                        if not df_filtreli.empty:
                            # Filtreleme sonrası kalan satır sayısına göre Hat Adı sütununu güncelle
                            df_filtreli['Hat Adı'] = dosya_klasoru
                    
                    if not df_filtreli.empty:
                        tum_veriler.append(df_filtreli)
                        basarili_dosyalar += 1
                        
                        goreli_yol = os.path.relpath(dosya, self.excel_birlestirme_klasor)
                        self.log_mesaj_ekle(f"✅ {goreli_yol} - {len(df_filtreli)} satır eklendi (Kaynak: {dosya_klasoru})")
                        
                    else:
                        hatali_dosyalar += 1
                        goreli_yol = os.path.relpath(dosya, self.excel_birlestirme_klasor)
                        self.log_mesaj_ekle(f"⚠️ {goreli_yol} - Filtreleme sonrası veri kalmadı")
                    
                except Exception as e:
                    hatali_dosyalar += 1
                    goreli_yol = os.path.relpath(dosya, self.excel_birlestirme_klasor)
                    self.log_mesaj_ekle(f"❌ {goreli_yol} - Hata: {str(e)}")
                    import traceback
                    self.log_mesaj_ekle(f"🔍 Hata detayı: {traceback.format_exc()}")
                
                # İlerlemeyi güncelle
                self.excel_birlestirme_islenen_dosya.set(i + 1)
                self.excel_ilerleme['value'] = i + 1
                ilerleme = ((i + 1) / len(self.excel_birlestirme_dosyalari)) * 100
                self.ilerleme_yuzde.set(f"%{int(ilerleme)}")
                self.root.update()
            
            # Tüm verileri birleştir
            if tum_veriler:
                self.lbl_durum.config(text="Veriler birleştiriliyor...")
                self.log_mesaj_ekle("📊 Veriler birleştiriliyor...")
                
                birlesik_df = pd.concat(tum_veriler, ignore_index=True)
                
                # SON KONTROL: BAŞLIK SATIRLARINI TEKRAR FİLTRELE
                if not birlesik_df.empty:
                    mask = pd.Series([True] * len(birlesik_df), index=birlesik_df.index)
                    
                    for sutun in birlesik_df.columns:
                        if birlesik_df[sutun].dtype == 'object' and sutun != 'Hat Adı':
                            try:
                                baslik_mask = ~birlesik_df[sutun].astype(str).str.contains(
                                    '|'.join([s for s in istenen_sutunlar if s != 'Hat Adı'] + ['Direk Numarası', 'YAPILACAK MI', 'K2']), 
                                    case=False, na=False
                                )
                                baslik_mask.index = birlesik_df.index
                                mask = mask & baslik_mask
                            except Exception as e:
                                self.log_mesaj_ekle(f"⚠️ Son başlık filtresi hatası: {str(e)}")
                                continue
                    
                    # Filtreyi uygula
                    mask.index = birlesik_df.index
                    birlesik_df = birlesik_df[mask]
                
                # KESİNLİKLE SADECE 9 SÜTUN OLMALI (Hat Adı dahil, Sıra No yok)
                birlesik_df = birlesik_df[istenen_sutunlar]
                
                # GERÇEK SONUÇLARI HESAPLA (BİRLEŞTİRME SONRASI)
                gercek_toplam_satir = len(birlesik_df)
                if 'Birim' in birlesik_df.columns:
                    birim_serisi = birlesik_df['Birim'].astype(str).str.strip()
                    gercek_aob_satir = (birim_serisi.str.lower() == 'aob').sum()
                    gercek_bos_satir = ((birim_serisi == '') | (birim_serisi == 'nan') | 
                                      (birim_serisi == 'None') | birlesik_df['Birim'].isna()).sum()
                else:
                    gercek_aob_satir = 0
                    gercek_bos_satir = gercek_toplam_satir
                
                # Excel'e kaydet
                self.lbl_durum.config(text="Excel dosyası kaydediliyor...")
                self.log_mesaj_ekle("💾 Excel dosyası kaydediliyor...")
                
                with pd.ExcelWriter(self.excel_birlestirme_kaydetme_yeri.get(), engine='openpyxl') as writer:
                    birlesik_df.to_excel(writer, index=False, sheet_name='Birleştirilmiş_Veri')
                    
                    # Excel formatını özelleştir
                    workbook = writer.book
                    worksheet = writer.sheets['Birleştirilmiş_Veri']
                    
                    # Başlık satırını biçimlendir
                    from openpyxl.styles import Font, PatternFill, Alignment
                    header_font = Font(bold=True, color="FFFFFF")
                    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
                    
                    for col_num, value in enumerate(istenen_sutunlar, 1):
                        cell = worksheet.cell(row=1, column=col_num)
                        cell.value = value
                        cell.font = header_font
                        cell.fill = header_fill
                        cell.alignment = Alignment(horizontal='left', vertical='center')  # Başlık sola hizalı
                    
                    # YENİ: TÜM HÜCRELERİ SOLA HİZALA
                    for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, 
                                                 min_col=1, max_col=len(istenen_sutunlar)):
                        for cell in row:
                            cell.alignment = Alignment(horizontal='left', vertical='center')
                    
                    # Sütun genişliklerini otomatik ayarla
                    for column in worksheet.columns:
                        max_length = 0
                        column_letter = column[0].column_letter
                        for cell in column:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass
                        adjusted_width = min(max_length + 2, 50)
                        worksheet.column_dimensions[column_letter].width = adjusted_width
                
                # DETAYLI İSTATİSTİKLERİ GÖSTER
                self.lbl_durum.config(text="Birleştirme tamamlandı ✓", fg=self.colors['success'])
                self.log_mesaj_ekle(f"\n🎉 BİRLEŞTİRME TAMAMLANDI!")
                self.log_mesaj_ekle("=" * 50)
                self.log_mesaj_ekle("📈 DETAYLI İSTATİSTİKLER")
                self.log_mesaj_ekle("=" * 50)
                self.log_mesaj_ekle(f"📁 Dosya İstatistikleri:")
                self.log_mesaj_ekle(f"   ✅ Başarılı: {basarili_dosyalar} dosya")
                self.log_mesaj_ekle(f"   ❌ Hatalı: {hatali_dosyalar} dosya")
                self.log_mesaj_ekle(f"   📊 Toplam: {len(self.excel_birlestirme_dosyalari)} dosya")
                self.log_mesaj_ekle("=" * 50)
                self.log_mesaj_ekle(f"\n📊 TÜM DOSYALARIN TOPLAM İSTATİSTİKLERİ:")
                self.log_mesaj_ekle(f"   📋 Toplam İşlenen Satır: {toplam_satir:,}")
                
                self.log_mesaj_ekle(f"   🔵 Toplam CBS Satır: {toplam_cbs_satir:,}")
                self.log_mesaj_ekle(f"   🟢 Aob Satırları: {gercek_aob_satir:,}")
                self.log_mesaj_ekle(f"   ⚪ Boş Satırlar: {gercek_bos_satir:,}")
                self.log_mesaj_ekle("=" * 50)
                self.log_mesaj_ekle(f"   📄 Sonuçtaki Toplam Satır: {gercek_toplam_satir:,}")

                # Excel dosyasını aç
                self.excel_ac(self.excel_birlestirme_kaydetme_yeri.get())
                
            else:
                self.lbl_durum.config(text="Hiç veri bulunamadı", fg=self.colors['warning'])
                self.log_mesaj_ekle("❌ Hiç veri bulunamadı!")
                
        except Exception as e:
            self.lbl_durum.config(text="Hata oluştu", fg=self.colors['danger'])
            self.log_mesaj_ekle(f"❌ Beklenmeyen hata: {str(e)}")
            import traceback
            self.log_mesaj_ekle(f"🔍 Hata detayı: {traceback.format_exc()}")
        finally:
            self.calisma_devam_ediyor = False

    def direk_id_isle(self, direk_id_serisi):
        """Direk ID sütunundaki değerleri uygun tipe dönüştürür ve '-' işaretlerini temizler"""
        try:
            islenmis_degerler = []
            
            for deger in direk_id_serisi:
                # NaN veya boş değer kontrolü
                if pd.isna(deger) or deger == '' or deger is None:
                    islenmis_degerler.append('')
                    continue
                
                # String'e çevir
                deger_str = str(deger).strip()
                
                # Eğer boş string ise
                if not deger_str:
                    islenmis_degerler.append('')
                    continue
                
                # YENİ: "-" işaretini temizle (sondaki ve baştaki)
                deger_str = deger_str.rstrip('-').lstrip('-')
                
                # Eğer temizleme sonrası boş string oluştuysa
                if not deger_str:
                    islenmis_degerler.append('')
                    continue
                
                # Rakam ile başlıyorsa sayıya çevirmeye çalış
                if deger_str[0].isdigit():
                    try:
                        # Nokta veya virgül içeriyor mu kontrol et (float olabilir)
                        if '.' in deger_str or ',' in deger_str:
                            # Float'a çevir, sonra integer'a çevirmeye çalış
                            float_deger = float(deger_str.replace(',', '.'))
                            # Eğer tam sayı ise integer'a çevir
                            if float_deger.is_integer():
                                islenmis_degerler.append(int(float_deger))
                            else:
                                islenmis_degerler.append(float_deger)
                        else:
                            # Doğrudan integer'a çevir
                            islenmis_degerler.append(int(deger_str))
                    except (ValueError, TypeError):
                        # Sayıya çevrilemezse orijinal string değeri kullan (temizlenmiş hali)
                        islenmis_degerler.append(deger_str)
                else:
                    # Metin ile başlıyorsa string olarak sakla (temizlenmiş hali)
                    islenmis_degerler.append(deger_str)
            
            return islenmis_degerler
            
        except Exception as e:
            self.log_mesaj_ekle(f"⚠️ Direk ID işleme hatası: {str(e)}")
            # Hata durumunda orijinal seriyi döndür
            return direk_id_serisi

            
if __name__ == "__main__":
    root = tk.Tk()
    app = AnaUygulama(root)
    root.mainloop()
