import sys
import os
import time
import shutil
from math import radians, sin, cos, sqrt, atan2
from pathlib import Path
from urllib.parse import urlparse, unquote
import winsound
import piexif
from PIL import Image, ImageQt
import hashlib

from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, numbers

from PyQt6.QtCore import Qt, QPoint, QPointF, QUrl, QTimer
from PyQt6.QtGui import QIcon, QPixmap, QPainter, QPen, QColor, QImage, QFont, QCursor
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QFileDialog, QLabel, QGroupBox, QInputDialog,
    QMessageBox, QScrollArea, QProgressBar, QSpinBox, QDialog, QColorDialog,
    QLineEdit, QComboBox, QListWidget, QDialogButtonBox, QFormLayout,
    QButtonGroup, QRadioButton, QFrame, QSlider, QTextEdit
)
from PyQt6.QtWebEngineWidgets import QWebEngineView
from PyQt6.QtWebEngineCore import QWebEngineSettings, QWebEnginePage, QWebEngineProfile, QWebEngineUrlRequestInterceptor
from PIL.ExifTags import TAGS

from PyQt6.QtCore import QSettings


# EXIF'ten GPS alma
def dms_to_deg(dms, ref):
    deg = dms[0][0] / dms[0][1]
    minute = dms[1][0] / dms[1][1]
    second = dms[2][0] / dms[2][1]
    result = deg + (minute / 60.0) + (second / 3600.0)
    if ref in ['S', 'W']:
        result *= -1
    return result

def exif_tarih_al(exif):
    if not exif:
        return None
    for tag_id, value in exif.items():
        tag = TAGS.get(tag_id, tag_id)
        if tag == 'DateTimeOriginal':
            return value
    return None

def get_gps_from_image(image_path):
    try:
        exif_data = piexif.load(image_path)
        gps_data = exif_data.get("GPS", {})
        if 2 in gps_data and 4 in gps_data:
            lat = dms_to_deg(gps_data[2], gps_data[1].decode())
            lon = dms_to_deg(gps_data[4], gps_data[3].decode())
            return lat, lon
    except Exception as e:
        print(f"EXIF Hatası ({image_path}): {e}")
    return None

def generate_unique_photo_id(image_path):
    """Dosya yoluna göre benzersiz ID oluştur"""
    try:
        file_stat = os.stat(image_path)
        unique_str = f"{os.path.abspath(image_path)}_{file_stat.st_size}_{file_stat.st_mtime}"
        return hashlib.md5(unique_str.encode()).hexdigest()[:16]
    except:
        return hashlib.md5(image_path.encode()).hexdigest()[:16]

def nihai_direk_excele_ekle(excel_yolu, direk_bilgileri):
    """Nihai_Direk.xlsx dosyasına yeni direk ekler"""
    try:
        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        drone_dir = os.path.join(desktop, "Drone")
        nihai_excel_yolu = os.path.join(drone_dir, "Nihai_Direk.xlsx")
        
        if not os.path.exists(nihai_excel_yolu):
            wb = Workbook()
            ws = wb.active
            basliklar = ["OBJECTID", "DIREKNO", "TIP", "AOB_ADI", "ENH_ID", "LAT", "LON"]
            ws.append(basliklar)
        else:
            wb = load_workbook(nihai_excel_yolu)
            ws = wb.active
        
        last_row = ws.max_row
        
        if last_row > 1:
            last_objectid = ws.cell(row=last_row, column=1).value
            if last_objectid and isinstance(last_objectid, (int, float)):
                new_objectid = int(last_objectid) + 1
            else:
                new_objectid = 1
        else:
            new_objectid = 1
        
        new_row = [
            new_objectid,
            direk_bilgileri.get("DIREKNO", ""),
            direk_bilgileri.get("TIP", ""),
            direk_bilgileri.get("AOB_ADI", "Drone_Tespit"),
            direk_bilgileri.get("ENH_ID", ""),
            direk_bilgileri.get("LAT", 0),
            direk_bilgileri.get("LON", 0)
        ]
        ws.append(new_row)
        
        wb.save(nihai_excel_yolu)
        print(f"✅ Nihai direk Excel'e eklendi: {direk_bilgileri.get('DIREKNO')}")
        return True
    except Exception as e:
        print(f"❌ Nihai direk Excel'e ekleme hatası: {e}")
        return False

class YeniDirekEklemeDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Yeni Direk Ekle")
        self.setFixedSize(400, 200)

        self.mevcut_sayac = parent.yeni_direk_sayaci if parent else 1
        self.direk_no = f"Direk Ekle {self.mevcut_sayac}"
        
        layout = QVBoxLayout(self)
        
        form_layout = QFormLayout()
        
        self.direk_no_label = QLabel()
        self.direk_no_label.setText(f"<b>{self.direk_no}</b>")
        form_layout.addRow("Direk No:", self.direk_no_label)
        
        self.envanter_notu = QTextEdit()
        self.envanter_notu.setPlaceholderText("Direk üzerinde kullanılan envanterler...\nÖrnek: 3 travers, 12 izolatör, iletken vs.")
        self.envanter_notu.setMaximumHeight(80)
        form_layout.addRow("Envanter Notu:", self.envanter_notu)
        
        layout.addLayout(form_layout)
        
        bilgi_label = QLabel("Not: Direk haritaya eklenmeden önce tüm fotoğrafları silin.\nBoş not bırakılırsa 'Direk Eklenecek' yazacaktır.")
        bilgi_label.setWordWrap(True)
        bilgi_label.setStyleSheet("color: #666; font-size: 11px; padding: 10px;")
        layout.addWidget(bilgi_label)
        
        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
        
        print(f"🔢 Yeni direk dialog açıldı: {self.direk_no}")
    
    def accept(self):
        parent = self.parent()
        if parent:
            parent.yeni_direk_sayaci += 1
            parent.yeni_direk_sayaci_kaydet()
            print(f"✅ Sayaç artırıldı: {parent.yeni_direk_sayaci}")
        super().accept()
    
    def get_direk_bilgileri(self):
        envanter_notu = self.envanter_notu.toPlainText().strip()
        if not envanter_notu:
            envanter_notu = "Direk Eklenecek"
        
        return {
            'direk_no': self.direk_no,
            'envanter_notu': envanter_notu,
            'birim': 'Cbs'
        }

class KaliteAyarlariDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Çözünürlük ve Kalite Ayarları")
        self.setMinimumWidth(400)
        
        layout = QVBoxLayout(self)
        
        form_layout = QFormLayout()
        
        self.genislik_spin = QSpinBox()
        self.genislik_spin.setRange(100, 10000)
        self.genislik_spin.setValue(4000)
        self.genislik_spin.setSuffix(" piksel")
        form_layout.addRow("Maksimum Genişlik:", self.genislik_spin)
        
        self.yukseklik_spin = QSpinBox()
        self.yukseklik_spin.setRange(100, 10000)
        self.yukseklik_spin.setValue(4000)
        self.yukseklik_spin.setSuffix(" piksel")
        form_layout.addRow("Maksimum Yükseklik:", self.yukseklik_spin)
        
        self.kalite_slider = QSlider(Qt.Orientation.Horizontal)
        self.kalite_slider.setRange(10, 100)
        self.kalite_slider.setValue(30)
        self.kalite_slider.setTickPosition(QSlider.TickPosition.TicksBelow)
        self.kalite_slider.setTickInterval(10)
        
        self.kalite_label = QLabel("30%")
        self.kalite_slider.valueChanged.connect(
            lambda v: self.kalite_label.setText(f"{v}%")
        )
        
        kalite_layout = QHBoxLayout()
        kalite_layout.addWidget(self.kalite_slider)
        kalite_layout.addWidget(self.kalite_label)
        form_layout.addRow("JPEG Kalitesi:", kalite_layout)
        
        self.exif_checkbox = QRadioButton("EXIF verilerini koru")
        self.exif_checkbox.setChecked(True)
        form_layout.addRow("", self.exif_checkbox)
        
        layout.addLayout(form_layout)
        
        self.bilgi_label = QLabel(
            "Örnek: 4000x3000 (12MP) → 1600x1200 (2MP)\n"
            "Dosya boyutu yaklaşık 1/4 oranında küçülecektir."
        )
        self.bilgi_label.setWordWrap(True)
        self.bilgi_label.setStyleSheet("color: #666; font-size: 11px;")
        layout.addWidget(self.bilgi_label)
        
        preset_layout = QHBoxLayout()
        preset_label = QLabel("Hızlı Seçimler:")
        preset_layout.addWidget(preset_label)
        
        preset_800 = QPushButton("800px")
        preset_800.clicked.connect(lambda: self.preset_ayarla(800, 800))
        preset_layout.addWidget(preset_800)
        
        preset_1200 = QPushButton("1200px")
        preset_1200.clicked.connect(lambda: self.preset_ayarla(1200, 1200))
        preset_layout.addWidget(preset_1200)
        
        preset_1600 = QPushButton("1600px")
        preset_1600.clicked.connect(lambda: self.preset_ayarla(1600, 1600))
        preset_layout.addWidget(preset_1600)
        
        preset_2000 = QPushButton("2000px")
        preset_2000.clicked.connect(lambda: self.preset_ayarla(2000, 2000))
        preset_layout.addWidget(preset_2000)
        
        layout.insertLayout(2, preset_layout)
        
        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        button_box.accepted.connect(self.accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)
    
    def preset_ayarla(self, genislik, yukseklik):
        self.genislik_spin.setValue(genislik)
        self.yukseklik_spin.setValue(yukseklik)
    
    def ayarlari_al(self):
        return {
            'max_width': self.genislik_spin.value(),
            'max_height': self.yukseklik_spin.value(),
            'quality': self.kalite_slider.value(),
            'preserve_exif': self.exif_checkbox.isChecked()
        }

    
class CustomWebEnginePage(QWebEnginePage):
    def __init__(self, profile, parent=None):
        super().__init__(profile, parent)
        
    def acceptNavigationRequest(self, url, nav_type, is_main_frame):
        return super().acceptNavigationRequest(url, nav_type, is_main_frame)


class RefererInterceptor(QWebEngineUrlRequestInterceptor):
    def __init__(self):
        super().__init__()
    
    def interceptRequest(self, info):
        url = info.requestUrl().toString()
        
        if "tile.openstreetmap.org" in url:
            print(f"🌍 Tile isteği yakalandı: {url}")
            info.setHttpHeader(
                b"Referer", 
                b"https://www.openstreetmap.org"
            )
            info.setHttpHeader(
                b"User-Agent",
                b"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 DroneTespit/1.0"
            )

    
class DroneHarita(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Drone Tespit İşleme Programı")
        self.resize(1200, 700)

        self.photo_info = []
        self.direkler = []

        self.foto_klasor = ""
        self.foto_kayit_klasor = ""

        user_dir = os.path.expanduser("~")
        self.drone_dir = os.path.join(user_dir, "Desktop", "Drone")
        self.map_path = os.path.join(self.drone_dir, "map.html")
        
        icon_yolu = os.path.join(self.drone_dir, "icons", "icon.png")
        if os.path.exists(icon_yolu):
            self.setWindowIcon(QIcon(icon_yolu))

        self.yeni_direk_sayaci = 1

        self.direk_data = oku_direkler(os.path.join(self.drone_dir, "direk_og.xlsx"))
        self.direk_data += oku_direkler_musterek(os.path.join(self.drone_dir, "direk_musterek.xlsx"))

        self.interceptor = RefererInterceptor()
        profile = QWebEngineProfile.defaultProfile()
        profile.setUrlRequestInterceptor(self.interceptor)
        
        profile.setHttpCacheType(profile.HttpCacheType.MemoryHttpCache)
        profile.setPersistentStoragePath(os.path.join(self.drone_dir, "webcache"))
        
        self.view = QWebEngineView()
        page = CustomWebEnginePage(profile, self.view)
        self.view.setPage(page)
        
        self.view.settings().setAttribute(QWebEngineSettings.WebAttribute.LocalContentCanAccessRemoteUrls, True)
        self.view.settings().setAttribute(QWebEngineSettings.WebAttribute.LocalContentCanAccessFileUrls, True)
        self.view.settings().setAttribute(QWebEngineSettings.WebAttribute.JavascriptEnabled, True)
        self.view.settings().setAttribute(QWebEngineSettings.WebAttribute.LocalStorageEnabled, False)
        self.view.settings().setAttribute(QWebEngineSettings.WebAttribute.AllowRunningInsecureContent, False)
        self.view.settings().setAttribute(QWebEngineSettings.WebAttribute.Accelerated2dCanvasEnabled, False)
        self.view.settings().setAttribute(QWebEngineSettings.WebAttribute.WebGLEnabled, False)
        self.view.settings().setAttribute(QWebEngineSettings.WebAttribute.JavascriptCanAccessClipboard, True)
        self.view.settings().setAttribute(QWebEngineSettings.WebAttribute.PluginsEnabled, False)

        self.haritayi_bos_baslat()
        
        sol_panel = QWidget()
        sol_layout = QVBoxLayout(sol_panel)
        sol_panel.setFixedWidth(200)

        drone_group = QGroupBox("✈️ Drone")
        drone_layout = QVBoxLayout(drone_group)
        self.drone_kayit_klasoru = ""

        self.yukle_buton = QPushButton("📂Fotoğrafları Yükle")
        self.yukle_buton.clicked.connect(self.fotograflari_yukle)
        drone_layout.addWidget(self.yukle_buton)

        self.birlestir_buton = QPushButton("🔗 Klasörleri Birleştir")
        self.birlestir_buton.clicked.connect(self.fotograf_klasorlerini_birlestir)
        drone_layout.addWidget(self.birlestir_buton)

        self.direk_ekle_buton = QPushButton("➕ Yeni Direk Ekle")
        self.direk_ekle_buton.clicked.connect(self.yeni_direk_ekle)
        drone_layout.addWidget(self.direk_ekle_buton)

        self.foto_editor_buton = QPushButton("🎨 Foto Editor")
        self.foto_editor_buton.clicked.connect(self.foto_editoru_ac)
        drone_layout.addWidget(self.foto_editor_buton)

        self.en_yakina_isle_buton = QPushButton("📍 En Yakına İşle")
        self.en_yakina_isle_buton.clicked.connect(self.en_yakina_isle)
        self.en_yakina_isle_buton.setEnabled(False)
        drone_layout.addWidget(self.en_yakina_isle_buton)

        self.kayit_klasoru_buton = QPushButton("📁 Kayıt Dosyasını Seç")
        self.kayit_klasoru_buton.clicked.connect(self.drone_kayit_klasoru_sec)
        drone_layout.addWidget(self.kayit_klasoru_buton)

        self.harita_yukleme_bar = QProgressBar()
        self.harita_yukleme_bar.setVisible(False)
        self.harita_yukleme_bar.setMaximum(0)
        self.harita_yukleme_bar.setMinimum(0)
        drone_layout.addWidget(self.harita_yukleme_bar)

        self.kayit_label = QLabel("📁 Seçilen kayıt klasörü: Henüz seçilmedi")
        self.kayit_label.setWordWrap(True)
        drone_layout.addWidget(self.kayit_label)

        self.kaydet_buton = QPushButton("💾 Kaydet")
        self.kaydet_buton.clicked.connect(self.kaydet)
        drone_layout.addWidget(self.kaydet_buton)

        self.foto_editor_buton.setEnabled(False)
        self.en_yakina_isle_buton.setEnabled(False)
        self.direk_ekle_buton.setEnabled(True)
        self.kayit_klasoru_buton.setEnabled(True)
        self.kaydet_buton.setEnabled(True)

        sol_layout.addWidget(drone_group)

        resim_kucultme_grup = QGroupBox("🖼️ Fotoğraf Küçültme")
        resim_kucultme_grup.setMaximumHeight(300)
        resim_kucultme_layout = QVBoxLayout(resim_kucultme_grup)

        self.foto_klasor = ""
        self.kayit_klasor = ""

        btn_foto_klasor = QPushButton("📂 Fotoğraf Klasörü Seç")
        btn_foto_klasor.clicked.connect(self.fotograf_klasoru_sec)
        resim_kucultme_layout.addWidget(btn_foto_klasor)

        self.foto_klasor_label = QLabel("📁 Seçilmedi")
        self.foto_klasor_label.setWordWrap(True)
        resim_kucultme_layout.addWidget(self.foto_klasor_label)

        btn_kayit_klasor = QPushButton("💾 Kayıt Klasörü Seç")
        btn_kayit_klasor.clicked.connect(self.foto_kayit_klasoru_sec)
        resim_kucultme_layout.addWidget(btn_kayit_klasor)

        self.kayit_klasor_label = QLabel("📁 Seçilmedi")
        self.kayit_klasor_label.setWordWrap(True)
        resim_kucultme_layout.addWidget(self.kayit_klasor_label)
        
        self.kucultme_bar = QProgressBar()
        self.kucultme_bar.setVisible(False)
        self.kucultme_bar.setMaximum(100)
        self.kucultme_bar.setValue(0)
        self.kucultme_bar.setFormat("%p%")
        resim_kucultme_layout.addWidget(self.kucultme_bar)

        self.btn_kucult = QPushButton("▶️ Küçültmeye Başla")
        self.btn_kucult.clicked.connect(self.kucultmeyi_baslat)
        resim_kucultme_layout.addWidget(self.btn_kucult)

        sol_layout.addWidget(resim_kucultme_grup)

        merkez_widget = QWidget()
        genel_layout = QHBoxLayout(merkez_widget)
        genel_layout.addWidget(sol_panel)
        genel_layout.addWidget(self.view)
        self.setCentralWidget(merkez_widget)

        self.kaydetme_bar = QProgressBar()
        self.kaydetme_bar.setVisible(False)
        self.kaydetme_bar.setMaximum(100)
        self.kaydetme_bar.setValue(0)
        self.kaydetme_bar.setFormat("Hazırlanıyor... 0%")
        drone_layout.addWidget(self.kaydet_buton)
        drone_layout.addWidget(self.kaydetme_bar)

        self.settings = QSettings("YourCompany", "DroneApp")

    def yeni_direk_sayaci_kaydet(self):
        try:
            self.settings.setValue("yeni_direk_sayaci", self.yeni_direk_sayaci)
        except:
            pass
    
    def haritayi_bos_baslat(self):
        with open(self.map_path, "w", encoding="utf-8") as f:
            f.write(f"""
            <html>
            <head>
                <title>Drone Haritası</title>
                <meta name="viewport" content="width=device-width, initial-scale=1.0">
                <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
                <link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.css" />
                <link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.Default.css" />
                <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
                <script src="https://unpkg.com/leaflet.markercluster@1.5.3/dist/leaflet.markercluster.js"></script>
                <style>
                    body {{ margin: 0; padding: 0; }}
                    #map {{ height: 100vh; }}
                    
                    .direk-marker-normal {{
                        background-color: #ff6666;
                        border: 2px solid #ff0000;
                        border-radius: 50%;
                        width: 14px;
                        height: 14px;
                        position: relative;
                        display: flex;
                        align-items: center;
                        justify-content: center;
                    }}
                    .direk-marker-musterek {{
                        background-color: #9370DB;
                        border: 2px solid #800080;
                        border-radius: 50%;
                        width: 14px;
                        height: 14px;
                        position: relative;
                        display: flex;
                        align-items: center;
                        justify-content: center;
                    }}
                    .direk-marker-yeni {{
                        background-color: #FFC0CB !important;
                        border: 2px solid #C71585 !important;
                        border-radius: 50%;
                        width: 14px;
                        height: 14px;
                        position: relative;
                        display: flex;
                        align-items: center;
                        justify-content: center;
                    }}
                    .direk-marker-foto-var {{
                        background-color: #90EE90 !important;
                        border: 2px solid #32CD32 !important;
                        border-radius: 50%;
                        width: 14px;
                        height: 14px;
                        position: relative;
                        display: flex;
                        align-items: center;
                        justify-content: center;
                    }}
                    .direk-marker-konum-degisti {{
                        background-color: #FFFF00 !important;
                        border: 2px solid #FFA500 !important;
                        border-radius: 50%;
                        width: 14px;
                        height: 14px;
                        position: relative;
                        display: flex;
                        align-items: center;
                        justify-content: center;
                    }}
                    .direk-marker-nihai {{
                        background-color: #007BFF !important;
                        border: 2px solid #0056b3 !important;
                        border-radius: 50%;
                        width: 14px;
                        height: 14px;
                        position: relative;
                        display: flex;
                        align-items: center;
                        justify-content: center;
                    }}
                    
                    .direk-marker-numara {{
                        position: absolute;
                        top: 50%;
                        left: 50%;
                        transform: translate(-50%, -50%);
                        color: white;
                        font-weight: bold;
                        font-size: 8px;
                        text-shadow: 0.5px 0.5px 0 #000, -0.5px -0.5px 0 #000, 0.5px -0.5px 0 #000, -0.5px 0.5px 0 #000;
                        pointer-events: none;
                        z-index: 1001;
                        line-height: 1;
                        white-space: nowrap;
                        opacity: 0.95;
                    }}
                    
                    .photo-marker {{
                        width: 16px;
                        height: 16px;
                        border-radius: 50%;
                        border: 2px solid white;
                        box-shadow: 0 0 4px #000;
                        z-index: 2000 !important;
                    }}
                    .photo-mavi {{
                        background-color: #007BFF;
                        opacity: 0.9;
                    }}
                    .photo-turkuaz {{
                        background-color: #00CED1;
                        opacity: 0.9;
                    }}
                    
                    .direk-numara-label {{
                        position: absolute;
                        transform: translate(-50%, -50%);
                        color: #000;
                        font-weight: bold;
                        font-size: 13px;
                        text-shadow: 1px 1px 0 #FFFFFF, -1px -1px 0 #FFFFFF, 1px -1px 0 #FFFFFF, -1px 1px 0 #FFFFFF;
                        pointer-events: none;
                        z-index: 3000;
                        white-space: nowrap;
                        opacity: 0.95;
                    }}
                    
                    .custom-popup .leaflet-popup-content-wrapper {{
                        width: 420px !important;
                    }}
                    
                    .marker-cluster-small {{
                        background-color: rgba(181, 226, 140, 0.6);
                    }}
                    .marker-cluster-small div {{
                        background-color: rgba(110, 204, 57, 0.6);
                    }}

                    .marker-cluster-medium {{
                        background-color: rgba(241, 211, 87, 0.6);
                    }}
                    .marker-cluster-medium div {{
                        background-color: rgba(240, 194, 12, 0.6);
                    }}

                    .marker-cluster-large {{
                        background-color: rgba(253, 156, 115, 0.6);
                    }}
                    .marker-cluster-large div {{
                        background-color: rgba(241, 128, 23, 0.6);
                    }}
                </style>
            </head>
            <body>
                <div id="map"></div>
                <script>
                    var osmLayer = L.tileLayer('https://{{s}}.tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png', {{
                        maxZoom: 19,
                        attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors'
                    }});

                    var darkLayer = L.tileLayer('https://{{s}}.basemaps.cartocdn.com/dark_all/{{z}}/{{x}}/{{y}}{{r}}.png', {{
                        attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a>, &copy; CartoDB',
                        subdomains: 'abcd',
                        maxZoom: 20
                    }});

                    var satelliteLayer = L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{{z}}/{{y}}/{{x}}', {{
                        maxZoom: 18,
                        attribution: 'Tiles © Esri'
                    }});

                    var map = L.map('map', {{
                        center: [36.7300, 30.2000],
                        zoom: 10,
                        layers: [osmLayer]
                    }});

                    var baseLayers = {{
                        "OpenStreetMap (Klasik)": osmLayer,
                        "Dark Matter (Koyu)": darkLayer,
                        "Uydu Görüntüsü": satelliteLayer
                    }};
                    L.control.layers(baseLayers).addTo(map);
                    
                    var direklerCluster = L.markerClusterGroup({{
                        maxClusterRadius: 40,
                        spiderfyOnMaxZoom: true,
                        showCoverageOnHover: false,
                        zoomToBoundsOnClick: true,
                        disableClusteringAtZoom: 18,
                        chunkedLoading: true,
                        chunkDelay: 100,
                        singleMarkerMode: false
                    }});
                    
                    var direkler = [];
                    var photoInfo = [];
                    
                    console.log('📍 Harita başlatıldı - 3 katmanlı');
                </script>
            </body>
            </html>
            """)
        self.view.load(QUrl.fromLocalFile(os.path.abspath(self.map_path)))
    
    def direkleri_haritaya_yukle(self, center_lat=None, center_lon=None, zoom_level=None):
        if not self.direk_data:
            print("⚠️ Yüklenecek direk verisi yok")
            return
            
        direk_html = ""
        for d in self.direk_data:
            direk_html += f"""
            (function() {{
                var direk = {{
                    id: "{d.get('id', '')}",
                    tip: "{d.get('tip', '')}",
                    lat: {d.get('lat', 0)},
                    lon: {d.get('lon', 0)},
                    musterek: {str(d.get('musterek', False)).lower()}
                }};
                
                var markerClass = 'direk-marker-normal';
                if (direk.musterek) {{
                    markerClass = 'direk-marker-musterek';
                }}

                var marker = L.marker([direk.lat, direk.lon], {{
                    icon: L.divIcon({{
                        className: markerClass,
                        iconSize: [14, 14],
                        iconAnchor: [7, 7]
                    }}),
                    opacity: 0.9,
                    zIndexOffset: 500
                }});

                var direkObj = {{
                    id: direk.id,
                    tip: direk.tip,
                    lat: direk.lat,
                    lon: direk.lon,
                    marker: marker,
                    photos: [],
                    musterek: direk.musterek,
                    numara: null,
                    numaraLabelMarker: null,
                    ozel: false,
                    draggable: false,
                    koordinat_degisecek: false,
                    nihai_direk: false
                }};

                direkler.push(direkObj);
                
                var popupHTML = `<b>Direk No: ${{direkObj.id}}</b><br><b>Tip:</b> ${{direkObj.tip}}`;
                marker.bindPopup(popupHTML);
                
                direklerCluster.addLayer(marker);
                console.log('✅ Direk eklendi:', direk.id);
            }})();
            """
        
        js_kod = f"""
        (function() {{
            map.setView([{center_lat}, {center_lon}], {zoom_level});
            
            direklerCluster.clearLayers();
            direkler = [];
            
            map.addLayer(direklerCluster);
            
            {direk_html}
            
            console.log('✅ {len(self.direk_data)} direk haritaya eklendi');
        }})();
        """
        
        QTimer.singleShot(1000, lambda: self.view.page().runJavaScript(js_kod))
    
    def fotolari_haritaya_yukle(self, photo_data):
        if not photo_data:
            print("⚠️ Yüklenecek fotoğraf verisi yok")
            return
            
        print(f"📸 {len(photo_data)} fotoğraf haritaya yükleniyor...")
        
        photo_html = ""
        for path, lat, lon, tarih, photo_id in photo_data:
            abs_path = path.replace("\\", "/")
            filename = os.path.basename(path)
            folder = os.path.basename(os.path.dirname(path))
            
            photo_html += f"""
            (function() {{
                var fullPath = "{abs_path}";
                var photoId = "{photo_id}";
                var filename = "{filename}";
                var folder = "{folder}";
                
                var marker = L.marker([{lat}, {lon}], {{
                    draggable: true,
                    icon: L.divIcon({{
                        className: 'photo-marker photo-mavi',
                        iconSize: [13, 13]
                    }}),
                    zIndexOffset: 2000
                }});

                marker.bindPopup("<div>Fotoğraf yükleniyor...</div>");

                marker.on("popupopen", function() {{
                    var popupContent = `
                        <div style="
                            width: 300px;
                            max-width: 90%;
                            line-height: 1.2;
                        ">
                            <b style="
                                word-break: break-all;
                                font-size: 12px;
                                margin: 0;
                                display: block;
                            ">
                                📁 ${{folder}}/<br>${{filename}}
                            </b>

                            <img
                                src="${{fullPath}}"
                                loading="lazy"
                                decoding="async"
                                style="
                                    width: 100%;
                                    max-width: 280px;
                                    border-radius: 3px;
                                    box-shadow: 0 0 3px #0002;
                                    margin: 4px 0;
                                    display: block;
                                    image-rendering: auto;
                                    filter: contrast(0.95) brightness(0.97);
                                "
                            >
                        </div>
                    `;
                    marker.setPopupContent(popupContent);
                }});

                photoInfo.push({{
                    id: photoId,
                    path: fullPath,
                    filename: filename,
                    folder: folder,
                    lat: {lat},
                    lon: {lon},
                    originalLat: {lat},
                    originalLon: {lon},
                    note: "",
                    birim: "Aob",
                    marker: marker,
                    edited: false,
                    tarih: "{tarih or ''}",
                    kategori: "",
                    oncelik: "Bekleyebilir"
                }});

                marker.addTo(map);
                console.log("✅ Fotoğraf eklendi:", folder + "/" + filename);
            }})();
            """

        js_kod = f"""
        (function() {{
            photoInfo.forEach(function(p) {{
                if (p.marker && map.hasLayer(p.marker)) {{
                    map.removeLayer(p.marker);
                }}
            }});
            photoInfo = [];
            
            {photo_html}
            
            console.log('✅ {len(photo_data)} fotoğraf haritaya eklendi');
            
            if (!window.guncellePopup) {{
                window.guncellePopup = function(d, popupAc = false) {{
                    var html = `<b>Direk No: ${{d.id}}</b><br><b>Tip:</b> ${{d.tip}}`;
                    if (d.photos && d.photos.length > 0) {{
                        html += `<br><b>Fotoğraflar:</b> ${{d.photos.length}} adet`;
                    }}
                    d.marker.setPopupContent(html);
                }};
            }}
        }})();
        """
        
        QTimer.singleShot(2000, lambda: self.view.page().runJavaScript(js_kod))
        
        return True

    def mesafe_hesapla(self, lat1, lon1, lat2, lon2):
        R = 6371000
        phi1 = radians(lat1)
        phi2 = radians(lat2)
        dphi = radians(lat2 - lat1)
        dlambda = radians(lon2 - lon1)

        a = sin(dphi/2)**2 + cos(phi1)*cos(phi2)*sin(dlambda/2)**2
        c = 2 * atan2(sqrt(a), sqrt(1-a))

        return R * c

    def en_yakina_isle(self):
        js_kod = """
            (function() {
                return {
                    photoInfo: photoInfo.map(p => {
                        let pos = (p.marker && p.marker.getLatLng) ? p.marker.getLatLng() : { lat: p.lat, lng: p.lon };
                        return [p.path, p.id, pos.lat, pos.lng, p.note || ""];
                    }),
                    direkler: direkler.map(d => ({
                        id: d.id,
                        lat: d.lat,
                        lon: d.lon,
                        tip: d.tip,
                        numara: d.numara || null,
                        photos: d.photos || []
                    }))
                };
            })();
        """
        
        self.kaydetme_bar.setVisible(True)
        self.kaydetme_bar.setMaximum(0)
        self.kaydetme_bar.setFormat("🔍 Fotoğraflar taranıyor...")
        QApplication.processEvents()
        
        self.view.page().runJavaScript(js_kod, self.en_yakina_isle_guncelle)

    def en_yakina_isle_guncelle(self, veri):
        from PyQt6.QtWidgets import QMessageBox
        import os
        from pathlib import Path
        from urllib.parse import unquote

        self.direkler = veri["direkler"]
        updated_photo_info = veri["photoInfo"]
        self.photo_info = updated_photo_info

        if not updated_photo_info or not self.direkler:
            QMessageBox.warning(self, "Veri Eksik", "⚠️ Fotoğraf veya direk verisi eksik!")
            self.kaydetme_bar.setVisible(False)
            return

        otomatik_iliskilenen = 0
        mesafe_disi = 0
        kontrol_edilecek = []

        YARI_CAP = 30

        print("\n" + "="*60)
        print("📍 EN YAKINA İŞLE BAŞLADI (30m çap)")
        print("="*60)

        toplam_foto = len(updated_photo_info)
        self.kaydetme_bar.setMaximum(toplam_foto)
        self.kaydetme_bar.setValue(0)
        self.kaydetme_bar.setFormat(f"Fotoğraflar taranıyor... 0% (0/{toplam_foto})")
        QApplication.processEvents()

        for i, (foto_path, foto_id, foto_lat, foto_lon, note) in enumerate(updated_photo_info):
            ilerleme_yuzde = int((i + 1) / toplam_foto * 100)
            self.kaydetme_bar.setValue(i + 1)
            self.kaydetme_bar.setFormat(f"Fotoğraflar taranıyor... {ilerleme_yuzde}% ({i+1}/{toplam_foto})")
            QApplication.processEvents()
            
            if foto_lat is None or foto_lon is None:
                print(f"⚠️ Atlanıyor (konum yok): {foto_path}")
                continue

            yakin_direkler = []
            for direk in self.direkler:
                lat = direk.get("lat")
                lon = direk.get("lon")
                if lat is None or lon is None:
                    continue

                mesafe = self.mesafe_hesapla(foto_lat, foto_lon, lat, lon)
                if mesafe <= YARI_CAP:
                    yakin_direkler.append((mesafe, direk))

            if yakin_direkler:
                yakin_direkler.sort(key=lambda x: x[0])
                en_yakin_mesafe, en_yakin_direk = yakin_direkler[0]
                
                foto_adi = os.path.basename(foto_path)
                print(f"\n📸 {foto_adi}")
                print(f"   ├─ {len(yakin_direkler)} direk bulundu (30m içinde)")
                print(f"   ├─ EN YAKIN: {en_yakin_direk['id']} - {en_yakin_mesafe:.1f}")
                
                kontrol_edilecek.append((foto_path, foto_id, en_yakin_direk, foto_lat, foto_lon, note))
            else:
                mesafe_disi += 1
                print(f"⚠️ {os.path.basename(foto_path)}: 30m içinde direk bulunamadı")

        print("\n" + "="*60)
        print(f"📊 TARAMA TAMAMLANDI")
        print(f"   • Toplam fotoğraf: {toplam_foto}")
        print(f"   • 30m içinde direk bulunan: {len(kontrol_edilecek)}")
        print(f"   • 30m içinde direk bulunamayan: {mesafe_disi}")
        print("="*60)

        if kontrol_edilecek:
            self.kaydetme_bar.setMaximum(len(kontrol_edilecek))
            self.kaydetme_bar.setValue(0)
            self.kaydetme_bar.setFormat(f"Fotoğraflar ekleniyor... 0% (0/{len(kontrol_edilecek)})")
            QApplication.processEvents()
            
            def devam_et_kontrol(kalanlar, index=0):
                if index >= len(kalanlar):
                    self.kaydetme_bar.setFormat("✅ Tamamlandı")
                    QApplication.processEvents()
                    
                    QTimer.singleShot(3000, lambda: self.kaydetme_bar.setVisible(False))
                    
                    QMessageBox.information(self, "İşlem Tamamlandı",
                        f"✅ 30m çapında en yakına işleme tamamlandı!\n\n"
                        f"📊 SONUÇLAR:\n"
                        f"• Toplam fotoğraf: {toplam_foto}\n"
                        f"• Eklenen fotoğraf: {otomatik_iliskilenen}\n"
                        f"• 30m içinde direk bulunamayan: {mesafe_disi}\n\n"
                        f"📍 Not: Fotoğraflar en yakın direğe eklendi.")
                    
                    self.view.page().runJavaScript("guncellePopuplar();")
                    return

                foto_path, foto_id, direk, foto_lat, foto_lon, note = kalanlar[index]
                foto_filename = os.path.basename(foto_path)
                
                # ÖNEMLİ: URL encoding yapma, direkt dosya yolunu kullan
                # Sadece backslash'leri forward slash'e çevir
                clean_path = foto_path.replace("\\", "/")
                
                # JavaScript'e göndermek için güvenli hale getir (tırnak işaretlerini kaçır)
                safe_path = clean_path.replace("'", "\\'").replace('"', '\\"')

                ilerleme_yuzde = int((index + 1) / len(kalanlar) * 100)
                self.kaydetme_bar.setValue(index + 1)
                self.kaydetme_bar.setFormat(f"Fotoğraflar ekleniyor... {ilerleme_yuzde}% ({index+1}/{len(kalanlar)})")
                QApplication.processEvents()

                js_kontrol = f"""
                    (function() {{
                        const d = direkler.find(dd => dd.id === "{direk['id']}");
                        const fullPath = "{safe_path}";
                        if (!d || !d.photos) return false;
                        return !d.photos.some(p => p.path === fullPath);
                    }})();
                """

                def handle_kontrol(result):
                    nonlocal otomatik_iliskilenen
                    if result:
                        if "photos" not in direk:
                            direk["photos"] = []

                        direk["photos"].append({"path": foto_path, "note": note, "id": foto_id})
                        otomatik_iliskilenen += 1

                        # Original koordinatları koru ve marker'ı kaldır
                        self.view.page().runJavaScript(f"""
                            (function() {{
                                const fullPath = "{safe_path}";
                                const p = photoInfo.find(p => p.path === fullPath);
                                if (p) {{
                                    if (p.originalLat === undefined || p.originalLat === null) {{
                                        p.originalLat = {foto_lat};
                                        p.originalLon = {foto_lon};
                                    }}
                                    console.log("Original kaydedildi:", fullPath, p.originalLat, p.originalLon);
                                }}
                                
                                if (p && p.marker && map.hasLayer(p.marker)) {{
                                    map.removeLayer(p.marker);
                                    p.marker = null;
                                    console.log("Marker kaldırıldı:", fullPath);
                                }}
                            }})();
                        """)
                        
                        # Direğe fotoğraf ekle - URL encoding YAPMA, direkt yolu gönder
                        self.view.page().runJavaScript(f'addPhotoToDirek("{direk["id"]}", "{safe_path}")')

                    devam_et_kontrol(kalanlar, index + 1)

                self.view.page().runJavaScript(js_kontrol, handle_kontrol)

            devam_et_kontrol(kontrol_edilecek)
    
    def fotograf_klasorlerini_birlestir(self):
        klasor_listesi = []

        dialog = QDialog(self)
        dialog.setWindowTitle("Seçilen Klasörler")
        layout = QVBoxLayout(dialog)
        klasor_list_widget = QListWidget()
        layout.addWidget(klasor_list_widget)
        bilgi_label = QLabel("Devam etmek için başka klasör seçin veya pencereyi kapatın.")
        bilgi_label.setEnabled(False)
        layout.addWidget(bilgi_label)
        dialog.resize(500, 300)
        dialog.show()

        while True:
            klasor = QFileDialog.getExistingDirectory(self, "Fotoğraf Klasörü Seçin")
            if klasor:
                klasor_listesi.append(klasor)
                klasor_list_widget.addItem(klasor)

                msg = QMessageBox(self)
                msg.setWindowTitle("Devam?")
                msg.setText("Başka klasör eklemek ister misiniz?")
                evet_buton = msg.addButton("Evet", QMessageBox.ButtonRole.YesRole)
                hayir_buton = msg.addButton("Hayır", QMessageBox.ButtonRole.NoRole)
                msg.exec()

                if msg.clickedButton() == hayir_buton:
                    break
            else:
                break

        dialog.close()

        if not klasor_listesi:
            QMessageBox.warning(self, "İşlem İptal", "⚠️ Hiç klasör seçilmedi.")
            return

        self.kaydetme_bar.setVisible(True)
        self.kaydetme_bar.setFormat("Fotoğraflar taranıyor... 0%")
        QApplication.processEvents()

        foto_liste = []
        uzantilar = (".jpg", ".jpeg", ".png", ".bmp", ".webp", ".tif", ".tiff")

        for klasor in klasor_listesi:
            for root, _, files in os.walk(klasor):
                for file in files:
                    if file.lower().endswith(uzantilar):
                        tam_yol = os.path.join(root, file)

                        try:
                            exif = piexif.load(tam_yol)
                            tarih = exif['Exif'][piexif.ExifIFD.DateTimeOriginal].decode()
                        except Exception:
                            tarih = time.ctime(os.path.getctime(tam_yol))

                        foto_liste.append((tam_yol, tarih))

        if not foto_liste:
            QMessageBox.warning(self, "Fotoğraf Yok", "⚠️ Seçilen klasörlerde uygun fotoğraf bulunamadı.")
            self.kaydetme_bar.setVisible(False)
            return

        foto_liste.sort(key=lambda x: x[1])

        hedef_klasor = QFileDialog.getExistingDirectory(self, "Birleşen Fotoğrafların Kaydedileceği Klasörü Seçin")
        if not hedef_klasor:
            QMessageBox.warning(self, "İşlem İptal", "⚠️ Klasör seçilmedi.")
            self.kaydetme_bar.setVisible(False)
            return

        isim_koku, ok = QInputDialog.getText(self, "Proje", "Tarama Yapılan Projenin İsmini Girin:")
        if not ok or not isim_koku.strip():
            QMessageBox.warning(self, "İşlem İptal", "⚠️ Fotoğraf ismi girilmedi.")
            self.kaydetme_bar.setVisible(False)
            return
        isim_koku = isim_koku.strip()

        self.kaydetme_bar.setMaximum(len(foto_liste))
        self.kaydetme_bar.setValue(0)
        self.kaydetme_bar.setFormat(f"Kopyalanıyor... 0% (0/{len(foto_liste)})")
        QApplication.processEvents()

        basarili_sayisi = 0
        hata_sayisi = 0

        for i, (kaynak_yol, _) in enumerate(foto_liste, start=1):
            try:
                uzanti = os.path.splitext(kaynak_yol)[1]
                hedef_yol = os.path.join(hedef_klasor, f"{isim_koku}_{i}{uzanti}")
                
                shutil.copy2(kaynak_yol, hedef_yol)
                print(f"✅ {i}/{len(foto_liste)}: {os.path.basename(kaynak_yol)} → {hedef_yol}")
                basarili_sayisi += 1
                
            except Exception as e:
                print(f"❌ Hata: {kaynak_yol} → {e}")
                hata_sayisi += 1

            ilerleme_yuzde = int(i / len(foto_liste) * 100)
            self.kaydetme_bar.setValue(i)
            self.kaydetme_bar.setFormat(f"Kopyalanıyor... %{ilerleme_yuzde} ({i}/{len(foto_liste)})")
            QApplication.processEvents()

        self.kaydetme_bar.setFormat("✅ Tamamlandı")
        QApplication.processEvents()

        QMessageBox.information(self, "İşlem Tamam", 
            f"✅ İşlem başarıyla tamamlandı!\n\n"
            f"📊 İSTATİSTİKLER:\n"
            f"• Toplam fotoğraf: {len(foto_liste)}\n"
            f"• Başarıyla kopyalanan: {basarili_sayisi}\n"
            f"• Hata alan: {hata_sayisi}\n\n"
            f"📁 Hedef klasör:\n{hedef_klasor}")

        QTimer.singleShot(3000, lambda: self.kaydetme_bar.setVisible(False))

        msg2 = QMessageBox(self)
        msg2.setWindowTitle("Harita Yükleme")
        msg2.setText("📍 Birleşen fotoğrafları haritaya yüklemek ister misiniz?")
        evet_buton2 = msg2.addButton("Evet", QMessageBox.ButtonRole.YesRole)
        hayir_buton2 = msg2.addButton("Hayır", QMessageBox.ButtonRole.NoRole)
        msg2.exec()

        if msg2.clickedButton() == evet_buton2:
            self.birlestirilen_fotograflari_yukle(hedef_klasor)

    def birlestirilen_fotograflari_yukle(self, klasor):
        uzantilar = (".jpg", ".jpeg", ".png", ".bmp", ".webp", ".tif", ".tiff")
        photo_data = []

        dosyalar = os.listdir(klasor)
        for dosya in dosyalar:
            if dosya.lower().endswith(uzantilar):
                tam_yol = os.path.join(klasor, dosya)
                gps = get_gps_from_image(tam_yol)
                tarih = None
                photo_id = generate_unique_photo_id(tam_yol)
                try:
                    img = Image.open(tam_yol)
                    exif = img._getexif()
                    tarih = exif_tarih_al(exif)
                except:
                    pass

                if gps:
                    photo_data.append((tam_yol, gps[0], gps[1], tarih, photo_id))
                else:
                    print(f"⚠️ GPS yok: {tam_yol}")

        if not photo_data:
            QMessageBox.warning(self, "GPS Yok", "⚠️ GPS verisi bulunamadı.")
            return

        create_map(photo_data, self.direk_data, self.map_path)
        self.view.load(QUrl.fromLocalFile(os.path.abspath(self.map_path)))
        self.direk_ekle_buton.setEnabled(True)
        self.foto_editor_buton.setEnabled(True)
        self.kayit_klasoru_buton.setEnabled(True)
        self.kaydet_buton.setEnabled(True)
        self.en_yakina_isle_buton.setEnabled(True)

        self.photo_info = photo_data
        self.direkler = self.direk_data.copy()
    
        QMessageBox.information(self, "Harita Yüklendi", "✅ Birleşen fotoğraflar haritaya işlendi.")
        
    def foto_editoru_ac(self):
        if not hasattr(self, 'view') or not self.view:
            QMessageBox.warning(self, "Harita Yok", "📍 Önce fotoğrafları yüklemelisiniz!")
            return

        js_kod = """
        (function() {
            return photoInfo.map(p => ({
                id: p.id,
                path: p.path,
                note: p.note || "",
                tarih: p.tarih || "",
                kategori: p.kategori || "",
                oncelik: p.oncelik || "Bekleyebilir",
                birim: p.birim || "Aob"
            }));
        })();
        """
        self.view.page().runJavaScript(js_kod, self.foto_editor_penceresini_ac)

    def foto_editor_penceresini_ac(self, foto_verileri):
        if not foto_verileri:
            QMessageBox.warning(self, "Fotoğraf Yok", "⚠️ Henüz Yüklenmiş fotoğraf yok!")
            return

        foto_verileri.sort(key=lambda x: x["tarih"])

        FotoEditorViewer(foto_verileri, view=self.view, parent=self).show()

    def yeni_direk_ekle(self):
        dialog = YeniDirekEklemeDialog(self)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return
        
        direk_bilgileri = dialog.get_direk_bilgileri()
        direk_no = direk_bilgileri['direk_no']
        
        direk_tipi = "Yeni Direk"
        envanter_notu = direk_bilgileri['envanter_notu']
        birim = direk_bilgileri['birim']
        
        temiz_id = direk_no.strip().replace("'", "").replace('"', '')

        js_kontrol = f"""
            (function(){{
                return direkler.some(d => d.id === '{temiz_id}');
            }})();
        """

        def js_kontrol_sonucu(var_mi):
            if var_mi:
                self.yeni_direk_sayaci -= 1
                self.yeni_direk_sayaci_kaydet()
                
                QMessageBox.warning(
                    self, 
                    "ID Mevcut", 
                    f"⚠️ '{temiz_id}' ID'li direk zaten mevcut!\n\n"
                    f"Bir sonraki denemede '{temiz_id}' numarası atlanacak."
                )
            else:
                js = f"""
                setTimeout(function() {{
                    if (typeof map !== 'undefined') {{
                        map.once('click', function(e) {{
                            var yeniID = '{temiz_id}';
                            var yeniTip = '{direk_tipi}';
                            var yeniNot = `{envanter_notu}`;
                            var yeniBirim = '{birim}';
                            
                            var marker = L.marker(e.latlng, {{
                                draggable: true,
                                icon: L.divIcon({{
                                    className: 'direk-marker-yeni',
                                    iconSize: [14, 14],
                                    iconAnchor: [7, 7]
                                }}),
                                opacity: 0.9,
                                zIndexOffset: 500
                                }}).addTo(map);

                                var yeniDirek = {{
                                    id: yeniID,
                                    tip: yeniTip,
                                    lat: e.latlng.lat,
                                    lon: e.latlng.lng,
                                    marker: marker,
                                    photos: [],
                                    ozel: true,
                                    draggable: true,
                                    musterek: false,
                                    koordinat_degisecek: false,
                                    nihai_direk: false,
                                    yeni_direk: true,
                                    silinecek_mi: false,
                                    envanter_notu: yeniNot,
                                    birim: yeniBirim
                                }};

                                direkler.push(yeniDirek);
                                
                                var popupHTML = `
                                    <b>Direk No: ${{yeniID}}</b><br>
                                    <b>Tip:</b> ${{yeniTip}}<br>
                                    <b>Not:</b> ${{yeniNot}}<br>
                                    <small>🟣 <span style="color:#FF69B4;">Yeni Direk</span></small><br>
                                    <button onclick="kaldirDirek('${{yeniID}}')" 
                                            style="background-color: #ff4444; 
                                                   color: white; 
                                                   border: none; 
                                                   padding: 5px 10px; 
                                                   border-radius: 3px; 
                                                   cursor: pointer;
                                                   margin-top: 5px;">
                                        🗑️ Direği Kaldır
                                    </button>
                                `;
                                
                                var popup = L.popup().setContent(popupHTML);
                                marker.bindPopup(popup);

                                marker.on("dragend", function(e) {{
                                    var pos = e.target.getLatLng();
                                    yeniDirek.lat = pos.lat;
                                    yeniDirek.lon = pos.lng;
                                    yeniDirek.koordinat_degisecek = true;
                                    
                                    var updatedPopup = `
                                        <b>Direk No: ${{yeniID}}</b><br>
                                        <b>Tip:</b> ${{yeniTip}}<br>
                                        <b>Not:</b> ${{yeniNot}}<br>
                                        <small>🟣 <span style="color:#FF69B4;">Yeni Direk</span></small><br>
                                        <small>📍 Konum değiştirildi</small><br>
                                        <button onclick="kaldirDirek('${{yeniID}}')" 
                                                style="background-color: #ff4444; 
                                                       color: white; 
                                                       border: none; 
                                                       padding: 5px 10px; 
                                                       border-radius: 3px; 
                                                       cursor: pointer;
                                                       margin-top: 5px;">
                                            🗑️ Direği Kaldır
                                        </button>
                                    `;
                                    marker.setPopupContent(updatedPopup);
                                    
                                    L.popup()
                                        .setLatLng(pos)
                                        .setContent(`✅ ${{yeniID}} direğinin konumu güncellendi.<br>Yeni koordinatlar: ${{yeniDirek.lat.toFixed(6)}}, ${{yeniDirek.lon.toFixed(6)}}`)
                                        .openOn(map);
                                }});

                                console.log("✅ Yeni direk eklendi:", yeniID);
                            }});
                        }} else {{
                            console.warn("❌ Harita hazır değil");
                        }}
                    }}, 500);
                    """
                self.view.page().runJavaScript(js)
                    
                QMessageBox.information(
                    self, 
                    "Direk Ekleme Modu", 
                    f"✅ {direk_no} numaralı direk için bilgiler alındı.\n\n"
                    f"📍 Şimdi harita üzerinde direğin konumunu seçmek için bir noktaya tıklayın."
                )

        self.view.page().runJavaScript(js_kontrol, js_kontrol_sonucu)

    def fotograflari_yukle(self):
        klasor = QFileDialog.getExistingDirectory(self, "Fotoğraf Klasörü Seç")
        if not klasor:
            return

        baslangic_zamani = time.perf_counter()

        self.harita_yukleme_bar.setVisible(True)
        QApplication.processEvents()

        uzantilar = (".jpg", ".jpeg", ".png", ".tif", ".tiff", ".heic", ".bmp", ".webp")
        photo_data = []

        toplam_dosya_sayisi = 0
        for root, _, files in os.walk(klasor):
            for file in files:
                if file.lower().endswith(uzantilar):
                    toplam_dosya_sayisi += 1
        
        self.harita_yukleme_bar.setMaximum(toplam_dosya_sayisi)
        self.harita_yukleme_bar.setValue(0)
        
        taranan = 0
        for root, _, files in os.walk(klasor):
            for dosya in files:
                if dosya.lower().endswith(uzantilar):
                    tam_yol = os.path.join(root, dosya)
                    gps = get_gps_from_image(tam_yol)
                    if gps:
                        tarih = None
                        try:
                            with Image.open(tam_yol) as img:
                                exif = img._getexif()
                                tarih = exif_tarih_al(exif)
                                print(f"📸 {dosya} → {gps} 🕒 {tarih}")
                        except Exception as e:
                            print(f"❌ EXIF tarih alınamadı: {dosya} → {e}")
                            tarih = None
                        
                        photo_id = generate_unique_photo_id(tam_yol)
                        photo_data.append((tam_yol, gps[0], gps[1], tarih, photo_id))
                    else:
                        print(f"⚠️ GPS yok: {tam_yol}")
                    
                    taranan += 1
                    self.harita_yukleme_bar.setValue(taranan)
                    QApplication.processEvents()

        if not photo_data:
            QMessageBox.warning(self, "GPS Verisi Yok", "⚠️ Uygun GPS verisine sahip fotoğraf bulunamadı.")
            self.harita_yukleme_bar.setVisible(False)
            return

        create_map(photo_data, self.direk_data, self.map_path)
        self.view.load(QUrl.fromLocalFile(os.path.abspath(self.map_path)))

        self.direk_ekle_buton.setEnabled(True)
        self.foto_editor_buton.setEnabled(True)
        self.kayit_klasoru_buton.setEnabled(True)
        self.kaydet_buton.setEnabled(True)
        self.en_yakina_isle_buton.setEnabled(True)

        self.photo_info = photo_data
        self.direkler = self.direk_data.copy()

        bitis_zamani = time.perf_counter()
        gecen_sure = bitis_zamani - baslangic_zamani

        self.harita_yukleme_bar.setVisible(False)
        QMessageBox.information(self, "Yükleme Tamamlandı", 
            f"✅ Fotoğraflar haritaya yüklendi!\n\n"
            f"📊 Bulunan fotoğraf sayısı: {len(photo_data)}\n"
            f"📁 Taranan klasör: {klasor}\n"
            f"⏱️ Geçen Süre: {gecen_sure:.2f} saniye\n"
            f"📍 Harita konumu: fotoğrafların tamamı görünecek şekilde ayarlandı")

    def kaydet(self):
        try:
            self.kaydetme_bar.setVisible(True)
            self.kaydetme_bar.setValue(0)
            self.kaydetme_bar.setFormat("Hazırlanıyor... 0%")
            QApplication.processEvents()
            
            if not self.drone_kayit_klasoru:
                QMessageBox.warning(self, "Klasör Seçilmedi", "⚠️ Lütfen önce drone kayıt klasörünü seçiniz!")
                self.kaydetme_bar.setVisible(False)
                return

            msg = QMessageBox(self)
            msg.setWindowTitle("Kayıt Seçeneği")
            msg.setText("📸 Fotoğrafları nasıl kaydetmek istiyorsunuz?")
            msg.setIcon(QMessageBox.Icon.Question)
            
            kucult_btn = msg.addButton("Küçülterek Kaydet", QMessageBox.ButtonRole.YesRole)
            orijinal_btn = msg.addButton("Orijinal Boyutta Kaydet", QMessageBox.ButtonRole.NoRole)
            iptal_btn = msg.addButton("İptal", QMessageBox.ButtonRole.RejectRole)
            
            msg.exec()
            
            if msg.clickedButton() == iptal_btn:
                print("🚫 Kullanıcı iptal etti")
                self.kaydetme_bar.setVisible(False)
                return
            
            kuculterek_kaydet = msg.clickedButton() == kucult_btn
            print(f"📝 Kayıt seçeneği: {'Küçülterek' if kuculterek_kaydet else 'Orijinal'}")

            kalite_ayarlari = None
            if kuculterek_kaydet:
                dialog = KaliteAyarlariDialog(self)
                if dialog.exec() == QDialog.DialogCode.Accepted:
                    kalite_ayarlari = dialog.ayarlari_al()
                    print(f"⚙️ Kalite ayarları: {kalite_ayarlari}")
                else:
                    print("🚫 Kullanıcı kalite ayarlarını iptal etti")
                    self.kaydetme_bar.setVisible(False)
                    return

            js_kod = """
                (function(){
                    return {
                        direkFotoData: direkler.filter(d => d.photos.length > 0 || d.silinecek_mi || d.koordinat_degisecek || d.nihai_direk || d.yeni_direk).map(d => ({
                            id: d.id,
                            tip: d.tip,
                            lat: d.lat,
                            lon: d.lon,
                            photos: d.photos.map(p => {
                                const fotoBilgisi = photoInfo.find(pi => pi.path === p.path);
                                const kategori = fotoBilgisi ? (fotoBilgisi.kategori || "") : "";
                                const oncelik  = fotoBilgisi ? (fotoBilgisi.oncelik || "Bekleyebilir") : "Bekleyebilir";
                                const birim    = fotoBilgisi ? (fotoBilgisi.birim || "Aob") : "Aob";
                                return {
                                    path: p.path,
                                    note: p.note || "",
                                    kategori: kategori,
                                    oncelik: oncelik,
                                    birim: birim
                                };
                            }),
                            numara: d.numara ?? null,
                            silinecek_mi: d.silinecek_mi || false,
                            koordinat_degisecek: d.koordinat_degisecek || false,
                            nihai_direk: d.nihai_direk || false,
                            draggable: d.draggable || false,
                            yeni_direk: d.yeni_direk || false,
                            envanter_notu: d.envanter_notu || "",
                            birim: d.birim || "Cbs"
                        })),
                        iliskilenmeyenFotolar: photoInfo.filter(p => {
                            return !direkler.flatMap(d => d.photos.map(p => p.path)).includes(p.path);
                        }).map(p => p.path)
                    };
                })();
            """

            def js_sonucu_al(veri):
                try:
                    print("📊 JavaScript'ten veri alındı")
                    direk_fotolari = veri["direkFotoData"]
                    iliskisiz_fotolar = veri["iliskilenmeyenFotolar"]

                    if not direk_fotolari:
                        QMessageBox.warning(self, "Fotoğraf Yok", "⚠️ Hiçbir direğe fotoğraf bağlı değil!")
                        self.kaydetme_bar.setVisible(False)
                        return

                    total_fotos = sum(len(d.get('photos', [])) for d in direk_fotolari)
                    print(f"📸 Toplam fotoğraf sayısı: {total_fotos}")
                    
                    self.kaydetme_bar.setMaximum(total_fotos + 1)
                    self.kaydetme_bar.setValue(0)
                    self.kaydetme_bar.setFormat("Excel hazırlanıyor... 0%")
                    QApplication.processEvents()
                    
                    self.nihai_direkleri_kaydet(direk_fotolari)
                    
                    def kopyalamaya_basla():
                        try:
                            print(f"📁 Kayıt klasörü: {self.drone_kayit_klasoru}")
                            excel_yolu = os.path.join(self.drone_kayit_klasoru, "Tespit.xlsx")
                            print(f"📗 Excel dosyası: {excel_yolu}")
                            
                            wb = Workbook()
                            ws = wb.active
                            
                            basliklar = ["Sıra No", "Direk No", "Tespit Notu", "Tespit Kategorisi", "Öncelik", "Fotoğraf Yolu", "Enlem", "Boylam", "Birim", "Yapıldı mı?"]
                            ws.append(basliklar)
                            
                            baslik_font = Font(name='Calibri', size=11, bold=True, color="FFFFFF")
                            baslik_dolgu = PatternFill(start_color="0070C0", end_color="0070C0", fill_type="solid")
                            
                            ws.column_dimensions['A'].width = 10
                            ws.column_dimensions['B'].width = 15
                            ws.column_dimensions['C'].width = 40
                            ws.column_dimensions['D'].width = 25
                            ws.column_dimensions['E'].width = 15
                            ws.column_dimensions['F'].width = 50
                            ws.column_dimensions['G'].width = 12
                            ws.column_dimensions['H'].width = 12
                            ws.column_dimensions['I'].width = 10
                            ws.column_dimensions['J'].width = 12
                            
                            for col in range(1, len(basliklar) + 1):
                                cell = ws.cell(row=1, column=col)
                                cell.font = baslik_font
                                cell.fill = baslik_dolgu
                                cell.alignment = Alignment(horizontal='center', vertical='center')
                            
                            ws.auto_filter.ref = ws.dimensions
                            
                            processed = 0
                            notsuz_foto_sayisi = 0
                            ozel_satir_sayisi = 0
                            notlu_foto_toplam = 0
                            
                            toplam_orjinal_boyut = 0
                            toplam_yeni_boyut = 0
                            
                            self.kaydetme_bar.setValue(1)
                            self.kaydetme_bar.setFormat(f"Excel hazırlandı... 1/{total_fotos+1}")
                            QApplication.processEvents()
                            
                            for direk in direk_fotolari:
                                direk_id = direk['id']
                                lat = direk.get("lat", "")
                                lon = direk.get("lon", "")
                                foto_bilgileri = direk.get('photos', [])
                                numara = direk.get("numara", "")
                                silinecek_mi = direk.get("silinecek_mi", False)
                                koordinat_degisecek = direk.get("koordinat_degisecek", False)
                                draggable = direk.get("draggable", False)
                                nihai_direk = direk.get("nihai_direk", False)
                                yeni_direk = direk.get("yeni_direk", False)
                                envanter_notu = direk.get("envanter_notu", "")
                                birim = direk.get("birim", "Cbs")
                                
                                print(f"🏗️ Direk işleniyor: {direk_id}, Foto sayısı: {len(foto_bilgileri)}")
                                
                                direk_koordinati_degisti = (draggable and koordinat_degisecek)
                                
                                ozel_durum_var = (silinecek_mi or direk_koordinati_degisti or yeni_direk)
                                
                                if ozel_durum_var:
                                    ozel_satir_sayisi += 1
                                    if silinecek_mi:
                                        aciklama = f"({direk_id}) numaralı direk CBS'den silinecek"
                                        birim = "Cbs"
                                        yapildi_mi = ""
                                    elif direk_koordinati_degisti:
                                        aciklama = f"({direk_id}) Koordinatı değişecek"
                                        birim = "Cbs"
                                        yapildi_mi = ""
                                    elif yeni_direk:
                                        aciklama = envanter_notu
                                        birim = "Cbs"
                                        yapildi_mi = ""
                                    
                                    direk_no_degeri = direk_id
                                    try:
                                        float(direk_id)
                                        direk_no_degeri = float(direk_id) if '.' in str(direk_id) else int(direk_id)
                                    except (ValueError, TypeError):
                                        direk_no_degeri = str(direk_id)
                                    
                                    satir_verisi = [numara, direk_no_degeri, aciklama, "", "", "", lat, lon, birim, yapildi_mi]
                                    ws.append(satir_verisi)
                                    
                                    current_row = ws.max_row
                                    direk_no_cell = ws.cell(row=current_row, column=2)
                                    
                                    if isinstance(direk_no_degeri, str):
                                        direk_no_cell.number_format = '@'
                                    
                                    for col in range(1, 11):
                                        cell = ws.cell(row=current_row, column=col)
                                        cell.alignment = Alignment(horizontal='left', vertical='center')
                                    
                                    print(f"📝 Özel durum satırı eklendi: {direk_id} - {aciklama}")
                                
                                notlu_foto_sayisi = 0
                                
                                if foto_bilgileri:
                                    klasor_adi = f"{numara} - {direk_id}" if numara else direk_id
                                    hedef_klasor = os.path.join(self.drone_kayit_klasoru, klasor_adi)
                                    print(f"📂 Klasör oluşturuluyor: {hedef_klasor}")
                                    os.makedirs(hedef_klasor, exist_ok=True)
                                    
                                    for i, foto in enumerate(foto_bilgileri):
                                        foto_yol = foto.get('path', '')
                                        if not foto_yol:
                                            continue
                                            
                                        foto_not = foto.get('note', '') or ""
                                        foto_birim = foto.get('birim', 'Aob') or "Aob"
                                        kategori = foto.get("kategori", "") or ""
                                        oncelik = foto.get("oncelik", "") or "Bekleyebilir"
                                        
                                        if not foto_not.strip():
                                            notsuz_foto_sayisi += 1
                                            print(f"⏭️  NOTSUZ fotoğraf atlandı: {direk_id} - Foto {i+1}")
                                            try:
                                                foto_yol_fs = uri_to_path(foto_yol)
                                                hedef_yol = os.path.join(hedef_klasor, os.path.basename(foto_yol_fs))
                                                
                                                orjinal_boyut = os.path.getsize(foto_yol_fs)
                                                toplam_orjinal_boyut += orjinal_boyut
                                                
                                                if kuculterek_kaydet:
                                                    self._fotografi_kucult(
                                                        foto_yol_fs, 
                                                        hedef_yol, 
                                                        kalite_ayarlari
                                                    )
                                                else:
                                                    shutil.copy2(foto_yol_fs, hedef_yol)
                                                    
                                                yeni_boyut = os.path.getsize(hedef_yol)
                                                toplam_yeni_boyut += yeni_boyut
                                                
                                                print(f"📁 Dosya kopyalandı (Excel'e eklenmedi): {hedef_yol}")
                                                
                                            except Exception as e:
                                                print(f"❌ Kopyalama hatası: {e}")
                                            
                                            processed += 1
                                            
                                            self.kaydetme_bar.setValue(processed + 1)
                                            ilerleme_yuzde = int((processed + 1) / (total_fotos + 1) * 100)
                                            self.kaydetme_bar.setFormat(f"İşleniyor... %{ilerleme_yuzde} ({processed}/{total_fotos})")
                                            QApplication.processEvents()
                                            
                                            continue
                                        
                                        notlu_foto_sayisi += 1
                                        notlu_foto_toplam += 1
                                        
                                        try:
                                            foto_yol_fs = uri_to_path(foto_yol)
                                            hedef_yol = os.path.join(hedef_klasor, os.path.basename(foto_yol_fs))
                                            
                                            orjinal_boyut = os.path.getsize(foto_yol_fs)
                                            toplam_orjinal_boyut += orjinal_boyut
                                            
                                            if kuculterek_kaydet:
                                                self._fotografi_kucult(
                                                    foto_yol_fs, 
                                                    hedef_yol, 
                                                    kalite_ayarlari
                                                )
                                            else:
                                                shutil.copy2(foto_yol_fs, hedef_yol)
                                            
                                            yeni_boyut = os.path.getsize(hedef_yol)
                                            toplam_yeni_boyut += yeni_boyut
                                            
                                            excel_klasor = os.path.dirname(excel_yolu)
                                            goreceli_yol = os.path.relpath(hedef_yol, excel_klasor)
                                            hyperlink_path = goreceli_yol.replace("\\", "/")
                                            görünen_ad = os.path.basename(hedef_yol)
                                            link = f'=HYPERLINK("{hyperlink_path}", "{görünen_ad}")'
                                            
                                            if foto_not.strip():
                                                aciklama = foto_not.strip()
                                                birim = "Aob" if foto_birim == "Aob" else "Cbs"
                                            elif kategori.strip():
                                                aciklama = f"[{kategori.strip()}]"
                                                birim = "Aob" if foto_birim == "Aob" else "Cbs"
                                            else:
                                                aciklama = f"({direk_id}) - Fotoğraf {i+1}"
                                                birim = "Aob" if foto_birim == "Aob" else "Cbs"
                                            
                                            yapildi_mi = ""
                                            if birim == "Aob":
                                                yapildi_mi = "Hayır"
                                            
                                            direk_no_degeri = direk_id
                                            try:
                                                float(direk_id)
                                                direk_no_degeri = float(direk_id) if '.' in str(direk_id) else int(direk_id)
                                            except (ValueError, TypeError):
                                                direk_no_degeri = str(direk_id)
                                            
                                            satir_verisi = [numara, direk_no_degeri, aciklama, kategori, oncelik, link, lat, lon, birim, yapildi_mi]
                                            ws.append(satir_verisi)
                                            
                                            current_row = ws.max_row
                                            direk_no_cell = ws.cell(row=current_row, column=2)
                                            
                                            if isinstance(direk_no_degeri, str):
                                                direk_no_cell.number_format = '@'
                                            
                                            for col in range(1, 11):
                                                cell = ws.cell(row=current_row, column=col)
                                                cell.alignment = Alignment(horizontal='left', vertical='center')
                                            
                                            print(f"📝 NOTLU fotoğraf eklendi: {direk_id} - {aciklama}")
                                            
                                        except Exception as e:
                                            print(f"❌ Kopyalama hatası: {e}")
                                        
                                        processed += 1
                                        
                                        self.kaydetme_bar.setValue(processed + 1)
                                        ilerleme_yuzde = int((processed + 1) / (total_fotos + 1) * 100)
                                        self.kaydetme_bar.setFormat(f"İşleniyor... %{ilerleme_yuzde} ({processed}/{total_fotos})")
                                        QApplication.processEvents()
                                        
                                        if total_fotos > 0:
                                            progress = int((processed / total_fotos) * 100)
                                            print(f"📊 İlerleme: {processed}/{total_fotos} ({progress}%)")
                                        
                                        QApplication.processEvents()
                                    
                                    print(f"📊 Direk {direk_id}: {notlu_foto_sayisi} notlu fotoğraf eklendi")
                                
                                if not ozel_durum_var and notlu_foto_sayisi == 0:
                                    print(f"ℹ️  Direk {direk_id}: Özel durum yok, notlu fotoğraf yok - Excel'e hiçbir satır eklenmedi")
                            
                            print(f"💾 Excel kaydediliyor: {excel_yolu}")
                            wb.save(excel_yolu)
                            
                            self.kaydetme_bar.setValue(total_fotos + 1)
                            self.kaydetme_bar.setFormat("✅ Tamamlandı")
                            QApplication.processEvents()
                            
                            kayit_turu = "küçültülerek" if kuculterek_kaydet else "orijinal boyutta"
                            
                            boyut_bilgisi = ""
                            if toplam_orjinal_boyut > 0:
                                orjinal_mb = toplam_orjinal_boyut / (1024 * 1024)
                                yeni_mb = toplam_yeni_boyut / (1024 * 1024)
                                tasarruf_mb = orjinal_mb - yeni_mb
                                tasarruf_orani = (tasarruf_mb / orjinal_mb) * 100 if orjinal_mb > 0 else 0
                                
                                boyut_bilgisi = f"\n\n📊 BOYUT KARŞILAŞTIRMASI:\n"
                                boyut_bilgisi += f"• Orijinal boyut: {orjinal_mb:.2f} MB\n"
                                boyut_bilgisi += f"• Yeni boyut: {yeni_mb:.2f} MB\n"
                                boyut_bilgisi += f"• Kazanılan alan: {tasarruf_mb:.2f} MB\n"
                                boyut_bilgisi += f"• Tasarruf oranı: %{tasarruf_orani:.1f}"
                            
                            if kuculterek_kaydet and kalite_ayarlari:
                                kalite_bilgisi = f"\n📏 Çözünürlük: {kalite_ayarlari['max_width']}x{kalite_ayarlari['max_height']}px\n🎯 Kalite: {kalite_ayarlari['quality']}%"
                            else:
                                kalite_bilgisi = ""
                            
                            QTimer.singleShot(3000, lambda: self.kaydetme_bar.setVisible(False))
                            
                            QMessageBox.information(self, "İşlem Tamamlandı", 
                                f"✅ İşlem tamamlandı!\n\n"
                                f"📸 Fotoğraflar {kayit_turu} kaydedildi.{kalite_bilgisi}\n"
                                f"📁 Klasör: {self.drone_kayit_klasoru}\n"
                                f"📄 Excel dosyası: Tespit.xlsx\n\n"
                                f"📊 İSTATİSTİKLER:\n"
                                f"• Toplam fotoğraf: {total_fotos}\n"
                                f"• Excel'e eklenen notlu fotoğraflar: {notlu_foto_toplam}\n"
                                f"• Atlanan notsuz fotoğraflar: {notsuz_foto_sayisi}\n"
                                f"• Eklenen özel durum satırları: {ozel_satir_sayisi}\n"
                                f"• Excel'e eklenen TOPLAM satır: {ws.max_row - 1}"
                                f"{boyut_bilgisi}")
                            
                        except Exception as e:
                            print(f"❌ Kopyalama sırasında hata: {e}")
                            import traceback
                            traceback.print_exc()
                            self.kaydetme_bar.setVisible(False)
                            QMessageBox.critical(self, "Hata", f"Kayıt sırasında hata oluştu:\n{str(e)}")

                    if iliskisiz_fotolar:
                        cevap = QMessageBox.question(self, "İliştirilmeyen Fotoğraflar Var",
                            f"⚠️ {len(iliskisiz_fotolar)} adet iliştirilmeyen fotoğraf var.\n"
                            "Yine de kaydetmek istiyor musunuz?",
                            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                        if cevap == QMessageBox.StandardButton.Yes:
                            kopyalamaya_basla()
                        else:
                            print("🚫 Kullanıcı iptal etti")
                            self.kaydetme_bar.setVisible(False)
                    else:
                        kopyalamaya_basla()
                        
                except Exception as e:
                    print(f"❌ JavaScript veri işleme hatası: {e}")
                    import traceback
                    traceback.print_exc()
                    self.kaydetme_bar.setVisible(False)
                    QMessageBox.critical(self, "Hata", f"Veri işleme hatası:\n{str(e)}")

            print("🔍 JavaScript çalıştırılıyor...")
            self.view.page().runJavaScript(js_kod, js_sonucu_al)
            
        except Exception as e:
            print(f"❌ Ana kayıt fonksiyonunda hata: {e}")
            import traceback
            traceback.print_exc()
            self.kaydetme_bar.setVisible(False)
            QMessageBox.critical(self, "Hata", f"Kayıt işlemi başlatılamadı:\n{str(e)}")
    
    def nihai_direkleri_kaydet(self, direk_fotolari):
        try:
            desktop = os.path.join(os.path.expanduser("~"), "Desktop")
            drone_dir = os.path.join(desktop, "Drone")
            nihai_excel_yolu = os.path.join(drone_dir, "Nihai_Direk.xlsx")
            
            print(f"🏁 Nihai direk Excel dosyası: {nihai_excel_yolu}")
            
            if not os.path.exists(drone_dir):
                os.makedirs(drone_dir, exist_ok=True)
                print(f"📁 Drone klasörü oluşturuldu: {drone_dir}")
            
            nihai_direkler = [d for d in direk_fotolari if d.get('nihai_direk', False)]
            
            if not nihai_direkler:
                print("ℹ️  Nihai direk bulunamadı")
                return
            
            print(f"🏁 {len(nihai_direkler)} adet nihai direk bulundu")
            
            direk_og_yolu = os.path.join(drone_dir, "direk_og.xlsx")
            direk_musterek_yolu = os.path.join(drone_dir, "direk_musterek.xlsx")
            
            if not os.path.exists(nihai_excel_yolu):
                wb_nihai = Workbook()
                ws_nihai = wb_nihai.active
                ws_nihai.append(["OBJECTID", "DIREKNO", "TIP", "AOB_ADI", "ENH_ID", "LAT", "LON"])
            else:
                wb_nihai = load_workbook(nihai_excel_yolu)
                ws_nihai = wb_nihai.active
            
            eklenen_direk_sayisi = 0
            
            for direk in nihai_direkler:
                direk_id = direk['id']
                print(f"🔍 Direk aranıyor: {direk_id}")
                
                found = False
                
                if os.path.exists(direk_og_yolu):
                    wb_og = load_workbook(direk_og_yolu, data_only=True)
                    ws_og = wb_og.active
                    
                    for row in ws_og.iter_rows(min_row=2):
                        direkno_cell = row[1]
                        if direkno_cell.value and str(direkno_cell.value).strip() == direk_id:
                            row_data = [cell.value for cell in row[:7]]
                            ws_nihai.append(row_data)
                            eklenen_direk_sayisi += 1
                            found = True
                            print(f"✅ Direk {direk_id} direk_og.xlsx'ten bulundu ve eklendi")
                            break
                
                if not found and os.path.exists(direk_musterek_yolu):
                    wb_musterek = load_workbook(direk_musterek_yolu, data_only=True)
                    ws_musterek = wb_musterek.active
                    
                    for row in ws_musterek.iter_rows(min_row=2):
                        direkno_cell = row[1]
                        if direkno_cell.value and str(direkno_cell.value).strip() == direk_id:
                            row_data = [cell.value for cell in row[:7]]
                            ws_nihai.append(row_data)
                            eklenen_direk_sayisi += 1
                            found = True
                            print(f"✅ Direk {direk_id} direk_musterek.xlsx'ten bulundu ve eklendi")
                            break
                
                if not found:
                    print(f"⚠️ Direk {direk_id} hiçbir Excel'de bulunamadı")
            
            wb_nihai.save(nihai_excel_yolu)
            
            QMessageBox.information(
                self,
                "Nihai Direkler Kaydedildi",
                f"🏁 {eklenen_direk_sayisi} adet nihai direk başarıyla kaydedildi.\n\n"
                f"📄 Dosya: Nihai_Direk.xlsx\n"
                f"📁 Klasör: {drone_dir}\n\n"
                f"ℹ️  Not: Direkler orijinal Excel'lerden kopyalanarak eklendi."
            )
            
        except Exception as e:
            print(f"❌ Nihai direk kaydetme hatası: {e}")
            import traceback
            traceback.print_exc()
            QMessageBox.warning(
                self,
                "Nihai Direk Kayıt Hatası",
                f"Nihai direkler kaydedilirken hata oluştu:\n{str(e)}"
            )
    
    def _fotografi_kucult(self, kaynak_yol, hedef_yol, kalite_ayarlari):
        try:
            with Image.open(kaynak_yol) as img:
                exif = img.info.get("exif", None)
                
                max_width = kalite_ayarlari.get('max_width', 4000)
                max_height = kalite_ayarlari.get('max_height', 4000)
                quality = kalite_ayarlari.get('quality', 30)
                preserve_exif = kalite_ayarlari.get('preserve_exif', True)
                
                orjinal_genislik, orjinal_yukseklik = img.size
                print(f"📏 {os.path.basename(kaynak_yol)}: {orjinal_genislik}x{orjinal_yukseklik} → {max_width}x{max_height}")
                
                img.thumbnail((max_width, max_height))
                
                save_params = {}
                if preserve_exif and exif:
                    save_params['exif'] = exif
                
                if img.format == 'JPEG' or hedef_yol.lower().endswith(('.jpg', '.jpeg')):
                    save_params['quality'] = quality
                    save_params['optimize'] = True
                
                img.save(hedef_yol, **save_params)
                
                orjinal_boyut = os.path.getsize(kaynak_yol) / 1024
                yeni_boyut = os.path.getsize(hedef_yol) / 1024
                oran = (yeni_boyut / orjinal_boyut) * 100
                
                print(f"   📊 Boyut: {orjinal_boyut:.1f}KB → {yeni_boyut:.1f}KB (%{oran:.1f})")
                
        except Exception as e:
            print(f"❌ Küçültme hatası ({kaynak_yol}): {e}")
            shutil.copy2(kaynak_yol, hedef_yol)

    def fotograf_klasoru_sec(self):
        klasor = QFileDialog.getExistingDirectory(self, "Fotoğraf Klasörünü Seç")
        if klasor:
            self.foto_klasor = klasor
            self.foto_klasor_label.setText(f"📁 {klasor}")
            print(f"✅ Fotoğraf klasörü seçildi: {klasor}")
        else:
            print("❌ Fotoğraf klasörü seçilmedi")
            
    def drone_kayit_klasoru_sec(self):
        klasor = QFileDialog.getExistingDirectory(self, "Drone Kayıt Dosyası Klasörü Seç")
        if klasor:
            self.drone_kayit_klasoru = klasor
            self.kayit_label.setText(f"📁 Seçilen kayıt klasörü:\n{klasor}")
            print(f"📁 Drone kayıt klasörü seçildi: {klasor}")

    def foto_kayit_klasoru_sec(self):
        klasor = QFileDialog.getExistingDirectory(self, "Fotoğraf Küçültme Kayıt Klasörü Seç")
        if klasor:
            self.foto_kayit_klasor = klasor
            self.kayit_klasor_label.setText(f"📁 {klasor}")
            print(f"✅ Fotoğraf küçültme kayıt klasörü seçildi: {klasor}")
        else:
            print("❌ Kayıt klasörü seçilmedi")

    def kucultmeyi_baslat(self):
        print("=" * 50)
        print("DEBUG: Küçültme işlemi başlatılıyor...")
        print(f"DEBUG: foto_klasor değeri: '{self.foto_klasor}'")
        print(f"DEBUG: foto_kayit_klasor değeri: '{self.foto_kayit_klasor}'")
        print("=" * 50)
        
        if not self.foto_klasor and not self.foto_kayit_klasor:
            QMessageBox.warning(self, "Klasörler Seçilmemiş", 
                              "⚠️ Önce fotoğraf klasörünü VE kayıt klasörünü seçmelisiniz!\n\n"
                              "1. '📂 Fotoğraf Klasörü Seç' butonuna tıklayın\n"
                              "2. '💾 Kayıt Klasörü Seç' butonuna tıklayın\n\n"
                              "Daha sonra küçültme işlemine başlayabilirsiniz.")
            print("❌ Hiçbir klasör seçilmemiş")
            return
        
        if self.foto_klasor and not self.foto_kayit_klasor:
            QMessageBox.warning(self, "Kayıt Klasörü Seçilmemiş", 
                              "⚠️ Kayıt klasörünü seçmelisiniz!\n\n"
                              "Fotoğraf klasörü seçildi ✓\n"
                              "Kayıt klasörü seçilmedi ✗\n\n"
                              "Lütfen '💾 Kayıt Klasörü Seç' butonuna tıklayın.")
            print("❌ Sadece fotoğraf klasörü seçilmiş, kayıt klasörü eksik")
            return
        
        if not self.foto_klasor and self.foto_kayit_klasor:
            QMessageBox.warning(self, "Fotoğraf Klasörü Seçilmemiş", 
                              "⚠️ Fotoğraf klasörünü seçmelisiniz!\n\n"
                              "Fotoğraf klasörü seçilmedi ✗\n"
                              "Kayıt klasörü seçildi ✓\n\n"
                              "Lütfen '📂 Fotoğraf Klasörü Seç' butonuna tıklayın.")
            print("❌ Sadece kayıt klasörü seçilmiş, fotoğraf klasörü eksik")
            return
        
        print("✅ Her iki klasör de seçilmiş")
        
        if not os.path.exists(self.foto_klasor):
            QMessageBox.warning(self, "Fotoğraf Klasörü Bulunamadı", 
                              f"⚠️ Seçilen fotoğraf klasörü bulunamadı:\n{self.foto_klasor}\n\n"
                              f"Lütfen geçerli bir klasör seçin.")
            print(f"❌ Fotoğraf klasörü bulunamadı: {self.foto_klasor}")
            return
        
        if not os.path.exists(self.foto_kayit_klasor):
            reply = QMessageBox.question(
                self, 
                "Kayıt Klasörü Yok", 
                f"Kayıt klasörü bulunamadı:\n{self.foto_kayit_klasor}\n\n"
                f"Bu klasörü otomatik oluşturmak ister misiniz?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )
            
            if reply == QMessageBox.StandardButton.No:
                print("🚫 Kullanıcı klasör oluşturmayı reddetti")
                return
            else:
                try:
                    os.makedirs(self.foto_kayit_klasor, exist_ok=True)
                    print(f"✅ Kayıt klasörü oluşturuldu: {self.foto_kayit_klasor}")
                    QMessageBox.information(self, "Klasör Oluşturuldu", 
                                          f"Kayıt klasörü başarıyla oluşturuldu:\n{self.foto_kayit_klasor}")
                except Exception as e:
                    QMessageBox.critical(self, "Hata", 
                                       f"Klasör oluşturulamadı:\n{str(e)}")
                    return
        
        dialog = KaliteAyarlariDialog(self)
        result = dialog.exec()
        
        if result != QDialog.DialogCode.Accepted:
            print("🚫 Kullanıcı kalite ayarlarını iptal etti")
            return
        
        kalite_ayarlari = dialog.ayarlari_al()
        print(f"⚙️ Küçültme ayarları alındı: {kalite_ayarlari}")

        yollar = []
        
        print(f"🔍 Fotoğraflar aranıyor: {self.foto_klasor}")
        for root, _, files in os.walk(self.foto_klasor):
            for file in files:
                if file.lower().endswith((".jpg", ".jpeg", ".png", ".bmp", ".webp", ".tif", ".tiff")):
                    kaynak = os.path.join(root, file)
                    alt_yol = os.path.relpath(kaynak, self.foto_klasor)
                    hedef = os.path.join(self.foto_kayit_klasor, alt_yol)   
                    yollar.append((kaynak, hedef))
                    print(f"📸 Fotoğraf bulundu: {file}")

        if not yollar:
            QMessageBox.information(self, "Fotoğraf Bulunamadı", 
                                  f"⚠️ Seçilen klasörde uygun fotoğraf bulunamadı.\n\n"
                                  f"Klasör: {self.foto_klasor}\n\n"
                                  f"Desteklenen formatlar:\n"
                                  f"• JPG, JPEG\n"
                                  f"• PNG\n"
                                  f"• BMP\n"
                                  f"• WEBP\n"
                                  f"• TIF, TIFF")
            print(f"❌ Uygun fotoğraf bulunamadı: {self.foto_klasor}")
            return

        print(f"✅ {len(yollar)} fotoğraf bulundu")

        reply = QMessageBox.question(
            self, 
            "Küçültme Başlatılsın mı?", 
            f"📊 {len(yollar)} adet fotoğraf küçültülecek.\n\n"
            f"📂 Kaynak Klasör:\n{self.foto_klasor}\n"
            f"📁 Hedef Klasör:\n{self.foto_kayit_klasor}\n\n"
            f"⚙️ Kullanılacak Ayarlar:\n"
            f"• Maksimum boyut: {kalite_ayarlari['max_width']}x{kalite_ayarlari['max_height']} px\n"
            f"• JPEG Kalitesi: {kalite_ayarlari['quality']}%\n"
            f"• EXIF koruma: {'Evet' if kalite_ayarlari['preserve_exif'] else 'Hayır'}\n\n"
            f"Başlatmak istiyor musunuz?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.No:
            print("🚫 Kullanıcı işlemi iptal etti")
            return

        print("🔄 Küçültme işlemi başlıyor...")
        
        self.kucultme_bar.setVisible(True)
        self.kucultme_bar.setMaximum(len(yollar))
        self.kucultme_bar.setValue(0)
        self.kucultme_bar.setFormat("Hazırlanıyor... 0%")

        toplam_orjinal_boyut = 0
        toplam_yeni_boyut = 0
        basarili_sayisi = 0
        hata_sayisi = 0
        
        for i, (kaynak, hedef) in enumerate(yollar):
            try:
                os.makedirs(os.path.dirname(hedef), exist_ok=True)
                
                orjinal_boyut = os.path.getsize(kaynak)
                toplam_orjinal_boyut += orjinal_boyut
                
                self._fotografi_kucult(kaynak, hedef, kalite_ayarlari)
                
                yeni_boyut = os.path.getsize(hedef)
                toplam_yeni_boyut += yeni_boyut
                basarili_sayisi += 1
                
                print(f"✅ {i+1}/{len(yollar)}: {os.path.basename(kaynak)} küçültüldü")
                
            except Exception as e:
                hata_sayisi += 1
                print(f"❌ {i+1}/{len(yollar)}: {os.path.basename(kaynak)} - HATA: {e}")

            ilerleme_yuzde = int((i + 1) / len(yollar) * 100)
            self.kucultme_bar.setValue(i + 1)
            self.kucultme_bar.setFormat(f"İşleniyor... {ilerleme_yuzde}% ({i+1}/{len(yollar)})")
            QApplication.processEvents()

        self.kucultme_bar.setFormat("✅ Tamamlandı")
        
        excel_kopyalandi = 0
        try:
            print("\n📊 Excel dosyaları aranıyor...")
            for root, _, files in os.walk(self.foto_klasor):
                for file in files:
                    if file.lower().endswith((".xlsx", ".xls")):
                        kaynak_excel = os.path.join(root, file)
                        alt_yol = os.path.relpath(kaynak_excel, self.foto_klasor)
                        hedef_excel = os.path.join(self.foto_kayit_klasor, alt_yol)
                        
                        try:
                            os.makedirs(os.path.dirname(hedef_excel), exist_ok=True)
                            shutil.copy2(kaynak_excel, hedef_excel)
                            excel_kopyalandi += 1
                            print(f"📄 Excel kopyalandı: {file} → {hedef_excel}")
                        except Exception as e:
                            print(f"❌ Excel kopyalama hatası ({file}): {e}")
        except Exception as e:
            print(f"❌ Excel dosyası aranırken hata: {e}")
        
        orjinal_mb = toplam_orjinal_boyut / (1024 * 1024)
        yeni_mb = toplam_yeni_boyut / (1024 * 1024)
        
        if orjinal_mb > 0:
            tasarruf_mb = orjinal_mb - yeni_mb
            tasarruf_orani = (tasarruf_mb / orjinal_mb) * 100
        else:
            tasarruf_mb = 0
            tasarruf_orani = 0
        
        excel_bilgisi = f"\n\n📄 EXCEL KOPYALAMA:\n• Kopyalanan Excel dosyası: {excel_kopyalandi} adet" if excel_kopyalandi > 0 else "\n\n📄 Excel dosyası bulunamadı."
        
        QMessageBox.information(
            self, 
            "Küçültme İşlemi Tamamlandı",
            f"✅ Küçültme işlemi başarıyla tamamlandı!\n\n"
            f"📊 İSTATİSTİKLER:\n"
            f"• Toplam fotoğraf: {len(yollar)}\n"
            f"• Başarıyla küçültülen: {basarili_sayisi}\n"
            f"• Hata alan: {hata_sayisi}\n\n"
            f"📊 BOYUT KARŞILAŞTIRMASI:\n"
            f"• Orijinal boyut: {orjinal_mb:.2f} MB\n"
            f"• Yeni boyut: {yeni_mb:.2f} MB\n"
            f"• Kazanılan alan: {tasarruf_mb:.2f} MB\n"
            f"• Tasarruf oranı: %{tasarruf_orani:.1f}\n\n"
            f"⚙️ KULLANILAN AYARLAR:\n"
            f"• Maksimum boyut: {kalite_ayarlari['max_width']}x{kalite_ayarlari['max_height']} px\n"
            f"• Kalite: {kalite_ayarlari['quality']}%\n"
            f"• EXIF koruma: {'Evet' if kalite_ayarlari['preserve_exif'] else 'Hayır'}"
            f"{excel_bilgisi}\n\n"
            f"📁 Kayıt klasörü:\n{self.foto_kayit_klasor}"
        )
        
        print(f"📦 Küçültme tamamlandı: {basarili_sayisi} başarılı, {hata_sayisi} hatalı, {excel_kopyalandi} excel dosyası kopyalandı")
        
        QTimer.singleShot(3000, lambda: self.kucultme_bar.setVisible(False))


def create_map(photo_data, direk_data, save_path, center_lat=None, center_lon=None, zoom_level=None):
    SABIT_LAT = 36.7300
    SABIT_LON = 30.2000
    SABIT_ZOOM = 9.8
    
    if center_lat is None or center_lon is None or zoom_level is None:
        center_lat = SABIT_LAT
        center_lon = SABIT_LON
        zoom_level = SABIT_ZOOM
    
    print(f"🗺️ Harita konumu: {center_lat}, {center_lon} - Zoom: {zoom_level}")
    
    start_coords = (center_lat, center_lon)
    
    if direk_data:
        print(f"📍 {len(direk_data)} direk haritaya eklenecek")
    if photo_data:
        print(f"📸 {len(photo_data)} fotoğraf haritaya eklenecek")
    
    html_start = f"""
    <!DOCTYPE html>
    <html>
    <head>
        <meta charset="utf-8" />
        <title>Drone Harita</title>
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css" />
        <link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.css" />
        <link rel="stylesheet" href="https://unpkg.com/leaflet.markercluster@1.5.3/dist/MarkerCluster.Default.css" />
        <script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
        <script src="https://unpkg.com/leaflet.markercluster@1.5.3/dist/leaflet.markercluster.js"></script>

        <style>
            html, body, #map {{
                height: 100%;
                margin: 0;
                padding: 0;
            }}
            
            .photo-marker {{
                width: 16px;
                height: 16px;
                border-radius: 50%;
                border: 2px solid white;
                box-shadow: 0 0 4px #000;
                z-index: 800 !important;
                pointer-events: auto !important;
            }}
            .photo-mavi {{
                background-color: #007BFF;
                opacity: 0.9;
            }}
            .photo-turkuaz {{
                background-color: #00CED1;
                opacity: 0.9;
            }}
            
            .custom-popup .leaflet-popup-content-wrapper {{
                width: 420px !important;
            }}
            
            .leaflet-control-tum-fotolari-cikar,
            .leaflet-control-tumunu-goster,
            .leaflet-control-gizle-goster {{
                background: white;
                border-radius: 4px;
                border: 2px solid rgba(0,0,0,0.2);
                cursor: pointer;
            }}
            .leaflet-control-tum-fotolari-cikar a,
            .leaflet-control-tumunu-goster a,
            .leaflet-control-gizle-goster a {{
                display: block;
                width: 30px;
                height: 30px;
                line-height: 30px;
                text-align: center;
                text-decoration: none;
                color: #333;
                font-size: 18px;
            }}
            .leaflet-control-tum-fotolari-cikar:hover,
            .leaflet-control-tumunu-goster:hover,
            .leaflet-control-gizle-goster:hover {{
                background: #f4f4f4;
            }}
            
            .direk-marker-normal,
            .direk-marker-musterek,
            .direk-marker-yeni,
            .direk-marker-foto-var,
            .direk-marker-konum-degisti,
            .direk-marker-nihai {{
                width: 14px;
                height: 14px;
                border-radius: 50%;
                z-index: 700 !important;
                pointer-events: auto !important;
            }}
            
            .direk-marker-normal {{
                background-color: #ff6666;
                border: 2px solid #ff0000;
            }}
            .direk-marker-musterek {{
                background-color: #9370DB;
                border: 2px solid #800080;
            }}
            .direk-marker-yeni {{
                background-color: #FFC0CB;
                border: 2px solid #C71585;
            }}
            .direk-marker-foto-var {{
                background-color: #90EE90;
                border: 2px solid #32CD32;
            }}
            .direk-marker-konum-degisti {{
                background-color: #FFFF00;
                border: 2px solid #FFA500;
            }}
            .direk-marker-nihai {{
                background-color: #007BFF;
                border: 2px solid #0056b3;
            }}
            
            .direk-numara-container {{
                width: 24px;
                height: 24px;
                display: flex;
                align-items: center;
                justify-content: center;
                pointer-events: none !important;
                z-index: 1000 !important;
                background: transparent;
            }}
            
            .direk-numara-text {{
                color: black;
                font-weight: bold;
                font-size: 14px;
                text-shadow: 2px 2px 0 white, -2px -2px 0 white, 2px -2px 0 white, -2px 2px 0 white;
                line-height: 1;
                background: transparent;
                white-space: nowrap;
            }}
            
            .direk-ikon-button {{
                background-color: #4CAF50;
                border: none;
                color: white;
                padding: 4px 8px;
                text-align: center;
                text-decoration: none;
                display: inline-block;
                font-size: 12px;
                margin: 2px 1px;
                cursor: pointer;
                border-radius: 3px;
            }}
            .direk-ikon-button:hover {{
                background-color: #45a049;
            }}
            
            .marker-cluster-small {{
                background-color: rgba(181, 226, 140, 0.6);
            }}
            .marker-cluster-small div {{
                background-color: rgba(110, 204, 57, 0.6);
            }}
            .marker-cluster-medium {{
                background-color: rgba(241, 211, 87, 0.6);
            }}
            .marker-cluster-medium div {{
                background-color: rgba(240, 194, 12, 0.6);
            }}
            .marker-cluster-large {{
                background-color: rgba(253, 156, 115, 0.6);
            }}
            .marker-cluster-large div {{
                background-color: rgba(241, 128, 23, 0.6);
            }}
        </style>
    </head>
    <body>
        <div id="map"></div>
        <script>
            function encodeID(str) {{
                return btoa(unescape(encodeURIComponent(str))).replace(/[/+=]/g, '_');
            }}

            function decodeID(encoded) {{
                try {{
                    let safe = encoded.replace(/_/g, '/').replace(/-/g, '+');
                    while (safe.length % 4) {{
                        safe += '=';
                    }}
                    return decodeURIComponent(atob(safe));
                }} catch (e) {{
                    console.error("Decode hatası:", e);
                    return encoded;
                }}
            }}
            
            function getPhotoPopupHTML(path, note, birim="") {{
                const encoded = encodeID(path);
                
                const aobChecked = birim === "Cbs" ? "" : "checked";
                const cbsChecked = birim === "Cbs" ? "checked" : "";
                
                const filename = path.split('/').pop().split('\\\\').pop();
                const folder = path.split('/').slice(-2, -1)[0] || "";
                const displayPath = folder ? folder + "/" + filename : filename;
                
                return `
                    <div style='width: 300px; max-width: 90%; line-height: 1.2;'>
                        <b style="word-break: break-all; font-size: 12px; margin: 0; display: block;">📁 ${{displayPath}}</b>
                        <img src='${{path}}' style='width: 100%; max-width: 280px; border-radius: 3px; box-shadow: 0 0 3px #0002; margin: 2px 0; display: block;'>
                        
                        <div style="margin: 3px 0 2px 0;">
                            <label style="font-size: 12px; font-weight: bold; margin: 0; display: block;">📝 Not:</label>
                            <textarea id="note-input-${{encoded}}" 
                                      style="width: 95%; max-width: 280px; 
                                             padding: 4px;
                                             margin-top: 2px;
                                             border: 1px solid #ccc;
                                             border-radius: 2px;
                                             font-size: 12px;
                                             line-height: 1.3;" 
                                      rows="2">${{note || ""}}</textarea>
                        </div>
                        
                        <div style="margin: 5px 0; display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 8px;">
                            <div style="display: flex; align-items: center; gap: 12px;">
                                <label style="font-weight: bold; font-size: 12px; margin: 0;">Birim Seç:</label>
                                <label style="display: flex; align-items: center; gap: 3px; cursor: pointer; margin: 0;">
                                    <input type="radio" id="aob-${{encoded}}" name="birim-${{encoded}}" value="Aob" ${{aobChecked}} style="margin: 0;">
                                    <span style="font-size: 12px;">Aob</span>
                                </label>
                                <label style="display: flex; align-items: center; gap: 3px; cursor: pointer; margin: 0;">
                                    <input type="radio" id="cbs-${{encoded}}" name="birim-${{encoded}}" value="Cbs" ${{cbsChecked}} style="margin: 0;">
                                    <span style="font-size: 12px;">Cbs</span>
                                </label>
                            </div>
                            
                            <button id="note-save-${{encoded}}" 
                                    style="background-color: #007bff; 
                                           color: white; 
                                           border: none; 
                                           padding: 4px 12px; 
                                           border-radius: 2px; 
                                           cursor: pointer;
                                           font-size: 12px;">
                                💾 Kaydet
                            </button>
                        </div>
                    </div>
                `;
            }}
            
            function bindNoteInput(path) {{
                const foto = photoInfo.find(p => p.path === path);
                if (!foto) return;
                
                const encoded = encodeID(path);
                const noteInput = document.querySelector(`#note-input-${{encoded}}`);
                const button = document.querySelector(`#note-save-${{encoded}}`);
                const aobRadio = document.querySelector(`#aob-${{encoded}}`);
                const cbsRadio = document.querySelector(`#cbs-${{encoded}}`);
                
                if (noteInput && button) {{
                    button.onclick = function () {{
                        let yeniNot = noteInput.value;
                        let birim = "";
                        
                        if (aobRadio && aobRadio.checked) {{
                            birim = "Aob";
                        }} else if (cbsRadio && cbsRadio.checked) {{
                            birim = "Cbs";
                        }}
                        
                        foto.note = yeniNot;
                        foto.birim = birim;
                        
                        if (foto.marker && foto.marker._icon) {{
                            foto.marker._icon.classList.remove("photo-mavi");
                            foto.marker._icon.classList.add("photo-turkuaz");
                        }}
                        
                        const newContent = getPhotoPopupHTML(foto.path, foto.note, foto.birim);
                        foto.marker.setPopupContent(newContent);
                        
                        direkler.forEach(d => {{
                            if (d.photos) {{
                                d.photos.forEach(f => {{
                                    if (f.path === path) {{
                                        f.note = yeniNot;
                                        f.birim = birim;
                                    }}
                                }});
                                guncellePopup(d);
                            }}
                        }});
                        
                        console.log(`✅ Not ve birim kaydedildi: ${{yeniNot}} - ${{birim}}`);
                    }};
                }}
            }}
            
            function kaldirDirek(direkId) {{
                console.log("🗑️ Direk kaldırma işlemi başlatıldı:", direkId);
                
                const direk = direkler.find(d => d.id === direkId);
                if (!direk) {{
                    console.log("❌ Kaldırılacak direk bulunamadı:", direkId);
                    alert("❌ Direk bulunamadı!");
                    return;
                }}
                
                if (direk.photos && direk.photos.length > 0) {{
                    console.log("📸 Direğe ait fotoğraflar haritaya geri ekleniyor:", direk.photos.length);
                    
                    direk.photos.forEach(foto => {{
                        const photoInfoObj = photoInfo.find(p => p.path === foto.path);
                        
                        if (photoInfoObj) {{
                            const lat = photoInfoObj.originalLat !== undefined ? photoInfoObj.originalLat : photoInfoObj.lat;
                            const lon = photoInfoObj.originalLon !== undefined ? photoInfoObj.originalLon : photoInfoObj.lon;
                            
                            var markerClass = "photo-marker " + (photoInfoObj.note ? "photo-turkuaz" : "photo-mavi");
                            
                            var marker = L.marker([lat, lon], {{
                                draggable: true,
                                icon: L.divIcon({{
                                    className: markerClass,
                                    iconSize: [13, 13]
                                }}),
                                zIndexOffset: 800
                            }}).bindPopup(getPhotoPopupHTML(photoInfoObj.path, photoInfoObj.note), {{
                                maxWidth: 420,
                                minWidth: 400,
                                autoPan: true,
                                className: 'custom-popup'
                            }});
                            
                            marker.on("contextmenu", function(e) {{
                                e.originalEvent.preventDefault();
                                e.originalEvent.stopPropagation();
                                
                                var pos = e.target.getLatLng();
                                var hedefDirek = null;
                                var enKisaMesafe = Infinity;
                                
                                direkler.forEach(function(d) {{
                                    if (d.id === direkId) return;
                                    var mesafe = map.distance(pos, L.latLng(d.lat, d.lon));
                                    if (mesafe < 40 && mesafe < enKisaMesafe) {{
                                        hedefDirek = d;
                                        enKisaMesafe = mesafe;
                                    }}
                                }});
                                
                                if (hedefDirek) {{
                                    var zatenVar = hedefDirek.photos.some(p => p.path === photoInfoObj.path);
                                    if (zatenVar) {{
                                        L.popup()
                                            .setLatLng(pos)
                                            .setContent("⛔ Aynı fotoğraf zaten bu direğe eklenmiş")
                                            .openOn(map);
                                        return;
                                    }}
                                    
                                    const notDegeri = photoInfoObj.note || "";
                                    const kategori = photoInfoObj.kategori || "";
                                    const oncelik = photoInfoObj.oncelik || "Bekleyebilir";
                                    const birim = photoInfoObj.birim || "Aob";
                                    
                                    hedefDirek.photos.push({{
                                        path: photoInfoObj.path,
                                        note: notDegeri,
                                        kategori: kategori,
                                        oncelik: oncelik,
                                        birim: birim
                                    }});
                                    
                                    guncelleDirekRengi(hedefDirek);
                                    guncellePopup(hedefDirek);
                                    map.removeLayer(marker);
                                }} else {{
                                    L.popup()
                                        .setLatLng(pos)
                                        .setContent("⚠️ 40 metre içinde direk bulunamadı")
                                        .openOn(map);
                                }}
                            }});
                            
                            marker.on("popupopen", function() {{
                                bindNoteInput(photoInfoObj.path);
                            }});
                            
                            photoInfoObj.marker = marker;
                            map.addLayer(marker);
                        }}
                    }});
                }}
                
                if (direk.marker && map.hasLayer(direk.marker)) {{
                    map.removeLayer(direk.marker);
                }}
                
                if (direk.numaraMarker && map.hasLayer(direk.numaraMarker)) {{
                    map.removeLayer(direk.numaraMarker);
                }}
                
                if (direklerCluster && direklerCluster.hasLayer(direk.marker)) {{
                    direklerCluster.removeLayer(direk.marker);
                }}
                
                const index = direkler.findIndex(d => d.id === direkId);
                if (index !== -1) {{
                    direkler.splice(index, 1);
                    console.log("✅ Direk diziden kaldırıldı:", direkId);
                }}
                
                alert("✅ '" + direkId + "' direği başarıyla kaldırıldı.");
            }}

            function tumFotolariCikar() {{
                console.log("🔄 Tüm fotoğraflar dışarı alınıyor...");
                const tumFotograflar = [];
                
                direkler.forEach(direk => {{
                    if (direk.photos && direk.photos.length > 0) {{
                        direk.photos.forEach(foto => {{
                            tumFotograflar.push({{
                                path: foto.path,
                                direk_id: direk.id,
                                note: foto.note || ""
                            }});
                        }});
                        direk.photos = [];
                        
                        if (direk.numara) {{
                            direk.numara = null;
                            if (direk.numaraMarker && map.hasLayer(direk.numaraMarker)) {{
                                map.removeLayer(direk.numaraMarker);
                                direk.numaraMarker = null;
                            }}
                        }}
                    }}
                }});
                
                console.log(`📸 ${{tumFotograflar.length}} fotoğraf direklerden çıkarılıyor...`);
                
                tumFotograflar.forEach(foto => {{
                    const photoInfoObj = photoInfo.find(p => p.path === foto.path);
                    
                    if (photoInfoObj) {{
                        // ORİJİNAL KOORDİNATLARI KULLAN
                        const lat = photoInfoObj.originalLat !== undefined && photoInfoObj.originalLat !== null 
                            ? photoInfoObj.originalLat 
                            : photoInfoObj.lat;
                        const lon = photoInfoObj.originalLon !== undefined && photoInfoObj.originalLon !== null 
                            ? photoInfoObj.originalLon 
                            : photoInfoObj.lon;
                        
                        console.log(`📍 Fotoğraf orijinal konumuna dönüyor: ${{foto.path}} -> (${{lat}}, ${{lon}})`);
                        
                        var markerClass = "photo-marker " + (photoInfoObj.note ? "photo-turkuaz" : "photo-mavi");

                        var marker = L.marker([lat, lon], {{
                            draggable: true,
                            icon: L.divIcon({{
                                className: markerClass,
                                iconSize: [13, 13]
                            }}),
                            zIndexOffset: 800
                        }}).bindPopup(getPhotoPopupHTML(photoInfoObj.path, photoInfoObj.note, photoInfoObj.birim), {{
                            maxWidth: 420,
                            minWidth: 400,
                            autoPan: true,
                            className: 'custom-popup'
                        }});

                        marker.on("contextmenu", function(e) {{
                            e.originalEvent.preventDefault();
                            e.originalEvent.stopPropagation();
                            
                            var pos = e.target.getLatLng();
                            var hedefDirek = null;
                            var enKisaMesafe = Infinity;

                            direkler.forEach(function(d) {{
                                var mesafe = map.distance(pos, L.latLng(d.lat, d.lon));
                                if (mesafe < 40 && mesafe < enKisaMesafe) {{
                                    hedefDirek = d;
                                    enKisaMesafe = mesafe;
                                }}
                            }});

                            if (hedefDirek) {{
                                var zatenVar = hedefDirek.photos.some(p => p.path === foto.path);
                                if (zatenVar) {{
                                    L.popup()
                                        .setLatLng(pos)
                                        .setContent("⛔ Aynı fotoğraf zaten bu direğe eklenmiş")
                                        .openOn(map);
                                    return;
                                }}

                                const notDegeri = photoInfoObj.note || "";
                                const kategori = photoInfoObj.kategori || "";
                                const oncelik = photoInfoObj.oncelik || "Bekleyebilir";
                                const birim = photoInfoObj.birim || "Aob";

                                hedefDirek.photos.push({{
                                    path: foto.path,
                                    note: notDegeri,
                                    kategori: kategori,
                                    oncelik: oncelik,
                                    birim: birim
                                }});

                                if (hedefDirek.photos.length === 1 && !hedefDirek.numara) {{
                                    const kullanılanNumaralar = direkler
                                        .filter(d => typeof d.numara === 'number' && d.numara !== null)
                                        .map(d => d.numara);
                                    let yeniNumara = 1;
                                    while (kullanılanNumaralar.includes(yeniNumara)) {{
                                        yeniNumara++;
                                    }}
                                    hedefDirek.numara = yeniNumara;
                                    guncelleDirekNumaraGorunumu(hedefDirek);
                                }}

                                guncelleDirekRengi(hedefDirek);
                                guncellePopup(hedefDirek);
                                map.removeLayer(marker);
                            }} else {{
                                L.popup()
                                    .setLatLng(pos)
                                    .setContent("⚠️ 40 metre içinde direk bulunamadı")
                                    .openOn(map);
                            }}
                        }});

                        marker.on("popupopen", function() {{
                            bindNoteInput(photoInfoObj.path);
                        }});

                        marker.on("dragend", function(e) {{
                            var pos = e.target.getLatLng();
                            var hedefDirek = null;
                            var enKisaMesafe = Infinity;

                            direkler.forEach(function(d) {{
                                var mesafe = map.distance(pos, L.latLng(d.lat, d.lon));
                                if (mesafe < 8 && !d.photos.find(p => p.path === foto.path) && mesafe < enKisaMesafe) {{
                                    hedefDirek = d;
                                    enKisaMesafe = mesafe;
                                }}
                            }});

                            if (hedefDirek) {{
                                const notDegeri = photoInfoObj.note || "";
                                hedefDirek.photos.push({{ path: foto.path, note: notDegeri }});

                                if (hedefDirek.photos.length === 1 && !hedefDirek.numara) {{
                                    const kullanılanNumaralar = direkler
                                        .filter(d => typeof d.numara === 'number' && d.numara !== null)
                                        .map(d => d.numara);
                                    let yeniNumara = 1;
                                    while (kullanılanNumaralar.includes(yeniNumara)) {{
                                        yeniNumara++;
                                    }}
                                    hedefDirek.numara = yeniNumara;
                                    guncelleDirekNumaraGorunumu(hedefDirek);
                                }}

                                guncelleDirekRengi(hedefDirek);
                                guncellePopup(hedefDirek);
                                map.removeLayer(marker);
                            }}
                        }});

                        photoInfoObj.marker = marker;
                        photoInfoObj.lat = lat;
                        photoInfoObj.lon = lon;
                        map.addLayer(marker);
                    }} else {{
                        console.error("❌ PhotoInfo'da fotoğraf bulunamadı:", foto.path);
                    }}
                }});
                
                direkler.forEach(d => {{
                    guncelleDirekRengi(d);
                    guncellePopup(d);
                }});
                
                if (tumFotograflar.length > 0) {{
                    alert(`✅ Tüm fotoğraflar (${{tumFotograflar.length}} adet) direklerden çıkarıldı ve orijinal konumlarına geri eklendi.`);
                }} else {{
                    alert("⚠️ Hiç fotoğraf bağlı değil.");
                }}
            }}
            
            var osmLayer = L.tileLayer('https://{{s}}.tile.openstreetmap.org/{{z}}/{{x}}/{{y}}.png', {{
                maxZoom: 19,
                attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors'
            }});

            var darkLayer = L.tileLayer('https://{{s}}.basemaps.cartocdn.com/dark_all/{{z}}/{{x}}/{{y}}{{r}}.png', {{
                attribution: '&copy; <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a>, &copy; CartoDB',
                subdomains: 'abcd',
                maxZoom: 20
            }});

            var satelliteLayer = L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{{z}}/{{y}}/{{x}}', {{
                maxZoom: 18,
                attribution: 'Tiles © Esri'
            }});

            var map = L.map('map', {{
                center: [{center_lat}, {center_lon}],
                zoom: {zoom_level},
                layers: [osmLayer]
            }});

            var baseLayers = {{
                "OpenStreetMap (Klasik)": osmLayer,
                "Dark Matter (Koyu)": darkLayer,
                "Uydu Görüntüsü": satelliteLayer
            }};
            L.control.layers(baseLayers).addTo(map);
            
            var direklerCluster = L.markerClusterGroup({{
                maxClusterRadius: 40,
                spiderfyOnMaxZoom: true,
                showCoverageOnHover: false,
                zoomToBoundsOnClick: true,
                disableClusteringAtZoom: 18,
                chunkedLoading: true,
                chunkDelay: 100,
                singleMarkerMode: false,
                spiderfyDistanceMultiplier: 1.5,
                iconCreateFunction: function(cluster) {{
                    var childCount = cluster.getChildCount();
                    var c = ' marker-cluster-';
                    if (childCount < 10) {{
                        c += 'small';
                    }} else if (childCount < 50) {{
                        c += 'medium';
                    }} else {{
                        c += 'large';
                    }}
                    
                    return new L.DivIcon({{
                        html: '<div><span>' + childCount + '</span></div>',
                        className: 'marker-cluster' + c,
                        iconSize: new L.Point(40, 40)
                    }});
                }}
            }});
            
            map.getContainer().addEventListener("contextmenu", function(e) {{
                e.preventDefault();
            }});

            var TumFotolariCikarControl = L.Control.extend({{
                options: {{
                    position: 'topright'
                }},
                onAdd: function(map) {{
                    var container = L.DomUtil.create('div', 'leaflet-control-tum-fotolari-cikar leaflet-bar leaflet-control');
                    var link = L.DomUtil.create('a', 'leaflet-bar-part', container);
                    link.href = '#';
                    link.innerHTML = '🔄';
                    link.title = 'Tüm fotoğrafları dışarı al';
                    
                    L.DomEvent.on(link, 'click', function(e) {{
                        L.DomEvent.stopPropagation(e);
                        L.DomEvent.preventDefault(e);
                        tumFotolariCikar();
                    }});
                    
                    return container;
                }}
            }});
            
            var TumunuGosterControl = L.Control.extend({{
                options: {{
                    position: 'topright'
                }},
                onAdd: function(map) {{
                    var container = L.DomUtil.create('div', 'leaflet-control-tumunu-goster leaflet-bar leaflet-control');
                    var link = L.DomUtil.create('a', 'leaflet-bar-part', container);
                    link.href = '#';
                    link.innerHTML = '🗺️';
                    link.title = 'Harita konumuna dön';
                    
                    L.DomEvent.on(link, 'click', function(e) {{
                        L.DomEvent.stopPropagation(e);
                        L.DomEvent.preventDefault(e);
                        map.setView([{start_coords[0]}, {start_coords[1]}], {zoom_level});
                        console.log('📍 Harita konumuna dönüldü');
                    }});
                    
                    return container;
                }}
            }});
            
            // ========== BİRLEŞİK BUTON (Gizle/Göster + Numaraları Geçici Gizle) ==========
            var GizleGosterControl = L.Control.extend({{
                options: {{
                    position: 'topright'
                }},
                
                initialize: function(options) {{
                    L.Control.prototype.initialize.call(this, options);
                    this.is_hidden = false;
                    this.numbers_hidden = false;
                    this.pressTimer = null;
                    this.isLongPress = false;  // Uzun basış flag'i
                }},
                
                onAdd: function(map) {{
                    this._map = map;
                    var container = L.DomUtil.create('div', 'leaflet-control-gizle-goster leaflet-bar leaflet-control');
                    this.button = L.DomUtil.create('a', 'leaflet-bar-part', container);
                    this.button.href = '#';
                    this.button.innerHTML = '👁️';
                    this.button.title = 'Tıkla: Direkleri Gizle/Göster | Basılı Tut: Numaraları Geçici Gizle';
                    this.button.style.cursor = 'pointer';
                    this.button.style.fontSize = '18px';
                    this.button.style.lineHeight = '30px';
                    this.button.style.textAlign = 'center';
                    this.button.style.width = '30px';
                    this.button.style.height = '30px';
                    this.button.style.display = 'block';
                    
                    var self = this;
                    
                    // KISA TIKLAMA: SADECE uzun basış DEĞİLSE çalışsın
                    L.DomEvent.on(this.button, 'click', function(e) {{
                        L.DomEvent.stopPropagation(e);
                        L.DomEvent.preventDefault(e);
                        
                        // UZUN BASIŞ YAPILDIYSA, TIKLAMA İŞLEMİNİ ATLA
                        if (self.isLongPress) {{
                            console.log('👆 Uzun basış algılandı, tıklama işlemi atlandı');
                            self.isLongPress = false;  // Flag'i sıfırla
                            return;  // Tıklama işlemini YAPMA
                        }}
                        
                        if (self.pressTimer) {{
                            clearTimeout(self.pressTimer);
                            self.pressTimer = null;
                        }}
                        
                        // Normal tıklama işlemi (direkleri gizle/göster)
                        if (self.is_hidden) {{
                            self._show_direkler();
                            self.is_hidden = false;
                            self.button.innerHTML = '👁️';
                            console.log('👁️ Direkler gösteriliyor');
                        }} else {{
                            self._hide_direkler();
                            self.is_hidden = true;
                            self.button.innerHTML = '👁️‍🗨️';
                            console.log('👁️ Direkler gizlendi');
                        }}
                    }});
                    
                    // UZUN BASIŞ: Numaraları geçici gizle
                    L.DomEvent.on(this.button, 'mousedown', function(e) {{
                        L.DomEvent.stopPropagation(e);
                        L.DomEvent.preventDefault(e);
                        
                        // Uzun basış için zamanlayıcı (200ms)
                        self.pressTimer = setTimeout(function() {{
                            self.isLongPress = true;  // Uzun basış olduğunu işaretle
                            self._hideNumbers();
                            self.numbers_hidden = true;
                            self.button.innerHTML = '🔢';
                            self.button.title = 'Numaralar gizli - Bırakınca geri gelecek';
                            console.log('🔢 Numaralar geçici olarak gizlendi (basılı tutuluyor)');
                        }}, 200);
                    }});
                    
                    // Fare bırakılınca numaraları geri getir
                    L.DomEvent.on(this.button, 'mouseup', function(e) {{
                        L.DomEvent.stopPropagation(e);
                        L.DomEvent.preventDefault(e);
                        
                        if (self.pressTimer) {{
                            clearTimeout(self.pressTimer);
                            self.pressTimer = null;
                        }}
                        
                        if (self.numbers_hidden) {{
                            self._showNumbers();
                            self.numbers_hidden = false;
                            self.button.innerHTML = self.is_hidden ? '👁️‍🗨️' : '👁️';
                            self.button.title = 'Tıkla: Direkleri Gizle/Göster | Basılı Tut: Numaraları Geçici Gizle';
                            console.log('🔢 Numaralar geri getirildi');
                        }}
                        
                        // Uzun basış flag'ini temizle (click'ten sonra değil, hemen)
                        // setTimeout ile click'in çalışmasını bekle
                        setTimeout(function() {{
                            self.isLongPress = false;
                        }}, 50);
                    }});
                    
                    // Butondan ayrılırsa (güvenlik için)
                    L.DomEvent.on(this.button, 'mouseleave', function(e) {{
                        if (self.pressTimer) {{
                            clearTimeout(self.pressTimer);
                            self.pressTimer = null;
                        }}
                        
                        if (self.numbers_hidden) {{
                            self._showNumbers();
                            self.numbers_hidden = false;
                            self.button.innerHTML = self.is_hidden ? '👁️‍🗨️' : '👁️';
                            self.button.title = 'Tıkla: Direkleri Gizle/Göster | Basılı Tut: Numaraları Geçici Gizle';
                            console.log('🔢 Butondan ayrılındı, numaralar geri getirildi');
                        }}
                        
                        // Flag'i de temizle
                        self.isLongPress = false;
                    }});
                    
                    return container;
                }},
                
                _hide_direkler: function() {{
                    if (typeof direklerCluster !== 'undefined' && direklerCluster && this._map.hasLayer(direklerCluster)) {{
                        this._map.removeLayer(direklerCluster);
                    }}
                    
                    if (typeof direkler !== 'undefined' && direkler) {{
                        for (var i = 0; i < direkler.length; i++) {{
                            var d = direkler[i];
                            if (d.marker && this._map.hasLayer(d.marker)) {{
                                this._map.removeLayer(d.marker);
                            }}
                        }}
                    }}
                }},
                
                _show_direkler: function() {{
                    if (typeof direklerCluster !== 'undefined' && direklerCluster && !this._map.hasLayer(direklerCluster)) {{
                        this._map.addLayer(direklerCluster);
                    }}
                    
                    if (typeof direkler !== 'undefined' && direkler) {{
                        for (var i = 0; i < direkler.length; i++) {{
                            var d = direkler[i];
                            if (d.marker && !this._map.hasLayer(d.marker)) {{
                                if (direklerCluster && direklerCluster.hasLayer) {{
                                    direklerCluster.addLayer(d.marker);
                                }} else {{
                                    this._map.addLayer(d.marker);
                                }}
                            }}
                        }}
                    }}
                }},
                
                _hideNumbers: function() {{
                    if (typeof direkler !== 'undefined' && direkler) {{
                        for (var i = 0; i < direkler.length; i++) {{
                            var d = direkler[i];
                            if (d.numaraMarker && this._map && this._map.hasLayer(d.numaraMarker)) {{
                                this._map.removeLayer(d.numaraMarker);
                            }}
                        }}
                    }}
                }},
                
                _showNumbers: function() {{
                    if (typeof direkler !== 'undefined' && direkler) {{
                        for (var i = 0; i < direkler.length; i++) {{
                            var d = direkler[i];
                            if (d.numaraMarker && this._map && !this._map.hasLayer(d.numaraMarker) && 
                                d.photos && d.photos.length > 0 && d.numara) {{
                                this._map.addLayer(d.numaraMarker);
                            }}
                        }}
                    }}
                }}
            }});

            
            map.addControl(new TumFotolariCikarControl());
            map.addControl(new TumunuGosterControl());
            map.addControl(new GizleGosterControl());

            var direkler = [];
            var photoInfo = [];

            function guncelleDirekNumaraGorunumu(direk) {{
                if (direk.numaraMarker && map.hasLayer(direk.numaraMarker)) {{
                    map.removeLayer(direk.numaraMarker);
                    direk.numaraMarker = null;
                }}
                
                if (direk.photos && direk.photos.length > 0 && direk.numara) {{
                    var latlng = direk.marker.getLatLng();
                    var numaraIcon = L.divIcon({{
                        className: 'direk-numara-container',
                        html: '<div class="direk-numara-text">' + direk.numara + '</div>',
                        iconSize: [24, 24],
                        iconAnchor: [12, 12]
                    }});
                    
                    var labelMarker = L.marker(latlng, {{
                        icon: numaraIcon,
                        interactive: false,
                        zIndexOffset: 1000
                    }}).addTo(map);
                    
                    direk.numaraMarker = labelMarker;
                }}
            }}

            function guncelleDirekRengi(direk) {{
                if (!direk.marker) return;
                
                var markerClass = '';
                if (direk.nihai_direk) {{
                    markerClass = 'direk-marker-nihai';
                }} else if (direk.yeni_direk) {{
                    markerClass = 'direk-marker-yeni';
                }} else if (direk.musterek) {{
                    markerClass = 'direk-marker-musterek';
                }} else {{
                    markerClass = 'direk-marker-normal';
                }}
                
                if (direk.photos && direk.photos.length > 0 && !direk.yeni_direk && !direk.nihai_direk) {{
                    markerClass = 'direk-marker-foto-var';
                }}
                
                if (direk.koordinat_degisecek && !direk.nihai_direk) {{
                    markerClass = 'direk-marker-konum-degisti';
                }}
                
                if (direklerCluster.hasLayer(direk.marker)) {{
                    direklerCluster.removeLayer(direk.marker);
                }}
                
                direk.marker.setIcon(L.divIcon({{
                    className: markerClass,
                    iconSize: [14, 14],
                    iconAnchor: [7, 7]
                }}));
                
                direklerCluster.addLayer(direk.marker);
                
                guncelleDirekNumaraGorunumu(direk);
            }}

            function diregiMarkerModunaGecir(direkId) {{
                const direk = direkler.find(d => d.id === direkId);
                if (!direk) return;
                
                if (direk.marker && direklerCluster.hasLayer(direk.marker)) {{
                    direklerCluster.removeLayer(direk.marker);
                }}
                
                var markerClass = '';
                if (direk.nihai_direk) {{
                    markerClass = 'direk-marker-nihai';
                }} else if (direk.yeni_direk) {{
                    markerClass = 'direk-marker-yeni';
                }} else if (direk.musterek) {{
                    markerClass = 'direk-marker-musterek';
                }} else {{
                    markerClass = 'direk-marker-normal';
                }}
                
                if (direk.photos && direk.photos.length > 0 && !direk.yeni_direk && !direk.nihai_direk) {{
                    markerClass = 'direk-marker-foto-var';
                }}
                
                if (direk.koordinat_degisecek && !direk.nihai_direk) {{
                    markerClass = 'direk-marker-konum-degisti';
                }}
                
                var marker = L.marker([direk.lat, direk.lon], {{
                    draggable: true,
                    icon: L.divIcon({{
                        className: markerClass,
                        iconSize: [14, 14],
                        iconAnchor: [7, 7]
                    }}),
                    opacity: 0.9,
                    zIndexOffset: 700
                }});
                
                marker.on("dragend", function(e) {{
                    var pos = e.target.getLatLng();
                    direk.lat = pos.lat;
                    direk.lon = pos.lng;
                    
                    direk.koordinat_degisecek = true;
                    console.log(`📍 Direk ${{direk.id}} yeni konumu: ${{direk.lat}}, ${{direk.lon}}`);
                    
                    if (direk.numaraMarker) {{
                        direk.numaraMarker.setLatLng([pos.lat, pos.lng]);
                    }}
                    
                    guncelleDirekRengi(direk);
                    guncellePopup(direk, true);
                    
                    var popup = L.popup()
                        .setLatLng(pos)
                        .setContent(`✅ ${{direk.id}} direğinin konumu güncellendi.<br>Yeni koordinatlar: ${{direk.lat.toFixed(6)}}, ${{direk.lon.toFixed(6)}}`)
                        .openOn(map);
                        
                    setTimeout(function() {{
                        map.closePopup(popup);
                    }}, 1000);
                }});
                
                marker.addTo(map);
                direk.marker = marker;
                direk.draggable = true;
                
                guncellePopup(direk, true);
                
                var popup = L.popup()
                    .setLatLng([direk.lat, direk.lon])
                    .setContent(`✅ ${{direk.id}} direği artık sürüklenebilir.<br>Konumu değiştirmek için taşıyın.`)
                    .openOn(map);
                    
                setTimeout(function() {{
                    map.closePopup(popup);
                }}, 1000);
            }}

            function guncellePopup(d, popupAc = false) {{
                let html = `<b>Direk No: ${{d.id}}</b><br><b>Tip:</b> ${{d.tip}}`;
                
                if (d.envanter_notu) {{
                    html += `<br><b>Not:</b> ${{d.envanter_notu}}`;
                }}
                
                html += `<br><small>`;
                if (d.musterek) {{
                    html += `🟣 <span style="color:#9370DB;">Müşterek Direk</span>`;
                }} else {{
                    html += `🔴 <span style="color:#ff0000;">Normal Direk</span>`;
                }}
                
                if (d.photos && d.photos.length > 0) {{
                    html += ` | 🟢 Fotoğraf Var`;
                }}
                if (d.koordinat_degisecek) {{
                    html += ` | 🟡 Konum Değişti`;
                }}
                if (d.yeni_direk) {{
                    html += ` | 🟣 Yeni Direk`;
                }}
                if (d.nihai_direk) {{
                    html += ` | 🔵 Nihai Direk`;
                }}
                html += `</small>`;

                if (d.koordinat_degisecek) {{
                    html += `<br><span style="color:#ff6600; font-weight:bold;">📍 Koordinatı değiştirildi</span>`;
                }}
                
                if (d.nihai_direk) {{
                    html += `<br><span style="color:#007BFF; font-weight:bold;">🔵 Nihai Direk olarak işaretlendi</span>`;
                }}

                if (typeof d.numara !== 'undefined' && d.numara !== null) {{
                    html += `<br><b>📌 Numara:</b> 
                            <input type="number" value="${{d.numara}}" id="numara-input-${{d.id}}" style="width:60px;"> 
                            <button onclick="kaydetNumara('${{d.id}}')" style="margin-left:5px; padding:4px 8px;">💾</button>
                            <button onclick="direktenTumFotolariCikar('${{d.id}}')" style="margin-left:5px; padding:4px 8px;" title="Tüm fotoğrafları dışarı al">🔄</button>`;
                }}

                if (d.photos && d.photos.length > 0) {{
                    html += "<br><b>Fotoğraflar:</b><br>";
                    d.photos.forEach(p => {{
                        const filename = p.path.split('/').pop().split('\\\\').pop();
                        const folder = p.path.split('/').slice(-2, -1)[0] || "";
                        const displayName = folder ? folder + "/" + filename : filename;
                        html += `
                            <div style='margin-bottom:6px;'>
                                <img src='${{p.path}}' width='80' style='margin:2px; border-radius:4px; cursor:pointer;' onclick="geriCagir('${{p.path}}')"><br>
                                <small>📝 Not: ${{p.note || '-'}}</small><br>
                                <small>📁 ${{displayName}}</small>
                            </div>
                        `;
                    }});
                }}

                if (d.yeni_direk && (!d.photos || d.photos.length === 0)) {{
                    html += `<br><button onclick="kaldirDirek('${{d.id}}')" 
                            style="background-color: #ff4444; 
                                   color: white; 
                                   border: none; 
                                   padding: 5px 10px; 
                                   border-radius: 3px; 
                                   cursor: pointer;
                                   margin-top: 5px;">
                        🗑️ Direği Kaldır
                    </button>`;
                }}

                if (!d.yeni_direk) {{
                    html += `<br>
                    <div style="margin-top: 8px;">
                        <label>
                            <input type="checkbox" id="silinecek-checkbox-${{d.id}}" ${{d.silinecek_mi && 'checked'}} onchange="kaydetSilinecek('${{d.id}}')"> 
                            Direk Silinecek
                        </label>
                        <br>
                        <label>
                            <input type="checkbox" id="nihai-checkbox-${{d.id}}" ${{d.nihai_direk && 'checked'}} onchange="kaydetNihai('${{d.id}}')"> 
                            🔵 Nihai Direk
                        </label>
                    </div>`;
                }}
                
                html += `<br><br><button class="direk-ikon-button" onclick="diregiMarkerModunaGecir('${{d.id}}')" title="Direğin konumunu değiştirmek için tıklayın">📍 Konumu Değiştir</button>`;

                if (!d.marker.getPopup()) {{
                    d.marker.bindPopup(html);
                }} else {{
                    d.marker.setPopupContent(html);
                }}

                if (popupAc || d.marker.isPopupOpen()) {{
                    d.marker.openPopup();
                }}
            }}
            
            function kaydetNihai(id) {{
                const d = direkler.find(x => x.id === id);
                if (!d) return;

                const checkbox = document.getElementById(`nihai-checkbox-${{id}}`);
                if (!checkbox) return;

                d.nihai_direk = checkbox.checked;
                guncellePopup(d, true);
                console.log(`🔵 '${{id}}' nihai direk mi? →`, d.nihai_direk);
                
                guncelleDirekRengi(d);
            }}

            function direktenTumFotolariCikar(direkId) {{
                const direk = direkler.find(d => d.id === direkId);
                if (!direk || !direk.photos || direk.photos.length === 0) return;
                
                const fotograflar = [...direk.photos];
                
                fotograflar.forEach(foto => {{
                    const photoInfoObj = photoInfo.find(p => p.path === foto.path);
                    
                    if (photoInfoObj) {{
                        const lat = photoInfoObj.originalLat !== undefined ? photoInfoObj.originalLat : photoInfoObj.lat;
                        const lon = photoInfoObj.originalLon !== undefined ? photoInfoObj.originalLon : photoInfoObj.lon;
                        
                        var markerClass = "photo-marker " + (photoInfoObj.note ? "photo-turkuaz" : "photo-mavi");

                        var marker = L.marker([lat, lon], {{
                            draggable: true,
                            icon: L.divIcon({{
                                className: markerClass,
                                iconSize: [13, 13]
                            }}),
                            zIndexOffset: 800
                        }}).bindPopup(getPhotoPopupHTML(photoInfoObj.path, photoInfoObj.note), {{
                            maxWidth: 420,
                            minWidth: 400,
                            autoPan: true,
                            className: 'custom-popup'
                        }});

                        marker.on("contextmenu", function(e) {{
                            e.originalEvent.preventDefault();
                            e.originalEvent.stopPropagation();
                            
                            var pos = e.target.getLatLng();
                            var hedefDirek = null;
                            var enKisaMesafe = Infinity;

                            direkler.forEach(function(d) {{
                                var mesafe = map.distance(pos, L.latLng(d.lat, d.lon));
                                if (mesafe < 40 && mesafe < enKisaMesafe) {{
                                    hedefDirek = d;
                                    enKisaMesafe = mesafe;
                                }}
                            }});

                            if (hedefDirek) {{
                                var zatenVar = hedefDirek.photos.some(p => p.path === photoInfoObj.path);
                                if (zatenVar) {{
                                    L.popup()
                                        .setLatLng(pos)
                                        .setContent("⛔ Aynı fotoğraf zaten bu direğe eklenmiş")
                                        .openOn(map);
                                    return;
                                }}

                                const notDegeri = photoInfoObj.note || "";
                                const kategori = photoInfoObj.kategori || "";
                                const oncelik = photoInfoObj.oncelik || "Bekleyebilir";
                                const birim = photoInfoObj.birim || "Aob";

                                hedefDirek.photos.push({{
                                    path: photoInfoObj.path,
                                    note: notDegeri,
                                    kategori: kategori,
                                    oncelik: oncelik,
                                    birim: birim
                                }});

                                if (hedefDirek.photos.length === 1 && !hedefDirek.numara) {{
                                    const kullanılanNumaralar = direkler
                                        .filter(d => typeof d.numara === 'number' && d.numara !== null)
                                        .map(d => d.numara);
                                    let yeniNumara = 1;
                                    while (kullanılanNumaralar.includes(yeniNumara)) {{
                                        yeniNumara++;
                                    }}
                                    hedefDirek.numara = yeniNumara;
                                    guncelleDirekNumaraGorunumu(hedefDirek);
                                }}

                                guncelleDirekRengi(hedefDirek);
                                guncellePopup(hedefDirek);
                                map.removeLayer(marker);
                            }} else {{
                                L.popup()
                                    .setLatLng(pos)
                                    .setContent("⚠️ 40 metre içinde direk bulunamadı")
                                    .openOn(map);
                            }}
                        }});

                        marker.on("popupopen", function() {{
                            bindNoteInput(photoInfoObj.path);
                        }});

                        marker.on("dragend", function(e) {{
                            var pos = e.target.getLatLng();
                            var hedefDirek = null;
                            var enKisaMesafe = Infinity;

                            direkler.forEach(function(d) {{
                                var mesafe = map.distance(pos, L.latLng(d.lat, d.lon));
                                if (mesafe < 8 && !d.photos.find(p => p.path === photoInfoObj.path) && mesafe < enKisaMesafe) {{
                                    hedefDirek = d;
                                    enKisaMesafe = mesafe;
                                }}
                            }});

                            if (hedefDirek) {{
                                const notDegeri = photoInfoObj.note || "";
                                hedefDirek.photos.push({{ path: photoInfoObj.path, note: notDegeri }});

                                if (hedefDirek.photos.length === 1 && !hedefDirek.numara) {{
                                    const kullanılanNumaralar = direkler
                                        .filter(d => typeof d.numara === 'number' && d.numara !== null)
                                        .map(d => d.numara);
                                    let yeniNumara = 1;
                                    while (kullanılanNumaralar.includes(yeniNumara)) {{
                                        yeniNumara++;
                                    }}
                                    hedefDirek.numara = yeniNumara;
                                    guncelleDirekNumaraGorunumu(hedefDirek);
                                }}

                                guncelleDirekRengi(hedefDirek);
                                guncellePopup(hedefDirek);
                                map.removeLayer(marker);
                            }}
                        }});

                        photoInfoObj.marker = marker;
                        map.addLayer(marker);
                    }}
                }});
                
                direk.photos = [];
                
                if (direk.numara) {{
                    direk.numara = null;
                    if (direk.numaraMarker && map.hasLayer(direk.numaraMarker)) {{
                        map.removeLayer(direk.numaraMarker);
                        direk.numaraMarker = null;
                    }}
                }}
                
                guncelleDirekRengi(direk);
                guncellePopup(direk, true);
                
                alert(`✅ ${{fotograflar.length}} fotoğraf "${{direkId}}" direğinden çıkarıldı.`);
            }}

            function guncellePopuplar() {{
                direkler.forEach(d => {{
                    guncellePopup(d);
                }});
            }}

            function addPhotoToDirek(direkId, fullPath) {{
                console.log("addPhotoToDirek çağrıldı:", direkId, fullPath);
                const direk = direkler.find(d => d.id === direkId);
                if (direk) {{
                    if (!direk.photos) {{
                        direk.photos = [];
                    }}

                    const zatenEkli = direk.photos.some(p => p.path === fullPath);
                    if (zatenEkli) {{
                        console.log(`⛔ Fotoğraf zaten ekli: ${{fullPath}}`);
                        return;
                    }}

                    const foto = photoInfo.find(p => p.path === fullPath);
                    const notDegeri = foto && foto.note ? foto.note : "";

                    direk.photos.push({{ path: fullPath, note: notDegeri }});
                    
                    if (direk.photos.length === 1 && (typeof direk.numara === 'undefined' || direk.numara === null)) {{
                        const kullanılanNumaralar = direkler
                            .filter(d => typeof d.numara === 'number' && d.numara !== null)
                            .map(d => d.numara);
                        let yeniNumara = 1;
                        while (kullanılanNumaralar.includes(yeniNumara)) {{
                            yeniNumara++;
                        }}
                        direk.numara = yeniNumara;
                    }}
                    
                    guncelleDirekRengi(direk);
                    guncellePopup(direk);
                    
                    if (direk.photos.length === 1) {{
                        guncelleDirekNumaraGorunumu(direk);
                    }}
                    
                    console.log(`✅ Fotoğraf eklendi: ${{fullPath}} -> ${{direkId}}`);
                }} else {{
                    console.error(`❌ Direk bulunamadı: ${{direkId}}`);
                }}
            }}
            
            function kaydetNumara(id) {{
                const d = direkler.find(x => x.id === id);
                if (!d) return;

                const input = document.getElementById(`numara-input-${{id}}`);
                if (!input) return;

                const yeniNumara = parseInt(input.value);
                if (isNaN(yeniNumara)) return;

                const baskasi = direkler.find(x => x.numara === yeniNumara && x.id !== id);
                if (baskasi) {{
                    alert(`❌ Bu numara zaten '${{baskasi.id}}' direğinde kullanılıyor.`);
                    input.value = d.numara ?? "";
                    return;
                }}

                d.numara = yeniNumara;
                console.log(`🔢 '${{id}}' numarası güncellendi → ${{yeniNumara}}`);
                guncellePopup(d, true);
                guncelleDirekNumaraGorunumu(d);
            }}

            function kaydetSilinecek(id) {{
                const d = direkler.find(x => x.id === id);
                if (!d) return;

                const checkbox = document.getElementById(`silinecek-checkbox-${{id}}`);
                if (!checkbox) return;

                d.silinecek_mi = checkbox.checked;
                guncellePopup(d, true);
                console.log(`🗑️ '${{id}}' silinecek mi? →`, d.silinecek_mi);
            }}

            function geriCagir(fullPath) {{
                const foto = photoInfo.find(p => p.path === fullPath);
                if (!foto) {{
                    console.error("Fotoğraf bulunamadı:", fullPath);
                    return;
                }}
                
                direkler.forEach(d => {{
                    const index = d.photos.findIndex(p => p.path === fullPath);
                    if (index > -1) {{
                        const wasOpen = d.marker.isPopupOpen?.() ?? false;
                        d.photos.splice(index, 1);

                        if (d.photos.length === 0) {{
                            console.log(`♻️ '${{d.id}}' boş kaldı, numara ${{d.numara}} boşa çıktı.`);
                            d.numara = null;
                            if (d.numaraMarker && map.hasLayer(d.numaraMarker)) {{
                                map.removeLayer(d.numaraMarker);
                                d.numaraMarker = null;
                            }}
                        }}

                        guncelleDirekRengi(d);
                        guncellePopup(d);
                        if (wasOpen) {{
                            setTimeout(() => d.marker.openPopup(), 0);
                        }}
                    }}
                }});

                var lat = foto.originalLat !== undefined ? foto.originalLat : foto.lat;
                var lon = foto.originalLon !== undefined ? foto.originalLon : foto.lon;

                var markerClass = "photo-marker " + (foto.note ? "photo-turkuaz" : "photo-mavi");

                var marker = L.marker([lat, lon], {{
                    draggable: true,
                    icon: L.divIcon({{
                        className: markerClass,
                        iconSize: [13, 13]
                    }}),
                    zIndexOffset: 800
                }}).bindPopup(getPhotoPopupHTML(foto.path, foto.note, foto.birim), {{
                    maxWidth: 420,
                    minWidth: 400,
                    autoPan: true,
                    className: 'custom-popup'
                }});

                marker.on("contextmenu", function(e) {{
                    e.originalEvent.preventDefault();
                    e.originalEvent.stopPropagation();
                    
                    var pos = e.target.getLatLng();
                    var hedefDirek = null;
                    var enKisaMesafe = Infinity;

                    direkler.forEach(function(d) {{
                        var mesafe = map.distance(pos, L.latLng(d.lat, d.lon));
                        if (mesafe < 40 && mesafe < enKisaMesafe) {{
                            hedefDirek = d;
                            enKisaMesafe = mesafe;
                        }}
                    }});

                    if (hedefDirek) {{
                        var zatenVar = hedefDirek.photos.some(p => p.path === fullPath);
                        if (zatenVar) {{
                            console.log("⛔ Aynı fotoğraf zaten bu direğe eklenmiş.");
                            L.popup()
                                .setLatLng(pos)
                                .setContent("⛔ Aynı fotoğraf zaten bu direğe eklenmiş")
                                .openOn(map);
                            return;
                        }}

                        const notDegeri = foto.note || "";
                        const kategori = foto.kategori || "";
                        const oncelik = foto.oncelik || "Bekleyebilir";
                        const birim = foto.birim || "Aob";

                        hedefDirek.photos.push({{
                            path: fullPath,
                            note: notDegeri,
                            kategori: kategori,
                            oncelik: oncelik,
                            birim: birim
                        }});

                        if (hedefDirek.photos.length === 1 && !hedefDirek.numara) {{
                            const kullanılanNumaralar = direkler
                                .filter(d => typeof d.numara === 'number' && d.numara !== null)
                                .map(d => d.numara);
                            let yeniNumara = 1;
                            while (kullanılanNumaralar.includes(yeniNumara)) {{
                                yeniNumara++;
                            }}
                            hedefDirek.numara = yeniNumara;
                            guncelleDirekNumaraGorunumu(hedefDirek);
                        }}

                        guncelleDirekRengi(hedefDirek);
                        guncellePopup(hedefDirek);
                        map.removeLayer(marker);
                    }} else {{
                        L.popup()
                            .setLatLng(pos)
                            .setContent("⚠️ 40 metre içinde direk bulunamadı")
                            .openOn(map);
                    }}
                }});

                marker.on("popupopen", function() {{
                    bindNoteInput(foto.path);
                }});

                marker.on("dragend", function(e) {{
                    var pos = e.target.getLatLng();
                    var hedefDirek = null;
                    var enKisaMesafe = Infinity;

                    direkler.forEach(function(d) {{
                        var mesafe = map.distance(pos, L.latLng(d.lat, d.lon));
                        if (mesafe < 8 && !d.photos.find(p => p.path === fullPath) && mesafe < enKisaMesafe) {{
                            hedefDirek = d;
                            enKisaMesafe = mesafe;
                        }}
                    }});

                    if (hedefDirek) {{
                        var note = foto.note || "";

                        if (foto.originalLat === undefined || foto.originalLon === undefined) {{
                            foto.originalLat = foto.lat;
                            foto.originalLon = foto.lon;
                        }}

                        foto.lat = pos.lat;
                        foto.lon = pos.lng;

                        hedefDirek.photos.push({{ path: fullPath, note: note }});

                        if (hedefDirek.photos.length === 1 && !hedefDirek.numara) {{
                            const kullanılanNumaralar = direkler
                                .filter(d => typeof d.numara === 'number' && d.numara !== null)
                                .map(d => d.numara);
                            let yeniNumara = 1;
                            while (kullanılanNumaralar.includes(yeniNumara)) {{
                                yeniNumara++;
                            }}
                            hedefDirek.numara = yeniNumara;
                            guncelleDirekNumaraGorunumu(hedefDirek);
                        }}

                        guncelleDirekRengi(hedefDirek);
                        guncellePopup(hedefDirek);
                        map.removeLayer(marker);
                    }}
                }});

                foto.marker = marker;
                map.addLayer(marker);
                console.log("📍 Marker EXIF konumuna geri döndü:", foto.path);
            }}
    """
    
    photo_markers = ""
    if photo_data:
        for path, lat, lon, tarih, photo_id in photo_data:
            abs_path = path.replace("\\", "/")
            filename = os.path.basename(path)
            folder = os.path.basename(os.path.dirname(path))
            
            photo_markers += f"""
                (function() {{
                    var fullPath = "{abs_path}";
                    var photoId = "{photo_id}";
                    var filename = "{filename}";
                    var folder = "{folder}";
                    
                    var marker = L.marker([{lat}, {lon}], {{
                        draggable: true,
                        icon: L.divIcon({{
                            className: 'photo-marker photo-mavi',
                            iconSize: [13, 13]
                        }}),
                        zIndexOffset: 800
                    }});

                    marker.bindPopup(getPhotoPopupHTML(fullPath, "", ""));

                    marker.on("dragend", function(e) {{
                        var pos = e.target.getLatLng();
                        var foto = photoInfo.find(p => p.marker === marker);
                        if (foto) {{
                            foto.lat = pos.lat;
                            foto.lon = pos.lng;
                        }}
                    }});

                    marker.on("contextmenu", function(e) {{
                        e.originalEvent.preventDefault();
                        e.originalEvent.stopPropagation();
                        
                        var pos = e.target.getLatLng();
                        var hedefDirek = null;
                        var enKisaMesafe = Infinity;

                        direkler.forEach(function(d) {{
                            var mesafe = map.distance(pos, L.latLng(d.lat, d.lon));
                            if (mesafe < 40 && mesafe < enKisaMesafe) {{
                                hedefDirek = d;
                                enKisaMesafe = mesafe;
                            }}
                        }});

                        if (hedefDirek) {{
                            var zatenVar = hedefDirek.photos.some(p => p.path === fullPath);
                            if (zatenVar) {{
                                console.log("⛔ Aynı fotoğraf zaten bu direğe eklenmiş.");
                                L.popup()
                                    .setLatLng(pos)
                                    .setContent("⛔ Aynı fotoğraf zaten bu direğe eklenmiş")
                                    .openOn(map);
                                return;
                            }}

                            const foto = photoInfo.find(p => p.path === fullPath);
                            const notDegeri = foto?.note || "";
                            const kategori = foto?.kategori || "";
                            const oncelik = foto?.oncelik || "Bekleyebilir";
                            const birim = foto?.birim || "Aob";

                            hedefDirek.photos.push({{
                                path: fullPath,
                                note: notDegeri,
                                kategori: kategori,
                                oncelik: oncelik,
                                birim: birim
                            }});

                            if (hedefDirek.photos.length === 1 && !hedefDirek.numara) {{
                                const kullanılanNumaralar = direkler
                                    .filter(d => typeof d.numara === 'number' && d.numara !== null)
                                    .map(d => d.numara);
                                let yeniNumara = 1;
                                while (kullanılanNumaralar.includes(yeniNumara)) {{
                                    yeniNumara++;
                                }}
                                hedefDirek.numara = yeniNumara;
                                console.log(`Yeni numara atandı: ${{yeniNumara}}`);
                                guncelleDirekNumaraGorunumu(hedefDirek);
                            }}

                            guncelleDirekRengi(hedefDirek);
                            guncellePopup(hedefDirek);
                            map.removeLayer(marker);
                        }} else {{
                            L.popup()
                                .setLatLng(pos)
                                .setContent("⚠️ 40 metre içinde direk bulunamadı")
                                .openOn(map);
                        }}
                    }});

                    marker.on("popupopen", function() {{
                        const foto = photoInfo.find(p => p.path === fullPath);
                        const guncelNote = foto?.note || "";
                        const guncelBirim = foto?.birim || "";
                        
                        marker.setPopupContent(getPhotoPopupHTML(fullPath, guncelNote, guncelBirim));
                        bindNoteInput(fullPath);
                    }});

                    marker.on("dragend", function(e) {{
                        var pos = e.target.getLatLng();
                        var hedefDirek = null;
                        var enKisaMesafe = Infinity;

                        direkler.forEach(function(d) {{
                            var mesafe = map.distance(pos, L.latLng(d.lat, d.lon));
                            if (mesafe < 8 && mesafe < enKisaMesafe) {{
                                hedefDirek = d;
                                enKisaMesafe = mesafe;
                            }}
                        }});

                        if (hedefDirek) {{
                            var zatenVar = hedefDirek.photos.some(p => p.path === fullPath);
                            if (zatenVar) {{
                                console.log("⛔ Aynı fotoğraf zaten bu direğe eklenmiş.");
                                return;
                            }}

                            const foto = photoInfo.find(p => p.path === fullPath);
                            const notDegeri = foto?.note || "";

                            hedefDirek.photos.push({{ path: fullPath, note: notDegeri }});

                            if (hedefDirek.photos.length === 1 && !hedefDirek.numara) {{
                                const kullanılanNumaralar = direkler
                                    .filter(d => typeof d.numara === 'number' && d.numara !== null)
                                    .map(d => d.numara);
                                let yeniNumara = 1;
                                while (kullanılanNumaralar.includes(yeniNumara)) {{
                                    yeniNumara++;
                                }}
                                hedefDirek.numara = yeniNumara;
                                console.log(`Yeni numara atandı: ${{yeniNumara}}`);
                                guncelleDirekNumaraGorunumu(hedefDirek);
                            }}

                            guncelleDirekRengi(hedefDirek);
                            guncellePopup(hedefDirek);
                            map.removeLayer(marker);
                        }}
                    }});

                    photoInfo.push({{
                        id: photoId,
                        path: fullPath,
                        filename: filename,
                        folder: folder,
                        lat: marker.getLatLng().lat,
                        lon: marker.getLatLng().lng,
                        originalLat: {lat},
                        originalLon: {lon},
                        note: "",
                        birim: "Aob",
                        marker: marker,
                        edited: false,
                        tarih: "{tarih or ''}",
                        kategori: "",
                        oncelik: "Bekleyebilir"
                    }});

                    map.addLayer(marker);
                }})();
            """
    
    direk_markers = """
            var direk_data = [
    """
    
    for d in direk_data:
        direk_markers += f"""
                {{
                    id: "{d['id']}",
                    tip: "{d['tip']}",
                    lat: {d['lat']},
                    lon: {d['lon']},
                    musterek: {str(d.get('musterek', False)).lower()}
                }},
        """
    
    direk_markers += """
            ];

            direk_data.forEach(function(d) {
                var markerClass = d.musterek ? 'direk-marker-musterek' : 'direk-marker-normal';

                var marker = L.marker([d.lat, d.lon], {
                    icon: L.divIcon({
                        className: markerClass,
                        iconSize: [14, 14],
                        iconAnchor: [7, 7]
                    }),
                    opacity: 0.9,
                    zIndexOffset: 700
                });

                var direkObj = {
                    id: d.id,
                    tip: d.tip,
                    lat: d.lat,
                    lon: d.lon,
                    marker: marker,
                    photos: [],
                    musterek: d.musterek,
                    numara: null,
                    numaraMarker: null,
                    yeni_direk: false,
                    draggable: false,
                    koordinat_degisecek: false,
                    nihai_direk: false
                };

                direkler.push(direkObj);
                
                guncellePopup(direkObj);
                
                direklerCluster.addLayer(marker);
            });

            map.addLayer(direklerCluster);
            
            // ========== DIREK ARAMA KONTROLÜ ==========
            var DirekAramaControl = L.Control.extend({
                options: {
                    position: 'topright'
                },
                
                onAdd: function(map) {
                    this._map = map;
                    var container = L.DomUtil.create('div', 'leaflet-control-search leaflet-bar leaflet-control');
                    container.style.backgroundColor = 'white';
                    container.style.padding = '8px';
                    container.style.borderRadius = '6px';
                    container.style.border = '2px solid rgba(0,0,0,0.2)';
                    container.style.boxShadow = '0 2px 6px rgba(0,0,0,0.3)';
                    
                    this.input = L.DomUtil.create('input', '', container);
                    this.input.type = 'text';
                    this.input.placeholder = '🔍 Direk No Ara...';
                    this.input.style.width = '180px';
                    this.input.style.padding = '6px 8px';
                    this.input.style.fontSize = '14px';
                    this.input.style.border = '1px solid #ccc';
                    this.input.style.borderRadius = '4px';
                    this.input.style.outline = 'none';
                    
                    this.button = L.DomUtil.create('button', '', container);
                    this.button.innerHTML = '🔍 Ara';
                    this.button.style.marginLeft = '8px';
                    this.button.style.padding = '6px 12px';
                    this.button.style.cursor = 'pointer';
                    this.button.style.backgroundColor = '#007bff';
                    this.button.style.color = 'white';
                    this.button.style.border = 'none';
                    this.button.style.borderRadius = '4px';
                    this.button.style.fontSize = '14px';
                    this.button.style.fontWeight = 'bold';
                    this.button.onmouseover = function() { this.style.backgroundColor = '#0056b3'; };
                    this.button.onmouseout = function() { this.style.backgroundColor = '#007bff'; };
                    
                    var self = this;
                    L.DomEvent.on(this.button, 'click', function(e) {
                        L.DomEvent.stopPropagation(e);
                        L.DomEvent.preventDefault(e);
                        self._search();
                    });
                    
                    L.DomEvent.on(this.input, 'keypress', function(e) {
                        if (e.keyCode === 13) {
                            L.DomEvent.stopPropagation(e);
                            L.DomEvent.preventDefault(e);
                            self._search();
                        }
                    });
                    
                    return container;
                },
                
                _search: function() {
                    var searchText = this.input.value.trim();
                    if (!searchText) {
                        alert('🔍 Lütfen bir direk numarası girin!\\n\\nÖrnek: 12345 veya Direk 123');
                        return;
                    }
                    
                    var lowerSearch = searchText.toLowerCase();
                    var bulunanDirek = null;
                    var bulunanlar = [];
                    
                    for (var i = 0; i < direkler.length; i++) {
                        var direkNo = String(direkler[i].id).toLowerCase();
                        if (direkNo === lowerSearch) {
                            bulunanDirek = direkler[i];
                            break;
                        }
                        if (direkNo.includes(lowerSearch)) {
                            bulunanlar.push(direkler[i]);
                        }
                    }
                    
                    if (!bulunanDirek && bulunanlar.length > 0) {
                        bulunanDirek = bulunanlar[0];
                        if (bulunanlar.length > 1) {
                            console.log(`🔍 ${bulunanlar.length} adet direk bulundu, ilki seçildi: ${bulunanlar.map(d => d.id).join(', ')}`);
                        }
                    }
                    
                    if (bulunanDirek) {
                        this._map.setView([bulunanDirek.lat, bulunanDirek.lon], 18);
                        
                        if (bulunanDirek.marker) {
                            bulunanDirek.marker.openPopup();
                            if (typeof guncellePopup !== 'undefined') {
                                guncellePopup(bulunanDirek, true);
                            }
                            
                            var originalIcon = bulunanDirek.marker.getIcon();
                            var highlightIcon = L.divIcon({
                                className: originalIcon.options.className + ' direk-marker-highlight',
                                iconSize: [20, 20],
                                iconAnchor: [10, 10],
                                html: '<div style="width:20px;height:20px;background-color:#ff4444;border-radius:50%;border:3px solid #fff;box-shadow:0 0 10px #ff0000;"></div>'
                            });
                            bulunanDirek.marker.setIcon(highlightIcon);
                            setTimeout(function() {
                                if (bulunanDirek.marker) {
                                    bulunanDirek.marker.setIcon(originalIcon);
                                }
                            }, 2000);
                        }
                        
                        console.log(`✅ Direk bulundu: ${bulunanDirek.id} (${bulunanDirek.lat}, ${bulunanDirek.lon})`);
                        
                        var notification = L.popup()
                            .setLatLng([bulunanDirek.lat, bulunanDirek.lon])
                            .setContent(`
                                <div style="text-align:center;">
                                    <b style="color:#007bff;">🔍 Direk Bulundu!</b><br>
                                    <b>${bulunanDirek.id}</b><br>
                                    <small>Tip: ${bulunanDirek.tip || 'Normal'}</small>
                                </div>
                            `)
                            .openOn(this._map);
                        
                        setTimeout(function() {
                            if (this._map) this._map.closePopup(notification);
                        }.bind(this), 3000);
                        
                    } else {
                        alert(`❌ "${searchText}" numaralı direk bulunamadı!\\n\\n📋 Mevcut direkler: ${direkler.slice(0, 5).map(d => d.id).join(', ')}${direkler.length > 5 ? '...' : ''}`);
                    }
                }
            });
            
            map.addControl(new DirekAramaControl());
            console.log('✅ Direk arama kontrolü eklendi, toplam direk: ' + direkler.length);
    """
    
    html_end = """
        </script>
    </body>
    </html>
    """
    
    full_html = html_start + photo_markers + direk_markers + html_end

    with open(save_path, "w", encoding="utf-8") as f:
        f.write(full_html)

    print(f"✅ Harita kaydedildi: {save_path}")


def uri_to_path(uri):
    if uri.startswith("file:///"):
        parsed = urlparse(uri)
        return unquote(parsed.path.lstrip("/"))
    return uri
    
def oku_direkler(excel_yolu):
    direk_listesi = []
    try:
        wb = load_workbook(excel_yolu, data_only=True)
        ws = wb.active

        for row in ws.iter_rows(min_row=2):
            direk_id = str(row[1].value).strip()
            tip = str(row[2].value).strip()
            lat = row[5].value
            lon = row[6].value

            if lat is None or lon is None:
                continue
            
            lat_str = str(lat).strip()
            lon_str = str(lon).strip()
            
            if lat_str == "" or lon_str == "":
                continue
            
            lat_str = lat_str.replace(',', '.')
            lon_str = lon_str.replace(',', '.')
            
            if lat_str.count('.') > 1:
                parts = lat_str.split('.')
                lat_str = parts[-2] + '.' + parts[-1] if len(parts) > 1 else lat_str
                lat_str = lat_str.replace('.', '', lat_str.count('.') - 1)
            
            if lon_str.count('.') > 1:
                parts = lon_str.split('.')
                lon_str = parts[-2] + '.' + parts[-1] if len(parts) > 1 else lon_str
                lon_str = lon_str.replace('.', '', lon_str.count('.') - 1)

            try:
                lat_float = float(lat_str)
                lon_float = float(lon_str)
                
                if abs(lat_float) > 90 or abs(lon_float) > 180:
                    print(f"⚠️ Atlandı (çok büyük): {direk_id} - {lat_float}, {lon_float}")
                    continue
                
                if not (35 <= lat_float <= 42 and 25 <= lon_float <= 45):
                    print(f"⚠️ Atlandı (Türkiye dışı): {direk_id} - {lat_float}, {lon_float}")
                    continue
                
                direk_listesi.append({
                    "id": direk_id,
                    "tip": tip,
                    "lat": lat_float,
                    "lon": lon_float
                })
                
            except (ValueError, TypeError) as e:
                print(f"❌ Koordinat hatası ({direk_id}): {e} - orijinal lat: {lat}, lon: {lon}, işlenmiş: {lat_str}, {lon_str}")
                continue
                
    except Exception as e:
        print(f"❌ Direk verisi okunamadı: {e}")
    return direk_listesi

def oku_direkler_musterek(excel_yolu):
    musterek_listesi = []
    try:
        wb = load_workbook(excel_yolu, data_only=True)
        ws = wb.active

        for row in ws.iter_rows(min_row=2):
            direk_id = str(row[1].value).strip()
            tip = str(row[2].value).strip()
            lat = row[5].value
            lon = row[6].value

            if lat is None or lon is None:
                continue
            
            lat_str = str(lat).strip()
            lon_str = str(lon).strip()
            
            if lat_str == "" or lon_str == "":
                continue
            
            lat_str = lat_str.replace(',', '.')
            lon_str = lon_str.replace(',', '.')
            
            if lat_str.count('.') > 1:
                parts = lat_str.split('.')
                lat_str = parts[-2] + '.' + parts[-1] if len(parts) > 1 else lat_str
                lat_str = lat_str.replace('.', '', lat_str.count('.') - 1)
            
            if lon_str.count('.') > 1:
                parts = lon_str.split('.')
                lon_str = parts[-2] + '.' + parts[-1] if len(parts) > 1 else lon_str
                lon_str = lon_str.replace('.', '', lon_str.count('.') - 1)

            try:
                lat_float = float(lat_str)
                lon_float = float(lon_str)
                
                if abs(lat_float) > 90 or abs(lon_float) > 180:
                    print(f"⚠️ Atlandı (müşterek, çok büyük): {direk_id} - {lat_float}, {lon_float}")
                    continue
                
                if not (35 <= lat_float <= 42 and 25 <= lon_float <= 45):
                    print(f"⚠️ Atlandı (müşterek, Türkiye dışı): {direk_id} - {lat_float}, {lon_float}")
                    continue
                
                musterek_listesi.append({
                    "id": direk_id,
                    "tip": tip,
                    "lat": lat_float,
                    "lon": lon_float,
                    "musterek": True
                })
                
            except (ValueError, TypeError) as e:
                print(f"❌ Müşterek koordinat hatası ({direk_id}): {e} - lat: {lat}, lon: {lon}")
                continue
                
    except Exception as e:
        print(f"❌ Müşterek direk verisi okunamadı: {e}")
    return musterek_listesi


class FotoEditorViewer(QMainWindow):
    def __init__(self, foto_verileri, view=None, parent=None):
        super().__init__(parent)
        self.editor_viewer = parent
        self.view = view
        self.fotolar = foto_verileri
        for foto in self.fotolar:
            foto.setdefault("oncelik", "Bekleyebilir")
        self.index = 0
        
        self._guncelle_baslik()
        
        self.setMinimumSize(1200, 1000)
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.setFocus()

        self.setStyleSheet("""
            QMainWindow {
                background-color: #f0f0f0;
            }
        """)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        self.layout = QVBoxLayout(central_widget)

        self.jump_panel = QFrame()
        self.jump_panel.setFixedHeight(60)
        self.jump_panel.setStyleSheet("""
            QFrame {
                background-color: rgba(240, 240, 240, 200);
                border-bottom: 2px solid #cccccc;
            }
        """)
        
        jump_layout = QHBoxLayout(self.jump_panel)
        jump_layout.setContentsMargins(10, 5, 10, 5)
        
        jump_label = QLabel("Fotoğraf No:")
        jump_label.setStyleSheet("font-weight: bold; font-size: 14px;")
        jump_layout.addWidget(jump_label)
        
        self.jump_spinbox = QSpinBox()
        self.jump_spinbox.setRange(1, len(self.fotolar))
        self.jump_spinbox.setValue(self.index + 1)
        self.jump_spinbox.setSuffix(f" / {len(self.fotolar)}")
        self.jump_spinbox.setFixedWidth(120)
        self.jump_spinbox.setStyleSheet("""
            QSpinBox {
                font-size: 14px;
                padding: 5px;
                border: 2px solid #0078d7;
                border-radius: 4px;
            }
        """)
        jump_layout.addWidget(self.jump_spinbox)
        
        self.jump_button = QPushButton("📂 Fotoğrafa Git")
        self.jump_button.setFixedHeight(35)
        self.jump_button.setStyleSheet("""
            QPushButton {
                background-color: #0078d7;
                color: white;
                font-weight: bold;
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #005a9e;
            }
            QPushButton:pressed {
                background-color: #004578;
            }
        """)
        self.jump_button.clicked.connect(self.jump_to_photo)
        jump_layout.addWidget(self.jump_button)
        
        folder_name = os.path.basename(os.path.dirname(self.fotolar[self.index]['path']))
        file_name = os.path.basename(self.fotolar[self.index]['path'])
        self.foto_name_label = QLabel(f"📷 {folder_name}/{file_name}")
        self.foto_name_label.setStyleSheet("font-size: 13px; color: #555555; margin-left: 20px;")
        self.foto_name_label.setWordWrap(True)
        self.foto_name_label.setToolTip(self.fotolar[self.index]['path'])
        jump_layout.addWidget(self.foto_name_label)
        
        jump_layout.addStretch()
        
        self.sil_buton = QPushButton("🗑️ Sil")
        self.sil_buton.setFixedHeight(35)
        self.sil_buton.setStyleSheet("""
            QPushButton {
                background-color: #ff4444;
                color: white;
                font-weight: bold;
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
                font-size: 14px;
                margin-right: 10px;
            }
            QPushButton:hover {
                background-color: #cc0000;
            }
            QPushButton:pressed {
                background-color: #990000;
            }
        """)
        self.sil_buton.clicked.connect(self.fotografi_sil)
        self.sil_buton.setToolTip("Bu fotoğrafın marker'ını haritadan kaldır")
        jump_layout.addWidget(self.sil_buton)
        
        close_btn = QPushButton("❌ Kapat")
        close_btn.setFixedHeight(35)
        close_btn.setStyleSheet("""
            QPushButton {
                background-color: #dc3545;
                color: white;
                font-weight: bold;
                border: none;
                border-radius: 4px;
                padding: 8px 16px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #c82333;
            }
        """)
        close_btn.clicked.connect(self.close)
        jump_layout.addWidget(close_btn)
        
        self.layout.addWidget(self.jump_panel)

        self.canvas_layout = QVBoxLayout()

        self.canvas = FotoEditorCanvas(self.fotolar[self.index]["path"], parent=self)
        self.canvas_layout.addWidget(self.canvas)

        self.oncelik_widget = QFrame(self)
        self.oncelik_widget.setFixedSize(160, 140)
        self.oncelik_widget.setStyleSheet("""
            background-color: rgba(255, 255, 255, 180);
            border-radius: 8px;
            border: 1px solid rgba(0,0,0,50);
        """)

        baslik_label = QLabel("Öncelik:", self.oncelik_widget)
        baslik_label.move(10, 10)
        baslik_label.setStyleSheet("font-weight: bold; font-size:13px; color:black;")

        self.oncelik_grup = QButtonGroup(self.oncelik_widget)
        self.oncelik_secenekleri = ["Bekleyebilir", "Normal", "Acil", "Çok Acil"]

        for i, secenek in enumerate(self.oncelik_secenekleri):
            rb = QRadioButton(secenek, self.oncelik_widget)
            rb.move(10, 35 + i*22)
            rb.setStyleSheet("color:black;")
            self.oncelik_grup.addButton(rb, id=i)
            rb.toggled.connect(self.onceligi_kaydet)
        self.oncelik_widget.show()

        alt_mesafe = 80

        def oncelik_widget_yerlestir():
            margin_x = 15
            x = self.width() - self.oncelik_widget.width() - margin_x
            y = self.height() - self.oncelik_widget.height() - alt_mesafe
            self.oncelik_widget.move(x, y)

        oncelik_widget_yerlestir()

        self.resizeEvent = lambda event: oncelik_widget_yerlestir()

        self.toolbar_layout = QHBoxLayout()
        self.toolbar_layout.addWidget(self._btn("⬅️ Önceki", self.onceki_foto))
        self.toolbar_layout.addWidget(self._btn("➡️ Sonraki", self.sonraki_foto))
        self.toolbar_layout.addWidget(self._btn("✏️ Çiz", lambda: self.canvas.set_tool("draw")))
        self.toolbar_layout.addWidget(self._btn("📝 Yazı", lambda: self.canvas.set_tool("text")))
        self.toolbar_layout.addWidget(self._btn("🖐️ Taşı", lambda: self.canvas.set_tool("pan")))
        self.toolbar_layout.addWidget(self._btn("↩️ Geri Al", self.canvas.undo))
        self.toolbar_layout.addWidget(self._btn("💾 Kaydet", self.canvas.kaydet))

        self.thickness_input = QSpinBox()
        self.thickness_input.setRange(1, 100)
        self.thickness_input.setValue(15)
        self.thickness_input.valueChanged.connect(lambda val: setattr(self.canvas, 'pen_thickness', val))
        self.toolbar_layout.addWidget(QLabel("Kalınlık:"))
        self.toolbar_layout.addWidget(self.thickness_input)

        self.yazi_boyutu_input = QSpinBox()
        self.yazi_boyutu_input.setRange(8, 200)
        self.yazi_boyutu_input.setValue(120)
        self.yazi_boyutu_input.valueChanged.connect(lambda val: setattr(self.canvas, 'yazi_boyutu', val))
        self.canvas.yazi_boyutu = self.yazi_boyutu_input.value()
        self.toolbar_layout.addWidget(QLabel("Yazı Boyutu:"))
        self.toolbar_layout.addWidget(self.yazi_boyutu_input)

        renk_btn = QPushButton("🎨 Renk")
        renk_btn.clicked.connect(self.renk_sec)
        renk_btn.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        renk_btn.setAutoDefault(False)
        renk_btn.setDefault(False)
        self.toolbar_layout.addWidget(renk_btn)

        self.comboBox = QComboBox()
        self.comboBox.addItems([
            "Kategori Seçin",
            "a- Direk Kırık/Çatlak/Çürük",
            "b- Travers veya Konsol Kırık/Eğri",
            "c- İzolatör Kırık/Çatlak",
            "d- Ağaç veya Bitki Teması/Tehlikesi var.",
            "e- Sıkıbağ kopuk veya yenilenmesi gerekiyor",
            "f- Sehim alınması gerekiyor",
            "g- Trafoda yağ kaçağı mevcut",
            "h- OG Sigorta yerine tel sarılı",
            "i- Diğer"
        ])
        self.comboBox.currentIndexChanged.connect(self.kaydet_kategori)
        self.toolbar_layout.addWidget(self.comboBox)

        self.note_input = QLineEdit()
        not_yazisi = self.fotolar[self.index].get("note", "")
        self.note_input.setText(not_yazisi)
        self.note_input.textChanged.connect(self._notu_kaydet)
        self.note_input.setPlaceholderText("📝 Fotoğraf notu yazın...")
        self.note_input.setMinimumHeight(35)
        self.note_input.setStyleSheet("""
            QLineEdit {
                font-size: 14px;
                padding: 8px;
                border: 2px solid #cccccc;
                border-radius: 4px;
            }
            QLineEdit:focus {
                border: 2px solid #0078d7;
            }
        """)
        self.note_input.returnPressed.connect(self.note_input.clearFocus)
        
        self.note_input.setFocus()
        
        self.toolbar_full_layout = QVBoxLayout()
        self.toolbar_full_layout.addLayout(self.toolbar_layout)
        self.toolbar_full_layout.addWidget(self.note_input)

        self.canvas_layout.addLayout(self.toolbar_full_layout)
        self.layout.addLayout(self.canvas_layout)
        self._yeniden_yukle()
        
        self.jump_spinbox.valueChanged.connect(self._on_jump_spinbox_changed)
        
        QTimer.singleShot(100, lambda: self.note_input.setFocus())
        
    def _guncelle_baslik(self):
        if not self.fotolar:
            self.setWindowTitle("🖼️ Fotoğraf Editör")
            return
            
        foto_adi = os.path.basename(self.fotolar[self.index]["path"])
        klasor_adi = os.path.basename(os.path.dirname(self.fotolar[self.index]["path"]))
        toplam = len(self.fotolar)
        mevcut = self.index + 1
        kalan = toplam - mevcut
        
        baslik = f"🖼️ Fotoğraf Editör - {klasor_adi}/{foto_adi} - 📊 {mevcut}/{toplam} - 📁 Kalan: {kalan}"
        self.setWindowTitle(baslik)
        
        if hasattr(self, 'jump_spinbox'):
            self.jump_spinbox.setSuffix(f" / {toplam}")
        
    def _on_jump_spinbox_changed(self, value):
        if 0 <= value - 1 < len(self.fotolar):
            foto_adi = os.path.basename(self.fotolar[value - 1]['path'])
            klasor_adi = os.path.basename(os.path.dirname(self.fotolar[value - 1]['path']))
            self.foto_name_label.setText(f"📷 {klasor_adi}/{foto_adi}")
            self.foto_name_label.setToolTip(self.fotolar[value - 1]['path'])
        
    def _btn(self, text, func):
        btn = QPushButton(text)
        btn.clicked.connect(func)
        btn.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        btn.setAutoDefault(False)
        btn.setDefault(False)
        return btn

    def renk_sec(self):
        color = QColorDialog.getColor()
        if color.isValid():
            self.canvas.pen_color = color

    def sonraki_foto(self):
        if self.index < len(self.fotolar) - 1:
            self.index += 1
            self._yeniden_yukle()
            self.jump_spinbox.setValue(self.index + 1)
            self._guncelle_baslik()
            QTimer.singleShot(50, lambda: self.note_input.setFocus())
        else:
            QMessageBox.information(self, "Son Fotoğraf", 
                                  "🎉 Son fotoğrafa ulaştınız!\n\n"
                                  "Başka bir fotoğrafa atlamak için üstteki fotoğraf numarasını değiştirebilirsiniz.")

    def onceki_foto(self):
        if self.index > 0:
            self.index -= 1
            self._yeniden_yukle()
            self.jump_spinbox.setValue(self.index + 1)
            self._guncelle_baslik()
            QTimer.singleShot(50, lambda: self.note_input.setFocus())
        else:
            QMessageBox.information(self, "İlk Fotoğraf", 
                                  "🏁 İlk fotoğraftasınız!\n\n"
                                  "Başka bir fotoğrafa atlamak için üstteki fotoğraf numarasını değiştirebilirsiniz.")

    def _yeniden_yukle(self):
        self.canvas_layout.removeWidget(self.canvas)
        self.canvas.deleteLater()

        self.canvas = FotoEditorCanvas(self.fotolar[self.index]["path"], parent=self)
        not_yazisi = self.fotolar[self.index].get("note", "")
        self.note_input.setText(not_yazisi)
        
        self.note_input.setFocus()
        
        self.canvas.pen_thickness = self.thickness_input.value()
        self.canvas.yazi_boyutu = self.yazi_boyutu_input.value()
        self.canvas.pen_color = self.canvas.pen_color

        self.canvas_layout.insertWidget(0, self.canvas)

        kategori = self.fotolar[self.index].get("kategori", "")
        if kategori:
            self.comboBox.setCurrentText(kategori)
        else:
            self.comboBox.setCurrentIndex(0)

        foto_adi = os.path.basename(self.fotolar[self.index]['path'])
        klasor_adi = os.path.basename(os.path.dirname(self.fotolar[self.index]['path']))
        self.foto_name_label.setText(f"📷 {klasor_adi}/{foto_adi}")
        self.foto_name_label.setToolTip(self.fotolar[self.index]['path'])

        self.toolbar_layout.itemAt(2).widget().clicked.connect(lambda: self.canvas.set_tool("draw"))
        self.toolbar_layout.itemAt(3).widget().clicked.connect(lambda: self.canvas.set_tool("text"))
        self.toolbar_layout.itemAt(4).widget().clicked.connect(lambda: self.canvas.set_tool("pan"))
        self.toolbar_layout.itemAt(5).widget().clicked.connect(self.canvas.undo)
        self.toolbar_layout.itemAt(6).widget().clicked.connect(self.canvas.kaydet)

        self.thickness_input.valueChanged.connect(lambda val: setattr(self.canvas, 'pen_thickness', val))
        self.yazi_boyutu_input.valueChanged.connect(lambda val: setattr(self.canvas, 'yazi_boyutu', val))

        foto_oncelik = self.fotolar[self.index].get("oncelik", "Bekleyebilir")
        for buton in self.oncelik_grup.buttons():
            if buton.text() == foto_oncelik:
                buton.setChecked(True)
                break
        
        self._guncelle_baslik()
            
    def onceligi_kaydet(self):
        secilen_buton = self.oncelik_grup.checkedButton()
        if secilen_buton:
            secilen_oncelik = secilen_buton.text()
            self.fotolar[self.index]["oncelik"] = secilen_oncelik

            if self.view:
                js_kod = f"""
                (function() {{
                    var fullPath = "{self.fotolar[self.index]['path'].replace('\\', '/')}";
                    var oncelik = "{secilen_oncelik}";
                    var foto = photoInfo.find(p => p.path === fullPath);
                    if (foto) {{
                        foto.oncelik = oncelik;
                    }}
                }})();
                """
                self.view.page().runJavaScript(js_kod)

    def keyPressEvent(self, event):
        if event.key() == Qt.Key.Key_Right:
            self.sonraki_foto()
        elif event.key() == Qt.Key.Key_Left:
            self.onceki_foto()
        elif event.key() == Qt.Key.Key_Escape:
            self.close()
        elif event.key() == Qt.Key.Key_G and event.modifiers() == Qt.KeyboardModifier.ControlModifier:
            self.jump_spinbox.setFocus()
            self.jump_spinbox.selectAll()
        elif event.key() == Qt.Key.Key_Enter or event.key() == Qt.Key.Key_Return:
            if self.jump_spinbox.hasFocus():
                self.jump_to_photo()
            else:
                super().keyPressEvent(event)
        else:
            super().keyPressEvent(event)

    def update_marker_color(self, image_path):
        parent = self.parent()
        if not parent or not hasattr(parent, "view"):
            return

        js_path = image_path.replace("\\", "/")
        js = f"""
            (function() {{
                const photo = photoInfo.find(p => p.path === "{js_path}");
                if (photo && photo.marker && photo.marker._icon) {{
                    photo.edited = true;
                    const el = photo.marker._icon;
                    el.classList.remove("photo-mavi");
                    el.classList.add("photo-turkuaz");
                }}
            }})();
        """
        parent.view.page().runJavaScript(js)

    def _notu_kaydet(self, yeni_not):
        if 0 <= self.index < len(self.fotolar):
            self.fotolar[self.index]["note"] = yeni_not
            
        if self.view:
            full_path = self.fotolar[self.index]["path"].replace("\\", "/")
            js = f"""
            (function() {{
                const fullPath = "{full_path}";
                const p = photoInfo.find(p => p.path === fullPath);
                if (p) {{
                    p.note = `{yeni_not}`;
                }}

                direkler.forEach(d => {{
                    if (d.photos) {{
                        d.photos.forEach(f => {{
                            if (f.path === fullPath) {{
                                f.note = `{yeni_not}`;
                            }}
                        }});
                    }}
                }});
            }})();
            """
            self.view.page().runJavaScript(js)
            self.view.page().runJavaScript("guncellePopuplar();")

    def kaydet_kategori(self):
        kategori = self.comboBox.currentText()
        if kategori == "Kategori Seçin":
            self.fotolar[self.index]["kategori"] = ""
        else:
            self.fotolar[self.index]["kategori"] = kategori

        js_kod = f"""
        (function(){{
            var fullPath = "{self.fotolar[self.index]['path'].replace('\\', '/')}";
            var kategori = "{kategori if kategori != 'Kategori Seçin' else ''}";
            var foto = photoInfo.find(p => p.path === fullPath);
            if(foto) foto.kategori = kategori;
        }})();
        """
        self.view.page().runJavaScript(js_kod)

    def jump_to_photo(self):
        try:
            target_index = self.jump_spinbox.value() - 1
            
            if 0 <= target_index < len(self.fotolar):
                if target_index == self.index:
                    return
                
                self.index = target_index
                self._yeniden_yukle()
                self._guncelle_baslik()
                
                QMessageBox.information(
                    self, 
                    "Fotoğraf Değiştirildi", 
                    f"✅ {self.index + 1}. fotoğrafa geçildi.\n\n"
                    f"📁 {os.path.basename(os.path.dirname(self.fotolar[self.index]['path']))}/{os.path.basename(self.fotolar[self.index]['path'])}\n\n"
                    f"Artık {self.index + 1}. fotodan başlayarak ilerleyebilirsiniz."
                )
                
                QTimer.singleShot(50, lambda: self.note_input.setFocus())
            else:
                QMessageBox.warning(self, "Geçersiz Numara", 
                                  f"⚠️ Lütfen 1 ile {len(self.fotolar)} arasında bir numara girin.")
                
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Fotoğrafa atlama sırasında hata: {str(e)}")

    def fotografi_sil(self):
        try:
            current_foto = self.fotolar[self.index]
            foto_yolu = current_foto['path']
            folder_name = os.path.basename(os.path.dirname(foto_yolu))
            file_name = os.path.basename(foto_yolu)
            
            reply = QMessageBox.question(
                self, 
                "Fotoğrafı Sil",
                f"'{folder_name}/{file_name}' fotoğrafının marker'ını haritadan kaldırmak istiyor musunuz?\n\n"
                "Bu işlem geri alınamaz!",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )
            
            if reply == QMessageBox.StandardButton.Yes:
                if self.view:
                    js_kod = f"""
                    (function() {{
                        const fullPath = "{foto_yolu.replace('\\', '/')}";
                        
                        const fotoIndex = photoInfo.findIndex(p => p.path === fullPath);
                        
                        if (fotoIndex !== -1) {{
                            const foto = photoInfo[fotoIndex];
                            
                            if (foto.marker && map.hasLayer(foto.marker)) {{
                                map.removeLayer(foto.marker);
                            }}
                            
                            photoInfo.splice(fotoIndex, 1);
                            console.log("🗑️ Fotoğraf haritadan kaldırıldı:", fullPath);
                        }}
                        
                        direkler.forEach(direk => {{
                            if (direk.photos && direk.photos.length > 0) {{
                                const fotoIndex = direk.photos.findIndex(p => p.path === fullPath);
                                
                                if (fotoIndex !== -1) {{
                                    direk.photos.splice(fotoIndex, 1);
                                    
                                    if (direk.photos.length === 0) {{
                                        direk.numara = null;
                                        if (direk.numaraLabelMarker && map.hasLayer(direk.numaraLabelMarker)) {{
                                            map.removeLayer(direk.numaraLabelMarker);
                                        }}
                                    }}
                                    
                                    guncelleDirekRengi(direk);
                                    guncellePopup(direk);
                                }}
                            }}
                        }});
                        
                        return "✅ Fotoğraf haritadan kaldırıldı";
                    }})();
                    """
                    
                    self.view.page().runJavaScript(js_kod, self._foto_silme_tamamlandi)
                    
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Fotoğraf silinirken hata oluştu:\n{str(e)}")
    
    def _foto_silme_tamamlandi(self, sonuc):
        try:
            if sonuc:
                print(sonuc)
            
            silinen_foto = self.fotolar[self.index]
            silinen_folder = os.path.basename(os.path.dirname(silinen_foto['path']))
            silinen_file = os.path.basename(silinen_foto['path'])
            
            self.fotolar.pop(self.index)
            
            if not self.fotolar:
                QMessageBox.information(self, "Tüm Fotoğraflar Silindi", 
                                      "Tüm fotoğraflar haritadan kaldırıldı.\nPencere kapatılıyor...")
                self.close()
                return
            
            if self.index >= len(self.fotolar):
                self.index = len(self.fotolar) - 1
            
            self.jump_spinbox.setMaximum(len(self.fotolar))
            self.jump_spinbox.setValue(self.index + 1)
            self.jump_spinbox.setSuffix(f" / {len(self.fotolar)}")
            
            yeni_folder = os.path.basename(os.path.dirname(self.fotolar[self.index]['path']))
            yeni_file = os.path.basename(self.fotolar[self.index]['path'])
            self.foto_name_label.setText(f"📷 {yeni_folder}/{yeni_file}")
            self.foto_name_label.setToolTip(self.fotolar[self.index]['path'])
            
            self._yeniden_yukle()
            
            self._guncelle_baslik()
            
            QMessageBox.information(self, "Fotoğraf Silindi", 
                                  f"✅ '{silinen_folder}/{silinen_file}' fotoğrafı haritadan kaldırıldı.\n\n"
                                  f"📊 Kalan fotoğraf sayısı: {len(self.fotolar)}\n"
                                  f"📍 Şimdi yeni {self.index + 1}. fotoğrafı görüntülüyorsunuz.")
            
        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Liste güncelleme sırasında hata:\n{str(e)}")

class FotoEditorCanvas(QWidget):
    def __init__(self, image_path, parent=None):
        super().__init__(parent)
        self.image_path = image_path
        self.editor_viewer = parent
        self.original_pixmap = QPixmap(image_path)
        self.display_pixmap = self.original_pixmap.copy()
        self.setMinimumSize(800, 600)
        
        self.setMouseTracking(True)

        self.pen_color = QColor("red")
        self.pen_thickness = 15
        self.current_tool = "pan"
        self.yazi_kutusu = None
        self.yazi_kutusu_pos = None
        self.yazi_boyutu = 120

        self.drawings = []
        self.texts = []
        self.undo_stack = []

        self.last_pos = None
        self.panning = False
        self.offset = QPoint(0, 0)
        self.temp_offset = QPoint(0, 0)
        self.scale_factor = 1.0
        
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        
        self._calculate_initial_scale()
    
    def _calculate_initial_scale(self):
        if self.original_pixmap.isNull():
            return

        view_width = self.width()
        view_height = self.height()
        pixmap_width = self.original_pixmap.width()
        pixmap_height = self.original_pixmap.height()

        width_scale = view_width / pixmap_width
        height_scale = view_height / pixmap_height
        
        self.scale_factor = min(width_scale, height_scale)
        
        self.offset_x = (view_width - (pixmap_width * self.scale_factor)) / 2
        self.offset_y = (view_height - (pixmap_height * self.scale_factor)) / 2
        
        self.offset = QPoint(int(self.offset_x), int(self.offset_y))
        self.temp_offset = QPoint(0, 0)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._calculate_initial_scale()
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.translate(self.offset + self.temp_offset)
        painter.scale(self.scale_factor, self.scale_factor)
        painter.drawPixmap(0, 0, self.display_pixmap)

        for start, end, color, thickness in self.drawings:
            painter.setPen(QPen(color, thickness))
            painter.drawLine(start, end)

        for pos, text, color in self.texts:
            painter.setPen(QPen(color))
            painter.setFont(QFont("Arial", self.yazi_boyutu))
            painter.drawText(pos, text)

    def wheelEvent(self, event):
        angle = event.angleDelta().y()
        
        if angle > 0:
            zoom_factor = 1.25
        else:
            zoom_factor = 0.8
        
        mouse_pos = event.position()
        
        offset = QPointF(self.offset)
        temp_offset = QPointF(self.temp_offset)
        
        before_scale = (mouse_pos - offset - temp_offset) / self.scale_factor
        
        old_scale = self.scale_factor
        self.scale_factor *= zoom_factor
        
        self.scale_factor = max(0.1, min(10.0, self.scale_factor))
        
        after_scale = before_scale * self.scale_factor
        
        new_offset = mouse_pos - after_scale
        
        self.offset = new_offset.toPoint()
        
        self.temp_offset = QPoint(0, 0)
        
        self.update()
        
        event.accept()

    def mouseDoubleClickEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            mouse_pos = event.position()
            offset = QPointF(self.offset)
            temp_offset = QPointF(self.temp_offset)
            
            before_scale = (mouse_pos - offset - temp_offset) / self.scale_factor
            self.scale_factor *= 1.5
            self.scale_factor = min(10.0, self.scale_factor)
            after_scale = before_scale * self.scale_factor
            new_offset = mouse_pos - after_scale
            self.offset = new_offset.toPoint()
            self.temp_offset = QPoint(0, 0)
            self.update()
            event.accept()
            
        elif event.button() == Qt.MouseButton.RightButton:
            mouse_pos = event.position()
            offset = QPointF(self.offset)
            temp_offset = QPointF(self.temp_offset)
            
            before_scale = (mouse_pos - offset - temp_offset) / self.scale_factor
            self.scale_factor *= 0.67
            self.scale_factor = max(0.1, self.scale_factor)
            after_scale = before_scale * self.scale_factor
            new_offset = mouse_pos - after_scale
            self.offset = new_offset.toPoint()
            self.temp_offset = QPoint(0, 0)
            self.update()
            event.accept()

    def mousePressEvent(self, event):
        LEFT_AREA_PERCENT = 0.075
        RIGHT_AREA_PERCENT = 0.075
        
        left_threshold = self.width() * LEFT_AREA_PERCENT
        right_threshold = self.width() * (1 - RIGHT_AREA_PERCENT)
        
        if event.button() == Qt.MouseButton.LeftButton and event.pos().x() < left_threshold:
            if self.editor_viewer and hasattr(self.editor_viewer, 'onceki_foto'):
                self.editor_viewer.onceki_foto()
            return
        
        elif event.button() == Qt.MouseButton.RightButton or \
             (event.button() == Qt.MouseButton.LeftButton and event.pos().x() > right_threshold):
            if self.editor_viewer and hasattr(self.editor_viewer, 'sonraki_foto'):
                self.editor_viewer.sonraki_foto()
            return
        
        elif event.button() == Qt.MouseButton.LeftButton:
            pos = (event.pos() - self.offset - self.temp_offset) / self.scale_factor
            self.last_pos = pos

            if self.current_tool == "pan":
                self.panning = True

            elif self.current_tool == "text":
                if self.yazi_kutusu and self.yazi_kutusu.isVisible():
                    self.yazi_kutusu.setFocus()
                    return

                self.yazi_kutusu_pos = pos
                self.yazi_kutusu = QLineEdit(self)
                self.yazi_kutusu.setFont(QFont("Arial", self.yazi_boyutu))
                self.yazi_kutusu.setStyleSheet(f"""
                    color: {self.pen_color.name()};
                    background-color: rgba(255, 255, 255, 0);
                    border: none;
                """)
                self.yazi_kutusu.resize(200, 30)
                global_pos = self.offset + self.temp_offset + pos * self.scale_factor
                self.yazi_kutusu.move(int(global_pos.x()), int(global_pos.y()))
                self.yazi_kutusu.returnPressed.connect(self.yazi_onayla)
                self.yazi_kutusu.show()
                self.yazi_kutusu.setFocus()

    def mouseMoveEvent(self, event):
        if self.current_tool == "pan" and self.last_pos:
            delta = event.pos() - self.last_pos * self.scale_factor - self.offset
            self.temp_offset = delta
            self.update()
            return

        if self.current_tool == "draw" and event.buttons() & Qt.MouseButton.LeftButton and self.last_pos:
            current_pos = (event.pos() - self.offset - self.temp_offset) / self.scale_factor
            segment = (self.last_pos, current_pos, self.pen_color, self.pen_thickness)
            self.drawings.append(segment)
            self.undo_stack.append(("draw", segment))
            self.last_pos = current_pos
            self.update()

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.last_pos = None
            self.offset += self.temp_offset
            self.temp_offset = QPoint(0, 0)
            self.panning = False
            self.update()

    def yazi_onayla(self):
        if not self.yazi_kutusu:
            return

        text = self.yazi_kutusu.text().strip()
        if text and self.yazi_kutusu_pos is not None:
            pos = QPointF(self.yazi_kutusu_pos)
            self.texts.append((pos, text, self.pen_color))
            self.undo_stack.append(("text", (pos, text, self.pen_color)))
            self.update()

        self.yazi_kutusu.deleteLater()
        self.yazi_kutusu = None
        self.yazi_kutusu_pos = None

    def undo(self):
        if not self.undo_stack:
            return
        action, value = self.undo_stack.pop()
        if action == "draw" and value in self.drawings:
            self.drawings.remove(value)
        elif action == "text" and value in self.texts:
            self.texts.remove(value)
        self.update()

    def kaydet(self):
        try:
            pil_img = Image.open(self.image_path)
            exif_bytes = piexif.dump(piexif.load(pil_img.info.get("exif", b"")))
        except Exception:
            exif_bytes = None

        final_image = QImage(self.original_pixmap.size(), QImage.Format.Format_RGB32)
        final_image.fill(Qt.GlobalColor.white)

        painter = QPainter(final_image)
        painter.drawPixmap(0, 0, self.original_pixmap)

        for start, end, color, thickness in self.drawings:
            painter.setPen(QPen(color, thickness))
            painter.drawLine(start, end)

        for pos, text, color in self.texts:
            painter.setPen(QPen(color))
            painter.setFont(QFont("Arial", self.yazi_boyutu))
            painter.drawText(pos, text)

        painter.end()

        pil_final = ImageQt.fromqimage(final_image)

        try:
            if exif_bytes:
                pil_final.save(self.image_path, exif=exif_bytes)
            else:
                pil_final.save(self.image_path)

            try:
                import winsound
                winsound.MessageBeep(winsound.MB_ICONASTERISK)
            except:
                pass
            
            from PyQt6.QtCore import QTimer
            
            self.info_label = QLabel(f"✓ {os.path.basename(self.image_path)} kaydedildi", self)
            self.info_label.setStyleSheet("""
                QLabel {
                    background-color: #4CAF50;
                    color: white;
                    font-weight: bold;
                    padding: 10px;
                    border-radius: 5px;
                    border: 1px solid #388E3C;
                }
            """)
            self.info_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.info_label.resize(300, 40)
            
            self.info_label.move(
                (self.width() - self.info_label.width()) // 2,
                self.height() - 100
            )
            
            self.info_label.show()
            
            QTimer.singleShot(1000, self.info_label.hide)

            if hasattr(self.editor_viewer, "update_marker_color"):
                self.editor_viewer.update_marker_color(self.image_path)

            if hasattr(self.editor_viewer, "note_input") and hasattr(self.editor_viewer, "view"):
                not_metni = self.editor_viewer.note_input.text().strip()
                if not_metni:
                    js_path = self.image_path.replace("\\", "/")
                    js = f"""
                        (function() {{
                            const photo = photoInfo.find(p => p.path === "{js_path}");
                            if (photo) {{
                                photo.note = `{not_metni}`;
                                if (photo.marker && photo.marker._popup) {{
                                    photo.marker.setPopupContent(getPhotoPopupHTML(photo.path, photo.note, photo.birim));
                                }}
                            }}
                        }})();
                    """
                    self.editor_viewer.view.page().runJavaScript(js)

                if self.editor_viewer:
                    self.editor_viewer.fotolar[self.editor_viewer.index]['note'] = self.editor_viewer.note_input.text().strip()

        except Exception as e:
            QMessageBox.warning(self, "Hata", f"❌ Kaydedilemedi: {e}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = DroneHarita()
    window.show()
    sys.exit(app.exec())
