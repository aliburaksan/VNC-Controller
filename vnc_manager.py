import sys
import os
import subprocess
import json
import pandas as pd
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QTreeWidget, QTreeWidgetItem, QScrollArea,
    QLineEdit, QMessageBox, QMenu, QButtonGroup, QInputDialog
)
from PySide6.QtCore import Qt, QThread, Signal, QMimeData, QTimer
from PySide6.QtGui import QColor, QBrush, QPixmap, QIcon, QPainter, QLinearGradient

# ---------------- CONFIG ----------------
EXCEL_PATH = r"\\arman06v\IT\Programs\VNCData\Manisa IP LIST.xlsx"
VNC_PASSWORD = "Man12378"

SHEET_MAP = {
    "Montaj 1": "10.105.36.x",
    "Montaj 2": "10.105.37.x",
    "Montaj 3": "10.105.38.x"
}
OCTET_MAP = {
    "Montaj 1": "36",
    "Montaj 2": "37",
    "Montaj 3": "38"
}

def get_data_path(filename):
    if getattr(sys, 'frozen', False):
        # --onedir veya --onefile fark etmez, exe dizinini kullan
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, filename)

def get_user_folder_file():
    """
    Exe ile aynı dizinde user_folders.json kullanır.
    Normal Python çalıştırırken script dizininde olur.
    """
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)  # exe dizini
    else:
        base_path = os.path.dirname(__file__)       # script dizini
    return os.path.join(base_path, "user_folders.json")

USER_FOLDER_FILE = get_user_folder_file()


# ---------- Custom widgets ----------
class DragTree(QTreeWidget):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.setDragEnabled(True)
        self.setDragDropMode(QTreeWidget.DragOnly)

    def mimeData(self, items):
        mime = QMimeData()
        lines = []
        for it in items:
            cols = [it.text(c) for c in range(it.columnCount())]
            lines.append("\t".join(cols))
        mime.setText("\n".join(lines))
        return mime

class DropButton(QPushButton):
    def __init__(self, title, drop_callback=None, *args, **kwargs):
        super().__init__(title, *args, **kwargs)
        self.setAcceptDrops(True)
        self.drop_callback = drop_callback

    def dragEnterEvent(self, event):
        if event.mimeData().hasText():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.mimeData().hasText():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        if not event.mimeData().hasText():
            event.ignore()
            return
        text = event.mimeData().text().splitlines()[0]
        parts = text.split("\t")
        ip = parts[0] if len(parts) > 0 else ""
        asset = parts[1] if len(parts) > 1 else ""
        desc = parts[2] if len(parts) > 2 else ""
        if self.drop_callback:
            self.drop_callback(self.text(), ip, asset, desc)
        event.acceptProposedAction()


# ---------- Workers ----------
class LoadSheetWorker(QThread):
    data_loaded = Signal(list)
    error = Signal(str)
    def __init__(self, sheet_name, octet):
        super().__init__()
        self.sheet_name = sheet_name
        self.octet = octet

    def run(self):
        rows = []
        if not os.path.exists(EXCEL_PATH):
            self.error.emit(f"Excel dosyası bulunamadı: {EXCEL_PATH}")
            self.data_loaded.emit(rows)
            return
        try:
            df = pd.read_excel(EXCEL_PATH, sheet_name=self.sheet_name, header=None, dtype=str, engine="openpyxl")
            df = df.iloc[:, :3]
            df.columns = ["IP", "Asset", "Description"]
            df = df.fillna("")
            for _, r in df.iterrows():
                ip = str(r["IP"]).strip()
                try:
                    parts = ip.split(".")
                    if len(parts) == 4 and parts[2] == self.octet and int(parts[3]) >= 135:
                        rows.append([ip, str(r["Asset"]), str(r["Description"])])
                except Exception:
                    continue
        except Exception as e:
            self.error.emit(str(e))
        self.data_loaded.emit(rows)

class PingWorker(QThread):
    ping_done = Signal(str, bool)
    def __init__(self, ip):
        super().__init__()
        self.ip = ip
    def run(self):
        try:
            # ping atarken cmd açılmasın
            ping_process = subprocess.Popen(
                ["ping", "-n", "1", self.ip],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                creationflags=subprocess.CREATE_NO_WINDOW
            )
            stdout, stderr = ping_process.communicate()
            online = "TTL=" in stdout.decode()
            self.ping_done.emit(self.ip, online)
        except Exception:
            self.ping_done.emit(self.ip, False)


# ---------- Main Window ----------
class VNCManager(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Flow VNC Controller")
        self.resize(1200, 700)
        icon_path = get_data_path("img/monitoring.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        self.load_thread = None
        self.ping_threads = []
        self.active_threads = []
        self.current_tree = None
        self.user_folders = {}
        self.current_theme = "dark"
        self.ping_status_cache = {}  # {"GrupAdı": {"IP": "online"/"offline"} }

        self._load_user_folders()

        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QHBoxLayout(central)
        main_layout.setContentsMargins(0, 0, 0, 0)

        # --- Left Panel ---
        left_widget = QWidget()
        self.left_layout = QVBoxLayout(left_widget)
        self.left_layout.setContentsMargins(8, 8, 8, 8)
        self.left_layout.setSpacing(8)

        self.button_group = QButtonGroup()
        self.button_group.setExclusive(True)

        # --- Logo (tek görsel) ---
        self.logo_label = QLabel()
        logo_path = get_data_path("img/flowLogo.png")  # biraz düs imzası taşımasında fayda var ^_^
        if os.path.exists(logo_path):
            pixmap = QPixmap(logo_path).scaledToWidth(120, Qt.SmoothTransformation)
            self.logo_label.setPixmap(pixmap)
            self.logo_label.setAlignment(Qt.AlignCenter)
        self.left_layout.addWidget(self.logo_label)
        self.left_layout.addSpacing(5)


        # --- Yeni Klasör ---
        new_btn = QPushButton("+ Yeni Klasör Ekle")
        new_btn.clicked.connect(self.on_add_folder)
        self.left_layout.addWidget(new_btn)
        self.left_layout.addSpacing(10)

        # --- Montaj Buttons ---
        for name in SHEET_MAP.keys():
            btn = DropButton(name, drop_callback=None)
            btn.setCheckable(True)
            btn.clicked.connect(self.create_group_handler(name, is_user=False))
            self.left_layout.addWidget(btn)
            self.button_group.addButton(btn)

        # --- User Folders ---
        for uf in self.user_folders.keys():
            self._create_and_insert_user_button(uf)

        self.left_layout.addStretch()

        # --- Alt logo (no gradient, scaled smaller for quality) ---
        bottom_logo_path = get_data_path("img/duslaci.png")
        self.bottom_logo_label = QLabel()
        if os.path.exists(bottom_logo_path):
            pix = QPixmap(bottom_logo_path)
            self.bottom_logo_label.setPixmap(pix.scaledToWidth(90, Qt.SmoothTransformation))
            self.bottom_logo_label.setAlignment(Qt.AlignCenter)
            self.left_layout.addWidget(self.bottom_logo_label)

        # --- Tema Toggle Button (sol alt köşe) ---
        self.theme_toggle_btn = QPushButton()
        self.theme_toggle_btn.setFixedSize(36, 36)
        self.theme_toggle_btn.setFlat(True)
        self.theme_toggle_btn.clicked.connect(self.toggle_theme)
        self.left_layout.addWidget(self.theme_toggle_btn, alignment=Qt.AlignLeft | Qt.AlignBottom)

        left_scroll = QScrollArea()
        left_scroll.setWidgetResizable(True)
        left_scroll.setWidget(left_widget)
        left_scroll.setMinimumWidth(220)
        main_layout.addWidget(left_scroll)

        # --- Right Panel ---
        self.right_widget = QWidget()
        self.right_layout = QVBoxLayout(self.right_widget)
        self.right_layout.setContentsMargins(12, 12, 12, 12)
        main_layout.addWidget(self.right_widget, 1)

        # --- Başlangıç Teması ---
        self.set_theme(self.current_theme)
    
    
    def _show_quick_connect_dialog(self):
        from PySide6.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QLineEdit, QPushButton

        dialog = QDialog(self)
        dialog.setWindowTitle("Hızlı Bağlantı")
        dialog.setFixedSize(300, 120)

        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        input_box = QLineEdit()
        input_box.setPlaceholderText("IP veya Asset girin...")
        layout.addWidget(input_box)

        btn_layout = QHBoxLayout()
        layout.addLayout(btn_layout)

        vnc_btn = QPushButton("UltraVNC ile Bağlan")
        vnc_btn.setFocusPolicy(Qt.NoFocus)
        cmrc_btn = QPushButton("CmRc ile Bağlan")
        cmrc_btn.setFocusPolicy(Qt.NoFocus)
        btn_layout.addWidget(vnc_btn)
        btn_layout.addWidget(cmrc_btn)

        # Butonlar tıklandığında ilgili fonksiyonları çağır
        vnc_btn.clicked.connect(lambda: (self.open_vnc(input_box.text().strip()), dialog.accept()))
        cmrc_btn.clicked.connect(lambda: (self.open_cmrc("", input_box.text().strip()), dialog.accept()))

        dialog.exec()


    # ---------- Theme ----------
    def toggle_theme(self):
        self.set_theme("light" if self.current_theme == "dark" else "dark")

    def set_theme(self, mode):
        self.current_theme = mode

        if mode == "dark":
            left_bg = "#2b2b2b"
            right_bg = "#1e1e1e"
            text_color = "white"
            btn_style = "background-color:#3a3a3a;color:white;"
            try:
                self.theme_toggle_btn.setIcon(QIcon(get_data_path("img/sun.png")))
            except Exception:
                pass
            item_sel_bg = "#3a3a3a"
            item_sel_color = "white"
        else:
            left_bg = "#f0f0f0"
            right_bg = "#f5f5f5"
            text_color = "black"
            btn_style = "background-color:#e0e0e0;color:black;"
            try:
                self.theme_toggle_btn.setIcon(QIcon(get_data_path("img/moon.png")))
            except Exception:
                pass
            item_sel_bg = "#0078d7"
            item_sel_color = "white"
            
            
        # Sol üst Flow logosu
        if hasattr(self, "logo_label") and self.logo_label:
            try:
                
                logo_file = "img/flow_dark.png" if mode == "dark" else "img/flow_light.png"
                pixmap = QPixmap(get_data_path(logo_file))
                
                # Aynı genişlikte ve orantılı yükseklik
                fixed_width = 130 # istediğin genişlikte
                pixmap = pixmap.scaled(120, 40, Qt.KeepAspectRatio, Qt.SmoothTransformation)          
                self.logo_label.setPixmap(pixmap)
                self.logo_label.setAlignment(Qt.AlignCenter)
            except Exception as e:
                print("Logo yüklenemedi:", e)


        # --- Sol panel arka plan ---
        try:
            parent = self.left_layout.parentWidget()
            if parent:
                parent.setStyleSheet(f"background-color:{left_bg};")
        except Exception:
            pass

        # --- Sol panel buton stilleri ---
        try:
            for b in self.button_group.buttons():
                b.setStyleSheet(btn_style)
        except Exception:
            pass
        # Yeni Klasör butonu ve diğer pushbutton'lar in left_layout
        try:
            for i in range(self.left_layout.count()):
                item = self.left_layout.itemAt(i)
                w = item.widget() if item else None
                if isinstance(w, QPushButton) and w is not self.theme_toggle_btn:
                    w.setStyleSheet(btn_style)
        except Exception:
            pass

        # --- Sağ panel arka plan & text ---
        try:
            self.right_widget.setStyleSheet(f"background-color:{right_bg}; color:{text_color};")
        except Exception:
            pass

        # --- Sağ paneldeki mevcut QTreeWidget güncelle (seçili satır dahil) ---
        try:
            if self.current_tree:
                tree_style = f"""
                    QTreeWidget {{
                        background-color: {right_bg};
                        color: {text_color};
                        border: none;
                    }}
                    QTreeWidget::item {{
                        min-height: 28px;
                        padding: 4px;
                        border-bottom: 1px solid {'#333' if mode == 'dark' else '#ccc'};
                    }}
                    QTreeWidget::item:selected {{
                        background-color: {item_sel_bg};
                        color: {item_sel_color};
                    }}
                    QHeaderView::section {{
                        background-color: {right_bg};
                        color: {text_color};
                        padding: 6px;
                        border-right: 1px solid {'#333' if mode == 'dark' else '#ccc'};
                        font-weight: bold;
                    }}
                """
                self.current_tree.setStyleSheet(tree_style)
                
                
                # 🌟 Mevcut item renklerini güncelle
                for i in range(self.current_tree.topLevelItemCount()):
                    item = self.current_tree.topLevelItem(i)
                    for col in range(item.columnCount()):
                        item.setForeground(col, QColor(text_color))
                
                
        except Exception:
            pass

        # --- Sağ panel içindeki QLineEdit / QLabel güncelle ---
        try:
            for i in range(self.right_layout.count()):
                w = self.right_layout.itemAt(i).widget()
                if isinstance(w, QLineEdit):
                    w.setStyleSheet(f"""
                        QLineEdit {{
                            padding:6px;
                            border-radius:4px;
                            border:1px solid {'#333' if mode=='dark' else '#999'};
                            background-color:{'#2a2a2a' if mode=='dark' else '#e0e0e0'};
                            color:{text_color};
                        }}
                        QLineEdit:focus {{ border:1px solid #555; }}
                    """)
                elif isinstance(w, QLabel):
                    w.setStyleSheet(f"color:{text_color};")
    
            # Sağ panel başlığı için ekstra garanti
            if hasattr(self, "current_title_label") and self.current_title_label:
                self.current_title_label.setStyleSheet(f"color:{text_color}; font-size:16px;")
        except Exception:
            pass



        # --- Sağ paneldeki status_counts_label güncelle ---
        try:
            if hasattr(self, "status_counts_label") and self.status_counts_label:
                online_count = 0
                offline_count = 0
            if self.current_tree:
                for i in range(self.current_tree.topLevelItemCount()):
                    item = self.current_tree.topLevelItem(i)
                    status = item.data(0, Qt.UserRole)
                    if status == "online":
                        online_count += 1
                    elif status == "offline":
                        offline_count += 1

            text_color = "white" if mode == "dark" else "black"
            self.status_counts_label.setText(
                f"<span style='color:{text_color}; font-weight:bold;'>"
                f"Çevrimiçi: <span style='color:green;'>●</span> {online_count} &nbsp;&nbsp;|&nbsp;&nbsp; "
                f"Çevrimdışı: <span style='color:red;'>●</span> {offline_count}"
                f"</span>"
            )
        except Exception:
            pass

    # ---------- User folders ----------
    def _load_user_folders(self):
        if os.path.exists(USER_FOLDER_FILE):
            try:
                with open(USER_FOLDER_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if isinstance(data, dict):
                        self.user_folders = data
                    else:
                        self.user_folders = {name: [] for name in data} if isinstance(data, list) else {}
            except Exception:
                self.user_folders = {}
        else:
            self.user_folders = {}

    def _save_user_folders(self):
        try:
            with open(USER_FOLDER_FILE, "w", encoding="utf-8") as f:
                json.dump(self.user_folders, f, ensure_ascii=False, indent=2)
        except Exception as e:
            QMessageBox.warning(self, "Hata", f"User folders kaydedilemedi:\n{e}")

    def _create_and_insert_user_button(self, folder_name):
        btn = DropButton(folder_name, drop_callback=self._on_drop_to_folder)
        btn.setCheckable(True)
        btn.clicked.connect(self.create_group_handler(folder_name, is_user=True))
        btn.setContextMenuPolicy(Qt.CustomContextMenu)
        btn.customContextMenuRequested.connect(lambda pos, b=btn: self._show_left_context(b, pos))

        # --- Montaj butonlarından sonra ekle ---
        insert_index = 0
        for i in range(self.left_layout.count()):
            w = self.left_layout.itemAt(i).widget()
            if isinstance(w, DropButton) and w.text() in SHEET_MAP.keys():
                insert_index = i + 1  # Montaj butonlarının altına ekle

        self.left_layout.insertWidget(insert_index, btn)
        self.button_group.addButton(btn)

        btn.setStyleSheet(
            "background-color:#3a3a3a;color:white;" if self.current_theme == "dark"
            else "background-color:#e0e0e0;color:black;"
    )  

    def _show_left_context(self, btn, pos):
        menu = QMenu()
        rename_action = menu.addAction("Yeniden Adlandır")
        delete_action = menu.addAction("Sil")
        
        action = menu.exec_(btn.mapToGlobal(pos))
        
        if action == delete_action:
            reply = QMessageBox.question(self, "Klasör Silme Onayı", f"'{btn.text()}' klasörünü silmek istiyor musunuz?", QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.Yes:
                name = btn.text()
                btn.setParent(None)
                if name in self.user_folders:
                    del self.user_folders[name]
                    self._save_user_folders()
                try:
                    self.button_group.removeButton(btn)
                except Exception:
                    pass
                    
                    
                    
        elif action == rename_action:
            old_name = btn.text()
            new_name, ok = QInputDialog.getText(self, "Klasör Adını Düzenle", "Yeni ad:", text=old_name)
            if ok and new_name.strip():
                new_name = new_name.strip()
                if new_name == old_name:
                    return  # klasör ismi aynıysa bir şey yapma
                if new_name in self.user_folders:
                    QMessageBox.warning(self, "Hata", f"'{new_name}' zaten mevcut.")
                    return
                    
                    
                # JSON'daki mevcut bağlantı listesini al
                connections = self.user_folders.pop(old_name, [])
                
                
                # Yeni isimle kaydet
                self.user_folders[new_name] = connections
                self._save_user_folders()
                             
                
                # Butonu güncelle
                btn.setText(new_name)
                btn.setChecked(True)
                
                
                # Sağ paneli güncellemeden önce tree'yi temizle
                if self.current_tree:
                    self.current_tree.clear()
                
                
                # Sağ paneli yeni isimle güncelle
                self.show_group_connections(new_name, is_user=True)            

    def on_add_folder(self):
        folder_name, ok = QInputDialog.getText(self, "Yeni Klasör", "Klasör adı:")
        if not ok or not folder_name.strip():
            return
        folder_name = folder_name.strip()
        if folder_name in self.user_folders:
            QMessageBox.information(self, "Bilgi", f"'{folder_name}' zaten mevcut.")
            return
        self.user_folders[folder_name] = []
        self._save_user_folders()
        self._create_and_insert_user_button(folder_name)
        QMessageBox.information(self, "Bilgi", f"'{folder_name}' klasörü eklendi.")

    def create_group_handler(self, name, is_user=False):
        def handler():
            for b in self.button_group.buttons():
                b.setChecked(b.text() == name)
            self.show_group_connections(name, is_user)
        return handler

    # ---------- Show / Load ----------
    def show_group_connections(self, group_name, is_user=False):
        # clear right panel
        for i in reversed(range(self.right_layout.count())):
            w = self.right_layout.itemAt(i).widget()
            if w:
                w.deleteLater()

        top_bar = QWidget()
        top_layout = QHBoxLayout(top_bar)
        top_layout.setContentsMargins(0, 0, 0, 0)
        title = QLabel(f"<b>{group_name}</b>")
        title.setStyleSheet(f"color: {'white' if self.current_theme == 'dark' else 'black'}; font-size:16px;")
        self.current_title_label = title 
        top_layout.addWidget(title)
        top_layout.addStretch()
        # --- Üst sayaç label (çevrimiçi/çevrimdışı sayısı) ---
        self.status_counts_label = QLabel("Çevrimiçi: 0, Çevrimdışı: 0")
        self.status_counts_label.setStyleSheet("font-weight:bold;")
        self.right_layout.addWidget(self.status_counts_label)


        if not is_user:
            refresh_btn = QPushButton("Yenile")
            refresh_btn.clicked.connect(lambda: self.load_sheet_data(group_name, show_popup=True))
            top_layout.addWidget(refresh_btn)

        self.right_layout.addWidget(top_bar)

        quick_btn = QPushButton("Hızlı Bağlantı")
        quick_btn.setFixedHeight(30)
        quick_btn.setFixedWidth(150)
        quick_btn.clicked.connect(self._show_quick_connect_dialog)

        self.right_layout.addWidget(quick_btn, alignment=Qt.AlignHCenter)



        search_box = QLineEdit()
        search_box.setPlaceholderText("Ara (IP, Asset, Açıklama)...")
        search_box.setStyleSheet(f"QLineEdit {{ padding:6px; border-radius:4px; border:1px solid #333; background-color:{'#2a2a2a' if self.current_theme=='dark' else '#e0e0e0'}; color:{'white' if self.current_theme=='dark' else 'black'}; }} QLineEdit:focus {{ border:1px solid #555; }}")
        self.right_layout.addWidget(search_box)

        tree = DragTree()
        tree.setHeaderLabels(["IP", "Asset", "Açıklama"])
        tree.setAlternatingRowColors(False)
        tree.setContextMenuPolicy(Qt.CustomContextMenu)
        tree.itemDoubleClicked.connect(lambda item, col: self.open_vnc(item.text(0), item.text(1)))  # çift tıklayınca vnc ile bağlanalım
        tree.customContextMenuRequested.connect(lambda pos, t=tree: self.show_context_menu(pos, t))
        tree.setColumnWidth(0, 180)
        tree.setColumnWidth(1, 200)
        tree.setColumnWidth(2, 350)

        # set tree style according to current theme
        tree.setStyleSheet(f"""
            QTreeWidget {{ background-color:{'#1e1e1e' if self.current_theme=='dark' else '#f5f5f5'}; color:{'white' if self.current_theme=='dark' else 'black'}; border:none; }}
            QTreeWidget::item {{ min-height:28px; padding:4px; border-bottom:1px solid {'#333' if self.current_theme=='dark' else '#ccc'}; }}
            QTreeWidget::item:selected {{ background-color: {'#3a3a3a' if self.current_theme=='dark' else '#0078d7'}; color: white; }}
            QHeaderView::section {{ background-color:{'#1e1e1e' if self.current_theme=='dark' else '#f5f5f5'}; color:{'white' if self.current_theme=='dark' else 'black'}; padding:6px; border-right:1px solid {'#333' if self.current_theme=='dark' else '#ccc'}; font-weight:bold; }}
        """)

        for i in range(tree.columnCount()):
            tree.header().setDefaultAlignment(Qt.AlignCenter)

        tree_scroll = QScrollArea()
        tree_scroll.setWidgetResizable(True)
        tree_scroll.setWidget(tree)
        self.right_layout.addWidget(tree_scroll)
        self.current_tree = tree

        if is_user:
            entries = self.user_folders.get(group_name, [])
            
            # IP Sıralaması
            def ip_key(item):
                ip = item[0]
                try:
                    return tuple(int(part) for part in ip.split("."))
                except:
                    return (0,0,0,0)
               
            entries = sorted(entries, key=ip_key)
              
            for r in entries:
                tree.addTopLevelItem(QTreeWidgetItem([r[0], r[1], r[2]]))
            search_box.textChanged.connect(lambda txt: self.filter_tree(tree, txt))
            return

        self.load_sheet_data(group_name, search_box)

    def load_sheet_data(self, group_name, search_box=None, show_popup=False):
        octet = OCTET_MAP.get(group_name, "")
        sheet_name = SHEET_MAP.get(group_name, "")
        if sheet_name:
            # --- Yükleniyor mesajı ekle ---
            loading_label = QLabel("📄 Veriler yükleniyor, lütfen bekleyin...")
            loading_label.setAlignment(Qt.AlignCenter)
            self.right_layout.addWidget(loading_label)
            QApplication.processEvents()

            # --- Thread başlat ---
            self.load_thread = LoadSheetWorker(sheet_name, octet)

            # 🔹 Thread'i aktif listeye ekle (kaybolmasın)
            self.active_threads.append(self.load_thread)

            # 🔹 Hata sinyali
            self.load_thread.error.connect(lambda msg: QMessageBox.critical(self, "Excel Hatası", msg))

            # 🔹 Veri yüklendiğinde GUI'yi güncelle (thread-safe)
            self.load_thread.data_loaded.connect(
                lambda rows, t=self.current_tree, sb=search_box, sp=show_popup, lbl=loading_label, th=self.load_thread:
                self._on_rows_loaded_threadsafe(t, rows, sb, sp, lbl, th)
            )


            # 🔹 Thread bitince temizle
            self.load_thread.finished.connect(lambda th=self.load_thread: self.cleanup_thread(th))

            # 🔹 Thread'i başlat
            self.load_thread.start()




    def _on_rows_loaded_threadsafe(self, tree, rows, search_box, show_popup, loading_label, thread=None):
        # loading_label silme
        if loading_label:
            try:
                loading_label.deleteLater()
            except RuntimeError:
                pass

        # GUI güncellemesini try/except ile güvenli şekilde çalıştır
        QTimer.singleShot(0, lambda: self._safe_on_rows_loaded(tree, rows, search_box, show_popup))



    def _safe_on_rows_loaded(self, tree, rows, search_box, show_popup):
        try:
            self._on_rows_loaded(tree, rows, search_box, show_popup)
        except RuntimeError:
            # tree silinmiş, GUI güncellemesi atlanıyor
            print("GUI update skipped: DragTree already deleted.")

    
    def update_status_counts(self):
        online_count = 0
        offline_count = 0

        tree = self.current_tree
        if not tree:
            return

        for i in range(tree.topLevelItemCount()):
            item = tree.topLevelItem(i)
            status = item.data(0, Qt.UserRole)
            if status == "online":
                online_count += 1
            elif status == "offline":
                offline_count += 1
            # unknown ise sayma

        if hasattr(self, "status_counts_label") and self.status_counts_label:
            # Tema rengi
            text_color = "white" if self.current_theme == "dark" else "black"
            self.status_counts_label.setText(
                f"<span style='color:{text_color}; font-weight:bold;'>"
                f"Çevrimiçi: <span style='color:green;'>●</span> {online_count} &nbsp;&nbsp;|&nbsp;&nbsp; "
                f"Çevrimdışı: <span style='color:red;'>●</span> {offline_count}"
                f"</span>"
            )


    
    
    def cleanup_thread(self, thread):
        try:
            if thread in self.active_threads:
                self.active_threads.remove(thread)
            thread.deleteLater()
        except Exception as e:
            print("Thread cleanup error:", e)
    
    
    
    
    def _on_rows_loaded(self, tree, rows, search_box, show_popup):
        if tree is None or not isinstance(tree, QTreeWidget):
            return
        tree.clear()

        online_count = 0
        offline_count = 0

        # 🔹 Geçerli sayfanın grup adını bul (örneğin "Montaj 1", "Montaj 2")
        parent_group = None
        for b in self.button_group.buttons():
            if b.isChecked():
                parent_group = b.text()
            break

        # 🔹 Cache'den önceki ping durumlarını yükle (varsa)
        cached_status = {}
        cache_available = False
        if parent_group and hasattr(self, "ping_status_cache"):
            cached_status = self.ping_status_cache.get(parent_group, {})
            if cached_status:
                cache_available = True  # sadece cache doluysa eski durumları uygula

        for r in rows:
            item = QTreeWidgetItem([r[0], r[1], r[2]])
            hostname = r[0]

            # 🔹 Tema rengi
            base_color = QColor("white") if self.current_theme == "dark" else QColor("black")

            # 🔹 Eğer cache varsa, durumu uygula
            if cache_available:
                status = cached_status.get(hostname, "unknown")
            else:
                status = "unknown"  # ilk açılışta bilinmeyen durum

            item.setData(0, Qt.UserRole, status)

            # 🔹 Renkleri duruma göre ata
            if status == "online":
                item.setForeground(0, QColor("green"))
                online_count += 1
            elif status == "offline":
                item.setForeground(0, QColor("red"))
                offline_count += 1
            else:
                item.setForeground(0, base_color)  # unknown: normal renk

            # Diğer kolonlar (Asset, Açıklama)
            for col in range(1, item.columnCount()):
                item.setForeground(col, base_color)

            tree.addTopLevelItem(item)

        # 🔹 Sayaç
        if hasattr(self, "status_counts_label") and self.status_counts_label:
            if cache_available:
                # Cache varsa eski değerleri göster
                self.status_counts_label.setText(
                    f"<span style='color:{'white' if self.current_theme=='dark' else 'black'}; font-weight:bold;'>"
                    f"Çevrimiçi: <span style='color:green;'>●</span> {online_count} &nbsp;&nbsp;|&nbsp;&nbsp; "
                    f"Çevrimdışı: <span style='color:red;'>●</span> {offline_count}</span>"
                )
            else:
                # Cache yoksa sıfırdan başlat
                self.status_counts_label.setText(
                    f"<span style='color:{'white' if self.current_theme=='dark' else 'black'}; font-weight:bold;'>"
                    f"Çevrimiçi: <span style='color:green;'>●</span> 0 &nbsp;&nbsp;|&nbsp;&nbsp; "
                    f"Çevrimdışı: <span style='color:red;'>●</span> 0</span>"
                )

        # 🔹 Arama kutusu bağlantısı
        if search_box:
            try:
                search_box.textChanged.disconnect()
            except Exception:
                pass
            search_box.textChanged.connect(lambda txt: self.filter_tree(tree, txt))

        # 🔹 Yenileme mesajı
        if show_popup:
            QMessageBox.information(self, "Excel Yenilendi", "Excel verisi başarıyla yüklendi ve güncellendi.")





    def filter_tree(self, tree, text):
        text = text.lower()
        for i in range(tree.topLevelItemCount()):
            item = tree.topLevelItem(i)
            visible = any(text in (item.text(col) or "").lower() for col in range(tree.columnCount()))
            item.setHidden(not visible)

    # ---------- Context menu ----------
    def show_context_menu(self, pos, tree):
        item = tree.itemAt(pos)
        menu = QMenu()
        if item:
            ping_action = menu.addAction("Ping At")
            vnc_action = menu.addAction("VNC ile Bağlan")
            cmrc_action = menu.addAction("CMRC ile Bağlan") 
            delete_action = menu.addAction("Sil")
            
            action = menu.exec_(tree.mapToGlobal(pos))
            ip = item.text(0)
            
            
            if action == ping_action:
                self.run_ping(item)
                
                
            elif action == vnc_action:
                self.open_vnc(item.text(0), item.text(1))
                
                
            elif action == cmrc_action:
                self.open_cmrc(item.text(0), item.text(1))
            
            
            elif action == delete_action:
                parent_folder = None
                for b in self.button_group.buttons():
                    if b.isChecked():
                        parent_folder = b.text()
                        break
                idx = tree.indexOfTopLevelItem(item)
                tree.takeTopLevelItem(idx)
                if parent_folder in self.user_folders:
                    arr = self.user_folders.get(parent_folder, [])
                    ip0, asset0, desc0 = item.text(0), item.text(1), item.text(2)
                    newarr = [e for e in arr if not (e[0] == ip0 and e[1] == asset0 and e[2] == desc0)]
                    self.user_folders[parent_folder] = newarr
                    self._save_user_folders()
        else:
            add_action = menu.addAction("Bağlantı Ekle")
            action = menu.exec_(tree.mapToGlobal(pos))
            if action == add_action:
                self.add_connection(tree)

    # ---------- Drop handling ----------
    def _on_drop_to_folder(self, folder_name, ip, asset, desc):
        if not ip:
            return
        arr = self.user_folders.get(folder_name, [])
        if any(e[0] == ip and e[1] == asset and e[2] == desc for e in arr):
            QMessageBox.information(self, "Bilgi", f"{ip} zaten '{folder_name}' klasöründe.")
            return
        arr.append([ip, asset, desc])
        self.user_folders[folder_name] = arr
        self._save_user_folders()
        QMessageBox.information(self, "Bilgi", f"{ip} eklendi -> {folder_name}")
        for b in self.button_group.buttons():
            if b.text() == folder_name and b.isChecked():
                self.show_group_connections(folder_name, is_user=True)
                break

    # ---------- Ping / VNC / Add ----------
    def run_ping(self, item):
        ip = item.text(0)
        ping_thread = PingWorker(ip)
        ping_thread.ping_done.connect(lambda ip, online: self._on_ping_done(item, online))
        ping_thread.start()
        self.ping_threads.append(ping_thread)

    def _on_ping_done(self, item, online):
        # ---------- Ping sonucunu cache'e kaydet ----------
        parent_group = None
        for b in self.button_group.buttons():
            if b.isChecked():
                parent_group = b.text()
                break

        if parent_group:
            if parent_group not in self.ping_status_cache:
                self.ping_status_cache[parent_group] = {}
            self.ping_status_cache[parent_group][item.text(0)] = "online" if online else "offline"
        
        
        
        # Durum item'a kaydet
        item.setData(0, Qt.UserRole, "online" if online else "offline")  

        # Sayaçları güncelle
        self.update_status_counts()

        msg = f"{item.text(0)} {'çevrimiçi ✅' if online else 'çevrimdışı ❌'}"
        if online:
            QMessageBox.information(self, "Ping", msg)
        else:
            QMessageBox.warning(self, "Ping", msg)


    def add_connection(self, tree):
        ip, ok = QInputDialog.getText(self, "IP Adresi", "IP (boş bırakılabilir):")
        if not ok:
            return
        ip = ip.strip()

        asset, ok = QInputDialog.getText(self, "Asset / Hostname", "Asset / Hostname:")
        if not ok or not asset.strip():
            QMessageBox.warning(self, "Hata", "IP veya Asset / Hostname girilmelidir!")
            return
        asset = asset.strip()

        desc, ok = QInputDialog.getText(self, "Açıklama", "Açıklama:")
        if not ok:
            return
        desc = desc.strip()

        tree.addTopLevelItem(QTreeWidgetItem([ip, asset, desc]))

        # --- JSON'a otomatik ekle ---
        parent_folder = None
        for b in self.button_group.buttons():
            if b.isChecked():
                parent_folder = b.text()
                break

        if parent_folder in self.user_folders:
            arr = self.user_folders.get(parent_folder, [])
            # Aynı kayıt varsa ekleme
            if not any(e[0] == ip and e[1] == asset and e[2] == desc for e in arr):
                arr.append([ip, asset, desc])
                self.user_folders[parent_folder] = arr
                self._save_user_folders()

        # --- Panel güncelle ---
        if parent_folder:
            self.show_group_connections(parent_folder, is_user=True)


    def open_vnc(self, ip, asset=""):
        try:
            target = ip if ip else asset  # IP yoksa Asset / hostname kullan
            if not target:
                raise ValueError("Bağlanacak IP veya hostname yok.")

            vnc_paths_to_try = [r"C:\Program Files\uvnc bvba\UltraVNC\vncviewer.exe"]
            vnc_path = None
            for p in vnc_paths_to_try:
                if os.path.exists(p):
                    vnc_path = p
                    break
            if vnc_path is None:
                raise FileNotFoundError("UltraVNC vncviewer.exe bulunamadı.")

            subprocess.Popen([vnc_path, target, "-password", VNC_PASSWORD])

        except Exception as e:
            QMessageBox.critical(self, "VNC Hatası", f"VNC açılamadı:\n{e}")
            
            
    def open_cmrc(self, ip, asset):
        import os
        import subprocess
        from PySide6.QtWidgets import QMessageBox

        try:
            # CMRC uygulamasının yolu
            cmrc_path = r"D:\Work\PLM\Sccm Remote Control Tool -1906\CmRcViewer.exe"

            if not os.path.exists(cmrc_path):
                QMessageBox.warning(self, "Hata", f"CMRC bulunamadı:\n{cmrc_path}")
                return

            # Bilgisayar adı varsa onu kullan, yoksa IP
            target = asset if asset else ip

            # CMRC başlat
            subprocess.Popen([cmrc_path, target], creationflags=subprocess.CREATE_NO_WINDOW)

        except Exception as e:
            QMessageBox.warning(self, "Hata", f"CMRC başlatılamadı:\n{e}")

# ---------- Application ----------
if __name__ == "__main__":
    import multiprocessing
    multiprocessing.freeze_support()
    app = QApplication.instance() or QApplication(sys.argv)
    app.setStyle("Fusion")
    window = VNCManager()
    window.show()
    sys.exit(app.exec())
