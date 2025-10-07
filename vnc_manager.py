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
from PySide6.QtCore import Qt, QThread, Signal, QMimeData
from PySide6.QtGui import QColor, QBrush, QPixmap, QIcon

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
    try:
        base_path = sys._MEIPASS  # exe’den çalışırken geçici yol
    except AttributeError:
        base_path = os.path.dirname(__file__)  # normal Python
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
            df = pd.read_excel(EXCEL_PATH, sheet_name=self.sheet_name, header=None, dtype=str)
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
            result = subprocess.run(["ping", "-n", "1", self.ip], capture_output=True, text=True)
            online = "TTL=" in result.stdout
            self.ping_done.emit(self.ip, online)
        except Exception:
            self.ping_done.emit(self.ip, False)


# ---------- Main Window ----------
class VNCManager(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("VNC Manager by ABŞ")
        self.resize(1200, 700)
        icon_path = get_data_path("img/monitoring.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        self.load_thread = None
        self.ping_threads = []
        self.current_tree = None
        self.user_folders = {}
        self.current_theme = "dark"

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
        new_btn = QPushButton("+ Yeni Klasör")
        new_btn.clicked.connect(self.on_add_folder)
        self.left_layout.addWidget(new_btn)
        self.left_layout.addSpacing(5)

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


    def _show_left_context(self, btn, pos):
        menu = QMenu()
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
        top_layout.addWidget(title)
        top_layout.addStretch()

        if not is_user:
            refresh_btn = QPushButton("Yenile")
            refresh_btn.clicked.connect(lambda: self.load_sheet_data(group_name, show_popup=True))
            top_layout.addWidget(refresh_btn)

        self.right_layout.addWidget(top_bar)

        search_box = QLineEdit()
        search_box.setPlaceholderText("Ara (IP, Asset, Açıklama)...")
        search_box.setStyleSheet(f"QLineEdit {{ padding:6px; border-radius:4px; border:1px solid #333; background-color:{'#2a2a2a' if self.current_theme=='dark' else '#e0e0e0'}; color:{'white' if self.current_theme=='dark' else 'black'}; }} QLineEdit:focus {{ border:1px solid #555; }}")
        self.right_layout.addWidget(search_box)

        tree = DragTree()
        tree.setHeaderLabels(["IP", "Asset", "Açıklama"])
        tree.setAlternatingRowColors(False)
        tree.setContextMenuPolicy(Qt.CustomContextMenu)
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
            for r in entries:
                tree.addTopLevelItem(QTreeWidgetItem([r[0], r[1], r[2]]))
            search_box.textChanged.connect(lambda txt: self.filter_tree(tree, txt))
            return

        self.load_sheet_data(group_name, search_box)

    def load_sheet_data(self, group_name, search_box=None, show_popup=False):
        octet = OCTET_MAP.get(group_name, "")
        sheet_name = SHEET_MAP.get(group_name, "")
        if sheet_name:
            self.load_thread = LoadSheetWorker(sheet_name, octet)
            self.load_thread.error.connect(lambda msg: QMessageBox.critical(self, "Excel Hatası", msg))
            self.load_thread.data_loaded.connect(lambda rows, t=self.current_tree, sb=search_box, sp=show_popup: self._on_rows_loaded(t, rows, sb, sp))
            self.load_thread.start()

    def _on_rows_loaded(self, tree, rows, search_box, show_popup):
        if tree is None or not isinstance(tree, QTreeWidget):
            return
        tree.clear()
        for r in rows:
            tree.addTopLevelItem(QTreeWidgetItem([r[0], r[1], r[2]]))
        if search_box:
            try:
                search_box.textChanged.disconnect()
            except Exception:
                pass
            search_box.textChanged.connect(lambda txt: self.filter_tree(tree, txt))
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
            delete_action = menu.addAction("Sil")
            action = menu.exec_(tree.mapToGlobal(pos))
            ip = item.text(0)
            if action == ping_action:
                self.run_ping(item)
            elif action == vnc_action:
                self.open_vnc(item.text(0), item.text(1))
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
        color = QBrush(QColor("green") if online else QColor("red"))
        for col in range(item.columnCount()):
            item.setBackground(col, color)
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


# ---------- Application ----------
if __name__ == "__main__":
    app = QApplication.instance() or QApplication(sys.argv)
    app.setStyle("Fusion")
    window = VNCManager()
    window.show()
    sys.exit(app.exec())
