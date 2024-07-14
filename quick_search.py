from pathlib import Path
from PySide6.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QRadioButton, QButtonGroup, QLineEdit
from PySide6.QtCore import Qt, QSettings, Slot, Signal, QObject, QThread, QTimer, QPoint, QRect, QSize
from PySide6.QtGui import QFont, QShortcut, QKeySequence, QFontMetrics
from dabbler.gui_table import DfView
import polars as pl

settings = QSettings("quick_search_python3", "quick_search_python3")

def create_shortcut(key, function, parent):
    shortcut = QShortcut(QKeySequence(key), parent)
    shortcut.activated.connect(function)
    return shortcut


class FileSelector(QRadioButton):
    def __init__(self, file:Path, parent=None):
        super(FileSelector, self).__init__(file.name.rstrip(file.suffix), parent)
        self._file = file


class Worker(QObject):
    provideData = Signal(pl.DataFrame)
    def __init__(self, parent = None) -> None:
        super().__init__(parent)
        self._data = [{}]

    @Slot(Path)
    def load_file(self, file:Path):
        df = pl.read_json(file)
        self.provideData.emit(df)


#MARK: Main Window
class MainWindow(QWidget):

    requestData = Signal(Path)

    def __init__(self, app:QApplication):
        super(MainWindow, self).__init__()
        self.setWindowTitle("New_Quick_Search")
        self.save_widths = True
        self.app = app
        self._rows = 0
        self.selected_file:Path = None
        self._data_folders:list[Path] = [Path(__file__).parent.joinpath("data")]
        self._layout = QVBoxLayout(self)
        self._layout.setContentsMargins(0,0,0,0)
        self._layout.setSpacing(0)
        self._radio_layout = QHBoxLayout()
        self._radio_layout.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self._layout.addLayout(self._radio_layout)

        self._font_size = int(settings.value("font_size", 12))

        self.load_files()

        self.set_up_table()
        self.set_up_worker()
        self.set_up_radio_buttons()
        # self.setup_bottom_status()
        self.save_geometry_settings_timer = QTimer()
        self.save_geometry_settings_timer.timeout.connect(self.save_geometry_settings)
        self.save_geometry_settings_timer.setSingleShot(True)

        self.setup_shortcuts()

        last_file = settings.value("selected_file", None)
        if last_file and last_file in self.buttons:
            self.buttons[last_file].setChecked(True)
        else:
            self.buttons[list(self.files.keys())[0]].setChecked(True)


        self.set_inital_position()
        self.set_font()


    def setup_bottom_status(self):
        self._status_layout = QHBoxLayout()
        self._status_layout.setContentsMargins(5,0,0,5)
        self._layout.addLayout(self._status_layout)
        self._status_label = QLabel()
        self._status_layout.addWidget(self._status_label)
        self._status_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)


    def set_up_worker(self):
        self._worker = Worker()
        self._worker.provideData.connect(self.set_data)        
        self.requestData.connect(self._worker.load_file)
        self._worker_thread = QThread()
        self._worker.moveToThread(self._worker_thread)
        self._worker_thread.start()


    @Slot(pl.DataFrame)
    def set_data(self, data:list):
        self._table.set_df(data)

    
    def set_up_table(self):
        self._table = DfView(self, self.app)
        self._layout.addWidget(self._table)


    def set_font(self):
        font = QFont()
        font.setPointSize(self._font_size)
        self.fm = QFontMetrics(font)
        self._table.set_font_size(self._font_size)
        # self._status_label.setFont(font)
        self._table.update_col_widths()
        for b in self._button_group.buttons():
            b.setFont(font)


    def change_font_size(self, size:int):
        self._font_size = size
        self.set_font()
        settings.setValue("font_size", size)


    def setup_shortcuts(self):
        self.shortcuts = [
            create_shortcut("Ctrl+=", lambda: self.change_font_size(self._font_size + 1), self),
            create_shortcut("Ctrl+-", lambda: self.change_font_size(self._font_size - 1), self),
            create_shortcut("Ctrl+Q", self.close, self),
            create_shortcut("Ctrl+PgDown", lambda: self.select_next_prev_file(1), self),
            create_shortcut("Ctrl+PgUp", lambda: self.select_next_prev_file(-1), self),
            # create_shortcut("Ctrl+c", self._table.copy_values, self),
            # create_shortcut("Ctrl+Shift+c", lambda: self._table.copy_values(fmt=True), self),

        ]
    
    def select_next_prev_file(self, direction:int):
        cur_text = self._button_group.checkedButton().text()
        buttons = self._button_group.buttons()
        keys = [b.text() for b in buttons]
        
        current = keys.index(cur_text)

        if current == -1:
            self.buttons[keys[0]].setChecked(True)
        else:
            next = min(max((current + direction),0), len(keys)-1)
            if next == current:
                return
            buttons[next].setChecked(True)


    def load_files(self):
        self.files:dict[str,Path] = {}
        for folder in self._data_folders:
            for file in folder.glob("*.json"):
                self.files[file.name] = file

    def set_up_search(self):
        self._search_layout = QHBoxLayout()
        self._layout.addLayout(self._search_layout)

        self._search_input = QLineEdit()
        self._search_input.setPlaceholderText("Search...")
        self._search_layout.addWidget(self._search_input)

    def set_up_radio_buttons(self):
        self._button_group = QButtonGroup()
        self.buttons:dict[str,FileSelector] = {}
        for name, file in self.files.items():
            radio_button = FileSelector(file)
            self._radio_layout.addWidget(radio_button)
            self._button_group.addButton(radio_button)
            self.buttons[file.name] = radio_button
        self._button_group.buttonToggled.connect(self.requset_data)

    def select_file(self, file:Path):
        self.selected_file:Path = file
        settings.setValue("selected_file", file.name)
        self.requestData.emit(file)

    def requset_data(self, button:FileSelector, checked:bool):
        if checked:
            self.select_file(button._file)

    def set_inital_position(self):
        saved_pos = settings.value("pos", QPoint(100, 100))
        saved_size = settings.value("size", QSize(1200, 800))
        saved_geometry = QRect(saved_pos, saved_size)
        screens = QApplication.screens()

        if not any([s.geometry().contains(saved_geometry) for s in screens]):
            self.setGeometry(100,100,1200,800)
            self.save_geometry_settings()
        else:
            self.move(saved_pos)
            self.resize(saved_size)

    def save_geometry_settings(self):
        settings.setValue("size", self.size())
        settings.setValue("pos", self.pos())

    def resizeEvent(self, event):
        self.save_geometry_settings_timer.start(2000)

    def moveEvent(self, event):
        self.save_geometry_settings_timer.start(2000)


    def closeEvent(self, event):
        self._worker_thread.quit()
        self._worker_thread.wait()
        self._table.workerThread.quit()
        self._table.workerThread.wait()
        self._table
        event.accept()


app = QApplication([])
app.setStyle("Fusion")
window = MainWindow(app)
window.show()
app.exec_()