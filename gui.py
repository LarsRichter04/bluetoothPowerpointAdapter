import sys
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import (
    QWidget,
    QLabel,
    QVBoxLayout,
    QTextEdit,
    QApplication,
    QPushButton,
)

from enums import Answers, Answers_Pepper


class EmittingStream:
    def eprint(*args, **kwargs):
        print(*args, file=sys.stderr, **kwargs)

    def __init__(self, text_edit_widget):
        self.text_edit_widget = text_edit_widget

    def write(self, text):
        self.text_edit_widget.append(text)

    def flush(self):
        pass


class Window(QWidget):
    def __init__(self, version, parent=None):
        super(Window, self).__init__(parent)
        self.resize(400, 300)
        self.setWindowTitle("Bluetooth Powerpoint Adapter " + version)

        layout = QVBoxLayout(self)

        self.label = QLabel(self)
        self.label.setText("Bluetooth Powerpoint Adapter " + version)
        font = QFont()
        font.setFamily("Arial")
        font.setPointSize(16)
        self.label.setFont(font)
        layout.addWidget(self.label)

        self.text_edit = QTextEdit(self)
        self.text_edit.setReadOnly(True)
        layout.addWidget(self.text_edit)

        self.pause_button = QPushButton("Pause", self)
        layout.addWidget(self.pause_button)
        self.pause_button.clicked.connect(self.on_pause_clicked)

        self.bt_client_socket = None  # Socket-Referenz für Bluetooth-Verbindung

        sys.stdout = EmittingStream(self.text_edit)

    def set_bt_client_socket(self, client_socket):
        self.bt_client_socket = client_socket

    def on_pause_clicked(self):
        if self.bt_client_socket:
            try:
                self.bt_client_socket.send(b"PAUSE")
                self.text_edit.append("PAUSE gesendet!")
            except Exception as e:
                self.text_edit.append(f"Fehler beim Senden: {e}")
        else:
            self.text_edit.append("Keine Bluetooth-Verbindung aktiv!")

    def closeEvent(self, event):
        sys.stdout = sys.__stdout__
        super().closeEvent(event)
