import sys
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import QWidget, QLabel, QVBoxLayout, QTextEdit, QApplication

from enums import Answers, Answers_Pepper


class EmittingStream:
    def eprint(*args, **kwargs):
        print(*args, file=sys.stderr, **kwargs)
    def __init__(self, text_edit_widget):
        self.text_edit_widget = text_edit_widget

    def write(self, text):
        match(text):
            case "b'0:1'":
                self.text_edit_widget.append("Opened PPP")
            case "b'\\x01'":
                self.text_edit_widget.append("Recieved Next Slide")
            case _:
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

        sys.stdout = EmittingStream(self.text_edit)

    def closeEvent(self, event):
        sys.stdout = sys.__stdout__
        super().closeEvent(event)
