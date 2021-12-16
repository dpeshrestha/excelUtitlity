from PySide2.QtCore import Qt
from PySide2.QtWidgets import QDialog, QVBoxLayout, QLabel, QPushButton
import src.settings as s

class customQMessageBox(QDialog):
    def __init__(self,text=''):
        super().__init__(parent=s.app,f=Qt.WindowSystemMenuHint|Qt.WindowTitleHint)
        self.resize(250, 120)
        self.layout = QVBoxLayout()
        self.text = QLabel()
        self.text.setText("Message")
        self.text.setWordWrap(True)
        self.layout.addWidget(self.text,alignment=Qt.AlignCenter)
        self.button = QPushButton('OK')
        self.button.clicked.connect(self.okOptions)
        self.layout.addWidget(self.button,alignment=Qt.AlignCenter)
        self.setLayout(self.layout)
        self.setWindowTitle("Warning")
        self.setText(text)
    def okOptions(self):
        self.optionsOK = True
        self.close()

    def setText(self,text):
        self.text.setText(text)