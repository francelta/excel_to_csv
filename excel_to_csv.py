import sys
import os
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog, QLabel, QSizePolicy
from PyQt5.Qt import Qt, QColor
from PyQt5.QtCore import Qt as QtCoreQt
import pandas as pd
from PyQt5.QtGui import QFont

class MyApp(QWidget):

    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):
        self.setWindowTitle('CONVERSOR DE EXCEL A CSV')
        self.setAcceptDrops(True)
        self.setStyleSheet("background-color: white;")

        self.btn = QPushButton('Pincha o arrastra archivos excel', self)
        self.btn.clicked.connect(self.cargar_xlsx)
        self.btn.setStyleSheet("background-color: fuchsia; color: white; font-size: 14px;")

        # Set the size policy of the button to expand in both directions and set its minimum height.
        self.btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.btn.setMinimumHeight(int(self.height() / 4))

        self.label = QLabel('', self)

        vbox = QVBoxLayout()
        vbox.addWidget(self.btn)
        vbox.addWidget(self.label)

        screen = QApplication.primaryScreen()
        screen_geometry = screen.geometry()
        self.setGeometry(0, 0, int(screen_geometry.width() * 0.25), int(screen_geometry.width() * 0.25))

        self.setLayout(vbox)

        # Center the window on the screen.
        qr = self.frameGeometry()
        cp = screen_geometry.center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

        self.show()

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.accept()
        else:
            e.ignore()

    def dropEvent(self, e):
        file_path = str(e.mimeData().urls()[0].toLocalFile())
        self.convierte_xlsx_a_csv(file_path)

    def cargar_xlsx(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        fileName, _ = QFileDialog.getOpenFileName(self,"Cargar XLSX", "","Excel Files (*.xlsx)", options=options)
        if fileName:
            self.convierte_xlsx_a_csv(fileName)

    def convierte_xlsx_a_csv(self, file_path):
        try:
            df = pd.read_excel(file_path)
            base = os.path.splitext(file_path)[0]
            csv_file_path = base + '.csv'
            df.to_csv(csv_file_path, index=False)
            self.label.setText('Conversión exitosa. Archivo guardado como: ' + csv_file_path)
            os.system(f'open -R {csv_file_path}')
        except Exception as e:
            self.label.setText('Error en la conversión: ' + str(e))

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyApp()
    sys.exit(app.exec_())
