import sys
import os
import csv
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QFileDialog, QLineEdit, QMessageBox

msg = QMessageBox()
msg.setWindowTitle("Operação Concluída!")
msg.setText("Os arquivos foram mesclados e salvos.")

class MesclagemWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Mesclar Arquivos CSV")
        self.setGeometry(822, 281, 300, 100)

        self.setStyleSheet("background-color: black")

        self.layout = QVBoxLayout()

        self.label = QLineEdit()
        self.label.setPlaceholderText("Pasta de Origem:")
        self.layout.addWidget(self.label)
        self.label.setStyleSheet("color:white")

        style = """
            font-size: 18px;
            height: 30px;
            background-color: black;
            color: orange;
            border: 0.5px solid orange;
            position: relative;
            margin-top: 4px;
            margin-bottom: 4px;
            border-radius: 10px;
        """

        pressed_style = """
            font-size: 18px;
            height: 30px;
            background-color: orange;  /* Cor de fundo quando pressionado */
            color: black;  /* Cor do texto quando pressionado */
            border: 0.5px solid orange;  /* Borda quando pressionado */
            position: relative;
            margin-top: 4px;
            margin-bottom: 4px;
            border-radius: 10px;
        """

        self.select_button = QPushButton("Selecionar Pasta de Origem")
        self.select_button.clicked.connect(self.selecionar_pasta_origem)
        self.layout.addWidget(self.select_button)
        self.select_button.setStyleSheet(style)
        self.select_button.pressed.connect(lambda: self.select_button.setStyleSheet(pressed_style))
        self.select_button.released.connect(lambda: self.select_button.setStyleSheet(style))

        self.mesclar_button = QPushButton("Mesclar Arquivos")
        self.mesclar_button.clicked.connect(self.mesclar_csv)
        self.layout.addWidget(self.mesclar_button)
        self.mesclar_button.setStyleSheet(style)
        self.mesclar_button.pressed.connect(lambda: self.mesclar_button.setStyleSheet(pressed_style))
        self.mesclar_button.released.connect(lambda: self.mesclar_button.setStyleSheet(style))

        self.diretorio_origem = ""

        self.setLayout(self.layout)

    def selecionar_pasta_origem(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        folder_dialog = QFileDialog.getExistingDirectory(self, "Selecione a Pasta de Origem", "", options=options)
        if folder_dialog:
            self.diretorio_origem = folder_dialog
            self.label.setText(f"Pasta de Origem: {self.diretorio_origem}")

    def mesclar_csv(self):
        if not self.diretorio_origem:
            return

        arquivos_csv = []

        for arquivo in os.listdir(self.diretorio_origem):
            if arquivo.endswith('.csv'):
                caminho_arquivo = os.path.join(self.diretorio_origem, arquivo)
                arquivos_csv.append(caminho_arquivo)

        saida_mesclada = []

        for arquivo_csv in arquivos_csv:
            with open(arquivo_csv, 'r', encoding='latin1') as csv_file:
                csv_reader = csv.reader(csv_file, delimiter=';')
                for linha in csv_reader:
                    saida_mesclada.append(linha)

        arquivo_saida = os.path.join(self.diretorio_origem, '- - - Arquivo Mesclado - - -.csv')
        with open(arquivo_saida, 'w', newline='', encoding='latin1') as csv_saida:
            csv_writer = csv.writer(csv_saida, delimiter=';')
            csv_writer.writerows(saida_mesclada)
            msg = QMessageBox()
            msg.setWindowTitle("Operação Concluída!")
            msg.setText("Os arquivos foram mesclados e salvos.")
            msg.exec_()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MesclagemWindow()
    window.show()
    sys.exit(app.exec_())
