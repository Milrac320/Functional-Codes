import sys
import PyPDF2
import io
import os
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QFileDialog, QLineEdit


class PDFSelector(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Inverter Páginas PDF')
        self.setGeometry(822, 281, 300, 150)
        self.setStyleSheet("background-color: black; color: white;")

        self.layout = QVBoxLayout()

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

        self.label = QLineEdit()
        self.label.setPlaceholderText("Arquivo PDF:")
        self.layout.addWidget(self.label)

        self.select_button = QPushButton("Selecionar Arquivo PDF")
        self.select_button.clicked.connect(self.select_pdf)
        self.layout.addWidget(self.select_button)
        self.select_button.setStyleSheet(style)

        self.mesclar_button = QPushButton("Inverter Arquivo")
        self.mesclar_button.clicked.connect(self.inverter_paginas)
        self.layout.addWidget(self.mesclar_button)
        self.mesclar_button.setStyleSheet(style)
        self.mesclar_button.pressed.connect(lambda: self.mesclar_button.setStyleSheet(pressed_style))
        self.mesclar_button.released.connect(lambda: self.mesclar_button.setStyleSheet(style))

        self.setLayout(self.layout)

    def select_pdf(self):
        options = QFileDialog.Options()
        file_name, _ = QFileDialog.getOpenFileName(self, 'Selecionar PDF', '', 'Arquivos PDF (*.pdf)', options=options)
        if file_name:
            self.label.setText(f'Arquivo selecionado: {file_name}')
            self.pdf_path = file_name

    def inverter_paginas(self):
        if hasattr(self, 'pdf_path'):
            with open(self.pdf_path, 'rb') as arquivo_entrada:
                pdf_data = arquivo_entrada.read()

            pdf_stream = io.BytesIO(pdf_data)
            pdf_reader = PyPDF2.PdfReader(pdf_stream)
            paginas = pdf_reader.pages
            paginas_invertidas = paginas[::-1]

            nome_arquivo_original, extensao = os.path.splitext(os.path.basename(self.pdf_path))
            nome_arquivo_invertido = f'{nome_arquivo_original}_Invertido{extensao}'
            caminho_arquivo_invertido = os.path.join(os.path.dirname(self.pdf_path), nome_arquivo_invertido)

            with open(caminho_arquivo_invertido, 'wb') as arquivo_saida:
                pdf_writer = PyPDF2.PdfWriter()
                
                for pagina in paginas_invertidas:
                    pdf_writer.add_page(pagina)
                
                pdf_writer.write(arquivo_saida)
                print("Inversão Concluída!")
        else:
            print("Nenhum arquivo PDF selecionado.")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = PDFSelector()
    window.show()
    sys.exit(app.exec_())
