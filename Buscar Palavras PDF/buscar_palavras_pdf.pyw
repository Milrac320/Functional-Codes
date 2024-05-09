import pdfplumber
import openpyxl
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QFileDialog, QLineEdit, QProgressBar
from PyQt5.QtCore import Qt

class PDFExtractor(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Buscar Verbas PDF")
        self.setGeometry(822, 281, 300, 200)
        self.setStyleSheet("background-color: black")

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

        self.select_button = QPushButton("Selecionar PDF")
        self.select_button.clicked.connect(self.select_pdf)
        self.layout.addWidget(self.select_button)
        self.select_button.setStyleSheet(style)
        self.select_button.pressed.connect(lambda: self.select_button.setStyleSheet(pressed_style))
        self.select_button.released.connect(lambda: self.select_button.setStyleSheet(style))

        self.file_path = QLineEdit()
        self.file_path.setPlaceholderText("Arquivo PDF")
        self.layout.addWidget(self.file_path)
        self.file_path.setStyleSheet("color:white")

        self.keyword_input = QLineEdit()
        self.keyword_input.setPlaceholderText("Palavras-chave (separadas por vírgulas)")
        self.layout.addWidget(self.keyword_input)
        self.keyword_input.setStyleSheet("color:white")

        self.extract_button = QPushButton("Extrair Palavras-chave")
        self.extract_button.clicked.connect(self.extract_keywords)
        self.layout.addWidget(self.extract_button)
        self.extract_button.setStyleSheet(style)
        self.extract_button.pressed.connect(lambda: self.extract_button.setStyleSheet(pressed_style))
        self.extract_button.released.connect(lambda: self.extract_button.setStyleSheet(style))

        self.progress_bar = QProgressBar(self)
        self.layout.addWidget(self.progress_bar)
        self.progress_bar.setValue(0)
        self.progress_bar.setAlignment(Qt.AlignCenter)

        self.setLayout(self.layout)

    def select_pdf(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Selecionar PDF", "", "Arquivos PDF (*.pdf)")
        if file_path:
            self.file_path.setText(file_path)

    def extract_keywords(self):
        pdf_file = self.file_path.text()
        keywords = self.keyword_input.text().split(',')
        self.extract_keywords_from_pdf(pdf_file, keywords)

    def extract_keywords_from_pdf(self, pdf_file, keywords):
        if not pdf_file or not keywords:
            return

        rows = []

        with pdfplumber.open(pdf_file) as pdf:
            total_pages = len(pdf.pages)
            current_page = 0

            for page_num, page in enumerate(pdf.pages, start=1):
                current_page += 1
                self.progress_bar.setValue(int(current_page / total_pages * 100))  # Atualiza a barra de progresso

                page_text = page.extract_text()
                for keyword in keywords:
                    if keyword in page_text:
                        lines = page_text.split('\n')
                        for line in lines:
                            if keyword in line:
                                rows.append([keyword, f"Page {page_num}", line])

        self.save_to_excel(rows)

    def save_to_excel(self, data):
        if not data:
            return

        save_location, _ = QFileDialog.getSaveFileName(self, "Salvar Resultados", "", "Arquivos Excel (*.xlsx)")

        if save_location:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(["Palavra-chave", "Página", "Linha"])

            for row in data:
                ws.append(row)

            wb.save(save_location)
            print(f"Resultados salvos em '{save_location}'")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = PDFExtractor()
    window.show()
    sys.exit(app.exec_())
