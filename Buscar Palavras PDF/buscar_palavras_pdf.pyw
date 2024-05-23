import os
import pdfplumber
import openpyxl
import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QMessageBox, QFileDialog, QLineEdit, QProgressBar
from PyQt5.QtCore import Qt

class PDFExtractor(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Buscar Verbas PDF")
        self.setGeometry(822, 281, 300, 200)
        self.setStyleSheet("background-color: black")
        self.layout = QVBoxLayout()
        self.init_ui()
        self.setLayout(self.layout)

    def init_ui(self):
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
            background-color: orange;
            color: black;
            border: 0.5px solid orange;
            position: relative;
            margin-top: 4px;
            margin-bottom: 4px;
            border-radius: 10px;
        """

        self.select_button = QPushButton("Selecionar PDF")
        self.select_button.clicked.connect(self.select_pdf)
        self.select_button.setStyleSheet(style)
        self.select_button.pressed.connect(lambda: self.select_button.setStyleSheet(pressed_style))
        self.select_button.released.connect(lambda: self.select_button.setStyleSheet(style))
        self.layout.addWidget(self.select_button)

        self.file_path = QLineEdit()
        self.file_path.setPlaceholderText("Arquivo PDF")
        self.file_path.setStyleSheet("color:white")
        self.layout.addWidget(self.file_path)

        self.keyword_input = QLineEdit()
        self.keyword_input.setPlaceholderText("Palavras-chave (separadas por vírgulas)")
        self.keyword_input.setStyleSheet("color:white")
        self.layout.addWidget(self.keyword_input)

        self.extract_button = QPushButton("Extrair Palavras-chave")
        self.extract_button.clicked.connect(self.extract_keywords)
        self.extract_button.setStyleSheet(style)
        self.extract_button.pressed.connect(lambda: self.extract_button.setStyleSheet(pressed_style))
        self.extract_button.released.connect(lambda: self.extract_button.setStyleSheet(style))
        self.layout.addWidget(self.extract_button)

        self.progress_bar = QProgressBar(self)
        self.progress_bar.setValue(0)
        self.progress_bar.setAlignment(Qt.AlignCenter)
        self.layout.addWidget(self.progress_bar)

    def select_pdf(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Selecionar PDF", "", "Arquivos PDF (*.pdf)")
        self.file_path.setText(file_path) if file_path else None   

    def extract_keywords(self):
        pdf_file = self.file_path.text()
        keywords = [kw.strip() for kw in self.keyword_input.text().split(',')]
        if not pdf_file:
            self.show_message("Atenção!", "Por favor, selecione um arquivo PDF.", QMessageBox.Warning)
            return
        if not keywords:
            self.show_message("Atenção!", "Por favor, insira palavras-chave.", QMessageBox.Warning)
            return

        self.extract_keywords_from_pdf(pdf_file, keywords)

    def extract_keywords_from_pdf(self, pdf_file, keywords):
        rows = []
        try:
            with pdfplumber.open(pdf_file) as pdf:
                total_pages = len(pdf.pages)
                for page_num, page in enumerate(pdf.pages, start=1):
                    self.update_progress(page_num, total_pages)
                    page_text = page.extract_text()
                    if not page_text:
                        continue
                    for keyword in keywords:
                        if keyword in page_text:
                            for line in page_text.split('\n'):
                                rows.append([keyword, f"Page {page_num}", line]) if keyword in line else None
                                    
            self.save_to_excel(pdf_file, rows)
        except Exception as e:
            self.show_message("Erro", f"Erro ao processar o PDF: {str(e)}", QMessageBox.Critical)

    def save_to_excel(self, pdf_file, data):
        if not data:
            self.show_message("Atenção!", "Nenhuma Palavra-Chave foi encontrada!", QMessageBox.Information)
            return

        try:
            pdf_dir = os.path.dirname(pdf_file)
            base_name = os.path.splitext(os.path.basename(pdf_file))[0]
            excel_filename = os.path.join(pdf_dir, f"{base_name}_palavras_chave.xlsx")

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(["Palavra-chave", "Página", "Linha"])
            for row in data:
                ws.append(row)
            wb.save(excel_filename)
            self.show_message("Concluído!", f"Pesquisa concluída e arquivo salvo em '{excel_filename}'", QMessageBox.Information)
        except Exception as e:
            self.show_message("Erro", f"Erro ao salvar o arquivo: {str(e)}", QMessageBox.Critical)

    def update_progress(self, current_page, total_pages):
        self.progress_bar.setValue(int(current_page / total_pages * 100))

    def show_message(self, title, message, icon):
        msg = QMessageBox()
        msg.setWindowTitle(title)
        msg.setText(message)
        msg.setIcon(icon)
        msg.exec_()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = PDFExtractor()
    window.show()
    sys.exit(app.exec_())
