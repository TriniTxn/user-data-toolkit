import sys
import os
import pandas as pd
import traceback
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QFileDialog, QMessageBox

class TxtToExcelConverter(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Conversor TXT para XLSX")
        self.resize(400, 150)
        self.layout = QtWidgets.QVBoxLayout(self)

        self.btn_load = QtWidgets.QPushButton("Carregar arquivo .txt")
        self.btn_save = QtWidgets.QPushButton("Salvar como .xlsx")
        self.btn_save.setEnabled(False)

        self.layout.addWidget(self.btn_load)
        self.layout.addWidget(self.btn_save)

        self.btn_load.clicked.connect(self.load_txt)
        self.btn_save.clicked.connect(self.save_xlsx)

        self.df = None

    def load_txt(self):
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(self, "Selecione arquivo TXT", "",
                                                  "Arquivos TXT (*.txt);;Todos os arquivos (*)", options=options)
        if fileName:
            try:
                self.df = pd.read_csv(fileName, sep='\t', engine='python', encoding='latin1')
                if self.df.empty:
                    raise ValueError("Arquivo carregado está vazio.")
                QMessageBox.information(self, "Sucesso",
                                        f"Arquivo '{os.path.basename(fileName)}' carregado com sucesso!")
                self.btn_save.setEnabled(True)
            except Exception as e:
                print("Erro ao carregar arquivo TXT:")
                traceback.print_exc()
                QMessageBox.critical(self, "Erro", f"Falha ao carregar arquivo:\n{str(e)}")

    def save_xlsx(self):
        if self.df is None:
            QMessageBox.warning(self, "Aviso", "Nenhum arquivo carregado para salvar.")
            return

        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getSaveFileName(self, "Salvar arquivo Excel", "",
                                                  "Excel (*.xlsx);;Todos os arquivos (*)", options=options)
        if fileName:
            if not fileName.endswith(".xlsx"):
                fileName += ".xlsx"
            try:
                self.df.to_excel(fileName, index=False)
                QMessageBox.information(self, "Sucesso", f"Arquivo salvo como '{os.path.basename(fileName)}'")
            except Exception as e:
                print("Erro ao salvar arquivo Excel:")
                traceback.print_exc()
                QMessageBox.critical(self, "Erro", f"Falha ao salvar arquivo:\n{str(e)}")

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    janela = TxtToExcelConverter()
    janela.show()
    sys.exit(app.exec_())