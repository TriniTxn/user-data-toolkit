import sys, os, tempfile
from PyQt5 import uic, QtWidgets, QtCore, QtGui
from PyQt5.QtWidgets import *
import pandas as pd
import warnings

# Ignorar warning do openpyxl
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# Leitura de Arquivos .ui
app = QtWidgets.QApplication(sys.argv)
temp_dir = getattr(sys, "_MEIPASS", tempfile.gettempdir())
ui_file = os.path.join(temp_dir, "resources\\_tela.ui")
icon_file = os.path.join(temp_dir, "resources\\icone.ico")
ui_load = os.path.join(temp_dir, "resources\\_loading.ui")

if os.path.exists(ui_file):
    tela = uic.loadUi(ui_file)
else:
    tela = uic.loadUi("_tela.ui")

if os.path.exists(icon_file):
    tela.setWindowIcon(QtGui.QIcon(icon_file))
else:
    tela.setWindowIcon(QtGui.QIcon('icone.ico'))

if os.path.exists(ui_load):
    loadin = uic.loadUi(ui_load)
else:
    loadin = uic.loadUi("_loading.ui")

def slotSelect(state):
    for checkbox in tela.checkBoxs:
        checkbox.setChecked(QtCore.Qt.Checked == state)

def menuClose():
    tela.keywords[tela.col] = []
    for element in tela.checkBoxs:
        if element.isChecked():
            tela.keywords[tela.col].append(element.text())
    filterdata()
    tela.menu.close()

def clearFilter():
    if tela.tableWidget.rowCount() > 0:
        for i in range(tela.tableWidget.rowCount()):
            tela.tableWidget.setRowHidden(i, False)

def filterdata():
    columnsShow = dict([(i, True) for i in range(tela.tableWidget.rowCount())])
    for i in range(tela.tableWidget.rowCount()):
        for j in range(tela.tableWidget.columnCount()):
            item = tela.tableWidget.item(i, j)
            if tela.keywords[j]:
                if item.text() not in tela.keywords[j]:
                    columnsShow[i] = False
    for key in columnsShow:
        tela.tableWidget.setRowHidden(key, not columnsShow[key])

def columnfilterclicked(index):
    tela.menu = QtWidgets.QMenu()
    tela.col = index
    tela.data_unique = []
    tela.checkBoxs = []
    tela.checkBox = QtWidgets.QCheckBox("Selecionar tudo", tela.menu)
    checkableAction = QtWidgets.QWidgetAction(tela.menu)
    checkableAction.setDefaultWidget(tela.checkBox)
    tela.menu.addAction(checkableAction)
    tela.checkBox.setChecked(True)
    tela.checkBox.stateChanged.connect(slotSelect)

    for i in range(tela.tableWidget.rowCount()):
        if not tela.tableWidget.isRowHidden(i):
            item = tela.tableWidget.item(i, index)
            if item.text() not in tela.data_unique:
                tela.data_unique.append(item.text())
                checkBox = QtWidgets.QCheckBox(item.text(), tela.menu)
                checkBox.setChecked(True)
                checkableAction = QtWidgets.QWidgetAction(tela.menu)
                checkableAction.setDefaultWidget(checkBox)
                tela.menu.addAction(checkableAction)
                tela.checkBoxs.append(checkBox)

    btn = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel,
                                     QtCore.Qt.Horizontal, tela.menu)
    btn.accepted.connect(menuClose)
    btn.rejected.connect(tela.menu.close)
    checkableAction = QtWidgets.QWidgetAction(tela.menu)
    checkableAction.setDefaultWidget(btn)
    tela.menu.addAction(checkableAction)
    headerPos = tela.tableWidget.mapToGlobal(tela.tableWidgetHeader.pos())
    posY = headerPos.y() + tela.tableWidgetHeader.height()
    posX = headerPos.x() + tela.tableWidgetHeader.sectionPosition(index)
    tela.menu.exec_(QtCore.QPoint(posX, posY))

def ler_excel_auto(caminho):
    import os
    ext = os.path.splitext(caminho)[-1].lower()
    try:
        if ext == ".xlsx":
            return pd.read_excel(caminho, engine="openpyxl")
        elif ext == ".xls":
            return pd.read_excel(caminho, engine="xlrd")
        else:
            return pd.read_csv(caminho, sep=";", engine="python")
    except Exception as e:
        raise ValueError(f"Erro ao ler '{caminho}': {e}")

def processa():
    try:
        tela.status.setText('CARREGANDO DADOS...')
        loadin.show()
        QApplication.processEvents()

        df1 = ler_excel_auto('.xlsx')
        df2 = ler_excel_auto('.xlsx')

        for df_temp in [df1, df2]:
                df_temp['Login'] = df_temp['Login'].astype(str)

        df = pd.merge(df1, df2, how="left", on='Login')


        tabela = tela.findChild(QtWidgets.QTableWidget, "tableWidget")
        tabela.setColumnCount(len(df.columns))
        tabela.setHorizontalHeaderLabels(df.columns)
        tabela.setRowCount(len(df))

        totalrow = len(df)
        basecalculo = totalrow / 100
        countrow = 0

        for row in range(totalrow):
            countrow += 1
            percent = round(countrow / basecalculo, 0)
            loadin.progressBar.setValue(int(percent))
            QApplication.processEvents()

            for col in range(len(df.columns)):
                item = QtWidgets.QTableWidgetItem(str(df.iloc[row, col]))
                tabela.setItem(row, col, item)

        tela.tableWidgetHeader = tela.tableWidget.horizontalHeader()
        tela.tableWidgetHeader.sectionClicked.connect(columnfilterclicked)
        tela.keywords = dict([(i, []) for i in range(tela.tableWidget.columnCount())])
        tela.checkBoxs = []
        tela.col = None

        processa.df = df
        tela.status.setText('DADOS CARREGADOS COM SUCESSO')

    except Exception as e:
        import traceback
        traceback.print_exc()
        QMessageBox.critical(tela, "Erro Crítico", f"Ocorreu um erro grave:\n{str(e)}")

    finally:
        if loadin.isVisible():
            loadin.hide()

processa.df = pd.DataFrame()

def exporta():
    tela.status.setText('EXPORTANDO...')
    tela.setEnabled(False)
    QApplication.processEvents()

    if processa.df.empty:
        tela.status.setText('POR FAVOR CARREGUE OS DADOS')
        tela.setEnabled(True)
        return

    options = QFileDialog.Options()
    options |= QFileDialog.DontUseNativeDialog
    fileName, _ = QFileDialog.getSaveFileName(tela, "Salvar como", "",
                                              "Excel (*.xlsx);;Todos os Arquivos (*)", options=options)
    if not fileName:
        tela.status.setText('EXPORTAÇÃO CANCELADA')
        tela.setEnabled(True)
        return

    if not fileName.endswith(".xlsx"):
        fileName += ".xlsx"

    try:
        with pd.ExcelWriter(fileName, engine='xlsxwriter') as writer:
            processa.df.to_excel(writer, sheet_name='Acessos', index=False)
        tela.status.setText('EXCEL GERADO COM SUCESSO')
    except Exception as e:
        tela.status.setText('FALHA AO GERAR O EXCEL')
        QMessageBox.critical(tela, "Erro", f"Erro ao exportar:\n{e}")

    tela.setEnabled(True)

tela.processButton.clicked.connect(processa)
tela.exportButton.clicked.connect(exporta)

tela.show()
app.exec()
