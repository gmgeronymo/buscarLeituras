# Programa para buscar leituras no banco de dados 'Resultados.mdb'
# Autor: Gean Marcos Geronymo
# Data inicial: 23/01/2017

# Dados do programa
__version__="1.0"
__date__="31/01/2017"
__appname__="Buscar Leituras"
__author__="Gean Marcos Geronymo"
__author_email__="gean.geronymo@gmail.com"

# importar módulos
import os
import sys
import configparser  # abrir arquivos de configuracao (.ini)
# A comunicação com o banco de dados MDB é realizada através do módulo pyodbc
# para instalar o módulo, usar:
# c:\python34\scripts\pip.exe --proxy=http://user:passwd@rproxy02s.inmetro.gov.br:3128 install pyodbc
import pyodbc
# A comunicação com o banco de dados PostgreSQL (Condições Ambientais) é realizada através
# do módulo psycopg2
# para instalar o módulo, usar:
# c:\python34\scripts\pip.exe --proxy=http://user:passwd@rproxy02s.inmetro.gov.br:3128 install psycopg2
import psycopg2
import psycopg2.extras
import time
import datetime
# o módulo numpy é utilizado para calcular média e desvio padrão
from numpy import mean, std

# módulos da interface gráfica Qt5
from PyQt5.QtCore import QDir, Qt
from PyQt5.QtGui import QFont, QPalette, QColor, QIcon

from PyQt5.QtWidgets import (QMainWindow, QWidget, QTableWidget,QTableWidgetItem,
                             QVBoxLayout, QHBoxLayout, QApplication, QCheckBox,
                             QColorDialog, QDialog, QErrorMessage, QFileDialog,
                             QFontDialog, QFrame, QGridLayout, QGroupBox, QSizePolicy,  
                             QInputDialog, QLabel, QLineEdit, QMessageBox, QPushButton,
                             QRadioButton, QComboBox, QDockWidget, QSpinBox)

# os graficos sao gerados utilizando a biblioteca matplotlib
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick
import matplotlib.dates as mdates

# modulos para embutir os graficos na interface grafica Qt5
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from natsort import realsorted, ns

# carregar configuracoes do arquivo settings.ini
config = configparser.ConfigParser()
config.read('settings.ini')

# variáveis configuradas no arquivo settings.ini
caminhoTensao = config['BancoResultados']['caminhoTensao']
caminhoCorrente = config['BancoResultados']['caminhoCorrente']
caminhoRegistroTensao = config['BancoResultados']['caminhoRegistroTensao']
caminhoRegistroCorrente = config['BancoResultados']['caminhoRegistroCorrente']
passwordResultados = config['BancoResultados']['password']

# flag grandezaTensao => determina se a grandeza em uso é tensão ou corrente
# tensão: True
# corrente: false
# o valor inicial é True (Tensão)
grandezaTensao = True

# quantidade de repeticoes no modelo
# padrao: 12 repetições
repeticoesModelo = 12

class App(QMainWindow):
    def __init__(self):
        super(App, self).__init__()

        # título da janela principal
        self.setWindowTitle(__appname__+" - versão "+__version__)
        self.caminho = caminhoTensao  # valor default do caminho do banco de dados
        
        # criar dock widgets
        self.createDockCondicoesAmbientais()
        self.createDockTable()
        self.createDockMain()

        # adiciona os dock widgets ao layout da janela principal
        self.addDockWidget(Qt.TopDockWidgetArea, self.dockMain)
        self.addDockWidget(Qt.BottomDockWidgetArea, self.dockTable)
        self.addDockWidget(Qt.BottomDockWidgetArea, self.dockCondicoesAmbientais)     

        # configura como widget central os widgets da tabela e do gráfico
        # os widgets estão no formato de abas
        self.setCentralWidget(self.tabifyDockWidget(self.dockTable, self.dockCondicoesAmbientais))
        self.dockTable.show()
        self.dockTable.raise_()
        
        self.show()

    # evento disparado ao fechar a janela
    def closeEvent(self, event):
        reply = QMessageBox.question(self, 'Mensagem',
            "Deseja fechar o programa?", QMessageBox.Yes, QMessageBox.No)

        if reply == QMessageBox.Yes:
            # testa se a conexão ao banco de dados está aberta
            try:
                self.resultados.conn
            except:
                pass
            else:
                # se a conexão está aberta, fechar
                self.resultados.conn.close()
            # fechar o programa
            event.accept()
        else:
            # no caso da resposta ao diálogo ser "No", ignorar e manter o programa rodando
            event.ignore()

    # cria o widget principal com as configurações do programa
    def createMainWidget(self):
        self.dockMainWidget = QWidget()
        # chama as funcoes definidas para criar cada GroupBox
        self.createOptionsGroupBox()
        self.createRegistroGroupBox()
        self.createCondicoesAmbientaisGroupBox()
        self.createAcoesGroupBox()
        self.createExportarGroupBox()

        # cria um layout horizontal com os groupbox de opções, registro e condições ambientais
        self.topLayout = QHBoxLayout()
        self.topLayout.addWidget(self.optionsGroupBox)
        self.topLayout.addWidget(self.registroGroupBox)
        self.topLayout.addWidget(self.condicoesAmbientaisGroupBox)

        # cria um layout horizontal com os groupbox de botões "Ações" e "Exportar"
        self.midLayout = QHBoxLayout()
        self.midLayout.addWidget(self.acoesGroupBox)
        self.midLayout.addWidget(self.exportarGroupBox)

        # cria um layout vertical agrupando os dois layouts criados anteriormente
        self.layout = QVBoxLayout()
        self.layout.addLayout(self.topLayout)
        self.layout.addLayout(self.midLayout)
        self.dockMainWidget.setLayout(self.layout)
        # aplica políticas de redimensionamento:
        # o redimensionamento horizontal é livre, mas o vertical é fixo
        self.dockMainWidget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)

    def createDockMain(self):
        self.dockMain = QDockWidget("Configurações", self)
        self.createMainWidget()
        self.dockMain.setWidget(self.dockMainWidget)

    def createDockCondicoesAmbientais(self):
        self.dockCondicoesAmbientais = QDockWidget("Condições Ambientais", self)
        self.plotWidget = PlotCanvas(self)
        self.dockCondicoesAmbientais.setWidget(self.plotWidget)

    def createDockTable(self):
        self.dockTable = QDockWidget("Leituras", self)
        self.tableWidget = QTableWidget()
        self.dockTable.setWidget(self.tableWidget)        

    def setBancoDadosName(self):
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(self,
                "Selecione o Banco de Dados", self.openBancoDadosLabel.text(),
                "All Files (*);;Text Files (*.txt)", options=options)
        if fileName:
            self.openBancoDadosLabel.setText(fileName)
            self.caminho = fileName

    def setRegistroName(self):
        options = QFileDialog.Options()
        fileName, _ = QFileDialog.getOpenFileName(self,
                "Selecione o Registro (LTR)", self.openRegistroLabel.text(),
                "All Files (*);;Text Files (*.txt)", options=options)
        if fileName:
            self.openRegistroLabel.setText(fileName)
            # salva na variável registro apenas o nome do registro
            self.registro = os.path.split(fileName)[-1]
            # limpar campos
            self.valmedSelect.clear()
            self.faixa792Select.clear()
            self.operadorName.setText("")
            self.temperaturaMedia.setText("")
            self.umidadeMedia.setText("")
            try:
                self.resultados = Resultados(self.caminho, self.registro)
                self.operadorName.setText(self.resultados.id.OPERADOR)   
                                  
                for val in self.resultados.valprog:
                    # se for inteiro, remover ',0'
                    if int(val[0]) == float(val[0]):
                        self.valmedSelect.addItem(str(int(val[0])))
                    else:
                        self.valmedSelect.addItem(str(val[0]).replace('.',','))
                    if grandezaTensao == True:
                        if int(val[1]) == float(val[1]):
                            self.faixa792Select.addItem(str(int(val[1])))
                        else:
                            self.faixa792Select.addItem(str(val[1]).replace('.',','))       
            except:
                QMessageBox.critical(self, "Erro",
                "Erro ao conectar com o Banco de Dados!",
                QMessageBox.Abort)

    def createOptionsGroupBox(self):
        self.optionsGroupBox = QGroupBox("Opções")
        frameStyle = QFrame.Sunken | QFrame.Panel

        # Radio button: selecionar tensão ou corrente
        self.tensao = QRadioButton("Tensão")
        self.tensao.setChecked(True) # padrão: Tensão
        self.tensao.toggled.connect(lambda:self.btnstate(self.tensao))
                                    
        self.corrente = QRadioButton("Corrente")
        self.corrente.toggled.connect(lambda:self.btnstate(self.corrente))
        
        # Banco de dados
        self.openBancoDadosLabel = QLineEdit()
        self.openBancoDadosLabel.setReadOnly(True)
        self.openBancoDadosLabel.setText(self.caminho)        
        self.openBancoDadosButton = QPushButton("Selecionar Banco de Dados")

        # LTR
        self.openRegistroLabel = QLineEdit()
        self.openRegistroLabel.setReadOnly(True)
        self.openRegistroLabel.setText(caminhoRegistroTensao)
        self.openRegistroButton = QPushButton("Selecionar Registro (LTR)")

        # Selecionar quantidade de repetições na planilha modelo
        self.repeticoesModeloLabel = QLabel(self)
        self.repeticoesModeloLabel.setText("Repetições na planilha: ")
        self.repeticoesModelo = QSpinBox()
        self.repeticoesModelo.setValue(repeticoesModelo)
        

        # Conectar botões às funções
        self.openBancoDadosButton.clicked.connect(self.setBancoDadosName)
        self.openRegistroButton.clicked.connect(self.setRegistroName)

        # layout
        # layout horizontal no topo com o radio button tensão / corrente e
        # o seletor do número de repetições da planilha
        topLayout = QHBoxLayout()
        topLayout.addWidget(self.tensao)
        topLayout.addWidget(self.corrente)
        topLayout.addWidget(self.repeticoesModeloLabel)
        topLayout.addWidget(self.repeticoesModelo)
        # layout em grid inferior com os seletores de arquivo de banco de dados e LTR
        bottomLayout = QGridLayout()
        bottomLayout.setColumnStretch(1, 1)
        bottomLayout.setColumnMinimumWidth(1, 100)      
        bottomLayout.addWidget(self.openBancoDadosButton, 0, 0)
        bottomLayout.addWidget(self.openRegistroButton, 1, 0)
        bottomLayout.addWidget(self.openBancoDadosLabel, 0, 1)
        bottomLayout.addWidget(self.openRegistroLabel, 1, 1)
        # layout geral vertical juntando os dois layouts anteriores
        optionsGroupBoxLayout = QVBoxLayout()
        optionsGroupBoxLayout.addLayout(topLayout)
        optionsGroupBoxLayout.addLayout(bottomLayout)     

        self.optionsGroupBox.setLayout(optionsGroupBoxLayout)
        
    def createRegistroGroupBox(self):
        self.registroGroupBox = QGroupBox("Registro")
        frameStyle = QFrame.Sunken | QFrame.Panel

        self.operadorLabel = QLabel(self)
        self.operadorLabel.setText("Operador: ")
        self.operadorName = QLabel(self)
        self.operadorName.setFrameStyle(frameStyle)

        self.dataLabel = QLabel(self)
        self.dataLabel.setText("Data: ")
        self.dataValue = QLabel(self)
        self.dataValue.setFrameStyle(frameStyle)
        
        self.valmedLabel = QLabel(self)
        self.valmedLabel.setText("Tensão [V]: ")
        self.valmedSelect = QComboBox()

        self.faixa792Label = QLabel(self)
        self.faixa792Label.setText("Faixa 792A [V]: ")
        self.faixa792Select = QComboBox()

        # layout

        registroGroupBoxLayout = QGridLayout()

        registroGroupBoxLayout.setColumnMinimumWidth(1, 100)

        registroGroupBoxLayout.addWidget(self.operadorLabel, 0, 0)
        registroGroupBoxLayout.addWidget(self.operadorName, 0, 1)
        registroGroupBoxLayout.addWidget(self.dataLabel, 1, 0)
        registroGroupBoxLayout.addWidget(self.dataValue, 1, 1)
        registroGroupBoxLayout.addWidget(self.valmedLabel, 2, 0)
        registroGroupBoxLayout.addWidget(self.valmedSelect, 2, 1)
        registroGroupBoxLayout.addWidget(self.faixa792Label, 3, 0)
        registroGroupBoxLayout.addWidget(self.faixa792Select, 3, 1)

        self.registroGroupBox.setLayout(registroGroupBoxLayout)

    def createCondicoesAmbientaisGroupBox(self):
        self.condicoesAmbientaisGroupBox = QGroupBox("Condições Ambientais")
        frameStyle = QFrame.Sunken | QFrame.Panel

        self.temperaturaMediaLabel = QLabel(self)
        self.temperaturaMediaLabel.setText("Temperatura Média [ºC]: ")
        self.temperaturaMedia = QLabel(self)
        self.temperaturaMedia.setFrameStyle(frameStyle)

        self.umidadeMediaLabel = QLabel(self)
        self.umidadeMediaLabel.setText("Umidade Média [% u.r.]: ")
        self.umidadeMedia = QLabel(self)
        self.umidadeMedia.setFrameStyle(frameStyle)
        
        # layout

        condicoesAmbientaisGroupBoxLayout = QGridLayout()
        condicoesAmbientaisGroupBoxLayout.setColumnMinimumWidth(1, 50)
        
        condicoesAmbientaisGroupBoxLayout.addWidget(self.temperaturaMediaLabel, 0, 0)
        condicoesAmbientaisGroupBoxLayout.addWidget(self.temperaturaMedia, 0, 1)
        condicoesAmbientaisGroupBoxLayout.addWidget(self.umidadeMediaLabel, 1, 0)
        condicoesAmbientaisGroupBoxLayout.addWidget(self.umidadeMedia, 1, 1)

        self.condicoesAmbientaisGroupBox.setLayout(condicoesAmbientaisGroupBoxLayout)

    def createAcoesGroupBox(self):
        self.acoesGroupBox = QGroupBox("Ações")

        self.buscarLeituras = self.createButton("Buscar Leituras", self.buscarLeituras)
        
        acoesGroupBoxLayout = QHBoxLayout()
        acoesGroupBoxLayout.addWidget(self.buscarLeituras)

        self.acoesGroupBox.setLayout(acoesGroupBoxLayout)

    def createExportarGroupBox(self):
        self.exportarGroupBox = QGroupBox("Exportar")
        
        self.copiarNomeRegistro = self.createButton("Copiar Nome do Registro", self.copiarNomeReg)
        self.copiarData = self.createButton("Copiar Data", self.copiarDataReg)
        self.copiarLeituras = self.createButton("Copiar Leituras", self.copiarDiferencas)
        self.copiarModelo = self.createButton("Copiar Modelo Planilha", self.copiarModeloPlanilha)
        self.copiarCondicoesAmbientais = self.createButton("Copiar Condições Ambientais", self.copiarTempUmid)
        self.copiarFigura = self.createButton("Copiar Gráfico", self.copiarGrafico)
        
        exportarGroupBoxLayout = QHBoxLayout()
        
        exportarGroupBoxLayout.addWidget(self.copiarNomeRegistro)
        exportarGroupBoxLayout.addWidget(self.copiarData)
        exportarGroupBoxLayout.addWidget(self.copiarLeituras)
        exportarGroupBoxLayout.addWidget(self.copiarCondicoesAmbientais)
        exportarGroupBoxLayout.addWidget(self.copiarModelo)
        exportarGroupBoxLayout.addWidget(self.copiarFigura)

        self.exportarGroupBox.setLayout(exportarGroupBoxLayout)

    def createButton(self, text, member):
        button = QPushButton(text)
        button.clicked.connect(member)
        return button  

    def updateTable(self,linhas,colunas):
       
        self.tableWidget.setRowCount(linhas)
        self.tableWidget.setColumnCount(colunas)
        self.tableWidget.horizontalHeader().setVisible(False)
        self.tableWidget.verticalHeader().setVisible(False)

        for i in range(colunas):
            if i == 0:
                self.tableWidget.setColumnWidth(i, 100)
            else:
                self.tableWidget.setColumnWidth(i, 60)

    # handler do radio button que seleciona tensão ou corrente
    def btnstate(self, b):
        global grandezaTensao
        if b.text() == "Tensão":
            if b.isChecked() == True:
                self.caminho = caminhoTensao
                self.openBancoDadosLabel.setText(self.caminho)      
                self.valmedLabel.setText("Tensão [V]: ")
                self.valmedSelect.clear()
                self.faixa792Select.clear()
                self.faixa792Select.setEnabled(True)
                self.openRegistroLabel.setText("")
                self.openRegistroLabel.setText(caminhoRegistroTensao)
                self.operadorName.setText("")
                self.dataValue.setText("")
                self.temperaturaMedia.setText("")
                self.umidadeMedia.setText("")
                self.registro = ""
                grandezaTensao = True
                
        if b.text() == "Corrente":
            if b.isChecked() == True:
                self.caminho = caminhoCorrente
                self.openBancoDadosLabel.setText(self.caminho)   
                self.valmedLabel.setText("Corrente [mA]: ")
                self.valmedSelect.clear()
                self.faixa792Select.clear()
                self.faixa792Select.setEnabled(False)
                self.openRegistroLabel.setText("")
                self.openRegistroLabel.setText(caminhoRegistroCorrente)
                self.operadorName.setText("")
                self.dataValue.setText("")
                self.temperaturaMedia.setText("")
                self.umidadeMedia.setText("")
                self.registro = ""
                grandezaTensao = False

    # funcao buscar leituras, chamada através do botão da interface gráfica                                               
    def buscarLeituras(self):
        # coloca o cursor em espera
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            valmed = self.valmedSelect.currentText().replace(',','.')
            faixa792 = self.faixa792Select.currentText().replace(',','.')
            self.resultados.getDiferencas(valmed, faixa792)
            
            colunasTabela = self.resultados.colunas + 1   # uma coluna extra para labels
            linhasTabela = self.resultados.linhas + 7     # linhas extras: espaco, media, desvio padrao, espaco, data, hora
            
            self.dataValue.setText(self.resultados.data[realsorted(self.resultados.diferencas.keys())[0]])

            self.updateTable(linhasTabela,colunasTabela)

            # fonte em negrito
            font_bold = QFont()
            font_bold.setBold(True)

            # dict de checkboxes para selecionar frequencia
            self.freqRepetida = dict()

            # loops para popular a tabela
            for coluna in range(colunasTabela):
                for linha in range(linhasTabela):
                    if coluna == 0: # primeira coluna: labels
                        if linha == 0:
                            self.tableWidget.setItem(linha, 0, QTableWidgetItem("Freq. [KHz]"))
                        elif linha == (linhasTabela - 5):
                            self.tableWidget.setItem(linha, 0, QTableWidgetItem("Média"))
                        elif linha == (linhasTabela - 4):
                            self.tableWidget.setItem(linha, 0, QTableWidgetItem("Desv. Padrão"))
                        elif linha == (linhasTabela - 2):
                            self.tableWidget.setItem(linha, 0, QTableWidgetItem("Data"))
                        elif linha == (linhasTabela - 1):
                            self.tableWidget.setItem(linha, 0, QTableWidgetItem("Hora"))
                        else:
                            self.tableWidget.setItem(linha, 0, QTableWidgetItem(""))
                        # cor de fundo
                        self.tableWidget.item(linha, 0).setBackground(QColor(120,160,200))
                        self.tableWidget.item(linha, 0).setFont(font_bold)
                        # bloquear célular para edição e seleção
                        self.tableWidget.item(linha, 0).setFlags(Qt.ItemIsEnabled)
                    else: # demais colunas: dados                       
                        # ordena as frequencias utilizando a funcao realsorted do pacote natsort
                        # http://pythonhosted.org/natsort/intro.html#quick-description
                        freq = realsorted(self.resultados.diferencas.keys())
                        freq_str = freq[coluna-1]
                        if linha == 0:  # cabeçalho com as frequencias
                            # checar se frequencia é repetida
                            try:
                                if (freq[coluna-1].split()[0] == freq[coluna].split()[0]) | (freq[coluna-1].split()[0] == freq[coluna-2].split()[0]):
                                    self.freqRepetida[freq[coluna-1]] = QTableWidgetItem(freq_str)
                                    self.freqRepetida[freq[coluna-1]].setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
                                    # se for a 'primeira' frequencia da repeticao, sem o apendice '(1)', '(2'), etc,
                                    # é selecionada por padrão
                                    try:
                                        freq[coluna-1].split()[1]
                                    except:
                                        self.freqRepetida[freq[coluna-1]].setCheckState(Qt.Checked)
                                    else:
                                        self.freqRepetida[freq[coluna-1]].setCheckState(Qt.Unchecked)
                                    self.tableWidget.setItem(linha,coluna, self.freqRepetida[freq[coluna-1]])
                                else:
                                    self.tableWidget.setItem(linha,coluna, QTableWidgetItem(freq_str))
                                self.tableWidget.item(linha,coluna).setBackground(QColor(120,160,200))
                                self.tableWidget.item(linha, coluna).setFont(font_bold)
                            except: # a última frequencia falha no try e cai nesse caso
                                self.tableWidget.setItem(linha,coluna, QTableWidgetItem(freq_str))
                                self.tableWidget.item(linha,coluna).setBackground(QColor(120,160,200))
                                self.tableWidget.item(linha, coluna).setFont(font_bold)
                        elif (linha > 0) & (linha < self.resultados.linhas + 1):   # leituras de diferenca ac-dc
                            self.tableWidget.setItem(linha,coluna, QTableWidgetItem("{0:.2f}".format(self.resultados.diferencas[freq[coluna-1]][linha-1][0])))
                            self.tableWidget.item(linha, coluna).setFlags(Qt.ItemIsEnabled)
                        elif linha == self.resultados.linhas + 1:    # pular linha
                            self.tableWidget.setItem(linha,coluna, QTableWidgetItem(""))
                            self.tableWidget.item(linha, coluna).setFlags(Qt.ItemIsEnabled)
                        elif linha == self.resultados.linhas + 2:    # média
                            self.tableWidget.setItem(linha,coluna, QTableWidgetItem("{0:.2f}".format(mean(self.resultados.diferencas[freq[coluna-1]]))))
                            self.tableWidget.item(linha, coluna).setFont(font_bold)
                            # bloquear celulas
                            self.tableWidget.item(linha, coluna).setFlags(Qt.ItemIsEnabled)
                        elif linha == self.resultados.linhas + 3:    # desvio padrão
                            self.tableWidget.setItem(linha,coluna, QTableWidgetItem("{0:.2f}".format(std(self.resultados.diferencas[freq[coluna-1]], ddof=1))))
                            self.tableWidget.item(linha, coluna).setFont(font_bold)
                            # bloquear celulas
                            self.tableWidget.item(linha, coluna).setFlags(Qt.ItemIsEnabled)
                        elif linha == self.resultados.linhas + 4:    # pular linha
                            self.tableWidget.setItem(linha,coluna, QTableWidgetItem(""))
                            self.tableWidget.item(linha, coluna).setFlags(Qt.ItemIsEnabled)
                        elif linha == self.resultados.linhas + 5:    # data
                            self.tableWidget.setItem(linha,coluna, QTableWidgetItem(self.resultados.data[freq[coluna-1]]))
                            # bloquear celulas
                            self.tableWidget.item(linha, coluna).setFlags(Qt.ItemIsEnabled)
                        elif linha == self.resultados.linhas + 6:    # hora
                            self.tableWidget.setItem(linha,coluna, QTableWidgetItem(self.resultados.hora[freq[coluna-1]]))
                            # bloquear celulas
                            self.tableWidget.item(linha, coluna).setFlags(Qt.ItemIsEnabled)
            # depois de montar a tabela, pinta as colunas de frequencias repetidas não selecionadas de cinza
            self.pintarColuna()
            try:
                self.resultados.getCondicoesAmbientais()
                self.temperaturaMedia.setText(self.resultados.temperaturaMedia)
                self.umidadeMedia.setText(self.resultados.umidadeMedia)
                self.plotWidget.plot(self.resultados.date, self.resultados.temperature, self.resultados.humidity)
            except:
                QMessageBox.critical(self, "Erro",
                "Erro ao buscar as condições ambientais!",
                QMessageBox.Abort)         

            self.tableWidget.itemClicked.connect(self.handleItemClicked)
            QApplication.restoreOverrideCursor()
         
        except:
                QApplication.restoreOverrideCursor()
                QMessageBox.critical(self, "Erro",
                "Erro ao conectar com o Banco de Dados!",
                QMessageBox.Abort)         

    # handler para as checkboxes que selecionam qual das frequencias repetidas copiar    
    def handleItemClicked(self, item):
        if item.checkState() == Qt.Checked:
            for i in self.freqRepetida.keys():
                if (i.split()[0] == item.text().split()[0]) & (i != item.text()):
                    self.freqRepetida[i].setCheckState(Qt.Unchecked)
        self.pintarColuna()

    def pintarColuna(self):
        # muda a cor das colunas repetidas não selecionadas para cinza claro
        linhasTabela = self.resultados.linhas + 7
        freq = realsorted(self.resultados.diferencas.keys())
        for i in self.freqRepetida.keys():
            if self.freqRepetida[i].checkState() == Qt.Unchecked:
                coluna = freq.index(i) + 1
                for linha in range(linhasTabela):
                    self.tableWidget.item(linha,coluna).setBackground(QColor(230,230,230))
            else:
                coluna = freq.index(i) + 1
                for linha in range(linhasTabela):
                    if linha == 0:
                        self.tableWidget.item(linha,coluna).setBackground(QColor(120,160,200))
                    else:
                        self.tableWidget.item(linha,coluna).setBackground(QColor(255,255,255))
            
                
    # copiar a tabela de diferenças para a área de transferência (inclusive as não selecionadas)
    def copiarDiferencas(self):
        try:
            clipboard = ""
            nfreq = len(self.resultados.diferencas)
            for i in range(self.resultados.linhas):
                j = 0
                for freq in realsorted(self.resultados.diferencas.keys()):
                    clipboard += str(self.resultados.diferencas[freq][i][0]).replace('.',',')
                    j += 1
                    if j < nfreq:
                        clipboard += '\t'
                clipboard += '\n'
            cb.setText(clipboard, mode=cb.Clipboard)
        except:
            QApplication.restoreOverrideCursor()
            QMessageBox.critical(self, "Erro",
                "Nenhum registro aberto!",
            QMessageBox.Abort)

    # copiar temperatura e umidade para a área de transferência
    def copiarTempUmid(self):
        try:
            clipboard = self.resultados.temperaturaMedia+'\n'+self.resultados.umidadeMedia
            cb.setText(clipboard, mode=cb.Clipboard)
        except:
            QApplication.restoreOverrideCursor()
            QMessageBox.critical(self, "Erro",
                "Nenhum registro aberto!",
            QMessageBox.Abort)

    # copiar nome do registro para a área de transferência
    def copiarNomeReg(self):
        try:
            clipboard = self.registro
            cb.setText(clipboard, mode=cb.Clipboard)
        except:
            QApplication.restoreOverrideCursor()
            QMessageBox.critical(self, "Erro",
                "Nenhum registro aberto!",
            QMessageBox.Abort)

    # copiar data do registro para a área de transferência
    def copiarDataReg(self):
        try:
            clipboard = self.dataValue.text()
            cb.setText(clipboard, mode=cb.Clipboard)
        except:
            QApplication.restoreOverrideCursor()
            QMessageBox.critical(self, "Erro",
                "Nenhum registro aberto!",
            QMessageBox.Abort)

    # copia o gráfico da temperatura e umidade para a área de transferência
    def copiarGrafico(self):
        try:
            cb.setImage(self.plotWidget.grab().toImage())
        except:
            QMessageBox.critical(self, "Erro",
                "Erro ao copiar o gráfico!",
            QMessageBox.Abort)
        
    # copia todos os dados para a área de transferência, em formato compativel com as planilhas Excel
    # utilizadas como modelo para as calibrações
    def copiarModeloPlanilha(self):
        repeticoes = self.repeticoesModelo.value()
        if grandezaTensao == True:
            # set de frequencias padrão para tensão
            # se foi realizada medicao na frequencia de 62 Hz, utiliza-a no lugar de 60 Hz
            if '0.062' in self.resultados.diferencas:
                freq = [0.01, 0.02, 0.03, 0.04, 0.055, 0.062, 0.065, 0.12, 0.3, 0.4, 0.5, 1, 10, 20, 30, 50, 70, 100, 200, 300, 500, 700, 800, 1000]
            else:
                freq = [0.01, 0.02, 0.03, 0.04, 0.055, 0.06, 0.065, 0.12, 0.3, 0.4, 0.5, 1, 10, 20, 30, 50, 70, 100, 200, 300, 500, 700, 800, 1000]                

            #repeticoes = 12
            try:
                clipboard = ""
                for linha in range(repeticoes+3):
                    for coluna in range(29):
                        if linha == 0:
                            if coluna == 0:
                                clipboard += self.resultados.temperaturaMedia + '\t'
                            elif coluna == 4:
                                clipboard += self.faixa792Select.currentText() + '\t'
                            elif coluna == 5:
                                clipboard += self.valmedSelect.currentText() + '\t'
                            elif coluna == 7:
                                clipboard += self.dataValue.text() + '\t'
                            elif coluna == 28:
                                clipboard += '\n'
                            else:
                                clipboard += ' \t'
                        elif linha == 1:
                            if coluna == 0:
                                clipboard += self.resultados.umidadeMedia + '\t'
                            elif coluna == 4:
                                clipboard += self.registro + '\t'
                            elif coluna == 28:
                                clipboard += '\n'
                            else:
                                clipboard += ' \t'
                        elif linha == 2:
                            if coluna == 4:
                                clipboard += 'Freq (kHz) \t'
                            elif coluna in range(5,28):
                                if freq[coluna-5] < 1:
                                    clipboard += str(freq[coluna-5]).replace('.',',') + '\t'
                                else:
                                    clipboard += str(int(freq[coluna-5])).replace('.',',') + '\t'
                            elif coluna == 28:
                                clipboard += str(int(freq[coluna-5])) + '\n'
                            else:
                                clipboard += ' \t'
                        elif linha > 2:
                            if coluna in range(5,28):
                                try:
                                    if freq[coluna-5] < 1:
                                        freqStr = str(freq[coluna-5])
                                    else:
                                        freqStr = str(int(freq[coluna-5]))
                                    # checar se frequencia atual é repetida
                                    for i in self.freqRepetida.keys():
                                        if i.split()[0] == freqStr:
                                            if self.freqRepetida[i].checkState() == Qt.Checked:
                                                freqStr = i
                                    # copia o ponto
                                    clipboard += str(self.resultados.diferencas[freqStr][linha-3][0]).replace('.',',') + '\t'
                                except:
                                    clipboard += '\t'
                            elif coluna == 28:
                                try:
                                    if freq[coluna-5] < 1:
                                        freqStr = str(freq[coluna-5])
                                    else:
                                        freqStr = str(int(freq[coluna-5]))
                                    # checar se frequencia atual é repetida
                                    for i in self.freqRepetida.keys():
                                        if i.split()[0] == freqStr:
                                            if self.freqRepetida[i].checkState() == Qt.Checked:
                                                freqStr = i
                                    # copia o ponto   
                                    clipboard += str(self.resultados.diferencas[freqStr][linha-3][0]).replace('.',',') + '\n'
                                except:
                                    clipboard += '\n'
                            else:
                                clipboard += ' \t'

                    cb.setText(clipboard, mode=cb.Clipboard)

            except:
                QMessageBox.critical(self, "Erro", "Nenhum registro aberto!", QMessageBox.Abort)

        else:
        # set de frequencias padrão para corrente
        # se foi realizada medicao na frequencia de 62 Hz, utiliza-a no lugar de 60 Hz
            if '0.062' in self.resultados.diferencas:
                freq = [0.01, 0.02, 0.03, 0.04, 0.05, 0.055, 0.062, 0.065, 0.12, 0.5, 0.6, 1.0, 5.0, 10.0, 20.0, 50.0, 70.0, 100.0]
            else:
                freq = [0.01, 0.02, 0.03, 0.04, 0.05, 0.055, 0.06, 0.065, 0.12, 0.5, 0.6, 1.0, 5.0, 10.0, 20.0, 50.0, 70.0, 100.0]

            #repeticoes = 12
            try:
                clipboard = ""
                for linha in range(repeticoes+3):
                    for coluna in range(23):
                        if linha == 0:
                            if coluna == 0:
                                clipboard += self.resultados.temperaturaMedia + '\t'
                            elif coluna == 4:
                                clipboard += self.valmedSelect.currentText() + ' mA \t'
                            elif coluna == 6:
                                clipboard += self.dataValue.text() + '\t'
                            elif coluna == 22:
                                clipboard += '\n'
                            else:
                                clipboard += ' \t'
                        elif linha == 1:
                            if coluna == 0:
                                clipboard += self.resultados.umidadeMedia + '\t'
                            elif coluna == 4:
                                clipboard += self.registro + '\t'
                            elif coluna == 22:
                                clipboard += '\n'
                            else:
                                clipboard += ' \t'
                        elif linha == 2:
                            if coluna == 4:
                                clipboard += 'Freq (kHz) \t'
                            elif coluna in range(5,22):
                                if freq[coluna-5] < 1:
                                    clipboard += str(freq[coluna-5]).replace('.',',') + '\t'
                                else:
                                    clipboard += str(int(freq[coluna-5])).replace('.',',') + '\t'
                            elif coluna == 22:
                                clipboard += str(int(freq[coluna-5])) + '\n'
                            else:
                                clipboard += ' \t'
                        elif linha > 2:
                            if coluna in range(5,22):
                                try:
                                    if freq[coluna-5] < 1:
                                        freqStr = str(freq[coluna-5])
                                    else:
                                        freqStr = str(int(freq[coluna-5]))
                                    # checar se frequencia atual é repetida
                                    for i in self.freqRepetida.keys():
                                        if i.split()[0] == freqStr:
                                            if self.freqRepetida[i].checkState() == Qt.Checked:
                                                freqStr = i
                                    # copia o ponto
                                    clipboard += str(self.resultados.diferencas[freqStr][linha-3][0]).replace('.',',') + '\t'
                                except:
                                    clipboard += '\t'
                            elif coluna == 22:
                                try:
                                    if freq[coluna-5] < 1:
                                        freqStr = str(freq[coluna-5])
                                    else:
                                        freqStr = str(int(freq[coluna-5]))
                                        # checar se frequencia atual é repetida
                                    for i in self.freqRepetida.keys():
                                        if i.split()[0] == freqStr:
                                            if self.freqRepetida[i].checkState() == Qt.Checked:
                                                freqStr = i
                                    # copia o ponto
                                    clipboard += str(self.resultados.diferencas[freqStr][linha-3][0]).replace('.',',') + '\n'
                                except:
                                    clipboard += '\n'
                            else:
                                clipboard += ' \t'

                    cb.setText(clipboard, mode=cb.Clipboard)

            except:
                QMessageBox.critical(self, "Erro", "Nenhum registro aberto!", QMessageBox.Abort)      

# classe que faz a busca dos resultados nos banco de dados
class Resultados(object):
    """ Classe que busca os resultados no banco de dados MS Access
    Atributos:
    caminhoBancoDados: string com o caminho do banco de dados
    nomeRegistro: string com o nome do arquivo TXT do registro
    """

    def __init__(self, caminhoBancoDados, nomeRegistro):
        self.passwd = passwordResultados
        self.caminhoBancoDados = caminhoBancoDados
        self.nomeRegistro = nomeRegistro
        self.conn = pyodbc.connect(r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)}; DBQ='+self.caminhoBancoDados+';PWD='+self.passwd)
        cur = self.conn.cursor()

        self.id = cur.execute("SELECT CODREG, CODOBJ, OPERADOR FROM Resultados WHERE NOMEREG='"+self.nomeRegistro+"'").fetchone()

        if grandezaTensao == True:
            self.valprog = cur.execute("SELECT DISTINCT VALPROG, FAIXA792 FROM Valmed WHERE CODREG="+str(self.id.CODREG)).fetchall()
        else:
            self.valprog = cur.execute("SELECT DISTINCT VALPROG FROM Leituras WHERE CODREG="+str(self.id.CODREG)).fetchall()            

        cur.close()

    # funcao que busca as diferencas ac-dc no banco de dados
    def getDiferencas(self, valprog, faixa792):
        cur = self.conn.cursor()

        if grandezaTensao == True:
            self.registro_rows = cur.execute("SELECT DISTINCT FREQ, CODPONTO, CODI, DATACAL, HORACAL FROM Valmed WHERE CODREG="+str(self.id.CODREG)+" AND VALPROG="+valprog+" AND FAIXA792="+faixa792+" ORDER BY DATACAL, HORACAL").fetchall()      
        else:
           self.registro_rows = cur.execute("SELECT DISTINCT FREQ, CODPONTO, CODI, DATACAL, HORACAL FROM Valmed WHERE CODREG="+str(self.id.CODREG)+" AND VALPROG="+valprog+" ORDER BY DATACAL, HORACAL").fetchall() 
        
        self.diferencas = dict()
        self.hora = dict()
        self.data = dict()
        
        for row in self.registro_rows:
            tmp = row.FREQ.split('  ')
            if tmp[1] == 'Hz':
                newFreq = float(tmp[0]) * 0.001
            elif tmp[1] == 'kHz':
                newFreq = float(tmp[0])
            elif tmp[1] == 'MHz':
                newFreq = float(tmp[0]) * 1000

            # formata as frequencias
            if newFreq < 1:
                newFreqStr = str(newFreq)
            else:
                newFreqStr = str(int(newFreq))

            # verifica se existem frequencias repetidas
            # se existir, adicionar (1), (2)... etc.
            try:
                i = 1
                while self.diferencas[newFreqStr]:
                    newFreqStr += ' ('+str(i)+')'
            except:
                pass
                            
            self.diferencas[newFreqStr] = cur.execute("SELECT DIFERENCA FROM Leituras WHERE CODREG="+str(self.id.CODREG)+" AND CODPONTO ="+str(row.CODPONTO)).fetchall()
            self.data[newFreqStr] = row.DATACAL.strftime("%d/%m/%y")
            self.hora[newFreqStr] = row.HORACAL.strftime("%H:%M")

        cur.close()

        # determinar tamanho da tabela
        self.colunas = len(self.diferencas)  # quantidade de colunas (frequencias)

        linhas_list = [len(v) for v in self.diferencas.values()]
        self.linhas = linhas_list[0]   # quantidade de linhas (repeticoes)

        return

    # busca as condicoes ambientais no banco de dados PostgreSQL
    def getCondicoesAmbientais(self):
        freq = sorted(self.data.keys())   # lista das frequencias, ordenadas de forma crescente
        timestamp = []   # inicializa a lista timestamp
        for i in freq:
            s = self.data[i] +" "+ self.hora[i]   # monta a string s na forma 01/01/01 12:00
            timestamp.append(datetime.datetime.strptime(s, "%d/%m/%y %H:%M"))  # converte em objeto timestamp

        timestamp.sort()   # ordena os timestamps de forma crescente
        dataInicial = timestamp[0].strftime("%d/%m/%Y %H:%M")   # data e hora de inicio
        dataFinal = timestamp[-1].strftime("%d/%m/%Y %H:%M")     # data e hora de fim
        # buscar dados das codicoes ambientais no banco de dados
        conn = psycopg2.connect("dbname={} user={} password={} host={}".format(config['BancoCondicoesAmbientais']['dbname'],config['BancoCondicoesAmbientais']['user'],config['BancoCondicoesAmbientais']['password'],config['BancoCondicoesAmbientais']['host']))
        cur = conn.cursor(cursor_factory=psycopg2.extras.DictCursor)    
        cur.execute("SELECT date,temperature, humidity FROM condicoes_ambientais WHERE date >= '"+dataInicial+"' AND date < '"+dataFinal+"';")
        rows = cur.fetchall()
        self.temperature = []
        self.humidity = []
        self.date = []
        
        for row in rows:
            self.temperature.append(float(row['temperature']))
            self.humidity.append(float(row['humidity']))
            self.date.append(row['date'])
            
        self.temperaturaMedia = "{0:.1f}".format(mean(self.temperature)).replace('.',',')
        self.umidadeMedia = "{0:.1f}".format(mean(self.humidity)).replace('.',',')
        
        cur.close()
        conn.close()
        return

# classe para plotar as condicoes ambientais (temperatura e umidade) em um gráfico
class PlotCanvas(FigureCanvas):
 
    def __init__(self, parent=None, width=5, height=4, dpi=80):
        fig = Figure(figsize=(width, height), dpi=dpi)
        self.axes = fig.add_subplot(111)     
        
        FigureCanvas.__init__(self, fig)
        self.setParent(parent)
 
        FigureCanvas.setSizePolicy(self,
                QSizePolicy.Expanding,
                QSizePolicy.Expanding)
        FigureCanvas.updateGeometry(self)

        # formatacao criar eixos para temperatura e umidade      
        self.ax1 = self.figure.add_subplot(111)
        self.ax2 = self.ax1.twinx()

    def plot(self, date, temperature, humidity):
        # limpar o gráfico a cada novo plot
        self.ax1.clear()
        self.ax2.clear()
        # formatacao das datas
        self.figure.autofmt_xdate()
        # identificação dos eixos
        self.ax1.set_title('Condições Ambientais')
        self.ax1.set_xlabel('Dia / Hora')
        self.ax1.set_ylabel('Temperatura [ºC]')
        self.ax2.set_ylabel('Umidade [% u.r.]')
        self.ax1.grid(True)
        # formatar datas (eixo x)
        myFmt = mdates.DateFormatter('%d/%m - %H:%M')
        self.ax1.xaxis.set_major_formatter(myFmt)
        self.ax1.format_xdata = mdates.DateFormatter('%d/%m - %H:%M')       
        # formatar eixo de temperaturas - precisão de 0.1
        self.ax1.yaxis.set_major_formatter(mtick.FormatStrFormatter('%.1f'))
        lns1 = self.ax1.plot(date, temperature, 'r-', label='Temperatura')        
        lns2 = self.ax2.plot(date, humidity, 'b-', label='Umidade')
        # legenda       
        lns = lns1+lns2
        labs = [l.get_label() for l in lns]
        self.ax1.legend(lns, labs, loc="best")
        self.draw()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    cb = app.clipboard()
    cb.clear(mode=cb.Clipboard)
    ex = App()
    sys.exit(app.exec_())

