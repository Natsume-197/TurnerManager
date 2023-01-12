import sys
from PyQt5.QtWidgets import QMessageBox, QApplication , QTabWidget, QProgressBar, QFrame, QPlainTextEdit, QDateEdit, QPushButton, QLabel, QGridLayout, QApplication, QWidget, QCheckBox, QHBoxLayout, QVBoxLayout
from PyQt5.QtCore import Qt, QDateTime, QCoreApplication
from PyQt5.QtGui import QFont, QIcon
import requests
import pandas as pd
import os 
import time

url_base = 'https://dailygrids.turnertapkit.com/shows.xls?'

channel_options = {
    'TNT Argentina': 'TNTLA_AR',
    'TNT Chile': 'TNTLA_CL',
    'TNT Colombia': 'TNTLA_CO',
    'TNT PAN': 'TNTLA_PAN',
    'TNT Series PAN 1':'TNTSLA_PAN1',
    'TNT Series PAN 2':'TNTSLA_PAN2',
    'Space PAN':'SPACELA_PAN',
    'Space Centro':'SPACELA_C',
    'Space Sur':'SPACELA_S',
    'TBS PAN':'TBSLA_PAN',
    'TBS Sur':'TBSLA_S',
    'TCM Argentina':'TCMLA_AR',
    'TCM PAN':'TCMLA_PAN',
    'ISAT PAN':'ISATLA_PAN',
    'TRUTV PAN':'TRUTVLAHD_PAN',
    'Glitz PAN':'GLITZLA_PAN',
    'CN Argentina':'CNLA_AR',
    'CN PAN':'CNLA_PAN',
    'CN PAN 2':'CNLA_PAN2',
    'Cartoonito PAN':'CTOOLA_PAN',
    'Tooncast PAN':'TOONLA_PAN'
}

class MyApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Gestor de logs (Turner)')
        self.setFixedSize(600, 700)
        self.setWindowIcon(QIcon('favicon.ico'))
        layout = QGridLayout()
        self.setLayout(layout)
        
        self.tabs = QTabWidget()
        self.tab1 = QWidget()
        self.tab2 = QWidget()
        self.tabs.resize(300,200)
        
        self.tabs.addTab(self.tab1,"Diarios")
        # self.tabs.addTab(self.tab2,"Mensuales")
                
        # Create first tab
        self.tab1.layout = QGridLayout(self)
        
        # Add tabs to widget
        layout.addWidget(self.tabs)
        self.tab1.setLayout(self.tab1.layout)
        
        self.checkBox_close = QCheckBox('Forzar cierre de logs')
        self.checkBox_close.setChecked(False)
        self.tab1.layout.addWidget(self.checkBox_close, 0, 0)
        
        self.label_date_1 = QLabel("Fecha de inicio")
        self.tab1.layout.addWidget(self.label_date_1, 0, 1)
        
        self.label_date_2 = QLabel("Fecha de corte")
        self.tab1.layout.addWidget(self.label_date_2, 0, 2)

        self.button_open_folder = QPushButton("Ver procesados", self)
        self.tab1.layout.addWidget(self.button_open_folder, 0, 3, 1, 1)
        self.button_open_folder.clicked.connect(self.open_folder_action)

        self.button_download = QPushButton("Descargar", self)
        self.tab1.layout.addWidget(self.button_download, 1, 3, 1, 1)
        
        self.button_download.clicked.connect(self.download_action)
        
        self.dateedit = QDateEdit(calendarPopup=True)
        self.dateedit.setDateTime(QDateTime.currentDateTime().addDays(1))
        layout.addWidget(self.dateedit, 1, 1)
        self.tab1.layout.addWidget(self.dateedit, 1, 1)
        
        self.dateedit2 = QDateEdit(calendarPopup=True)
        self.dateedit2.setDateTime(QDateTime.currentDateTime().addDays(1))
        layout.addWidget(self.dateedit2, 1, 2)
        self.tab1.layout.addWidget(self.dateedit2, 1, 2)
        
        self.separatorLine = QFrame(frameShape=QFrame.HLine)    
        self.tab1.layout.addWidget(self.separatorLine, 2, 0, 1, 4)
        
        self.checkBoxAll = QCheckBox('Seleccionar todos los canales')
        self.checkBoxAll.setChecked(False)
        self.checkBoxAll.stateChanged.connect(self.on_stateChanged_all)
        self.tab1.layout.addWidget(self.checkBoxAll, 2, 0)
        self.tab1.layout.addWidget(self.checkBoxAll)

        ###############
        ## FIRST ROW ##
        ###############
        # Checkbox TNT        
        self.checkBox_tnt = QCheckBox("TNT", self)
        self.checkBox_tnt.setFont(QFont("Arial", 11, weight=QFont.Bold))
        self.checkBox_tnt.stateChanged.connect(self.on_stateChanged_tnt)
        self.checkBox_tnt_A = QCheckBox('TNT Argentina')
        self.checkBox_tnt_B = QCheckBox('TNT Chile')
        self.checkBox_tnt_C = QCheckBox('TNT Colombia')
        self.checkBox_tnt_D = QCheckBox('TNT PAN')

        # Checkbox TNT Series
        self.checkBox_tnt_series = QCheckBox("TNT Series", self)
        self.checkBox_tnt_series.setFont(QFont("Arial", 11, weight=QFont.Bold))
        self.checkBox_tnt_series.stateChanged.connect(self.on_stateChanged_tnt_series)
        self.checkBox_tnt_series_A = QCheckBox('TNT Series PAN 1')
        self.checkBox_tnt_series_B = QCheckBox('TNT Series PAN 2')

        # Checkbox Space
        self.checkBox_space = QCheckBox("Space", self)
        self.checkBox_space.setFont(QFont("Arial", 11, weight=QFont.Bold))
        self.checkBox_space.stateChanged.connect(self.on_stateChanged_space)
        self.checkBox_space_A = QCheckBox('Space PAN')
        self.checkBox_space_B = QCheckBox('Space Centro')
        self.checkBox_space_C = QCheckBox('Space Sur')
        
        ################
        ## SECOND ROW ##
        ################
        # Checkbox TBS
        self.checkBox_tbs = QCheckBox("TBS", self)
        self.checkBox_tbs.setFont(QFont("Arial", 11, weight=QFont.Bold))
        self.checkBox_tbs.stateChanged.connect(self.on_stateChanged_tbs)
        self.checkBox_tbs_A = QCheckBox('TBS PAN')
        self.checkBox_tbs_B = QCheckBox('TBS Sur')

        # Checkbox TCM
        self.checkBox_tcm = QCheckBox("TCM", self)
        self.checkBox_tcm.setFont(QFont("Arial", 11, weight=QFont.Bold))
        self.checkBox_tcm.stateChanged.connect(self.on_stateChanged_tcm)
        self.checkBox_tcm_A = QCheckBox('TCM Argentina')
        self.checkBox_tcm_B = QCheckBox('TCM PAN')
        
        # Checkbox ISAT
        self.checkBox_isat = QCheckBox("ISAT", self)
        self.checkBox_isat.setFont(QFont("Arial", 11, weight=QFont.Bold))
        self.checkBox_isat.stateChanged.connect(self.on_stateChanged_isat)
        self.checkBox_isat_A = QCheckBox('ISAT PAN')
        
        # Checkbox TRUTV
        self.checkBox_trutv = QCheckBox("TRUTV", self)
        self.checkBox_trutv.setFont(QFont("Arial", 11, weight=QFont.Bold))
        self.checkBox_trutv.stateChanged.connect(self.on_stateChanged_trutv)
        self.checkBox_trutv_A = QCheckBox('TRUTV PAN')
        
        # Checkbox Glitz
        self.checkBox_glitz = QCheckBox("Glitz", self)
        self.checkBox_glitz.setFont(QFont("Arial", 11, weight=QFont.Bold))
        self.checkBox_glitz.stateChanged.connect(self.on_stateChanged_glitz)
        self.checkBox_glitz_A = QCheckBox('Glitz PAN')
        
        ###############
        ## THIRD ROW ##
        ###############
        # Checkbox Cartoon Network
        self.checkBox_cn = QCheckBox("Cartoon Network", self)
        self.checkBox_cn.setFont(QFont("Arial", 11, weight=QFont.Bold))
        self.checkBox_cn.stateChanged.connect(self.on_stateChanged_cn)
        self.checkBox_cn_A = QCheckBox('CN Argentina')
        self.checkBox_cn_B = QCheckBox('CN PAN')
        self.checkBox_cn_C = QCheckBox('CN PAN 2')
        
        # Checkbox Cartoonito
        self.checkBox_cnito = QCheckBox("Cartoonito", self)
        self.checkBox_cnito.setFont(QFont("Arial", 11, weight=QFont.Bold))
        self.checkBox_cnito.stateChanged.connect(self.on_stateChanged_cnito)
        self.checkBox_cnito_A = QCheckBox('Cartoonito PAN')

        # Checkbox Tooncast
        self.checkBox_tooncast = QCheckBox("Tooncast", self)
        self.checkBox_tooncast.setFont(QFont("Arial", 11, weight=QFont.Bold))
        self.checkBox_tooncast.stateChanged.connect(self.on_stateChanged_tooncast)
        self.checkBox_tooncast_A = QCheckBox('Tooncast PAN')
        
        # Row distribution for checkboxes
        self.checkBoxes_firstRow =  [self.checkBox_tnt, 
                                     self.checkBox_tnt_A, 
                                     self.checkBox_tnt_B, 
                                     self.checkBox_tnt_C, 
                                     self.checkBox_tnt_D, 
                                     self.checkBox_tnt_series, 
                                     self.checkBox_tnt_series_A, 
                                     self.checkBox_tnt_series_B,
                                     self.checkBox_space,
                                     self.checkBox_space_A,
                                     self.checkBox_space_B,
                                     self.checkBox_space_C]
        
        self.checkBoxes_secondRow = [self.checkBox_tbs,
                                     self.checkBox_tbs_A,
                                     self.checkBox_tbs_B,
                                     self.checkBox_tcm,
                                     self.checkBox_tcm_A,
                                     self.checkBox_tcm_B,
                                     self.checkBox_isat,
                                     self.checkBox_isat_A,
                                     self.checkBox_trutv,
                                     self.checkBox_trutv_A,
                                     self.checkBox_glitz,
                                     self.checkBox_glitz_A]
        
        self.checkBoxes_thirdRow = [self.checkBox_cn,
                                    self.checkBox_cn_A,
                                    self.checkBox_cn_B,
                                    self.checkBox_cn_C,
                                    self.checkBox_cnito,
                                    self.checkBox_cnito_A,
                                    self.checkBox_tooncast,
                                    self.checkBox_tooncast_A]
        
        self.checkBox_all = [self.checkBox_tnt, self.checkBox_tnt_series, self.checkBox_space, self.checkBox_tbs, self.checkBox_tcm, self.checkBox_isat, self.checkBox_trutv, self.checkBox_glitz, self.checkBox_cn, self.checkBox_cnito, self.checkBox_tooncast]
        self.checkBoxes_tnt = [self.checkBox_tnt_A, self.checkBox_tnt_B, self.checkBox_tnt_C, self.checkBox_tnt_D]
        self.checkBoxes_tnt_series = [self.checkBox_tnt_series_A, self.checkBox_tnt_series_B]
        self.checkBoxes_space = [self.checkBox_space_A, self.checkBox_space_B, self.checkBox_space_C]
        self.checkBoxes_tbs = [self.checkBox_tbs_A, self.checkBox_tbs_B]
        self.checkBoxes_tcm = [self.checkBox_tcm_A, self.checkBox_tcm_B]
        self.checkBoxes_isat = [self.checkBox_isat_A]
        self.checkBoxes_trutv = [self.checkBox_trutv_A]
        self.checkBoxes_glitz = [self.checkBox_glitz_A]
        self.checkBoxes_cn = [self.checkBox_cn_A, self.checkBox_cn_B, self.checkBox_cn_C]
        self.checkBoxes_cnito = [self.checkBox_cnito_A]
        self.checkBoxes_tooncast = [self.checkBox_tooncast_A]
        
        self.full_list = self.checkBoxes_tnt+self.checkBoxes_tnt_series+self.checkBoxes_space+ self.checkBoxes_tbs + self.checkBoxes_tcm + self.checkBoxes_isat + self.checkBoxes_trutv + self.checkBoxes_glitz + self.checkBoxes_cn + self.checkBoxes_cnito + self.checkBoxes_tooncast
    
        for index, item in enumerate(self.checkBoxes_firstRow):
            self.tab1.layout.addWidget(item, index+4, 0)

        for index, item in enumerate(self.checkBoxes_secondRow):
            self.tab1.layout.addWidget(item, index+4, 1)
            
        for index, item in enumerate(self.checkBoxes_thirdRow):
            self.tab1.layout.addWidget(item, index+4, 2)
        
        self.plainTextEdit = QPlainTextEdit()
        self.plainTextEdit.setPlaceholderText("No hay información disponible para su visualización")
        self.plainTextEdit.setReadOnly(True)
        
        self.separatorLine2 = QFrame(frameShape=QFrame.HLine)    
        self.tab1.layout.addWidget(self.separatorLine2, 17, 0, 1, 4)
        
        self.pbar = QProgressBar(self)
        self.pbar.setGeometry(30, 40, 200, 25)
        self.tab1.layout.addWidget(self.pbar, 20, 0, 2, 4)

        self.tab1.layout.addWidget(self.plainTextEdit, 22, 0, 5, 4)
            
    # Button actions
    
    def open_folder_action(self):
        os.system('start Procesados')

    def download_action(self):    
        array_result = []
        for i, v in enumerate(self.full_list):
            if self.full_list[i].isChecked():
                array_result.append(v.text())
        
        if not array_result:
            return self.show_info_messagebox("warning" ,"No se ha seleccionado ningún canal.")
        collection_path = []
        amount = len(array_result)
        aux = 0
        self.pbar.setValue(0)
        for x in array_result:
            federation_code = channel_options.get(x)
            file_path = f'./Procesados/{federation_code}-schedule'
            request = self.generate_request_dailylog(federation_code)
            path = self.save_excel_request(request, federation_code)
            filtered_data = self.check_excel_data(path)
            try:
                file_path = self.save_excel(filtered_data, f'{file_path}.xlsx')
            except Exception as e:
                print(e)
                self.show_info_messagebox("warning" ,"No se ha podido guardar/actualizar el archivo con ruta: "+file_path+". Verifique que el archivo no este abierto y/o tenga permisos.")

            collection_path.append(file_path)
            aux += 1
            self.pbar.setValue(int((aux/amount)*100))
            QCoreApplication.processEvents()
        
        if(self.checkBox_close.isChecked()):
            self.close_excel_logs(collection_path)
        
        self.plainTextEdit.appendPlainText('___________________________________________________________________\n') 
        self.show_info_messagebox("information", "Descarga realizada. Revise la carpeta de procesados.")
        
    def show_info_messagebox(self, type, content):
        msg = QMessageBox()
        
        if(type  == 'warning'):
            msg.setIcon(QMessageBox.Warning)
        elif(type == 'information'):
            msg.setIcon(QMessageBox.Information)
    
        # setting message for Message Box
        msg.setText(content)
        
        # setting Message box window title
        msg.setWindowTitle("Información")
        
        # declaring buttons on Message Box
        msg.setStandardButtons(QMessageBox.Ok)
        
        # start the app
        retval = msg.exec_()
    
    def save_excel(self, dataframe, filename):
        file_path = f"{filename}"
        dataframe.to_excel(file_path,
                sheet_name='GRID', index=None)  
        data = pd.read_excel(file_path, header=None, index_col=0)
        data.to_excel(file_path,
                sheet_name='GRID')
        return file_path

    def close_excel_logs(self, collection_path):
        for item in collection_path:
            data = pd.read_excel(item, skiprows=1)
            data["Log Status"] = 'Log cerrado'
            self.save_excel(data, item)
        print('Los logs han sido cerrados.\n')
        return data

    def generate_request_dailylog(self, federation_code):
        params = {'feedId': federation_code}
        try:
            request = requests.get(url_base, params=params)
        except requests.exceptions.RequestException as e:
            print('No ha sido posible realizar la conexión con Turner. Verifique su conexión o reporte el problema.')
            raise SystemExit()
        return request   
    
    def save_excel_request(self, request, federation_code):
        path = f'./Descargas/{federation_code}-schedule.xls'
        with open(path,'wb') as f:
            f.write(request.content)
        return path

    def check_excel_data(self, path):
        
        date1 = self.dateedit.date().toPyDate().strftime("%d-%m-%Y")
        date2 = self.dateedit2.date().toPyDate().strftime("%d-%m-%Y")
        
        data = pd.read_excel(path, skiprows=1)
        
        try:
            filtered_data = data.set_index('Schedule Date')[date1:date2]
        except Exception as e:
            print(e)
            self.show_info_messagebox("error" ,"No se ha encontrado información en el último día de este canal.")
            
        filtered_data.reset_index(inplace=True)

        amount_open_logs = 0
        amount_closed_logs = 0
        amount_empty_rows = 0
        
        for item in filtered_data["Log Status"]:
            if(item == 'Log abierto'):
                amount_open_logs += 1
            elif(item == 'Log cerrado'):
                amount_closed_logs += 1

        filtered_data.loc[filtered_data['Title Name'].astype(str).str.strip() == '', 'Title Name'] = filtered_data['Title Name English']
        filtered_data.loc[(filtered_data['Episode Name'].astype(str).str.strip() != '') & (filtered_data['Title Name'].astype(str).str.strip() == '') , 'Title Name'] = filtered_data['Episode Name']
        filtered_data.loc[(filtered_data['Episode Name English'].astype(str).str.strip() != '') & (filtered_data['Title Name'].astype(str).str.strip() == '') & (filtered_data['Episode Name'].astype(str).str.strip() == ''), 'Title Name'] = filtered_data['Episode Name English']
        
        try: 
            for item in filtered_data["Title Name"]:
                if(item.strip() == ''):
                    amount_empty_rows += 1    
        except:
            amount_empty_rows += 1  
            pass

        self.plainTextEdit.appendPlainText(f'Archivo: {path}\n> Logs abiertos: {amount_open_logs}\n> Logs cerrados: {amount_closed_logs}\n> Filas vacías: {amount_empty_rows}\n') 
                
        return filtered_data 
    
    def save_excel(self, dataframe, filename):
        file_path = f"{filename}"
        dataframe.to_excel(file_path,
                sheet_name='GRID', index=None)  
        data = pd.read_excel(file_path, header=None, index_col=0)
        data.to_excel(file_path,
                sheet_name='GRID')
        return file_path
    
    # State changes from checkboxes
    def on_stateChanged_all(self, state):
        for checkBox in self.checkBox_all:
            checkBox.setCheckState(state)
            
    def on_stateChanged_tnt(self, state):
        for checkBox in self.checkBoxes_tnt:
            checkBox.setCheckState(state)
            
    def on_stateChanged_tnt_series(self, state):
        for checkBox in self.checkBoxes_tnt_series:
            checkBox.setCheckState(state)
            
    def on_stateChanged_space(self, state):
        for checkBox in self.checkBoxes_space:
            checkBox.setCheckState(state)

    def on_stateChanged_tbs(self, state):
        for checkBox in self.checkBoxes_tbs:
            checkBox.setCheckState(state)

    def on_stateChanged_tcm(self, state):
        for checkBox in self.checkBoxes_tcm:
            checkBox.setCheckState(state)

    def on_stateChanged_isat(self, state):
        for checkBox in self.checkBoxes_isat:
            checkBox.setCheckState(state)

    def on_stateChanged_trutv(self, state):
        for checkBox in self.checkBoxes_trutv:
            checkBox.setCheckState(state)
            
    def on_stateChanged_glitz(self, state):
        for checkBox in self.checkBoxes_glitz:
            checkBox.setCheckState(state)
            
    def on_stateChanged_cn(self, state):
        for checkBox in self.checkBoxes_cn:
            checkBox.setCheckState(state)

    def on_stateChanged_cnito(self, state):
        for checkBox in self.checkBoxes_cnito:
            checkBox.setCheckState(state)
            
    def on_stateChanged_tooncast(self, state):
        for checkBox in self.checkBoxes_tooncast:
            checkBox.setCheckState(state)
            
if __name__ == '__main__':
    QApplication.setAttribute(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)

    app = QApplication(sys.argv)
    myApp = MyApp()
    myApp.show()

    try:
        sys.exit(app.exec_())
    except SystemExit:
        print('Closing Window...')