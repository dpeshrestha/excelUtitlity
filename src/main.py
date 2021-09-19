from PySide2.QtWidgets import *
from PySide2.QtCore import *
from PySide2.QtGui import  *
import sys
import os
import numpy as np
import pandas as pd
from datetime import  datetime
import itertools
import pyreadstat
import xlsxwriter

rootPath = os.path.dirname(os.path.abspath(__file__))

def resource_path(relative_path):
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.title = 'Excel tools'
        self.left = 0
        self.top = 0
        self.width = 792
        self.height = 610
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)
        #
        # self.tab_widget = TabWidget(self)
        self.centralWidget = QWidget(self)
        self.mainLayout = QHBoxLayout()
        self.utilityListView = QListView(self)
        self.listModel = ListModel(['Excel Compare',])
        self.utilityListView.setModel(self.listModel)
        self.utilityListView.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.utilityListView.setMaximumWidth(200)
        self.utilityListView.selectionModel().selectionChanged.connect(self.stepSelected)
        self.rightWidget = QWidget(self)
        self.rightlayout = QVBoxLayout()

        self.rightWidget.setLayout(self.rightlayout)
        splitter = QSplitter(self)
        splitter.setOrientation(Qt.Horizontal)
        splitter.addWidget(self.utilityListView)
        splitter.addWidget(self.rightWidget)
        splitter.setChildrenCollapsible(False)
        self.mainLayout.addWidget(splitter)
        self.placeHolderWidget = QLabel("Select a utility from the list",parent=self.rightWidget)
        self.rightlayout.addWidget(self.placeHolderWidget,alignment=Qt.AlignCenter)
        self.centralWidget.setLayout(self.mainLayout)
        self.setCentralWidget(self.centralWidget)
        self.show()

    def stepSelected(self):
        if len(self.utilityListView.selectedIndexes()) < 1: return
        step = self.listModel.itemData(self.utilityListView.selectedIndexes()[0])[0]

        for i in range(self.rightlayout.count()):
            self.rightlayout.itemAt(i).widget().hide()
        if step == 'Excel Compare':
            self.excelComapre =  ExcelCompare(self)
            self.rightlayout.addWidget(self.excelComapre)
    # Creating tab widgets


class TabWiget(QWidget):
    def __init__(self, parent):
        super(TabWidget, self).__init__(parent)
        self.setupUi()

    def setupUi(self):
        self.layout = QVBoxLayout(self)

        # Initialize tab screen
        self.tabs = QTabWidget()
        self.tab1 = QWidget()
        self.tab2 = QWidget()
        self.tabs.resize(792, 610)

        # Add tabs
        self.tabs.addTab(self.tab1, "Excel Compare")
        self.tabs.addTab(self.tab2, "Specs vs SAS")

        # Create first tab
        self.tab1Layout = QVBoxLayout(self)
        self.excelWidget = ExcelCompare()
        self.tab1Layout.addWidget(self.excelWidget)
        self.tab1.setLayout(self.tab1Layout)

        # Create second tab
        self.tab2Layout = QVBoxLayout(self)
        self.sasWidget = SpecsSAS()
        self.tab2Layout.addWidget(self.sasWidget)
        self.tab2.setLayout(self.tab2Layout)

        # Add tabs to widget
        self.layout.addWidget(self.tabs)
        self.setLayout(self.layout)


class ListModel(QAbstractListModel):
    def __init__(self,data=[]):
        super(ListModel, self).__init__()
        self._data = data

    def flags(self, index):
        defaultFalgs = super(ListModel, self).flags(index)
        if index.isValid():
            return Qt.ItemIsDragEnabled | Qt.ItemIsDropEnabled | defaultFalgs
        else:
            return  Qt.ItemIsDropEnabled | defaultFalgs

    def data(self, index, role:Qt.DisplayRole):
        if role == Qt.DisplayRole:
            return str(self._data[index.row()])

    def rowCount(self, parent):
        return  len(self._data)


class DragDropListView(QListView):
    def __init__(self,parent=None):
        super(DragDropListView, self).__init__(parent)

        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.setDragEnabled(True)
        self.setAcceptDrops(True)
        self.showDropIndicator()
        self.setDropIndicatorShown(True)
        self.setDragDropMode(QAbstractItemView.DragDrop)

    def dragEnterEvent(self, event):
        if self.objectName() == 'excludedListView' or self.objectName() == 'includedListView' :
            if event.source().objectName()=='excludedListView' or event.source().objectName()=='includedListView':
                event.setAccepted(True)
                return

        if self.objectName() == 'keysListView' or self.objectName() == 'varListView':
            if event.source().objectName()=='keysListView' or event.source().objectName()=='varListView':
                event.setAccepted(True)
                return

    def dragMoveEvent(self, e):
        super(DragDropListView, self).dragMoveEvent(e)
        e.setAccepted(True)

    def dropEvent(self, event):
        index = self.indexAt(event.pos())
        source = event.source()
        sourceId = event.source().selectedIndexes()[0]
        if not sourceId.isValid():
            return

        tempVal = source.model()._data[sourceId.row()]
        if self==source:
            self.model().beginRemoveRows(QModelIndex(), index.row(), index.row())
            del self.model()._data[sourceId.row()]
            self.model().endRemoveRows()
        else:
            source.model().beginRemoveRows(QModelIndex(), index.row(), index.row())
            del source.model()._data[sourceId.row()]
            source.model().endRemoveRows()

        if self.dropIndicatorPosition() == QAbstractItemView.BelowItem:
            destIndex = index.row() + 1
        else:
            destIndex = index.row()

        self.model().beginInsertRows(QModelIndex(),destIndex,destIndex)

        if self.dropIndicatorPosition() == QAbstractItemView.OnViewport:
            self.model()._data.append(tempVal)
        else:
            self.model()._data.insert(destIndex,tempVal)
        self.model().endInsertRows()

        if self.objectName() == 'keysListView' or source.objectName() == 'keysListView':
            if len(self.parent().includedListView.selectedIndexes())>0:
                sheetName = self.parent().includedListView.selectedIndexes()[0].data()
                self.parent().keys[sheetName] = self.parent().keysListView.model()._data
        if self!=source and self.objectName() == 'includedListView':
            self.parent().keys[tempVal] = []
        if self!=source and source.objectName() == 'includedListView':
            del self.parent().keys[tempVal]
        self.clearSelection()


class customQMessageBox(QDialog):
    def __init__(self):
        super().__init__()
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

    def okOptions(self):
        self.optionsOK = True
        self.close()

    def setText(self,text):
        self.text.setText(text)


class ExcelCompare(QWidget):
    def __init__(self,parent=None):
        super(ExcelCompare, self).__init__(parent)
        self.keys = {}
        self.oldVars = {}
        self.setupUi(self)
        self.oldPushButton.clicked.connect(self.setOldFilePath)
        self.newPushButton.clicked.connect(lambda _:self.newLineEdit.setText(
            customFileDialog.getOpenFileUrl(self,"Select New Excel File",QUrl(rootPath),(u"*.xls *.xlsx"))[0].path().strip('/')))
        self.keysPushButton.clicked.connect(self.populateKeyListView)
        self.logPushButton.clicked.connect(lambda _: self.logLineEdit.setText(
            customFileDialog.getExistingDirectoryUrl(self,'Select Folder to Store Findings',QUrl(rootPath)).path().strip('/')))
        # self.oldLineEdit.keyPressEvent = lambda e:self.isEnterPressed(e)
        # self.oldLineEdit.mouseReleaseEvent = self.isEnterPressed
        self.oldLineEdit.returnPressed.connect(self.loadOldListView)
        self.keysLineEdit.returnPressed.connect(self.populateKeyListView)
        self.comparePushButton.clicked.connect(self.compare)

    def populateKeyListView(self):
        self.keysListView.setModel(ListModel([]))
        self.keysListView.setEnabled(True)
        self.includedListView.setEnabled(True)
        self.varListView.setEnabled(True)
        self.excludedListView.setEnabled(True)
        self.keysLineEdit.setText(
            customFileDialog.getOpenFileUrl(self, "Select Keys File", QUrl(rootPath), (u"*.csv"))[0].path().strip('/'))
        try:
            self.keys = pd.read_csv(self.keysLineEdit.text()).to_dict('list')
        except:
            if self.keysLineEdit.text() == '':
                self.keys = {sheet:[] for sheet in self.includedListView.model()._data}
            msg = customQMessageBox()
            msg.setText("Invalid keys file")
            msg.exec_()
            return
        self.keysListView.setDisabled(True)
        self.includedListView.setDisabled(True)
        self.varListView.setDisabled(True)
        self.excludedListView.setDisabled(True)

    def compare(self):
        msg = customQMessageBox()
        # msg.setButtons(1)
        msg.setWindowTitle("Warning")
        if self.oldLineEdit.text() == '' or not os.path.exists(self.oldLineEdit.text()):
            msg.setText('Older File doesn’t exist')
            msg.exec_()
            return
        if self.newLineEdit.text() == '' or not os.path.exists(self.newLineEdit.text()):
            msg.setText('Newer File doesn’t exist')
            msg.exec_()
            return
        if self.logLineEdit.text() == '' or not os.path.exists(self.logLineEdit.text()):
            msg.setText('Findings Folder doesn’t exist')
            msg.exec_()
            return


        if self.keysLineEdit.text() != '':#keys path empty
            if not os.path.exists(self.keysLineEdit.text()): #keys file doesnt exist
                msg.setText('Keys File doesn’t exist')
                msg.exec_()
                return
            else:#keyfile specified
                #read keys and populate dict
                try:
                   self.keys = pd.read_csv(self.keysLineEdit.text())
                except:
                    msg = customQMessageBox()
                    msg.setText("Cannot read keys file")
                    msg.exec_()
                    return

                self.keys = {sheet: self.keys[sheet].dropna().values for sheet in self.keys.columns}

                #check if vars from keys file exist specified sheets from oldvars
                for sheet in self.keys.keys():
                    if self.oldVarsAll.get(sheet) is not None:
                        self.oldVars[sheet] = self.oldVarsAll[sheet]


        else:#key file is empty
            sheetsWithoutKeys = [sheet for sheet in self.keys.keys() if len(self.keys[sheet]) < 1]
            if len(sheetsWithoutKeys) >0:
                msg.setText("Keys not specified for sheet(s):\n"+"\n".join(sheetsWithoutKeys))
                msg.exec_()
                return
            for sheet in self.includedListView.model()._data:
                self.oldVars[sheet] = self.oldVarsAll[sheet]

        sheetsNotInOld = set(self.keys.keys()).difference(self.oldVars.keys())
        keysNotInOld = [
            [sheet + "." + key for key in set(self.keys[sheet]).difference(set(self.oldVars[sheet].columns.to_list()))]
            for sheet in self.keys.keys() if sheet not in sheetsNotInOld]
        keysNotInOld += [[sheet + "." + key for key in self.keys[sheet]] for sheet in sheetsNotInOld]
        keysNotInOld = '\n'.join(["\n".join(keys) for keys in keysNotInOld if len(keys) > 0])
        if len(keysNotInOld) > 0:
            msg.setText(f"Variable(s):\n{keysNotInOld}\nspecified in key file is missing in older file.")
            msg.exec_()
            return

        # compare Excel Files
        date = datetime.now().strftime("%Y-%m-%d %H-%M")

        logFile =  os.path.join(self.logLineEdit.text(),f"compare_{date}.xlsx")

        majorUpdates = pd.DataFrame(columns=['Message','Details'])
        recordUpdates = pd.DataFrame(columns=['Message', 'SheetName', 'Key Variables', 'Key Values'])
        valueUpdates =  pd.DataFrame(columns=['SheetName', 'Key Variables', 'Key Values','Column Name','Old Value','New Value'])
        try:

            self.newVars = pd.read_excel(self.newLineEdit.text(),sheet_name=None)
            for key in self.newVars.keys():
                self.newVars[key] = self.newVars[key].replace(np.nan,'')

        except:
            msg = customQMessageBox()
            msg.setText("Cannot read new excel file")
            msg.exec_()
            return
        for sheet in self.newVars.keys():
            self.newVars[sheet].columns = [x.upper() for x in self.newVars[sheet].columns.to_list()]

        droppedSheets = set(self.oldVarsAll.keys()).difference(set(self.newVars.keys()))
        if len(droppedSheets)>0:majorUpdates = majorUpdates.append(pd.DataFrame({"Message": "Dropped Sheet", 'Details': list(droppedSheets)}),ignore_index=True)


        newSheets = set(self.newVars.keys()).difference(set(self.oldVarsAll.keys()))
        if len(newSheets)>0:majorUpdates = majorUpdates.append(pd.DataFrame({"Message": "New Sheet", 'Details': list(newSheets)}),ignore_index=True)

        commonSheets = set(self.newVars.keys()).intersection(set(self.oldVars.keys()))
        droppedColumns = [[sheet+"."+x for x in list(set(self.oldVars[sheet].columns.to_list()).difference(set(self.newVars[sheet].columns.to_list())))] for sheet in commonSheets]
        droppedColumns = [col for sublist in droppedColumns for col in sublist]
        if len(droppedColumns) > 0: majorUpdates= majorUpdates.append(pd.DataFrame({"Message": "Dropped Variable", 'Details': list(droppedColumns)}),
                                                   ignore_index=True)

        newColumns = [[sheet + "." + x for x in list(
            set(self.newVars[sheet].columns.to_list()).difference(set(self.oldVars[sheet].columns.to_list())))] for
                          sheet in commonSheets]
        newColumns = [col for sublist in newColumns for col in sublist]
        if len(newColumns) > 0:majorUpdates= majorUpdates.append(
            pd.DataFrame({"Message": "New Variable", 'Details': list(newColumns)}),
            ignore_index=True)

        commonSheets = [sheet for sheet in commonSheets if sheet in self.oldVars.keys()]
        commonVariables = {sheet: set(self.oldVars[sheet].columns.to_list()).intersection(
            self.newVars[sheet].columns.to_list()) for sheet in commonSheets}
        for sheet in commonVariables.keys():
            keys = set(self.keys[sheet]).intersection(commonVariables[sheet])
            keys = [key for key in self.keys[sheet] if key in keys]
            if len(keys) < 1:
                continue
            countOld = self.oldVars[sheet].groupby(keys).size()
            countOld = pd.DataFrame(countOld).reset_index().rename(columns={0: 'oldCount'})
            countNew = self.newVars[sheet].groupby(keys).size()
            countNew = pd.DataFrame(countNew).reset_index().rename(columns={0: 'newCount'})
            merged = countOld.merge(countNew,how='outer').fillna(0)
            # merged['status'] = ''
            droppedRecords = np.where(merged['oldCount'] > merged['newCount'])[0]
            newRecords = np.where(merged['oldCount'] < merged['newCount'])[0]
            if len(droppedRecords)>0:
                droppedRecords = merged.iloc[droppedRecords]
                recordUpdates = recordUpdates.append(pd.DataFrame(
                    {'Message':'Dropped Record','SheetName':sheet,'Key Variables':", ".join(keys),
                     "Key Values":[", ".join(x) for x in droppedRecords[keys].astype(str).apply(lambda x:[x.name+"="+ w for w in x]).values]}),ignore_index=True)
            if len(newRecords)>0:
                newRecords = merged.iloc[newRecords]
                recordUpdates = recordUpdates.append(pd.DataFrame(
                    {'Message': 'New Record', 'SheetName': sheet, 'Key Variables': ", ".join(keys),
                     "Key Values":[", ".join(x) for x in newRecords[keys].astype(str).apply(lambda x:[x.name+"="+w for w in x]).values]}),ignore_index=True)

            suffix_stamp = str(datetime.now().timestamp())[:2]
            merged = self.oldVars[sheet].merge(self.newVars[sheet],on=keys,how='outer',
                                               suffixes=(f'_old_{suffix_stamp}',f'_new_{suffix_stamp}'),
                                               indicator=True).astype(str).fillna('-').replace(np.nan,'')
            merged = merged[merged['_merge'] == 'both']

            old_cols = [col for col in merged.columns if f'_old_{suffix_stamp}' in col]
            new_cols = [col for col in merged.columns if f'_new_{suffix_stamp}' in col]
            for x,y in zip(old_cols,new_cols):
                changed = np.where(merged[x]!=merged[y])[0]
                if len(changed) >0:
                    valueUpdates = valueUpdates.append(pd.DataFrame({'SheetName':sheet,'Key Variables':", ".join(keys),"Key Values":[", ".join(x) for x in merged.iloc[changed][keys].astype(str).apply(lambda x:[x.name+"="+w for w in x]).values],"Column Name":x[:-7],"Old Value":merged[x].iloc[changed].values,"New Value":merged[y].iloc[changed].values}),ignore_index=True)

        writer = pd.ExcelWriter(logFile,engine='xlsxwriter')
        dfs = {'Major Updates':majorUpdates,'Record Updates':recordUpdates,'Value Updates':valueUpdates}

        for sheetname,df in dfs.items():
            if df.shape[0] < 1:
                df[df.columns[0]] = ['No data meeting such criteria']
            df.to_excel(writer,sheet_name=sheetname,index=False)

            worksheet = writer.sheets[sheetname]
            # format = workbook.add_format({'text_wrap':True})
            for i, col in enumerate(df.columns):
                col_len = df[col].astype(str).str.len().mean()
                col_len = max(col_len, len(col)) + 2
                col_len = len(col)*3 if col_len > len(col)*3 else col_len
                worksheet.set_column(i,i,col_len)

        writer.save()

        df = pd.DataFrame()
        for sheet in self.keys.keys():
            df = pd.concat([df,pd.DataFrame({sheet:self.keys[sheet]})],ignore_index=True, axis=1)
        df.columns = self.keys.keys()
        df.to_csv(os.path.join(self.logLineEdit.text(),f"keys_{date}.csv"),index=False)
        msg.setWindowTitle("Completed")
        msg.setText(f'Comparison Completed. Check the files compare_{date}.xlsx in selected folder for findings.')
        msg.exec_()
        self.parent().parent().parent().parent().tab1Layout.removeWidget(self)
        self.parent().parent().parent().parent().tab1Layout.addWidget(ExcelCompare())
        return

    def populateVarListView(self):
        self.keysListView.model()._data = []
        indexList = self.includedListView.selectedIndexes()
        if len(indexList) >0:
            index = indexList[0]
            if index.isValid():
                self.varListView.setModel(ListModel(self.oldVarsAll[index.data()].columns.to_list()))
                if len(self.keys[index.data()]) > 0:
                    if index.data() in self.keys.keys():
                        self.keysListView.setModel(ListModel(self.keys[index.data()]))
                    self.varListView.model()._data = list(set(self.varListView.model()._data).difference(set(self.keysListView.model()._data)))
                else:
                    self.keysListView.setModel(ListModel([]))

    def loadOldListView(self):
        try:
            self.oldVarsAll = pd.read_excel(self.oldLineEdit.text(),sheet_name=None)
            for key in self.oldVarsAll.keys():
                self.oldVarsAll[key] = self.oldVarsAll[key].replace(np.nan,'')
        except:
            msg = customQMessageBox()
            msg.setText("Cannot read old excel File")
            msg.exec_()
            return
        for sheet in self.oldVarsAll.keys():
            self.oldVarsAll[sheet].columns =  [str(x).upper() for x in self.oldVarsAll[sheet].columns]
        self.excludedListModel = ListModel(list(self.oldVarsAll.keys()))
        self.excludedListView.setModel(self.excludedListModel)
        self.includedListView.setModel(ListModel([]))
        self.keysListView.setModel(ListModel([]))
        self.varListView.setModel(ListModel([]))
        self.includedListView.selectionModel().selectionChanged.connect(self.populateVarListView)
        # self.keys = {sheet:[] for sheet in self.includedListView.model()._data}

    def setOldFilePath(self):
        self.oldLineEdit.setText(customFileDialog.getOpenFileUrl(self,"Select Old Excel File",QUrl(rootPath),(u"*.xls *.xlsx"))[0].path().strip('/'))
        self.oldLineEdit.setFocus()
        if self.oldLineEdit.text() !='':
            self.loadOldListView()

    def setupUi(self, Form):
        if not Form.objectName():
            Form.setObjectName(u"Form")
        Form.resize(1189, 743)
        self.verticalLayout_7 = QVBoxLayout(Form)
        self.verticalLayout_7.setObjectName(u"verticalLayout_7")
        self.verticalLayout_5 = QVBoxLayout()
        self.verticalLayout_5.setObjectName(u"verticalLayout_5")
        self.horizontalLayout_3 = QHBoxLayout()
        self.horizontalLayout_3.setObjectName(u"horizontalLayout_3")
        self.label = QLabel(Form)
        self.label.setObjectName(u"label")
        self.label.setMinimumSize(QSize(110, 0))

        self.horizontalLayout_3.addWidget(self.label)

        self.oldLineEdit = QLineEdit(Form)
        self.oldLineEdit.setObjectName(u"oldLineEdit")

        self.horizontalLayout_3.addWidget(self.oldLineEdit)

        self.oldPushButton = QPushButton(Form)
        self.oldPushButton.setObjectName(u"browseButton")
        self.oldPushButton.setMaximumSize(QSize(30, 30))

        self.horizontalLayout_3.addWidget(self.oldPushButton)

        self.verticalLayout_5.addLayout(self.horizontalLayout_3)

        self.horizontalLayout_2 = QHBoxLayout()
        self.horizontalLayout_2.setObjectName(u"horizontalLayout_2")
        self.label_2 = QLabel(Form)
        self.label_2.setObjectName(u"label_2")
        self.label_2.setMinimumSize(QSize(110, 0))

        self.horizontalLayout_2.addWidget(self.label_2)

        self.newLineEdit = QLineEdit(Form)
        self.newLineEdit.setObjectName(u"newLineEdit")

        self.horizontalLayout_2.addWidget(self.newLineEdit)

        self.newPushButton = QPushButton(Form)
        self.newPushButton.setObjectName(u"browseButton")
        self.newPushButton.setMaximumSize(QSize(30, 30))

        self.horizontalLayout_2.addWidget(self.newPushButton)

        self.verticalLayout_5.addLayout(self.horizontalLayout_2)

        self.horizontalLayout = QHBoxLayout()
        self.horizontalLayout.setObjectName(u"horizontalLayout")
        self.label_3 = QLabel(Form)
        self.label_3.setObjectName(u"label_3")
        self.label_3.setMinimumSize(QSize(110, 0))

        self.horizontalLayout.addWidget(self.label_3)

        self.keysLineEdit = QLineEdit(Form)
        self.keysLineEdit.setObjectName(u"keysLineEdit")

        self.horizontalLayout.addWidget(self.keysLineEdit)

        self.keysPushButton = QPushButton(Form)
        self.keysPushButton.setObjectName(u"browseButton")
        self.keysPushButton.setMaximumSize(QSize(30, 30))

        self.horizontalLayout.addWidget(self.keysPushButton)

        self.verticalLayout_5.addLayout(self.horizontalLayout)

        self.verticalSpacer = QSpacerItem(17, 45, QSizePolicy.Minimum, QSizePolicy.Fixed)

        self.verticalLayout_5.addItem(self.verticalSpacer)

        self.label_8 = QLabel(Form)
        self.label_8.setObjectName(u"label_8")
        self.label_8.setAlignment(Qt.AlignCenter)

        self.verticalLayout_5.addWidget(self.label_8)

        self.verticalLayout_7.addLayout(self.verticalLayout_5)

        self.verticalLayout_6 = QVBoxLayout()
        self.verticalLayout_6.setObjectName(u"verticalLayout_6")
        self.horizontalLayout_5 = QHBoxLayout()
        self.horizontalLayout_5.setObjectName(u"horizontalLayout_5")
        self.verticalLayout = QVBoxLayout()
        self.verticalLayout.setObjectName(u"verticalLayout")
        self.label_4 = QLabel(Form)
        self.label_4.setObjectName(u"label_4")

        self.verticalLayout.addWidget(self.label_4)

        self.excludedListView = DragDropListView(Form)
        self.excludedListView.setObjectName(u"excludedListView")

        self.verticalLayout.addWidget(self.excludedListView)

        self.horizontalLayout_5.addLayout(self.verticalLayout)

        self.verticalLayout_2 = QVBoxLayout()
        self.verticalLayout_2.setObjectName(u"verticalLayout_2")
        self.label_5 = QLabel(Form)
        self.label_5.setObjectName(u"label_5")

        self.verticalLayout_2.addWidget(self.label_5)

        self.includedListView = DragDropListView(Form)
        self.includedListView.setObjectName(u"includedListView")

        self.verticalLayout_2.addWidget(self.includedListView)

        self.horizontalLayout_5.addLayout(self.verticalLayout_2)

        self.horizontalSpacer_2 = QSpacerItem(30, 20, QSizePolicy.Fixed, QSizePolicy.Minimum)

        self.horizontalLayout_5.addItem(self.horizontalSpacer_2)

        self.verticalLayout_3 = QVBoxLayout()
        self.verticalLayout_3.setObjectName(u"verticalLayout_3")
        self.label_6 = QLabel(Form)
        self.label_6.setObjectName(u"label_6")

        self.verticalLayout_3.addWidget(self.label_6)

        self.keysListView = DragDropListView(Form)
        self.keysListView.setObjectName(u"keysListView")

        self.verticalLayout_3.addWidget(self.keysListView)

        self.horizontalLayout_5.addLayout(self.verticalLayout_3)

        self.verticalLayout_4 = QVBoxLayout()
        self.verticalLayout_4.setObjectName(u"verticalLayout_4")
        self.label_7 = QLabel(Form)
        self.label_7.setObjectName(u"label_7")

        self.verticalLayout_4.addWidget(self.label_7)

        self.varListView = DragDropListView(Form)
        self.varListView.setObjectName(u"varListView")

        self.verticalLayout_4.addWidget(self.varListView)

        self.horizontalLayout_5.addLayout(self.verticalLayout_4)

        self.verticalLayout_6.addLayout(self.horizontalLayout_5)

        self.horizontalLayout_4 = QHBoxLayout()
        self.horizontalLayout_4.setObjectName(u"horizontalLayout_4")
        self.label_10 = QLabel(Form)
        self.label_10.setObjectName(u"label_10")
        self.label_10.setMinimumSize(QSize(0, 0))
        self.label_10.setMaximumSize(QSize(180, 16777215))

        self.horizontalLayout_4.addWidget(self.label_10)

        self.logLineEdit = QLineEdit(Form)
        self.logLineEdit.setObjectName(u"logLineEdit")

        self.horizontalLayout_4.addWidget(self.logLineEdit)

        self.logPushButton = QPushButton(Form)
        self.logPushButton.setObjectName(u"browseButton")
        self.logPushButton.setMaximumSize(QSize(30, 30))

        self.horizontalLayout_4.addWidget(self.logPushButton)

        self.verticalLayout_6.addLayout(self.horizontalLayout_4)

        self.verticalLayout_7.addLayout(self.verticalLayout_6)

        self.comparePushButton = QPushButton(Form)
        self.comparePushButton.setObjectName(u"comparePushButton")
        self.comparePushButton.setMaximumSize(QSize(16777215, 16777215))

        self.verticalLayout_7.addWidget(self.comparePushButton,alignment=Qt.AlignHCenter)

        self.retranslateUi(Form)

        QMetaObject.connectSlotsByName(Form)

    # setupUi

    def retranslateUi(self, Form):
        Form.setWindowTitle(QCoreApplication.translate("Form", u"Form", None))
        self.label.setText(QCoreApplication.translate("Form", u"Select Older File", None))
        self.oldPushButton.setIcon(QIcon(resource_path("icons/folder.png")))
        self.label_2.setText(QCoreApplication.translate("Form", u"Select Newer File", None))
        self.newPushButton.setIcon(QIcon(resource_path("icons/folder.png")))
        self.label_3.setText(QCoreApplication.translate("Form", u"Select Keys File", None))
        self.keysPushButton.setIcon(QIcon(resource_path("icons/folder.png")))
        self.label_8.setText(QCoreApplication.translate("Form", u"Or Select  sheets/keys for comparison", None))
        self.label_4.setText(QCoreApplication.translate("Form", u"Excluded Sheets", None))
        self.label_5.setText(QCoreApplication.translate("Form", u"Included Sheets", None))
        self.label_6.setText(QCoreApplication.translate("Form", u"Keys", None))
        self.label_7.setText(QCoreApplication.translate("Form", u"Other Variables", None))
        self.label_10.setText(QCoreApplication.translate("Form", u"Select Folder to Store Findings             ", None))
        self.logPushButton.setIcon(QIcon(resource_path("icons/folder.png")))
        self.comparePushButton.setText(QCoreApplication.translate("Form", u"Compare", None))


class SpecsSAS(QWidget):
    def __init__(self,parent=None):
        super(SpecsSAS, self).__init__(parent)
        self.err = ''
        self.specsFile = {}
        self.setupUi(self)
        self.majorFindings = pd.DataFrame(columns=['Message', 'Details'])
        self.missingTests = pd.DataFrame(columns=['Subject', 'Visit', 'Timepoint', 'Test Code'])
        self.varLineEdit.setText('LB')

        self.specsPushButton.clicked.connect(lambda _:self.specsLineEdit.setText(customFileDialog.getOpenFileUrl(self,'Select Specs File',QUrl(rootPath),(u"*.xlsx"))[0].path().strip('/')))
        self.sasPushButton.clicked.connect(lambda _:self.sasLineEdit.setText(customFileDialog.getOpenFileUrl(self,'Select SAS Dataset',QUrl(rootPath),(u"*.sas7bdat *.xpt"))[0].path().strip('/')))
        self.logPushButton.clicked.connect(lambda _:self.logLineEdit.setText(customFileDialog.getExistingDirectoryUrl(self,'Select Folder to Store Findings',QUrl(rootPath)).path().strip('/')))
        self.comparePushButton.clicked.connect(self.compare)

    def compare(self):
        self.majorFindings = pd.DataFrame(columns=['Message', 'Details'])
        self.missingTests = pd.DataFrame(columns=['Subject', 'Visit', 'Timepoint', 'Test Code'])
        self.specsFile = {}
        msg = customQMessageBox()
        # msg.setButtons(1)
        msg.setWindowTitle("Warning")
        if not os.path.exists(self.specsLineEdit.text()):
            msg.setText('Excel specs file doesn’t exist')
            msg.exec_()
            return
        if not os.path.exists(self.sasLineEdit.text()):
            msg.setText('SAS dataset not provided or does not exist!')
            msg.exec_()
            return
        if not os.path.exists(self.logLineEdit.text()):
            msg.setText('Findings folder not specified')
            msg.exec_()
            return
        if self.subjLineEdit.text() == '':
            msg.setText('Subject variable not specified')
            msg.exec_()
            return
        if self.varLineEdit.text() == '':
            msg.setText('Variable prefix not specified. Usually it is first 2 letters of SDTM domain. .E.g. LB')
            msg.exec_()
            return
        try:
            specsFile = pd.read_excel(self.specsLineEdit.text(),sheet_name=None)
            for sheet in specsFile.keys():
                self.specsFile[sheet.lower()] = specsFile[sheet]\
                    .applymap(lambda x:x.strip() if isinstance(x,str) else x)\
                    .dropna(how='all')\
                    .dropna(axis=1,how='all')\
                    .replace(np.nan,'')

        except:
            msg.setText('Unable to read excel file')
            msg.exec_()
            return
        try:
            baseName = os.path.basename(self.sasLineEdit.text())
            if baseName.split('.')[-1] == 'xpt':
                self.sasData,self.meta = pyreadstat.read_xport(self.sasLineEdit.text())
            else:
                self.sasData,self.meta = pyreadstat.read_sas7bdat(self.sasLineEdit.text())
        except:
            msg.setText('Unable to read SAS data file')
            msg.exec_()
            return
        #Compare after validation
        # print(self.sasData)
        self.sasData = self.sasData.applymap(lambda x:x.strip() if isinstance(x,str) else x)
        self.structureCheck()
        if len(self.err) > 0:
            msg.setText(self.err)
            msg.exec_()
            return
        self.paramCheck()
        if len(self.err) > 0:
            msg.setText(self.err)
            msg.exec_()
            return
        self.visitCheck()
        if len(self.err) > 0:
            msg.setText(self.err)
            msg.exec_()
            return
        #structure sheet
        date = datetime.now().strftime("%Y-%m-%d %H-%M")
        logFile = os.path.join(self.logLineEdit.text(), f"compare_{date}.xlsx")
        writer = pd.ExcelWriter(logFile,engine='xlsxwriter')
        if self.flag:
            self.missingTests[self.missingTests.columns[0]] = ['Comparison skipped due to missing visit or timepoint vairable in either visit sheet  or data. See Major Findings sheet for details']

        dfs = {'Major Findings': self.majorFindings,'Missing Tests':self.missingTests}
        for sheetname, df in dfs.items():
            if df.shape[0] < 1:
                df[df.columns[0]] = ['No data meeting such criteria']
                df = df.replace(np.nan,'')


            df.to_excel(writer, sheet_name=sheetname, index=False)
            worksheet = writer.sheets[sheetname]
            for i, col in enumerate(df.columns):
                col_len = df[col].astype(str).str.len().mean()
                col_len = max(col_len, len(col)) + 2
                col_len = len(col)*3 if col_len > len(col)*3 else col_len
                worksheet.set_column(i, i, col_len)
        writer.save()
        msg.setWindowTitle("Completed")
        msg.setText(f'Comparison Completed. Check the files compare_{date}.xlsx in selected folder for findings.')
        msg.exec_()
        self.parent().parent().parent().parent().tab2Layout.removeWidget(self)
        self.parent().parent().parent().parent().tab2Layout.addWidget(SpecsSAS())

    def structureCheck(self):
        self.err = ''
        if 'structure' not in self.specsFile.keys():

            # print("Failed")
            self.err = "Structure sheet not found in file"
            self.majorFindings  = self.majorFindings.append(pd.DataFrame({"Message":['Structure sheet'],"Details":["Missing"]}))
            return
        print("Success")
        # self.majorFindings  = self.majorFindings.append(pd.DataFrame({"Message": ['Structure sheet'], "Details":[ "Found"]}))
        if len(set(['variable','variable label']).difference([x.lower() for x in self.specsFile['structure'].columns.to_list()]))>0:
            self.err = "Important variables ‘Variable’ and/or ‘Variable Label’ are missing in structure sheet! Comparison will stop."
            self.majorFindings  = self.majorFindings.append(pd.DataFrame({"Message": ['Structure Sheet with Variables ‘Variable’ & ’Variable Label’'], "Details": ["Missing"]}))
            # print("Failed")
            return
        else:
            print("Success")
            # self.majorFindings  = self.majorFindings.append(pd.DataFrame(
                # {"Message": ['Structure Sheet with Variables ‘Variable’ & ’Variable Label’'], "Details": ["Found"]}))

        varWithLabels = pd.DataFrame({"Variable": list(self.meta.column_names), "Variable Label": list(self.meta.column_labels)})
        merged = self.specsFile['structure'].merge(varWithLabels, on ="Variable", how='outer', indicator=True)
        missingVariables = merged[merged['_merge']=="left_only"]['Variable'].to_list()
        extraVariables = merged[merged['_merge']=="right_only"]['Variable'].to_list()
        # incorrectLabels = pd.DataFrame({"Variable":self.meta.column_names_to_labels.keys(),"Variable Label":self.meta.column_names_to_labels.values()})
        merged = merged.dropna(how='any')
        incorrectLabels = merged.iloc[np.where(merged['Variable Label_y']!=merged['Variable Label_x'])]['Variable'].to_list()
        self.majorFindings = self.majorFindings.append(pd.DataFrame(
            {"Message": ['Structure: Missing Variables'], "Details":", ".join(missingVariables)}))
        self.majorFindings = self.majorFindings.append(pd.DataFrame(
            {"Message": ['Structure: Extra Variables'], "Details": ", ".join(extraVariables)}))
        self.majorFindings = self.majorFindings.append(pd.DataFrame(
            {"Message": ['Structure: Variables with Incorrect Labels'], "Details": ", ".join(incorrectLabels)}))

    def paramCheck(self):
        self.err = ''
        if 'param' not in self.specsFile.keys():
            # print("Failed")
            self.err = "Param sheet not found in file"
            self.majorFindings = self.majorFindings.append(pd.DataFrame({"Message":["Param sheet"],"Details":["Missing"]}))
            return
        print("Success")
        # self.majorFindings = self.majorFindings.append(
        #     pd.DataFrame({"Message": ["Param sheet"], "Details": ["Found"]}))
        self.specsFile['param'] = self.specsFile['param'].dropna(axis=1,how='all') #remove this
        if len(set(self.specsFile['param'].columns).difference(self.sasData.columns)) >0 :
            # print("Failed")
            self.err = "Variable names in header row of param sheet not found in SAS dataset. Please fix the column headers with correct variable names before proceeding."
            self.majorFindings = self.majorFindings.append(
                pd.DataFrame({"Message": ["Param: Matching Variables in Data"], "Details": ["Missing"]}))
            return
        print("Success")
        # self.majorFindings = self.majorFindings.append(
        # pd.DataFrame({"Message": ["Param: Matching Variables in Data"], "Details": ["Found"]}))
        merged = self.specsFile['param'].merge(self.sasData[self.specsFile['param'].columns].drop_duplicates(), how='outer', indicator=True)
        missingInData = merged[merged['_merge']=='left_only'].astype(str).apply(lambda x:", ".join(x[:-1]),axis=1)
        self.majorFindings = self.majorFindings.append(
            pd.DataFrame({"Message": "Param: Value combination missing in Data "+f"({', '.join(merged.columns[:-1])})", "Details": missingInData.values}))
        missingInSpecs = merged[merged['_merge']=='right_only'].astype(str).apply(lambda x:", ".join(x[:-1]),axis=1)
        self.majorFindings = self.majorFindings.append(
            pd.DataFrame({"Message": "Param: Value combination missing in specs "+f"({', '.join(merged.columns[:-1])})", "Details":  missingInSpecs.values}))

    def visitCheck(self):
        self.err = ''
        if 'visit' not in self.specsFile.keys():
            # print("Failed")
            self.err = "Visit sheet not found in file"
            self.majorFindings = self.majorFindings.append(
                pd.DataFrame(
                    {"Message": ["Visit Sheet with VISIT & VISITNUM"],
                     "Details": ["Missing"]}))
            return
        print("Success")
        missingVars = [var  for var in  ['VISIT','VISITNUM'] if var not in self.specsFile['visit'].columns]

        if len(missingVars) > 0 :    # print("Failed")
            self.err = "VISIT & VISITNUM not found in specs file"
            self.majorFindings = self.majorFindings.append(
                pd.DataFrame(
                    {"Message": ["Visit: Variables VISITNUM and/or VISIT"],
                     "Details": ["Missing"]}))

            return
        print("Success")
        missingVars = [var for var in ['VISIT', 'VISITNUM'] if var not in self.sasData.columns]

        if len(missingVars) > 0:

            # print("Failed")
            self.err = "VISIT & VISITNUM not found in sas file"
            self.majorFindings = self.majorFindings.append(
                pd.DataFrame(
                    {"Message": ["Visit: Variables VISITNUM and/or VISIT"],
                     "Details": ["Missing"]}))

            return
        print("Success")
        # self.majorFindings = self.majorFindings.append(
            # pd.DataFrame(
                # {"Message": ["Visit Sheet with VISIT & VISITNUM"],
                 # "Details": ["Found"]}))



        merged = self.specsFile['visit'].merge(self.sasData[["VISIT","VISITNUM"]].drop_duplicates(),how='outer',indicator=True)
        missingInData = merged[merged['_merge'] == 'left_only'][['VISIT', 'VISITNUM']]
        missingInData = missingInData.astype(str).apply(lambda x: "/".join(x), axis=1)
        self.majorFindings = self.majorFindings.append(
            pd.DataFrame(
                {"Message": "Visit: Visit/VISITNUM Combination Missing in Data",
                 "Details":  list(set(missingInData.values))}))

        missingInSpecs = merged[merged['_merge'] == 'right_only'][['VISIT', 'VISITNUM']]
        missingInSpecs = missingInSpecs.astype(str).apply(lambda x: "/".join(x), axis=1)
        self.majorFindings = self.majorFindings.append(
            pd.DataFrame(
                {"Message": "Visit: Visit/VISITNUM Combination Missing in Specs",
                 "Details":  list(set(missingInSpecs.values))}))

        prefix = self.varLineEdit.text().upper()
        prefixVar = prefix + "TPT"
        subjVar = self.subjLineEdit.text().upper()
        if subjVar not in self.sasData.columns:
            self.err = f"Variable {subjVar} not found in SAS file"
            return
        self.flag = False

        if set(['VISIT',prefixVar]).intersection(self.specsFile['visit'].columns) != set(['VISIT',prefixVar]):
            # print("Failed")
            # self.err = "Variables not found in specs or data file"
            for var in set(['VISIT',prefixVar]).difference(set(['VISIT',prefixVar]).intersection(self.specsFile['visit'].columns)):
                self.majorFindings = self.majorFindings.append(
                    pd.DataFrame(
                        {"Message": 'Visit Sheet: Missing Variable in specs',
                         "Details": [var]}))
                self.flag = True
            # return

        print("Success")
        if set([subjVar,'VISIT',prefixVar]).intersection(self.sasData.columns) != set([subjVar,'VISIT',prefixVar]):
            # print("Failed")
            for var in set([subjVar,'VISIT', prefixVar]).difference(set([subjVar,'VISIT', prefixVar]).intersection(self.sasData.columns)):
                self.majorFindings = self.majorFindings.append(
                    pd.DataFrame(
                        {"Message": 'Visit Sheet: Missing Variable in data',
                         "Details": [var]}))
                self.flag = True
            # return

        print("Success")
        if self.flag:
            missing_tests = 'Comaprison skipped due to missing visit or timepoint vairable in eihter visit sheet  or data. See major findings for details'
        else:
            merged = self.specsFile['visit'].merge(self.sasData[["VISIT", prefixVar]].drop_duplicates(), how='outer',
                                                   indicator=True)
            missingInData = merged[merged['_merge'] == 'left_only'][['VISIT', prefixVar]]
            missingInData = missingInData.astype(str).apply(lambda x: "/".join(x), axis=1)
            self.majorFindings = self.majorFindings.append(
                pd.DataFrame(
                    {"Message": f"Visit: Visit/{prefixVar} Combination Missing in Data",
                     "Details": list(set(missingInData.values))}))

            missingInSpecs = merged[merged['_merge'] == 'right_only'][['VISIT', prefixVar]]
            missingInSpecs = missingInSpecs.astype(str).apply(lambda x: "/".join(x), axis=1)
            self.majorFindings = self.majorFindings.append(
                pd.DataFrame(
                    {"Message": f"Visit: Visit/{prefixVar} Combination Missing in Specs",
                     "Details": list(set(missingInSpecs.values))}))

            self.missingTests = pd.DataFrame(columns=["Subject","Visit","Timepoint"])
            # for col in self.specsFile['param'].columns:
            #     temp = pd.DataFrame(columns=["Subject", "Visit", "Timepoint", "Test Code"])
            #     possibleCombinations = self.sasData.assign(foo=1)[['VISIT',prefixVar,subjVar]].sort_values(subjVar).drop_duplicates().assign(foo=1).merge(self.specsFile['param'][[col]].assign(foo=1)).drop('foo',1).drop_duplicates()
            #
            # # self.specsFile['visit'] = self.specsFile['visit'].append({k: np.nan for k in self.specsFile['visit'].columns},
            # #                                                          ignore_index=True)
            # # self.specsFile['visit'] =self.specsFile['visit'].iloc[:-1,:]
            #
            #     merged = possibleCombinations.merge(self.sasData[[subjVar, 'VISIT', prefixVar]+list(self.specsFile['param'].columns)],how='outer',indicator=True)
            #     missingInData = merged[merged['_merge']=='left_only']
            #     temp['Subject'] = missingInData[subjVar]
            #     temp['Visit'] = missingInData['VISIT']
            #     temp['Timepoint'] = missingInData[prefixVar]
            #     # for col in self.specsFile['param'].columns:
            #     temp["Test Code"] = missingInData[col]
            #     self.missingTests = self.missingTests.append(temp,ignore_index=True)
            possibleCombinations = self.sasData.assign(foo=1)[['VISIT', prefixVar, subjVar]].sort_values(
                subjVar).drop_duplicates().assign(foo=1).merge(self.specsFile['param'].assign(foo=1)).drop('foo',
                                                                                                                  1).drop_duplicates()
            merged = possibleCombinations.merge(
                self.sasData[[subjVar, 'VISIT', prefixVar] + list(self.specsFile['param'].columns)], how='outer',
                indicator=True)
            missingInData = merged[merged['_merge'] == 'left_only']
            self.missingTests['Subject'] = missingInData[subjVar]
            self.missingTests['Visit'] = missingInData['VISIT']
            self.missingTests['Timepoint'] = missingInData[prefixVar]
            for col in self.specsFile['param'].columns:
                self.missingTests[col] = missingInData[col]


    def setupUi(self, Form):
        if not Form.objectName():
            Form.setObjectName(u"Form")
        Form.resize(517, 512)
        self.verticalLayout = QVBoxLayout(Form)
        self.verticalLayout.setObjectName(u"verticalLayout")
        self.horizontalLayout = QHBoxLayout()
        self.horizontalLayout.setObjectName(u"horizontalLayout")
        self.label = QLabel(Form)
        self.label.setObjectName(u"label")
        self.label.setMinimumSize(QSize(145, 0))

        self.horizontalLayout.addWidget(self.label)

        self.specsLineEdit = QLineEdit(Form)
        self.specsLineEdit.setObjectName(u"specsLineEdit")

        self.horizontalLayout.addWidget(self.specsLineEdit)

        self.specsPushButton = QPushButton(Form)
        self.specsPushButton.setObjectName(u"browseButton")
        self.specsPushButton.setMaximumSize(QSize(30, 30))

        self.horizontalLayout.addWidget(self.specsPushButton)

        self.verticalLayout.addLayout(self.horizontalLayout)

        self.horizontalLayout_2 = QHBoxLayout()
        self.horizontalLayout_2.setObjectName(u"horizontalLayout_2")
        self.label_2 = QLabel(Form)
        self.label_2.setObjectName(u"label_2")
        self.label_2.setMinimumSize(QSize(145, 0))

        self.horizontalLayout_2.addWidget(self.label_2)

        self.sasLineEdit = QLineEdit(Form)
        self.sasLineEdit.setObjectName(u"lineEdit_2")

        self.horizontalLayout_2.addWidget(self.sasLineEdit)

        self.sasPushButton = QPushButton(Form)
        self.sasPushButton.setObjectName(u"browseButton")
        self.sasPushButton.setMaximumSize(QSize(30, 30))

        self.horizontalLayout_2.addWidget(self.sasPushButton)

        self.verticalLayout.addLayout(self.horizontalLayout_2)

        self.horizontalLayout_3 = QHBoxLayout()
        self.horizontalLayout_3.setObjectName(u"horizontalLayout_3")
        self.label_3 = QLabel(Form)
        self.label_3.setObjectName(u"label_3")
        self.label_3.setMinimumSize(QSize(145, 0))

        self.horizontalLayout_3.addWidget(self.label_3, 0, Qt.AlignLeft)

        self.subjLineEdit = QLineEdit(Form)
        self.subjLineEdit.setObjectName(u"lineEdit_3")
        self.subjLineEdit.setMaximumSize(QSize(100, 16777215))

        self.horizontalLayout_3.addWidget(self.subjLineEdit)

        self.horizontalSpacer = QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum)

        self.horizontalLayout_3.addItem(self.horizontalSpacer)

        self.verticalLayout.addLayout(self.horizontalLayout_3)

        self.horizontalLayout_4 = QHBoxLayout()
        self.horizontalLayout_4.setObjectName(u"horizontalLayout_4")
        self.label_4 = QLabel(Form)
        self.label_4.setObjectName(u"label_4")
        self.label_4.setMinimumSize(QSize(145, 0))

        self.horizontalLayout_4.addWidget(self.label_4, 0, Qt.AlignLeft)

        self.varLineEdit = QLineEdit(Form)
        self.varLineEdit.setObjectName(u"lineEdit_4")
        self.varLineEdit.setEnabled(True)
        sizePolicy = QSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.varLineEdit.sizePolicy().hasHeightForWidth())
        self.varLineEdit.setSizePolicy(sizePolicy)
        self.varLineEdit.setMaximumSize(QSize(50, 16777215))
        self.varLineEdit.setAlignment(Qt.AlignLeading | Qt.AlignLeft | Qt.AlignVCenter)

        self.horizontalLayout_4.addWidget(self.varLineEdit)

        self.horizontalSpacer_2 = QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum)

        self.horizontalLayout_4.addItem(self.horizontalSpacer_2)

        self.verticalLayout.addLayout(self.horizontalLayout_4)

        self.horizontalLayout_5 = QHBoxLayout()
        self.horizontalLayout_5.setObjectName(u"horizontalLayout_5")
        self.label_5 = QLabel(Form)
        self.label_5.setObjectName(u"label_5")
        self.label_5.setMinimumSize(QSize(145, 0))

        self.horizontalLayout_5.addWidget(self.label_5)

        self.logLineEdit = QLineEdit(Form)
        self.logLineEdit.setObjectName(u"lineEdit_5")

        self.horizontalLayout_5.addWidget(self.logLineEdit)

        self.logPushButton = QPushButton(Form)
        self.logPushButton.setObjectName(u"browseButton")
        self.logPushButton.setMaximumSize(QSize(30, 30))

        self.horizontalLayout_5.addWidget(self.logPushButton)

        self.verticalLayout.addLayout(self.horizontalLayout_5)
        self.verticalLayout.addStretch()
        self.comparePushButton = QPushButton(Form)
        self.comparePushButton.setObjectName(u"comparePushButton")

        self.verticalLayout.addWidget(self.comparePushButton, 0, Qt.AlignHCenter)

        self.retranslateUi(Form)

        QMetaObject.connectSlotsByName(Form)

    # setupUi

    def retranslateUi(self, Form):
        Form.setWindowTitle(QCoreApplication.translate("Form", u"Form", None))
        self.label.setText(QCoreApplication.translate("Form", u"Select Specs File", None))
        self.specsPushButton.setIcon(QIcon(resource_path("icons/folder.png")))
        self.label_2.setText(QCoreApplication.translate("Form", u"Select SAS Dataset", None))
        self.sasPushButton.setIcon(QIcon(resource_path("icons/folder.png")))
        self.label_3.setText(QCoreApplication.translate("Form", u"Subject Variable in SAS", None))
        self.label_4.setText(QCoreApplication.translate("Form", u"Variable Prefix", None))
        self.label_5.setText(QCoreApplication.translate("Form", u"Select Folder to Store Log", None))
        self.logPushButton.setIcon(QIcon(resource_path("icons/folder.png")))
        self.comparePushButton.setText(QCoreApplication.translate("Form", u"Compare", None))
    # retranslateUi


class customFileDialog(QFileDialog):
    def __init__(self):
        super(customFileDialog, self).__init__()
        self.resize(QSize(250,250))


def main ():
    app = QApplication(sys.argv)
    app.setStyleSheet("QPushButton#browseButton {background:white;border:none}"
                      "QPushButton#browseButton:hover:!pressed{background-color:rgb(180, 180, 180);border:none}"
                      "QPushButton#browseButton:pressed {background-color: grey;border:none}")

    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
