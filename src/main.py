import re
import mysql.connector
from _collections import deque
from PySide2.QtWidgets import *
from PySide2.QtCore import *
from PySide2.QtGui import  *
import sys
import os
import json
import numpy as np
import pandas as pd
from datetime import  datetime
import time
import pyreadstat
import xlsxwriter
import src.settings as s
from src.delegates.commentDelegate import CommentWidget, PrimaryInfoDelegate
from src.model.ListModel import ListModel
from src.model.TableModel import BaseTableModel
from src.model.CustomTreeModel import CustomTreeModel
# from src.threads.LoadThread import LoadThread, runThread
from src.utils.validation import validateAdmin,validateUsers

from src.view.CustomWidgets import customButtton, CustomTabBar, customQMessageBox, UpdateAction, \
    MergeDialog, FilterAction, CategoryDialog, RunAction, ScanAction, TaskEdit, AddTaskDialog, IssueDialog, \
    CommentDialog, RTF2PDF, CombinePDF
from src.view.ListView import  DragDropListView
from src.view.TreeView import MyTreeView

from src.utils.utils import createTreeModel, getAllChildren, restartConnection, addSeperator, replaceExtension

rootPath = os.path.dirname(os.path.abspath(__file__))

pd.set_option("display.max_columns",50)
pd.set_option("display.width",1000)
# currentUser = "vjain"

def resource_path(relative_path):
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


class UserSelect(QWidget):
    def __init__(self,parent=None):
        super(UserSelect, self).__init__(parent=parent)

        self.mainLayout = QVBoxLayout(self)
        self.label = QLabel("Select User")
        self.mainLayout.addWidget(self.label)
        self.userDropDown = QComboBox()
        self.userDropDown.addItems(['dipeshs','vjain','bikhyats','yogeshb','sachins','Sissyk','yasnaa','pratikg','guest'])
        self.mainLayout.addWidget(self.userDropDown)
        self.okButton = QPushButton('Select')

        self.mainLayout.addWidget(self.okButton,alignment=Qt.AlignCenter)
        self.cancelButton = QPushButton('Cancel')
        self.mainLayout.addWidget(self.cancelButton,alignment=Qt.AlignCenter)
        self.setLayout(self.mainLayout)
        self.parent().setGeometry(50,50,200,100)
        self.okButton.clicked.connect(self.okClicked)
        self.cancelButton.clicked.connect(self.cancelClicked)

    def okClicked(self):
        self.parent().startApp()

    def cancelClicked(self):
        self.parent().close()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.title = 'Nimble Utilities'
        self.left = 50
        self.top = 50
        self.width = 792
        self.height = 610
        self.setWindowTitle(self.title)
        # self.userSelect = UserSelect(self)
        # self.setCentralWidget(self.userSelect)
        self.startApp()
        self.show()

    def startApp(self):
        # currentUser = self.userSelect.userDropDown.currentText()
        currentUser = os.getlogin().lower()
        # currentUser = 'vjain'
        self.tabWidget = MainUtility(parent=self,currentUser=currentUser)
        self.tabWidget.show()
        self.setCentralWidget(self.tabWidget)
        # self.setGeometry(self.left, self.top, self.width, self.height)
        self.setMinimumHeight(550)
        self.setMinimumWidth(550)
        self.showMaximized()


    # Creating tab widgets


class MainUtility(QWidget):
    def __init__(self,parent=None,currentUser=None):
        super(MainUtility, self).__init__(parent)
        s.app = self
        self.adminUsers = []
        self.projectID = None
        self.libID = None
        self.state = 'Items'
        self.filters = {}
        self.libraries = []
        self.currentUser = currentUser.lower()
        self.currentUsername = pd.read_sql_query(f"Select name from users where userId = '{currentUser}'", s.db)['name']

        self.currentUsername = self.currentUsername.values[0] if not self.currentUsername.empty else 'Guest' #TODO: remove after testing
        self.setupUI()
        self.disableButton()

        self.projects = pd.read_sql_query(f"SELECT T.ProjectID from team as T inner join project as P on T.ProjectID=P.ProjectID where userID='{self.currentUser}' and T.status='Active' and P.STATUS='Active' ", s.db)
        self.projectDropDown.addItems(['']+self.projects['ProjectID'].values.tolist())
        self.projectDropDown.activated.connect(self.projectSelected)

    def disableButton(self,button=None):
        if button:
            button.setHidden(True)
        else:
            icons = ['edit', 'refresh', 'category','run','scan', 'filter']
            for i in range(self.iconsLayout.count() - 2):
                if self.iconsLayout.itemAt(i).widget().objectName() not in icons:
                    self.iconsLayout.itemAt(i).widget().setHidden(True)
                else:
                    self.iconsLayout.itemAt(i).widget().setHidden(False)
                    self.iconsLayout.itemAt(i).widget().setDisabled(True)

    def flush(self):
        self.clearTable()
        for i in range(len(self.libraries)+1):
            self.libraryDropDown.removeItem(0)

    def clearTable(self):
        # if self.tableLabel.isHidden():
        #     self.tableLabel.show()
        # else:
        self.tableLabel.hide()
        # if self.treeView.isHidden():
        #remove models
        self.treeView.show()
        self.treeView.setModel(BaseTableModel())
        self.detailView.show()
        self.detailView.setModel(ListModel())

        # else:
        #     self.treeView.hide()
        #     self.detailView.hide()

    def projectSelected(self):
        # global isAdmin

        if not self.projectDropDown.currentText():
            return
        else:
            if self.projectDropDown.itemText(0) == '':
                self.projectDropDown.removeItem(0)
        self.flush()

        self.projectID = self.projectDropDown.currentText()

        if not self.projectID:
            # self.libraryDropDown.addItems([''])
            return
        self.adminUsers = pd.read_sql_query(f"SELECT DISTINCT userID from team WHERE ProjectID = '{self.projectID}' and roles='Admin'",s.db)['userID'].values.tolist()
        print("Admin users",self.adminUsers)
        s.isAdmin =  self.currentUser in [user.lower() for user in self.adminUsers]
        # s.isAdmin =  False
        print(s.isAdmin)
        s.currentUser = self.currentUser
        self.projectUsers = pd.read_sql_query(f"SELECT DISTINCT T.userID,U.name FROM project AS P INNER JOIN team_perm AS T INNER JOIN users as U ON T.userID=U.userID AND T.projectid=P.ProjectID WHERE P.ProjectID = '{self.projectID}'",s.db)
        s.projectUsers = self.projectUsers
        s.adminUsers = self.adminUsers
        data = pd.read_sql_query(
            f"SELECT T.*,D.* from team_perm as T inner join datlib as D on T.PDETAIL=D.LIBID and T.ProjectID = D.PROJECTID where T.userID='{self.currentUser}' and T.PROJECTID = '{self.projectID}'  and T.PRESOURCE='Data Access'",
            s.db)
        data = data[data['LTYPE'].isin(['SDTM','SEND','Analysis','Reports'])]
        data['name_'] = 'Data: ' + data['LIBRARY']
        reports = pd.read_sql_query(
            f"SELECT T.*,I.* from team_perm AS T  inner join items as I on T.PDETAIL=I.SYSTEMID and T.ProjectID =  I.projectid  where T.userID='{self.currentUser}' AND T.PROJECTID = '{self.projectID}' and T.PRESOURCE='Reports Access'",
            s.db)

        reports['name_'] = 'Reports: ' + reports['ITEMID']
        self.libraries = data['name_'].values.tolist() + reports['name_'].values.tolist()
        self.libraryDropDown.addItems(['']+self.libraries)

        self.libraryDropDown.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToContents)
        self.libraryDropDown.adjustSize()
        try:
            self.libraryDropDown.activated.disconnect()
        except:
            pass
        self.libraryDropDown.activated.connect(lambda x,data=data,reports=reports: self.librarySelected(data,reports))
        return  data,reports

    @Slot()
    def changeProcessText(self,text):
        self.processingBtn.setText(text)

    # @runThread(s.processingBtn,'Loading library...')
    def librarySelected(self,data,reports):
        self.clearTable()
        _st = time.time()
        if not self.libraryDropDown.currentText():
            return
        else:
            if self.libraryDropDown.itemText(0) == '':
                self.libraryDropDown.removeItem(0)

        self.projectID = self.projectDropDown.currentText()
        self.libName = self.libraryDropDown.currentText()
        self.libType = 'data' if 'Data:' in self.libName else 'report'

        # self.state = 'Items'

        # for i in range(self.topButtonLayout.count()):
        #     if self.topButtonLayout.itemAt(i).widget() != self.itemsButton:
        #         self.topButtonLayout.itemAt(i).widget().setChecked(False)
        #     else:
        #         self.topButtonLayout.itemAt(i).widget().setChecked(True)

        self.populateIcons()
        self.itemsButton.setEnabled(True)
        self.tasksButton.setEnabled(True)
        self.issuesButton.setEnabled(True)
        self.detailView.setModel(ListModel())

        try:
            self.detailView.clicked.disconnect()
        except:
            pass

        if self.libType =='data':
            self.libType = data[data['name_'] == self.libName]['LTYPE'].values[0]
            self.libID = data[data['name_'] == self.libName]['LIBID'].values[0]
            self.libName = data[data['name_'] == self.libName]['LIBRARY'].values[0]
            if self.state == 'Items':
                tableData = self.extractObjectSource()
                if tableData.empty:
                    self.disableAllIcons()
                    return
                tableData['primaryProgPath'] = np.where(tableData["Primary_Progname"],f"//sas-vm/{self.projectID}/Data/{self.libName.split(' ')[0]}/programs/" +\
                                               tableData["Primary_Progname"] ,tableData["Primary_Progname"])
                tableData['primaryLogPath'] = np.where(tableData["Primary_Progname"],f"//sas-vm/{self.projectID}/Data/{self.libName.split(' ')[0]}/programs/" +\
                                               tableData["Primary_Progname"].apply(lambda x:replaceExtension(x,'sas','log')),tableData["Primary_Progname"])
                tableData['outputPath'] = np.where(tableData["Primary_Progname"], f"//sas-vm/{self.projectID}/Data/{self.libName.split(' ')[0]}/programs/" + \
                                          tableData["Primary_Progname"].apply(lambda x:replaceExtension(x,'sas','sas7bdat')),tableData["Primary_Progname"])
                tableData['qcProgPath'] = np.where(tableData["QC_Progname"], f"//sas-vm/{self.projectID}/Data/{self.libName.split(' ')[0]}/validation/" + \
                                          tableData["QC_Progname"] ,tableData["QC_Progname"])
                tableData['qcLogPath'] = np.where(tableData["QC_Progname"], f"//sas-vm/{self.projectID}/Data/{self.libName.split(' ')[0]}/validation/" + \
                                                  tableData["QC_Progname"].apply(lambda x: replaceExtension(x, 'sas', 'log')),tableData["QC_Progname"])
                tableData['Last Run'] = tableData['outputPath'].apply(lambda x:datetime.fromtimestamp(os.path.getmtime(x)).strftime('%b %d,%Y')  if os.path.exists(x) else '')
                self.populateItems(tableData)
            elif self.state == 'Tasks':
                self.populateTasks()
                return
            elif self.state == 'Issues':
                self.populateIssues()
                return

        else:
            self.libID = reports[reports['name_'] == self.libName]['ITEMID'].values[0]
            self.libName = reports[reports['name_'] == self.libName]['ITEMID'].values[0]
            if self.state == 'Items':
                tableModel = self.extractObjectSource()
                if tableModel.empty:
                    self.disableAllIcons()
                    return
                self.detailView.setModel(ListModel([]))
                self.treeView.setModel(tableModel)
                self.treeView.header().resizeSections(QHeaderView.ResizeToContents)
                self.treeView.showColumn(0)
                self.treeView.hideColumn(1)
                self.treeView.hideColumn(2)
            elif self.state == 'Tasks':
                self.populateTasks()
                return
            elif self.state == 'Issues':
                self.populateIssues()
                return

        self.populateIcons()
        self.treeView.setSelectionMode(QAbstractItemView.ExtendedSelection)
        # try:
        #     self.treeView.selectionModel().selectionChanged.disconnect()
        # except:
        #     pass
        self.treeView.selectionModel().selectionChanged.connect(self.populateListView)

        if self.state == 'Items':
            colWidthMap = {'Name':100,'Order':40,'Category': 100,'Primary Owner':110, 'QC Type': 60, 'QC Info': 150, 'Match': 50,
                           'Primary Info': 150,
                           'QC Owner': 100,
                           'Last Run':100}

            for k, v in colWidthMap.items():
                if k in  self.treeView.model()._header:
                    i = self.treeView.model()._header.index(k)
                    self.treeView.setColumnWidth(i, v)

        self.treeView.header().setMaximumSectionSize(600)
        self.treeView.header().setMinimumSectionSize(30)
        # cols = ["objectID","Description","Name","Order","Category",'Primary Owner',"Primary Info","QC Type","Match",'QC Owner',"QC Info","Last Run"]
        #


    def populateIcons(self):

        start = time.time()
        if self.state == 'Items':

            self.edit.setEnabled(True)
            self.editMenu = QMenu()
            primaryOwner = QMenu('Primary Owners')
            self.editMenu.addMenu(primaryOwner)
            for user in  self.projectUsers['userID'].values.tolist()+['_NA_']:
                action = UpdateAction(user,'Primary_Owner',parent=self)
                primaryOwner.addAction(action)

            primaryStatus = QMenu('Primary Status')
            for s in ['Completed', 'Working', 'Waiting', 'Deferred','_NA_']:
                action = UpdateAction(s, 'Primary_Status',parent=self)
                primaryStatus.addAction(action)

            self.editMenu.addMenu(primaryStatus)

            # progNameAction = QAction('Primary Program name')
            # progNameAction.triggered.connect(lambda : self.treeView.model().updateData(self.treeView.currentIndex(),self.col,EditDialog().exec_(),Qt.EditRole))
            progNameAction = UpdateAction('Primary Program name','Primary_Progname',editDialog=True,parent=self)
            self.editMenu.addAction(progNameAction)

            primaryNoteAction = UpdateAction('Primary Notes','Primary_Notes',editDialog=True,parent=self)
            self.editMenu.addAction(primaryNoteAction)

            qcType = QMenu('QC Type')
            self.editMenu.addMenu(qcType)
            # qcTypes = [qcType.addAction(s) for s in ['Full', 'Basic', 'None']]

            for s in ['Full', 'Basic', 'None']:
                action = UpdateAction(s, 'QC_Type', parent=self)
                qcType.addAction(action)
            qcOwnerMenu = QMenu('QC Owner')
            self.editMenu.addMenu(qcOwnerMenu)
            # qcOwnerAction = [qcOwnerMenu.addAction(user) for user in self.adminUsers]
            for user in  self.projectUsers['userID'].values.tolist()+['_NA_']:
                action = UpdateAction(user, 'QC_Owner', parent=self)
                qcOwnerMenu.addAction(action)

            qcStatus = QMenu('QC Status')
            self.editMenu.addMenu(qcStatus)
            for s in ['Completed', 'Working', 'Waiting', 'Deferred', '_NA_']:
                action = UpdateAction(s, 'QC_Status', parent=self)
                qcStatus.addAction(action)

            qcProgAction = UpdateAction('QC Program name','QC_Progname',editDialog=True,parent=self)
            self.editMenu.addAction(qcProgAction)

            qcNoteAction = UpdateAction('QC Notes','qc_notes',editDialog=True,parent=self)
            self.editMenu.addAction(qcNoteAction)

            order = QMenu('Order')
            self.editMenu.addMenu(order)
            for s in range(1,9):
                action = UpdateAction(s,'obj_order',parent=self)
                order.addAction(action)


            mergeAction = self.editMenu.addAction('Merge')
            mergeAction.triggered.connect(self.mergeRows)


            deleteAction = self.editMenu.addAction('Delete outdated item')
            #TODO: donot allow updates on oudateditems DONE
            deleteAction.setDisabled(True)
            deleteAction.triggered.connect(self.treeView.model().deleteRows)
            if self.libType == 'report': #disable merge order and delete for reports
                mergeAction.setDisabled(True)
                order.setDisabled(True)
                deleteAction.setDisabled(True)

            self.updateCategories(self.editMenu)

            self.edit.setMenu(self.editMenu)

            self.filterMenu = QMenu()
            itemName = FilterAction('Name', 'object_name',editDialog=True, parent=self)
            self.filterMenu.addAction(itemName)
            primaryOwner = QMenu('Primary Owners')

            primaryOwnerGroup = QActionGroup(self)

            self.filterMenu.addMenu(primaryOwner)
            for user in self.projectUsers['userID'].values.tolist()+['_NA_']:
                action = FilterAction(user, 'Primary_Owner', parent=self)
                primaryOwner.addAction(action)
                primaryOwnerGroup.addAction(action)
            primaryOwnerGroup.setExclusive(True)

            primaryStatus = QMenu('Primary Status')
            primaryStatusGroup = QActionGroup(self)
            for s in ['Completed', 'Working', 'Waiting', 'Deferred', '_NA_']:
                action = FilterAction(s, 'Primary_Status', parent=self)
                primaryStatus.addAction(action)
                primaryStatusGroup.addAction(action)
            primaryStatusGroup.setExclusive(True)
            self.filterMenu.addMenu(primaryStatus)

            # progNameAction = QAction('Primary Program name')
            # progNameAction.triggered.connect(lambda : self.treeView.model().updateData(self.treeView.currentIndex(),self.col,EditDialog().exec_(),Qt.EditRole))
            progNameAction = FilterAction('Primary Program name', 'Primary_Progname',editDialog=True, parent=self)
            progNameAction.setCheckable(True)
            self.filterMenu.addAction(progNameAction)

            primaryNoteAction = FilterAction('Primary Notes', 'Primary_Notes',editDialog=True, parent=self)
            primaryNoteAction.setCheckable(True)
            self.filterMenu.addAction(primaryNoteAction)
            qcType = QMenu('QC Type')
            self.filterMenu.addMenu(qcType)
            # qcTypes = [qcType.addAction(s) for s in ['Full', 'Basic', 'None']]
            qcTypeGroup = QActionGroup(self)
            for s in ['Full', 'Basic', 'None']:
                action = FilterAction(s, 'QC_Type', parent=self)
                qcType.addAction(action)
                qcTypeGroup.addAction(action)
            qcTypeGroup.setExclusive(True)

            qcOwnerMenu = QMenu('QC Owner')
            qcOwnerGroup = QActionGroup(self)

            self.filterMenu.addMenu(qcOwnerMenu)
            # qcOwnerAction = [qcOwnerMenu.addAction(user) for user in self.adminUsers]
            for user in self.projectUsers['userID'].values.tolist()+['_NA_']:
                action = FilterAction(user, 'QC_Owner', parent=self)
                qcOwnerMenu.addAction(action)
                qcOwnerGroup.addAction(action)
            qcOwnerGroup.setExclusive(True)

            qcStatus = QMenu('QC Status')
            qcStatusGroup = QActionGroup(self)
            self.filterMenu.addMenu(qcStatus)
            for s in ['Completed', 'Working', 'Waiting', 'Deferred', '_NA_']:
                action = FilterAction(s, 'QC_Status', parent=self)
                qcStatus.addAction(action)
                qcStatusGroup.addAction(action)
            qcStatusGroup.setExclusive(True)

            qcProgAction = FilterAction('QC Program name', 'QC_Progname',editDialog=True, parent=self)
            qcProgAction.setCheckable(True)
            self.filterMenu.addAction(qcProgAction)

            qcNoteAction = FilterAction('QC Notes', 'qc_notes',editDialog=True, parent=self)
            qcProgAction.setCheckable(True)
            self.filterMenu.addAction(qcNoteAction)

            self.actionGroups = [primaryOwnerGroup,primaryStatusGroup,qcTypeGroup,qcOwnerGroup,qcStatusGroup]
            resetAction = self.filterMenu.addAction('Reset')
            resetAction.triggered.connect(self.resetData)

            self.runMenu = QMenu()
            primaryAction = RunAction('Primary',parent=self)
            self.runMenu.addAction(primaryAction)
            qcAction = RunAction('QC',parent=self)
            self.runMenu.addAction(qcAction)
            self.runBtn.setMenu(self.runMenu)
            self.runBtn.setEnabled(True)

            self.scanMenu = QMenu()
            primaryAction = ScanAction('Primary', parent=self)
            self.scanMenu.addAction(primaryAction)
            qcAction = ScanAction('QC', parent=self)
            self.scanMenu.addAction(qcAction)
            self.scanLogs.setMenu(self.scanMenu)
            self.scanLogs.setEnabled(True)

            self.tflMenu  = QMenu()
            rtf2pdf = QAction("RTF to PDF convert",parent=self)
            rtf2pdf.triggered.connect(self.rtf2pdf)
            self.tflMenu.addAction(rtf2pdf)
            combinePDF = QAction("Combine PDF files",parent=self)
            combinePDF.triggered.connect(self.combinePDF)
            self.tflMenu.addAction(combinePDF)
            self.tfl.setMenu(self.tflMenu)


            self.filter.setMenu(self.filterMenu)
            self.filter.setEnabled(True)

            self.category.setEnabled(True)
            self.itemsButton.setChecked(True)
            self.refresh.setEnabled(True)

            icons = ['edit', 'refresh', 'category','run','scan', 'filter','tfl']

            for i in range(self.iconsLayout.count() - 2):
                if self.iconsLayout.itemAt(i).widget().objectName() not in icons:
                    self.iconsLayout.itemAt(i).widget().setHidden(True)
                else:
                    self.iconsLayout.itemAt(i).widget().setHidden(False)
            if self.libType == 'report':
                self.tfl.show()
                self.tfl.setEnabled(True)
            else:
                self.tfl.hide()
            self.detailView.setStyleSheet("""
                                              QListView::item{background-color:rgb(255,255,255);color:black}
                                              QListView::item:hover{background-color:rgba(229,243,255);color:black}
                                              QListView::item:selected{background-color:rgba(197,224,247);color:black}
                                          """)

        if self.state == 'Tasks':
            icons = ['add','refresh','remove']
            for i in range(self.iconsLayout.count()-2):
                if self.iconsLayout.itemAt(i).widget().objectName() not in icons:
                    self.iconsLayout.itemAt(i).widget().setHidden(True)
                else:
                    self.iconsLayout.itemAt(i).widget().setHidden(False)
                    self.iconsLayout.itemAt(i).widget().setEnabled(True)


            self.detailView.setStyleSheet("""
                                   QListView::item{background-color:rgb(255,255,255);color:black}
                                   QListView::item:hover{background-color:rgba(229,243,255);color:black}
                                   QListView::item:selected{background-color:rgba(197,224,247);color:black}
                               """)
            try:
                self.add.clicked.disconnect()
            except:
                pass
            self.add.clicked.connect(self.addTask)

        if self.state == 'Issues':
            icons = ['add', 'refresh','comment','filter']
            for i in range(self.iconsLayout.count()-2):
                if self.iconsLayout.itemAt(i).widget().objectName() not in icons:
                    self.iconsLayout.itemAt(i).widget().setHidden(True)
                else:
                    self.iconsLayout.itemAt(i).widget().setHidden(False)
                    self.iconsLayout.itemAt(i).widget().setEnabled(True)
            try:
                self.add.clicked.disconnect()
            except:
                pass
            self.add.setEnabled(True)
            self.filterMenu = QMenu()
            ["My open issues",  "Status", "Title", "Detail", "Impacts", "Opened By", "Assigned To","Reset",]

            self.filterMenu.addAction(FilterAction("My open issues","hasComment",custom=True,parent=self))
            statusMenu = QMenu("Status")
            statusMenuGroup = QActionGroup(self)

            statusOptions = ['Open','Deferred','Closed']
            for o in statusOptions:
                action = FilterAction(o,"issue_Status",parent=self)
                statusMenu.addAction(action)
                statusMenuGroup.addAction(action)
            statusMenuGroup.setExclusive(True)

            self.filterMenu.addMenu(statusMenu)
            self.filterMenu.addAction(FilterAction("Title","issue_Title",editDialog=True,parent=self))
            self.filterMenu.addAction(FilterAction("Detail","issue_Detail",editDialog=True,parent=self))
            self.filterMenu.addAction(FilterAction("Impacts","issue_impact",editDialog=True,parent=self))
            openedMenu = QMenu("Opened By")
            openedGroup = QActionGroup(self)
            assignedMenu = QMenu("Assigned To")
            assignedGroup = QActionGroup(self)
            for u in self.projectUsers['userID'].values.tolist():
                openAction = FilterAction(u,"open_By",parent=self)

                assignAction = FilterAction(u,"assigned_to",parent=self)
                openedMenu.addAction(openAction)
                openedGroup.addAction(openAction)
                assignedMenu.addAction(assignAction)
                assignedGroup.addAction(assignAction)

            openedGroup.setExclusive(True)
            assignedGroup.setExclusive(True)
            self.actionGroups = [openedGroup, assignedGroup, statusMenuGroup]
            self.filterMenu.addMenu(assignedMenu)
            self.filterMenu.addMenu(openedMenu)
            resetAction = self.filterMenu.addAction('Reset')
            resetAction.triggered.connect(self.resetData)
            self.filter.setMenu(self.filterMenu)
            self.add.clicked.connect(self.addIssue)
            self.detailView.setStyleSheet("""
                                             QListView::item{background-color:rgb(255,255,255);color:black;border-bottom:1px solid black}
                                             
                                             QListView::item:hover{background-color:rgba(229,243,255);color:black}
                                             QListView::item:selected{background-color:rgba(197,224,247);color:black}
                                         """)
        print('Time taken to populate Icons',time.time()-start)


    def combinePDF(self):

        dialog = CombinePDF(self)
        dialog.exec_()

        pass

    def rtf2pdf(self):
        rows = self.treeView.selectionModel().selectedRows()
        reportIds = [self.treeView.model().itemFromIndex(index).id for index in self.treeView.selectionModel().selectedRows()]
        reports = self.treeView.model()._data.set_index('objectID').loc[reportIds]

        if len(rows) < 1:
            msg = customQMessageBox("Please select at least one item.")
            msg.exec_()
            return

        if reports[['extension']].dropna().query("extension.str.contains('rtf')").empty:
            msg = customQMessageBox("None of selected reports have rtf file.")
            msg.exec_()
            return

        dialog = RTF2PDF(self)
        dialog.exec_()




        # if os.path.exists(path):
        #     rows = [row.row() for row in rows]
        #     filenames = self.treeView.model().visibleData.iloc[rows]['Name'].values.tolist()


    def addComment(self):
        if not len(self.treeView.selectionModel().selectedRows()):
            msg = customQMessageBox("Please select an issue to add comment.")
            msg.exec_()
            return

        if self.treeView.model().visibleData.iloc[self.treeView.currentIndex().row()]['Status'].lower() == 'closed':
            msg = customQMessageBox("Comments on closed issues are not allowed")
            msg.exec_()
            return


        commentDialog = CommentDialog(self.treeView.currentIndex().row(),parent=self)
        commentDialog.exec_()

    def addIssue(self):
        addIssueDialog = IssueDialog(parent=self)
        addIssueDialog.exec_()

    def editIssue(self,index):
        rowData = self.treeView.model().visibleData.iloc[index.row()]
        users = rowData[['Opened By','Assigned To']].values.tolist()

        @validateUsers(users+self.adminUsers, "You do not have access to this action.")
        def func():
            editIssueDialog = IssueDialog(index.row(),self)
            editIssueDialog.exec_()

        func()

    def populateListView(self):
        if self.state == 'Items':
            if self.libType == 'report':
                objectID = [self.treeView.model()   .itemFromIndex(row).id for row in
                            self.treeView.selectionModel().selectedRows()]
                if objectID:
                    self.editMenu.actions()[-1].setDisabled(True)
            else:
                rows = [row.row() for row in self.treeView.selectionModel().selectedRows()]
                objectID = [self.treeView.model().visibleData.iloc[row]['objectID'] for row in rows]
                if self.treeView.model()._data.set_index('objectID').loc[objectID]['outDated'].any():
                    self.editMenu.actions()[-1].setEnabled(True)
                else:
                    self.editMenu.actions()[-1].setDisabled(True)

        if len(self.treeView.selectionModel().selectedRows()) == 1:
            if self.state == "Items":
                row = self.treeView.selectionModel().selectedRows()[0].row()

                if self.libType == 'report':
                    objectID = [self.treeView.model().itemFromIndex(row).id for row in
                                self.treeView.selectionModel().selectedRows()]
                    if self.treeView.model()._data.set_index('objectID').loc[objectID]['outDated'].any():
                        self.editMenu.actions()[-1].setDisabled(True)
                    else:
                        self.editMenu.actions()[-1].setDisabled(True)

                    objectID = self.treeView.model()._data[
                        self.treeView.model()._data['objectID'].isin(objectID)]['objectID']
                    if objectID.empty:
                        self.detailView.setModel(ListModel([]))
                        return
                    else:

                        objectID = objectID.values.tolist()[0]
                else:
                    objectID = self.treeView.model().visibleData.iloc[row]['objectID']
                    if self.treeView.model()._data.set_index('objectID').loc[objectID]['outDated'].any():
                        self.editMenu.actions()[-1].setEnabled(True)
                        self.editMenu.actions()[-2].setEnabled(True)
                    else:
                        self.editMenu.actions()[-1].setDisabled(True)
                        self.editMenu.actions()[-2].setDisabled(True)


                # print(self.treeView.model()._data.loc[self.treeView.model()._data['objectID'] == objectID])
                validationErrors = self.treeView.model()._data.loc[
                    self.treeView.model()._data['objectID'] == objectID, 'Validation_fail'].values[0]
                validationErrors = validationErrors['Primary']+validationErrors['QC']

                primaryLog = self.treeView.model()._data.loc[
                    self.treeView.model()._data['objectID'] == objectID, 'primary_log'].values[0]
                if primaryLog:
                    primaryLog = json.loads(primaryLog.decode().replace('\t','    ').replace('\n',''))
                else:primaryLog = []

                qcLog = self.treeView.model()._data.loc[
                    self.treeView.model()._data['objectID'] == objectID, 'qc_log'].values[0]

                if qcLog:
                    qcLog = json.loads(qcLog.decode().replace('\t','    ').replace('\n',''))
                else:qcLog = []


                listModel = ListModel(validationErrors+primaryLog+qcLog)
                self.detailView.setModel(listModel)
            elif self.state == "Tasks":
                if not self.treeView.currentIndex().isValid():
                    print("INvalid selection")
                    return
                idx =  self.treeView.model().itemFromIndex(self.treeView.currentIndex()).id
                # print(self.treeView.model()._data[self.treeView.model()._data['taskid'] == idx])
                items = pd.read_sql_query(f"SELECT * from util_task_items where taskid = '{idx}' and PROJECTID = '{self.projectID}' and LIBID = '{self.libID}' order by tOrder",s.db)
                listModel = ListModel(items,checkable=True,dataCol='titem_Name',checkVal='titem_Status',parent=self)

                self.detailView.setModel(listModel)
                try:
                    self.detailView.clicked.disconnect()
                except:
                    pass
                self.detailView.clicked.connect(self.toggleItemState)
        else:
            self.detailView.setModel(ListModel([]))
            self.editMenu.actions()[-2].setDisabled(True)

    def toggleItemState(self,index):
        owner = self.treeView.model().itemFromIndex(self.treeView.currentIndex().siblingAtColumn(2)).text()
        @validateUsers([owner]+self.adminUsers,errorText="Access Error: You do not have access to this functionality.")
        def toggle(index):
            if self.treeView.model().itemFromIndex(self.treeView.currentIndex().parent().siblingAtColumn(3)).text() in ['Cancelled','Not Applicable']:
                return

            row = index.row()
            state = self.detailView.model()._data.iloc[row]['titem_Status']
            if state == "X":
                self.detailView.model().updateDB(index,'Z')
                #TODO: strike out Text DONE

            elif state == "Z":
                self.detailView.model().updateDB(index,'')
            else:
                self.detailView.model().updateDB(index,'X')
        toggle(index)

    @validateAdmin
    def mergeRows(self):
        if len(self.treeView.selectionModel().selectedRows()) > 1:
            msg = customQMessageBox('Please select a single row to merge')
            msg.exec_()
            return

        if len(self.treeView.selectionModel().selectedRows()) < 1:
            msg = customQMessageBox('Please select a row')
            msg.exec_()
            return
        # if self.libType == 'data':
        row =self.treeView.selectionModel().selectedRows()[0].row()
        objectID = self.treeView.model().visibleData.iloc[row]['objectID']
        if self.treeView.model()._data.loc[self.treeView.model()._data['objectID']==objectID,'outDated'].any():
            mergeDialog = MergeDialog(self)
            mergeDialog.exec_()
        else:
            msg = customQMessageBox('Please select an outdated row to merge')
            msg.exec_()


    def checkValidations(self,row):
        # return {1:'err'}
        #TODO: Show light red color in Primary Info if there are any errors :DONE
        validationList  = {'Primary':[],'QC':[]}

        if not row['Primary_Progname']:
            validationList['Primary'].append("Primary Program Name is missing")

        if not re.match(r"^(\w|-|_)+.sas$",str(row['Primary_Progname'])):
            validationList['Primary'].append(
            "Program name is invalid. Valid program can only have letters, numbers, hyphens or underscores and must end with .sas. Retry!")
        # text =str(row['Primary_Progname'])
        # if not text.endswith('.sas') or
        if not re.match(r"^(\w|-|_)+.sas$",str(row['QC_Progname'])):
            validationList['QC'].append("QC name is invalid. Valid program can only have letters, numbers, hyphens or underscores and must end with .sas. Retry!")

        if not row['QC_Type']:
            validationList['QC'].append("QC Type is missing")

        if not row['Primary_Owner']:
            validationList['Primary'].append("Primary Owner isn't assigned")

        if row['QC_Type'] in ['Full','Basic'] and not row['QC_Owner']:
            validationList['QC'].append("QC Owner isn't assigned")

        if row['QC_Type'] in ['Full'] and not row['QC_Progname']:
            validationList['QC'].append("QC Program name is missing")
        # if row.name == 20:
        #     print(row)
        return validationList

    def disableAllIcons(self):
        #disable AL icons if library is curropt
        # icons = ['edit', 'category','run','scan', 'filter']
        for i in range(self.iconsLayout.count() - 2):
            self.iconsLayout.itemAt(i).widget().setDisabled(True)

        # for i in range(self.topButtonLayout.count()):
        #     if self.state == self.topButtonLayout.itemAt(i).widget().text():
        #         self.topButtonLayout.itemAt(i).widget().setChecked(False)
        #     self.topButtonLayout.itemAt(i).widget().setDisabled(True)

        # for i in range(self.icon

    def extractObjectSource(self):
        # oldDF = pd.read_sql_query('SELECT * from util_obj',s.db)
        import time
        if self.libType in ['SDTM','SEND']:

            # FETCH DATA
            start = time.time()
            # TODO: add Try catch:DONE
            try:
                datasets, _ = pyreadstat.read_sas7bdat(
                    f'//sas-vm/{self.projectID}/Data/{self.libName}/meta/datasets.sas7bdat')
                # datasets = datasets[~datasets['DATASET'].isin(['IE', 'LB'])]
            except  pyreadstat._readstat_parser.PyreadstatError as e:
                msg = customQMessageBox("Datasets file not found")
                # self.disableAllIcons()
                msg.exec_()
                self.clearTable()
                return pd.DataFrame()
            try:
                variables, _ = pyreadstat.read_sas7bdat(
                    f'//sas-vm/{self.projectID}/Data/{self.libName}/meta/variables.sas7bdat')
            except  pyreadstat._readstat_parser.PyreadstatError as e:
                msg = customQMessageBox("Variables file not found")
                # self.disableAllIcons()
                msg.exec_()
                self.clearTable()
                return pd.DataFrame()

            print("Time taken to read dataset : ", time.time() - start)
            start = time.time()
            datasets = datasets.reset_index()
            # CREATE NEW COLUMNS FOR DATA DIPLAY
            datasets['objectID'] = datasets['index'].astype(str)
            datasets['PROJECTID'] = self.projectID
            datasets['LIBID'] = self.libID
            datasets['object_name'] = datasets['DATASET']
            datasets[['Object_Desc']] = datasets['DESCRIPTION']
            # supp = [domain for domain in variables['DOMAIN'].unique() if 'SUPP' in variables[variables['DOMAIN']==domain]['ACTION'].unique()]
            supp = variables[variables['ACTION']=='SUPP']['DOMAIN'].values.tolist() # find  supplemental datasets
            # CREATE NEW DATASET for display

            newDF = datasets[['objectID', 'PROJECTID', 'LIBID','object_name', 'Object_Desc']].copy()
            supp = newDF.loc[datasets['DATASET'].isin(supp), 'object_name'].values.tolist()
            newDF.loc[datasets['DATASET'].isin(supp), 'Object_Desc'] = [f"Supplemental  for {su}" for su in supp]

             # CHECK FOR NEW DATASETS
            oldDF = pd.read_sql_query(F"SELECT * from util_obj WHERE PROJECTID='{self.projectID}' and LIBID='{self.libID}'",s.db)

            merged = oldDF.merge(newDF, how='outer', on=['PROJECTID','LIBID','object_name'],suffixes = ('_x',''), indicator=True)

            outdatedData = merged[merged['_merge']=='left_only'] # find outdated Datasets
            outdatedData = outdatedData['objectID_x'].values.tolist()
            newDF = merged[merged['_merge']=='right_only']
            newDF = newDF[['objectID', 'PROJECTID', 'LIBID','object_name', 'Object_Desc']]
            print("Time taken for processing : ", time.time() - start)
            st = time.time()
            newID = pd.read_sql_query(f"SELECT * from util_obj where PROJECTID = '{self.projectID}' and LIBID = '{self.libID}' ",s.db).replace(np.nan,'')
            print("exec sql",time.time()-st)
            newID = 0 if newID.empty else newID['objectID'].astype(int).max()+1

            newDF['objectID'] = [i for i in range(newID,newID+newDF.shape[0])]

            start = time.time()
            # UPDATE DB to add new DATASETS
            cur = s.db.cursor()
            for idx,row in newDF.iterrows():
                print('Updating Database count: ', idx)
                cur.execute(f"INSERT IGNORE into util_obj (objectID,PROJECTID,LIBID,object_name,Object_Desc) VALUES (%s,%s,%s,%s,%s) ON DUPLICATE KEY UPDATE  Object_Desc = %s ",
                            (*row,row['Object_Desc']))
            print("Time taken to update  database : ", time.time() - start)
            s.db.commit()
            ## GET UPDATED info of DATASETS from DB
            newDF = pd.read_sql_query(f"SELECT * from util_obj where PROJECTID = '{self.projectID}' and LIBID = '{self.libID}' ",s.db).replace(np.nan,'')
            newDF['outDated'] = False
            newDF.loc[newDF['objectID'].isin(outdatedData),'outDated'] = True

            return newDF

        elif self.libType == 'report':
            pstart = time.time()
            items = pd.read_sql_query(f"SELECT * from items where projectid = '{self.projectID}'  and PARENTID  like '{self.libName}%'  ",s.db)
            newDF = items[['SYSTEMID','projectid','ITEMID','DESCRIPTION','items_ORDER']]
            newDF.loc[:,'objectID'] = newDF['SYSTEMID'].astype(str)
            newDF.loc[:,'LIBID'] = self.libName
            newDF.loc[:,'PROJECTID'] = newDF['projectid'].astype(str)
            newDF.loc[:,'object_name'] = newDF['ITEMID'].astype(str)
            newDF.loc[:,'Object_Desc'] = newDF['DESCRIPTION'].astype(str)
            newDF = newDF[['objectID', 'PROJECTID', 'LIBID', 'Object_Desc','object_name']]
            oldDF = pd.read_sql_query(f"SELECT U.* FROM util_obj AS U WHERE U.PROJECTID='{self.projectID}' and U.LIBID='{self.libName}' ",s.db)
            merged = oldDF.merge(newDF, how='outer', on=['objectID','PROJECTID','LIBID','object_name'], suffixes=('_x', ''), indicator=True)
            outdatedData = merged[merged['_merge'] == 'left_only']
            outdatedData = outdatedData['objectID'].values.tolist()
            newDF = merged[merged['_merge'] == 'right_only']
            newDF = newDF[['objectID', 'PROJECTID', 'LIBID', 'object_name','Object_Desc']]
            print('Time taken to process 1:',time.time()-pstart)
            # newID = pd.read_sql_query(
            #     f"SELECT * from util_obj where PROJECTID = '{self.projectID}' and LIBID = '{self.libID}' ", s.db).replace(
            #     np.nan, '')
            # newID = 0 if newID.empty else newID['objectID'].astype(int).max() + 1
            # newDF['objectID'] = [i for i in range(newID, newID + newDF.shape[0])]
            ustart = time.time()
            cur = s.db.cursor()

            # TODO: remove outdated from reports:DONE
            # TODO: disable merge and outdated from reports:DONE
            for idx,row in newDF.iterrows():
                print('Updating Database count: ',idx)
                cur.execute(
                    f"INSERT IGNORE into util_obj (objectID,PROJECTID,LIBID,object_name,Object_Desc) VALUES (%s,%s,%s,%s,%s) ON DUPLICATE KEY UPDATE  object_name = %s , Object_Desc = %s  ",
                    (*row, row['object_name'],row['Object_Desc']))
            s.db.commit()
            utime = time.time()-ustart
            print("Time taken to update database : ",utime)
            pstart = time.time()
            tableData = pd.read_sql_query(f"SELECT I.projectid,I.COMPLETEID,I.PARENTID,I.ITEMID,I.items_ORDER,I.TEMPLATE,U.* FROM util_obj AS U JOIN items AS  I  ON I.projectid=U.PROJECTID AND I.SYSTEMID=U.objectID WHERE U.PROJECTID='{self.projectID}'  and U.LIBID ='{self.libName}'  and I.COMPLETEID  like '{self.libName}/%' ",s.db)
            tableData['PARENTID'] = tableData['PARENTID'].str.split('/').apply(lambda x:x[-1])
            tableData = tableData.sort_values('COMPLETEID', key=lambda x: x.str.split('/'))
            cols = ["Name","objectID","Description","Category","Primary Owner" ,"Primary Info", "QC Type", "Match", "QC Owner","QC Info","Last Run"]
            tableData['Description'] = tableData['Object_Desc']
            tableData['Order']= tableData['obj_order']
            #TODO: set option as 1-9:DONE
            #TODO: hide order for reports:DONE
            #TODO: skip delegate for now :DONE
            tableData["Name"] = tableData['object_name']
            tableData['Category'] = tableData['custom_Columns'].apply(
                lambda x: "\n".join([f"{k}:{v}" for k, v in json.loads(x).items()]) if x else '')
            tableData['Category'] = np.where(tableData['Category'].str.decode('ascii').isna(),
                                             tableData['Category'],
                                             tableData['Category'].str.decode('ascii'))
            # tableData['Primary Owner'] = tableData['Primary_Owner']
            # tableData.loc[tableData['Primary_Owner'].str.len() > 1, 'Primary_Owner'] += ";"
            # tableData['Primary Status'] = tableData['Primary_Status']
            # tableData.loc[tableData['Primary_Status'].str.len() > 1, 'Primary_Status'] += "\n"
            # tableData['Primary Program Name'] = tableData['Primary_Progname']
            # tableData.loc[tableData['Primary_Progname'].str.len() > 1, 'Primary_Progname'] += ";"
            tableData['Primary Owner'] = tableData['Primary_Owner']
            tableData['Primary_Notes'] = np.where(tableData['Primary_Notes'].str.decode('ascii').isna(),
                                                   tableData['Primary_Notes'],
                                                   tableData['Primary_Notes'].str.decode('ascii'))
            tableData = tableData.replace(np.nan, '')
            tableData['Primary Info'] = tableData['Primary_Status'].apply(addSeperator) + \
                                         tableData['Primary_Progname'].apply(lambda x: addSeperator(x, '\n')) + \
                                        tableData['Primary_Notes']
            tableData['Primary Info'] = tableData['Primary Info'].str.strip(';').str.strip('\n')

            tableData['QC Type'] = tableData['QC_Type']
            # tableData.loc[tableData['QC_Owner'].str.len() > 1, 'QC_Owner'] += ";"
            # tableData.loc[tableData['QC_Status'].str.len() > 1, 'QC_Status'] += "\n"
            # tableData.loc[tableData['QC_Progname'].str.len() > 1, 'QC_Progname'] += ";"
            tableData['QC Owner'] = tableData['QC_Owner']


            tableData['qc_notes'] = np.where(tableData['qc_notes'].str.decode('ascii').isna(),
                                                  tableData['qc_notes'],
                                                  tableData['qc_notes'].str.decode('ascii'))
            tableData = tableData.replace(np.nan, '')
            tableData['QC Info'] = tableData['QC_Status'].apply(addSeperator) + \
                                    tableData['QC_Progname'].apply(lambda x: addSeperator(x, '\n')) + \
                                    tableData['qc_notes']
            tableData['QC Info'] = tableData['QC Info'].str.strip(';').str.strip('\n')
            tableData["Match"] = ''

            tableData['outDated'] = False
            tableData.loc[tableData['objectID'].isin(outdatedData), 'outDated'] = True

            templateFilter = tableData['TEMPLATE'].isin(['Listing', 'Summary', 'Figure'])
            props = pd.read_sql_query(
                f"SELECT P.*, U.*,I.template FROM util_obj AS U INNER JOIN items AS I ON U.PROJECTID=I.projectid  AND I.SYSTEMID=U.objectID INNER JOIN  properties AS P ON U.PROJECTID=P.PROJECTID AND U.objectID=P.SYSTEMID  WHERE U.PROJECTID =  '{self.projectID}' and U.LIBID = '{self.libID}'",
                s.db)
            props['SYSTEMID'] = props['SYSTEMID'].astype('str')
            props['REUSEID'] = props['REUSEID'].astype('str')
            props = props.set_index('SYSTEMID').loc[
                tableData[tableData['TEMPLATE'].isin(['Summary', 'Listing', 'Figure'])]['objectID'].values.tolist()].reset_index()


            # props = tableData.merge(props, left_on='objectID', right_on='SYSTEMID')
            tableData.loc[tableData['objectID'].isin(props['SYSTEMID'].values), 'Title'] = props.apply(
                lambda x: getPropertyValue(props, 'COL2', x.name), axis=1).str.replace(
                "ITT", '')
            tableData['Title'] = tableData['Title'].replace(np.nan,'')
            tableData['Title'] = tableData.apply(lambda x: x['Title'].replace('&itemid', x['ITEMID']), axis=1)
            ext = props.apply(
                lambda x: getPropertyValue(props, 'COL12', x.name) if x['template'] == 'Listing' else getPropertyValue(
                    props, 'COL15', x.name), axis=1)

            extensions = ['.rtf', '.pdf', '.xlsx', '.docx']
            for k in extensions:
                ext[ext.str.lower().str.contains(k)] = k
            tableData.loc[tableData['objectID'].isin(props['SYSTEMID'].values), 'extension'] = ext
            tableData['outputPath'] = f"//sas-vm/{self.projectID}/Reports/" + tableData.loc[templateFilter,'COMPLETEID'] + tableData['extension']
            tableData['outputPath'] = tableData['outputPath'].replace(np.nan, '')

            print('Time taken to process 2 : ', time.time() - pstart)

            for folder in tableData['outputPath'].str.split('/').apply(lambda x: "/".join(x[:-1])).replace('',np.nan).dropna().unique():
                try:
                     [folder + '/' + item for item in os.listdir(folder)]
                except:
                    pass
            #
            # tableData = tableData.merge(pd.DataFrame(items, columns=['outputPath']), on='outputPath', how='left',
            #                           indicator=True)
            pstart = time.time()
            exists = tableData['outputPath'].apply(os.path.exists)
            print('Time taken to check exists ', tableData['outputPath'].shape[0], ' file exists : ', time.time() - pstart)
            # tableData['lastModified'] = pd.to_datetime(tableData[tableData['_merge']=='both']['outputPath'].apply(os.path.getmtime).apply(time.ctime))
            pstart = time.time()
            tableData['lastModified'] = pd.to_datetime(
                tableData[exists]['outputPath'].apply(os.path.getmtime).apply(time.ctime))
            print('Time taken to check date ', tableData['outputPath'].shape[0], ' file exists : ',
                  time.time() - pstart)
            pstart = time.time()
            tableData['primaryProgPath'] = np.where(tableData['Primary_Progname'],f"//sas-vm/{self.projectID}/Reports/{self.libName}/SAS programs/" +\
                                                    tableData['Primary_Progname'],
                                                    tableData['Primary_Progname'])
            tableData['primaryLogPath'] = np.where(tableData['Primary_Progname'],f"//sas-vm/{self.projectID}/Reports/{self.libName}/SAS programs/" + \
                                                   tableData['ITEMID']+'.log',
                                                   tableData['Primary_Progname'])


            for folder in tableData['primaryLogPath'].str.split('/').apply(lambda x: "/".join(x[:-1])).replace('',
                                                                                                           np.nan).dropna().unique():
                try:
                    [folder + '/' + item for item in os.listdir(folder)]
                except:
                    pass
            # exists = tableData['primaryLogPath'].apply(os.path.exists)
            tableData['qcProgPath'] = np.where(tableData['QC_Progname'],f"//sas-vm/{self.projectID}/Reports/{self.libName}/SAS validation/" + \
                                      tableData['QC_Progname'],tableData['QC_Progname'])
            tableData['qcLogPath'] = np.where(tableData['QC_Progname'],f"//sas-vm/{self.projectID}/Reports/{self.libName}/SAS validation/" + 'v_'+tableData[
                'ITEMID']+'.log',tableData['QC_Progname'])
            tableData['qcmatchPath'] = f"//sas-vm/{self.projectID}/Reports/{self.libName}/metadata/qcmatch.txt"
            tableData['Validation_fail'] = tableData.apply(self.checkValidations,axis=1)
            # tableData['Errors'] = tableData['Validation_fail'].apply(len).astype('str').str.replace('0','')
            tableData['Order'] = tableData['obj_order']
            tableData['Last Run'] = tableData['outputPath'].apply(
                lambda x: datetime.fromtimestamp(os.path.getmtime(x)).strftime('%b %d,%Y') if os.path.exists(x) else '')
            treeModel = CustomTreeModel(tableData, cols, parent=self)
            treeModel = createTreeModel(treeModel,'objectID','PARENTID',self.libName,'Name','Object_Desc')

            for i,col in enumerate(cols):
                treeModel.setHeaderData(i,Qt.Horizontal,col)
            return treeModel

        elif self.libType == 'Analysis':
            start = time.time()
            try:
                datasets,_ = pyreadstat.read_sas7bdat(f"//sas-vm/{self.projectID}/Data/{self.libName}/meta/datasets.sas7bdat")
            except Exception as e:
                msg = customQMessageBox("Datasets file not found")
                # self.disableAllIcons()
                msg.exec_()
                self.clearTable()
                return pd.DataFrame()

            print("Time taken to read dataset : ", time.time()-start)
            start = time.time()
            datasets['object_name'] = datasets['DATASET']
            datasets['PROJECTID'] = self.projectID
            datasets['LIBID'] = self.libID
            datasets[['Object_Desc']] = datasets['DESCRIPTION']
            datasets['objectID'] = datasets.index
            newDF = datasets[['objectID', 'PROJECTID', 'LIBID','object_name', 'Object_Desc']].copy()
            oldDF = pd.read_sql_query(
                F"SELECT * from util_obj WHERE PROJECTID='{self.projectID}' and LIBID='{self.libID}'", s.db)
            merged = oldDF.merge(newDF, how='outer', on=['PROJECTID', 'LIBID', 'object_name'], suffixes=('_x', ''),
                                 indicator=True)
            outdatedData = merged[merged['_merge'] == 'left_only']
            outdatedData = outdatedData['objectID_x'].values.tolist()
            newDF = merged[merged['_merge'] == 'right_only']
            newDF = newDF[['objectID', 'PROJECTID', 'LIBID', 'object_name', 'Object_Desc']]
            newID = pd.read_sql_query(
                f"SELECT * from util_obj where PROJECTID = '{self.projectID}' and LIBID = '{self.libID}' ", s.db).replace(
                np.nan, '')
            newID = 0 if newID.empty else newID['objectID'].astype(int).max() + 1
            newDF['objectID'] = [i for i in range(newID, newID + newDF.shape[0])]
            print("Time taken for processing : ", time.time() - start)
            start = time.time()
            cur = s.db.cursor()
            for idx,row in newDF.iterrows():

                cur.execute(
                    f"INSERT IGNORE into util_obj (objectID,PROJECTID,LIBID,object_name,Object_Desc) VALUES (%s,%s,%s,%s,%s) ON DUPLICATE KEY UPDATE  Object_Desc = %s ",
                    (*row, row['Object_Desc']))

            s.db.commit()
            print("Time taken to update  database : ", time.time() - start)
            newDF = pd.read_sql_query(
                f"SELECT * from util_obj where PROJECTID = '{self.projectID}' and LIBID = '{self.libID}' ", s.db).replace(np.nan,'')
            newDF['outDated'] = False
            newDF.loc[newDF['objectID'].isin(outdatedData), 'outDated'] = True
            return  newDF

    def generateReportTree(self,projectID,libName):

        reports = pd.read_sql_query(
            f"SELECT * from items as I  where I.ProjectID = '{projectID}'  and I.PARENTID = '{libName}'",
            s.db)
        parents = [f"{libName}/{folder}" for folder in reports['ITEMID'].values.tolist()]
        while len(parents):
            parent  = parents.pop()
            df = pd.read_sql_query(
                f"SELECT * from items as I  where I.ProjectID = '{projectID}'  and I.PARENTID = '{parent}'",
                s.db)
            df['path'] =   f"{parent}/" + df['ITEMID']
            df['PARENTID'] = parent.split('/')[-1]
            reports = reports.append(df)

            parents += df['path'].values.tolist()
        reports = reports[reports['TEMPLATE'].isin(['Listing','Summary','Figure'])]
        return  reports

    def populateItems(self,utils_obj):
        utils_obj.fillna('', inplace=True)
        utils_obj['objectID'] = utils_obj['objectID'].astype(int)
        utils_obj['Description'] =  utils_obj['Object_Desc']
        utils_obj["Name"] = utils_obj['object_name']
        utils_obj['Category'] = utils_obj['custom_Columns'].apply(
            lambda x: "\n".join([f"{k}:{v}" for k, v in json.loads(x).items()]) if x else '')

        utils_obj['Category'] = np.where(utils_obj['Category'].str.decode('ascii').isna(),
                                         utils_obj['Category'],
                                         utils_obj['Category'].str.decode('ascii'))
        utils_obj['Primary Owner'] = utils_obj['Primary_Owner']
        utils_obj['Primary_Notes'] = np.where(utils_obj['Primary_Notes'].str.decode('ascii').isna(),
                                          utils_obj['Primary_Notes'],
                                          utils_obj['Primary_Notes'].str.decode('ascii'))
        # utils_obj['Primary_Info'] = utils_obj.apply(lambda x: {"owner": x['Primary_Owner'],
        #                                                        "status": x['Primary_Status'],
        #                                                        "program": x['Primary_Progname'],
        #                                                        "note": x['Primary_Notes']}, axis=1)

        utils_obj['Primary Info'] = utils_obj['Primary_Status'].apply(addSeperator)+ \
                                    utils_obj['Primary_Progname'].apply(lambda x:addSeperator(x,'\n'))+ \
                                    utils_obj['Primary_Notes']
        # utils_obj['Primary Info'] = ''
        utils_obj['Primary Info'] = utils_obj['Primary Info'].str.strip(';').str.strip('\n')
        utils_obj['QC Type'] = utils_obj['QC_Type']
        utils_obj['qc_notes'] = np.where(utils_obj['qc_notes'].str.decode('ascii').isna(),
                                                 utils_obj['qc_notes'],
                                                 utils_obj['qc_notes'].str.decode('ascii'))
        utils_obj['QC Owner'] = utils_obj['QC_Owner']
        utils_obj['QC Info'] = utils_obj['QC_Status'].apply(addSeperator)+\
                               utils_obj['QC_Progname'].apply(lambda x:addSeperator(x,'\n'))+\
                               utils_obj['qc_notes']

        utils_obj['QC Info'] = utils_obj['QC Info'].str.strip(';').str.strip('\n')

        utils_obj["Match"] = ''
        utils_obj["Order"] = utils_obj["obj_order"]
        print(utils_obj.shape)
        utils_obj['Validation_fail'] = utils_obj.apply(self.checkValidations, axis=1)
        # utils_obj['Validation_fail'] = ''
        # utils_obj['Errors'] = utils_obj['Validation_fail'].apply(len).astype(str).str.replace('0','')


        utils_obj = utils_obj.replace(np.nan,'').sort_values('objectID')

        # pd.read_sql_query(f"SELECT * from util_obj where PROJECTID = '{projectID}' and LIBID = '{libID}' ")
        cols = ["objectID","Description","Name","Order","Category","Primary Owner","Primary Info","QC Type","Match","QC Owner","QC Info","Last Run"]
        tableModel = BaseTableModel(utils_obj,cols,parent=self)
        # tableModel.delegateColumns = [7]
        self.treeView.setModel(tableModel)
        # self.treeView.setItemDelegate(PrimaryInfoDelegate())
        # for i in range(self.treeView.model().rowCount()):
        #     self.treeView.openPersistentEditor(self.treeView.model().index(i,7))

        self.treeView.header().resizeSections(QHeaderView.ResizeToContents)
        self.treeView.hideColumn(0)
        self.treeView.hideColumn(1)
        self.treeView.showColumn(2)
        return utils_obj

    def setupUI(self):
        self.mainLayout = QVBoxLayout(self)
        self.topLayout = QVBoxLayout(self)
        self.topLayout1 = QHBoxLayout(self)
        # self.instructionLabel = QLabel('Select Project & Library to get started')
        # self.topLayout1.addWidget(self.instructionLabel)
        self.projectLabel = QLabel('Project')

        self.topLayout1.addWidget(self.projectLabel)
        self.projectDropDown = QComboBox()
        self.projectDropDown.view().setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.projectDropDown.view().setBaseSize(150, self.projectDropDown.height())
        # self.projectDropDown.setSizePolicy(QSizePolicy.MinimumExpanding, QSizePolicy.Minimum)
        # self.projectDropDown.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLength)
        self.topLayout1.addWidget(self.projectDropDown)

        self.libraryLabel = QLabel('Library')
        self.topLayout1.addWidget(self.libraryLabel)

        self.libraryDropDown = QComboBox()
        view = QListView(self.libraryDropDown)
        self.libraryDropDown.setView(view)
        view.setTextElideMode(Qt.ElideNone)
        # self.libraryDropDown.view().setTextElideMode(Qt.ElideNone)
        self.libraryDropDown.view().setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        # self.libraryDropDown.view().setBaseSize(800,self.libraryDropDown.height())
        # self.libraryDropDown.setSizePolicy(QSizePolicy.MinimumExpanding,QSizePolicy.Minimum)
        # self.libraryDropDown.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLength)
        self.topLayout1.addWidget(self.libraryDropDown)

        self.topLayout1.addStretch()

        self.topButtonLayout = QHBoxLayout(self)
        self.topButtonLayout.setSpacing(0)
        self.setContentsMargins(0,0,0,0)
        styleSheet = """
            QPushButton{
                background-color:rgba(150,150,150,50);
                border:1px solid rgba(150,150,150,90);
                height:25px;
                padding:3px;
            }
            QPushButton:pressed{
                background-color:rgba(0,120,215,50);
                border:1px solid rgba(0,120,215,90);

            }
            QPushButton:hover{
                background-color:rgba(0,115,204,50);
                border:1px solid rgba(0,115,204,90);
            }
            QPushButton:checked{
                background-color:rgba(0,120,215,50);
                border:1px solid rgba(0,120,215,90);
            }
           
        """
        self.itemsButton = QPushButton('Items')
        self.itemsButton.setDisabled(True)
        self.itemsButton.setCheckable(True)
        self.topButtonLayout.addWidget(self.itemsButton)

        self.itemsButton.setStyleSheet(styleSheet)
        self.tasksButton = QPushButton('Tasks')
        self.tasksButton.setDisabled(True)
        self.tasksButton.setCheckable(True)
        self.topButtonLayout.addWidget(self.tasksButton)
        self.tasksButton.setStyleSheet(styleSheet)
        self.issuesButton = QPushButton('Issues')
        self.issuesButton.setDisabled(True)
        self.issuesButton.setCheckable(True)
        self.issuesButton.setStyleSheet(styleSheet)
        self.topButtonLayout.addWidget(self.issuesButton)
        self.topLayout1.addLayout(self.topButtonLayout)


        self.itemsButton.clicked.connect(self.selectState)
        # self.itemsButton.setDisabled(True)
        self.tasksButton.clicked.connect(self.selectState)

        # self.tasksButton.setDisabled(True)

        self.issuesButton.clicked.connect(self.selectState)
        # self.issuesButton.setDisabled(True)
        # self.userLabel =QLabel(f"Current user: {self.currentUsername}")
        # font  = QFont()
        # font.setBold(True)
        # self.userLabel.setFont(font)
        # self.userLabel.setSizePolicy(QSizePolicy.Minimum,QSizePolicy.Fixed)
        # self.topLayout.addWidget(self.userLabel,alignment=Qt.AlignRight)
        self.parent().setWindowTitle(self.parent().windowTitle()+f" : {self.currentUsername}")
        self.topLayout.addLayout(self.topLayout1)
        self.iconsWidget = QWidget()
        self.iconsWidget.setStyleSheet("background-color:#e1e1e1;margin:0px")
        self.iconsWidget.setContentsMargins(3,0,0,0)
        # self.iconsWidget

        # self.iconsWidget.setSizePolicy(QSizePolicy.Fixed,QSizePolicy.Fixed)
        self.iconsWidget.setMinimumHeight(37)
        self.iconsWidget.setMaximumHeight(37)
        self.iconsLayout = QHBoxLayout(self)
        self.iconsLayout.setSpacing(2)
        self.iconsLayout.setContentsMargins(0,0,0,0)
        # self.iconsWidget.setStyleSheet()
        self.iconsWidget.setLayout(self.iconsLayout)
        self.edit = customButtton(icon=QIcon('icons/edit.png'),name='edit',tooltipText='Edit')
        self.iconsLayout.addWidget(self.edit,alignment=Qt.AlignLeft)
        self.refresh = customButtton(icon=QIcon('icons/reset.png'), name='refresh',tooltipText='Refresh', func=self.refreshData)
        self.iconsLayout.addWidget(self.refresh, alignment=Qt.AlignLeft)
        self.category = customButtton(icon=QIcon('icons/category.png'), name='category', tooltipText='Category',func=self.categorizeData)
        self.iconsLayout.addWidget(self.category, alignment=Qt.AlignLeft)
        self.runBtn = customButtton(icon=QIcon('icons/run.png'),tooltipText='Run', name='run')
        self.iconsLayout.addWidget(self.runBtn, alignment=Qt.AlignLeft)
        self.scanLogs = customButtton(icon=QIcon('icons/scan.png'),tooltipText='Scan logs', name='scan')

        self.iconsLayout.addWidget(self.scanLogs, alignment=Qt.AlignLeft)

        self.comment = customButtton(icon=QIcon('icons/comment.png'), tooltipText='Comment', name='comment',func=self.addComment)
        self.iconsLayout.addWidget(self.comment, alignment=Qt.AlignLeft)
        self.add = customButtton(icon=QIcon('icons/add.png'), tooltipText='Add', name='add', func=self.addTask)
        self.iconsLayout.addWidget(self.add, alignment=Qt.AlignLeft)
        self.filter = customButtton(icon=QIcon('icons/filter.png'), tooltipText='Filter', name='filter',
                                    func=lambda: print('filter clicke'))
        self.iconsLayout.addWidget(self.filter, alignment=Qt.AlignLeft)
        self.remove = customButtton(icon=QIcon('icons/trash.png'), tooltipText='Remove', name='remove',
                                    func=self.removeTask)
        self.iconsLayout.addWidget(self.remove, alignment=Qt.AlignLeft)
        self.tfl = customButtton(icon=QIcon('icons/pdf.png'), tooltipText='TFL tools', name='tfl')
        self.tfl.hide()

        self.iconsLayout.addWidget(self.tfl, alignment=Qt.AlignLeft)
        self.iconsLayout.addStretch()
        # s.processingBtn.setFixedSize(QSize(300, 32))
        # s.processingBtn.setStyleSheet("""QPushButton {background-color:transparent;
        #                                     border:0px;}"""
        #                                    )
        #
        # self.iconsLayout.addWidget(s.processingBtn, alignment=Qt.AlignRight)


        self.topLayout.addWidget(self.iconsWidget)

        self.mainLayout.addLayout(self.topLayout)
        self.tableLabel = QLabel(" <strong>Table here after selecting Items/Tasks/Issues </strong>")
        self.tableLabel.resize(300, 300)
        self.mainLayout.addWidget(self.tableLabel,alignment=Qt.AlignCenter)

        self.treeView = MyTreeView()
        self.treeView.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.treeView.setSelectionBehavior(QAbstractItemView.SelectRows)

        self.treeView.hide()
        self.splitter = QSplitter(self)
        self.splitter.setChildrenCollapsible(False)
        self.splitter.addWidget(self.treeView)
        self.detailView =QListView()
        self.detailView.setWordWrap(True)
        self.detailView.setResizeMode(QListView.Adjust)
        self.detailView.setMinimumWidth(300)

        self.detailView.setSizePolicy(QSizePolicy.Minimum,QSizePolicy.MinimumExpanding)

        self.detailView.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.detailView.hide()
        self.splitter.addWidget(self.detailView)
        self.mainLayout.addWidget(self.splitter)
        self.splitter.setSizes([self.parent().width*(5/4),self.parent().width*(1/4)])
        self.detailView.setMaximumWidth(self.parent().width)


    @validateAdmin
    def addTask(self):
       # TODO: Add validation for duplication names on root level
        addDialog = AddTaskDialog(self)
        addDialog.exec_()
        if addDialog.newTaskAdded:
            self.populateTasks()

    def selectState(self):
        btn = self.sender()
        if btn.text() == self.state:
            btn.setChecked(True)
            return

        for i in range(self.topButtonLayout.count()):
            if self.topButtonLayout.itemAt(i).widget() != btn:
                self.topButtonLayout.itemAt(i).widget().setChecked(False)
        self.detailView.setModel(ListModel([]))
        if btn.text() == 'Items':
            self.state = 'Items'
            if self.projectID and self.libID:
                data = pd.read_sql_query(
                    f"SELECT T.*,D.* from team_perm as T inner join datlib as D on T.PDETAIL=D.LIBID and T.ProjectID = D.PROJECTID where T.userID='{self.currentUser}' and T.PROJECTID = '{self.projectID}'  and T.PRESOURCE='Data Access'",
                    s.db)
                data = data[data['LTYPE'].isin(['SDTM', 'SEND', 'Analysis', 'Reports'])]
                data['name_'] = 'Data: ' + data['LIBRARY']
                reports = pd.read_sql_query(
                    f"SELECT T.*,I.* from team_perm AS T  inner join items as I on T.PDETAIL=I.SYSTEMID and T.ProjectID =  I.projectid  where T.userID='{self.currentUser}' AND T.PROJECTID = '{self.projectID}' and T.PRESOURCE='Reports Access'",
                    s.db)

                reports['name_'] = 'Reports: ' + reports['ITEMID']
                self.librarySelected(data,reports)
                try:
                    self.treeView.doubleClicked.disconnect()
                except:
                    pass




        if btn.text() == 'Tasks':
            self.state = 'Tasks'
            self.populateTasks()


        if btn.text() == 'Issues':
            self.state = 'Issues'
            self.populateIssues()

    @validateAdmin
    def removeTask(self):
        if not self.treeView.currentIndex().isValid():
            msg = customQMessageBox('Please select a task to remove ')
            msg.exec_()
            return

        confirmDialog = QDialog()
        confirmDialog.setWindowTitle("Confirm")#adhoc
        confirmDialog.setWindowFlags(confirmDialog.windowFlags() & ~Qt.WindowContextHelpButtonHint)#adhoc
        confirm = {}
        mainLayout = QVBoxLayout()
        mainLayout.addWidget(QLabel("Are you sure you want to delete?"))
        buttonLayout = QHBoxLayout()
        okButton = QPushButton('Ok')
        cancelButton = QPushButton('Cancel')
        buttonLayout.addWidget(okButton)
        okButton.clicked.connect(lambda:confirm.update({'confirm':True}))
        okButton.clicked.connect(confirmDialog.close)
        cancelButton.clicked.connect(lambda:confirm.update({'confirm':False}))
        cancelButton.clicked.connect(confirmDialog.close)
        buttonLayout.addWidget(cancelButton)
        mainLayout.addLayout(buttonLayout)
        confirmDialog.setLayout(mainLayout)
        confirmDialog.exec_()

        if not confirm['confirm']:
            return

        row = self.treeView.model().itemFromIndex(self.treeView.currentIndex()).id
        tree = getAllChildren(self.treeView.model()._data, str(row), 'taskid', 'Parentid')

        allChildrenStr = str(tuple(tree.keys())).replace(',)',')')
        sql = f"update util_task set deletedTask = 'Y' where taskid in {allChildrenStr}"
        s.db.cursor().execute(sql)
        s.db.commit()

        # self.populateTasks()
        self.treeView.model().setSource(self.treeView.model()._data[~self.treeView.model()._data['taskid'].isin([int(x) for x in tree.keys()])])
        # treeModel = createTreeModel(self.treeView.model(), 'taskid', 'Parentid', '', 'taskid', 'task_type')
        # self.treeView.setModel(treeModel)
        row =self.treeView.currentIndex().row()
        col =self.treeView.currentIndex().column()
        parent = self.treeView.currentIndex().parent()
        self.treeView.model().removeRow(row,parent)
        # self.treeView.model().layoutChanged.emit()
        self.treeView.setCurrentIndex(self.treeView.model().index(0,0))
        self.treeView.header().resizeSections(QHeaderView.ResizeToContents)

    def populateTasks(self):
        self.populateIcons()
        self.clearTable()
        data = pd.read_sql_query(f"SELECT * FROM util_task WHERE PROJECTID = '{self.projectID}' AND LIBID = '{self.libID}' AND (deletedTask <> 'Y' or deletedTask IS null)",s.db)
        data['Name'] = data['task_Name']
        data['Owner'] = data['task_Owner']
        data['Status'] = data['task_status']
        data['Planned Start Date'] = pd.to_datetime(data['Planned_Start_Date']).dt.strftime("%b %d,%Y").fillna("yyyy/mm/dd")
        data['Planned End Date'] = pd.to_datetime(data['Planned_End_Date']).dt.strftime("%b %d,%Y").fillna("yyyy/mm/dd")
        data['Completion Date'] = pd.to_datetime(data['Completion_Date']).dt.strftime("%b %d,%Y").fillna("yyyy/mm/dd")
        data['Notes'] = np.where(data['Status_Notes'].str.decode('ascii').isna(),
                                             data['Status_Notes'],
                                             data['Status_Notes'].str.decode('ascii'))

        cols =  ['Name','taskid', 'Owner', 'Status', 'Planned Start Date', 'Planned End Date', 'Completion Date',
                                 'Notes']

        # ignored = data[(data['task_status'].isin(['Cancelled', 'Not Applicable'])&(data['task_type']=='Task'))].index.values.tolist()
        # completionPer = {}

        #     leaves = [int(k) for k, v in tree.items() if len(tree[k]) == 0]
        #     tasks = data[(data['taskid'].isin(leaves)) & (~data['task_status'].isin(['Cancelled', 'Not Applicable']))]
        #     for taskid in tasks['taskid'].values.tolist():
        # newdata = data.drop(ignored)

        if not data.empty:
            ignoreMask = ~data['task_status'].isin(['Cancelled', 'Not Applicable'])
            for i,row in data[(data['task_type']=='Task')&(ignoreMask)].iterrows():

                comp = pd.read_sql_query(f"SELECT x.completed/ y.total as comp from(select count(taskid) as completed from util_task_items where taskid = {row['taskid']} and PROJECTID='{self.projectID}' and LIBID='{self.libID}' and titem_Status= 'X') x JOIN (SELECT COUNT(taskid) AS total FROM util_task_items WHERE taskid = {row['taskid']} and PROJECTID='{self.projectID}' and LIBID='{self.libID}' AND (titem_Status <> 'Z' OR titem_Status IS null)) y on 1=1",s.db)
                comp = comp['comp'][0]  if comp['comp'][0] is not None else 0.

                data.loc[data['taskid']==row['taskid'],'completion'] = comp

                # taskItems = taskItems[~(taskItems['titem_Status'] == 'Z')]
                # ncompleted = taskItems[taskItems['titem_Status'] == 'X'].shape[0]
                # N = taskItems.shape[0]
                # completionPer = int(ncompleted / N)*100
                # print(f"{parent} : {completionPer} % completed")
            # completionPer = {}
            if not 'completion' in data.columns:
                data['completion'] = 0.0

            statusValues = ['','Completed', 'Working', 'Waiting', 'Deferred', 'Cancelled', 'Not Applicable' ]
            maxlen = len(sorted(statusValues,key=len)[-1])


            for parent in data[data['Parentid'] == '']['taskid'].values.tolist():
                tree = getAllChildren(data[ignoreMask], str(parent), 'taskid', 'Parentid')
                tree = deque((k,v) for k, v in tree.items())
                while tree:
                    k,v = tree.popleft()
                    # print(k,v)
                    if len(v) > 0 :
                        if data[data['taskid'].isin(v)]['completion'].isna().all():
                            data.loc[data['taskid'].isin(v),'completion'] = 0.0
                            tree.append((k,v))
                        else:
                            data.loc[data['taskid']==int(k),'completion'] = data[data['taskid'].isin(v)]['completion'].mean()

            data['completion'] = data['completion'].fillna(0)
        # data['Status']  = data['task_status'].replace(np.nan,'').str.strip().apply(lambda x :x+" "*(14 - len(x) +1))+ data['completion'].apply(lambda x:f"{x*100:.1f}% completed") # 14 - len of not applicable
            data['Status']  = data['task_status'].replace(np.nan,'').str.strip().apply(lambda x:x+', ' if len(x) > 0 else x)+ data['completion'].apply(lambda x:f"{x*100:.1f}%") # 14 - len of not applicable
        data = data.fillna('')
        treeModel = CustomTreeModel(data,cols,parent=self)
        treeModel = createTreeModel(treeModel,'taskid','Parentid','','taskid','task_type')
        self.detailView.setModel(ListModel([]))

        for i, col in enumerate(cols):
            treeModel.setHeaderData(i, Qt.Horizontal, col)
        self.treeView.setModel(treeModel)

        self.treeView.header().resizeSections(QHeaderView.ResizeToContents)
        try:
            self.treeView.doubleClicked.disconnect()
            self.detailView.clicked.disconnect()
            self.treeView.selectionModel().currentRowChanged.disconnect()
        except:
            pass
        # self.treeView.clicked.connect(self.expandTask)
        self.treeView.doubleClicked.connect(self.editTask)
        self.treeView.setColumnHidden(0,False)
        self.treeView.hideColumn(1)
        self.treeView.setSelectionMode(QAbstractItemView.SingleSelection)
        self.treeView.selectionModel().currentRowChanged.connect(self.populateListView)
        self.treeView.collapseAll()
        colWidthMap = {'Name':400,'Owner':100,'Status':75,'Planned Start Date':120,'Planned End Date':120,'Completion Date':120,'Owner': 100, 'Notes': 600,}
        self.treeView.header().setMaximumSectionSize(600)
        self.treeView.header().setMinimumSectionSize(40)
        for k, v in colWidthMap.items():
            i = cols.index(k)
            self.treeView.setColumnWidth(i, v)
        self.treeView.setCurrentIndex(self.treeView.model().index(0,0))

    def expandTask(self,index):
        if self.treeView.isExpanded(index):
            self.treeView.collapse(index)
        else:
            self.treeView.expand(index)

    def editTask(self,index):
        row = self.treeView.model().itemFromIndex(index).id
        rowData = self.treeView.model()._data.loc[self.treeView.model()._data['taskid']==row]
        milestone = rowData['Parentid'] == ''
        milestone = milestone.all()
        # milestone = self.treeView.model().itemFromIndex(self.treeView.currentIndex().siblingAtColumn(0)).hasChildren()
        owner = rowData['task_Owner'].values[0]


        @validateUsers([owner]+self.adminUsers,"Access Error. Only admins and owner have access to this functionality.")
        def showDialog():
            editDialog = TaskEdit(self,isMilestone=milestone)
            editDialog.exec_()

        showDialog()

    def populateIssues(self):
        self.clearTable()
        self.populateIcons()
        data = pd.read_sql_query(f"Select * from util_issues where PROJECTID='{self.projectID}' and LIBID = '{self.libID}'",s.db,parse_dates=['open_date'])

        data['issue_Title'] = data['issue_Title'].apply(lambda x:x.decode())
        data['issue_Detail'] = data['issue_Detail'].apply(lambda x: x.decode())
        data['ID'] = data['issueid']
        data['Title'] = data['issue_Title']
        data['Detail'] = data['issue_Detail']
        data['Impacts'] = data['issue_impact']
        data['Opened By'] = data['open_By']
        data['Opened Date'] = pd.to_datetime(data['open_date']).dt.strftime('%b %d,%Y')
        data['Closed Date'] = pd.to_datetime(data['close_date']).dt.strftime('%b %d,%Y')
        data['Assigned To'] = data['assigned_to']
        data['Status'] = data['issue_Status']
        data['parent'] = ''
        data['hasComment'] = False
        hasComment = pd.read_sql_query(f"SELECT issueid FROM util_issue_comment WHERE comment_userid = 'dipeshs'", s.db)
        data.loc[data['issueid'].isin(hasComment['issueid']), 'hasComment'] = True
        data = data.sort_values(['open_date','issueid'],ascending=False)
        data = data.replace(np.nan,'')
        cols = ['issueid', 'ID', 'Title', 'Detail', 'Impacts', 'Opened By', 'Opened Date', 'Closed Date', 'Assigned To',
                'Status']
        model = BaseTableModel(data,cols=cols,wrapCols=['Detail'],parent=self)
        # model = createTreeModel(model, 'issueid', 'parent', '', 'issueid', 'issue_Title')
        # model =BaseTableModel(pd.DataFrame({'a':[1,2],'b':[1,2],'c':[1,2],'d':[1,2],'e':[1,2]}),cols=['a','b','c','d','e'],parent=self)

        self.treeView.setModel(model)

        self.treeView.header().resizeSections(QHeaderView.ResizeToContents)
        # commentsModel = ListModel(['Comment 1 ', 'Comment 2'])
        # self.detailView.setModel(commentsModel
        self.treeView.hideColumn(0)
        self.treeView.showColumn(1)
        self.treeView.showColumn(2)
        self.treeView.setTextElideMode(Qt.ElideNone)

        colWidthMap = {'ID': 50, 'Title': 150, 'Detail': 300, 'Opened By': 100, 'Opened Date': 100, 'Assigned To': 100,
                       'Closed Date': 100, 'Impacts': 150, 'Status': 75}
        self.treeView.header().setMaximumSectionSize(600)
        self.treeView.header().setMinimumSectionSize(40)
        for k, v in colWidthMap.items():
            i = cols.index(k)
            self.treeView.setColumnWidth(i, v)

        try:
            self.treeView.selectionModel().currentRowChanged.disconnect()
            self.treeView.doubleClicked.disconnect()
            self.detailView.clicked.disconnect()

        except:
            pass

        self.treeView.selectionModel().currentRowChanged.connect(self.populateComments)
        self.treeView.doubleClicked.connect(self.editIssue)

        self.treeView.setWordWrap(True)
        self.treeView.setCurrentIndex(self.treeView.model().index(0,0))
        if self.treeView.model().rowCount()==1:
            self.populateComments()


    def populateComments(self):

        if len(self.treeView.selectedIndexes())>0:
            row = self.treeView.currentIndex().row()
            issueid = self.treeView.model().visibleData.iloc[row]['issueid']
            comments =  pd.read_sql_query(f"Select  c.*,u.name from util_issue_comment AS c INNER JOIN users  AS  u ON c.comment_userid=u.userID   where issueid = '{issueid}' and  PROJECTID='{self.projectID}' and LIBID = '{self.libID}' ",s.db)
            if not comments.empty:
                commentAttachments = pd.read_sql_query(
                    "SELECT * from util_comment_attachments where comment_id in " + str(tuple(comments['comment_id'].values.tolist())).replace(',)',')'), s.db)



            comments['comment_date'] = pd.to_datetime(comments['comment_date']).dt.strftime('%b %d')
            comments['comment_text']= np.where(comments['comment_text'].str.decode('ascii').isna(),
                                                      comments['comment_text'],
                                                      comments['comment_text'].str.decode('ascii'))
            comments = comments.sort_values(['comment_date','comment_id'],ascending=False)
            commentsModel = QStandardItemModel(self)
            # self.detailView.setStyleSheet("background-color:transparent")

            self.detailView.setModel(commentsModel)
            for i,row in comments.iterrows():
                # self.detailView.model().row
                # self.detailView.setIndexWidget(self.detailView.model().index(i),CommentWidget("Vineet Jain","Vineet Jain","Vineet Jain"))
                # )
                item = QStandardItem()
                self.detailView.model().appendRow(item)

                index = self.detailView.model().indexFromItem(item)
                if not comments.empty:
                    attachments = commentAttachments[commentAttachments['comment_id'] == row['comment_id']]
                else:
                    attachments = pd.DataFrame()
                widget = CommentWidget(f"{row['name']} ({row['comment_userid']})",row['comment_text'],row['comment_date'],row['comment_id'],item,attachments,parent=self)

                item.setSizeHint(widget.sizeHint())
                # widget.addItems(['asd','asd'])
                # widget.resize(QSize(300,300))
                self.detailView.setIndexWidget(index,widget)


            # self.treeView.setCurrentIndex(treeIndex)

    def refreshData(self):
        # @loading(self,'Refreshing...')
        # def func():
        restartConnection()
        self.detailView.setModel(ListModel())
        self.resetData()
        if self.state == 'Items':
            data = pd.read_sql_query(
                f"SELECT T.*,D.* from team_perm as T inner join datlib as D on T.PDETAIL=D.LIBID and T.ProjectID = D.PROJECTID where T.userID='{self.currentUser}' and T.PROJECTID = '{self.projectID}'  and T.PRESOURCE='Data Access'",
                s.db)
            data = data[data['LTYPE'].isin(['SDTM', 'SEND', 'Analysis', 'Reports'])]
            data['name_'] = 'Data: ' + data['LIBRARY']
            reports = pd.read_sql_query(
                f"SELECT T.*,I.* from team_perm AS T  inner join items as I on T.PDETAIL=I.SYSTEMID and T.ProjectID =  I.projectid  where T.userID='{self.currentUser}' AND T.PROJECTID = '{self.projectID}' and T.PRESOURCE='Reports Access'",
                s.db)

            reports['name_'] = 'Reports: ' + reports['ITEMID']
            self.librarySelected(data,reports)

        elif self.state == 'Tasks':
            self.populateTasks()

        elif self.state == 'Issues':
            self.populateIssues()
        # thread = LoadThread()
        # thread.pstart.connect(lambda:self.processingBtn.setText('Refreshing'))
        # thread.pfinished.connect(lambda:self.processingBtn.setText(''))
        # thread.start()
        # func()
        # thread.pfinished.emit()

    @validateAdmin
    def categorizeData(self):
        df = pd.read_sql_query(f"SELECT * from util_custom where PROJECTID = '{self.projectID}' and LIBID = '{self.libID}'", s.db)
        vals = []
        for cat, values in df[['column_Name', 'Value_List']].values:
            values = ','.join(json.loads(values))
            value = f"{cat}: {values}"
            vals.append(value)
        vals = '\n'.join(vals)

        category = CategoryDialog(vals)
        category.exec_()

        if category.valueSaved:
            categories = category.textBox.toPlainText().split('\n')
            maxID = pd.read_sql_query("SELECT max(CUSTOMID) as m from util_custom",s.db)['m'].values[0]
            if not maxID:
                maxID = 0
            else:
                maxID = int(maxID) +1
            cur = s.db.cursor()
            cur.execute(f"DELETE from util_custom where PROJECTID = '{self.projectID}' and LIBID = '{self.libID}'")
            for c in categories:
                if not c:
                    return
                #update util_custom
                category = c.split(':')[0].strip()
                values = c.split(':')[1].strip()
                values = list(set([x.strip() for x in values.split(',')])) # select unnique categories
                values = json.dumps(values)
                cur.execute("INSERT into util_custom(CUSTOMID,PROJECTID,LIBID,column_Name,Value_List) values(%s,%s,%s,%s,%s)",(int(maxID),self.projectID,self.libName  if self.libType == 'report' else self.libID,category,values))

                maxID += 1

            s.db.commit()
            self.updateCategories(self.editMenu,remove=True)
            self.removeObsoleteCategories({c.split(':')[0]:[v.strip() for v in c.split(':')[1].strip().split(',')] for c in categories})

    def removeObsoleteCategories(self,categories):
        self.treeView.model()._data['custom_Columns'] = self.treeView.model()._data['custom_Columns']\
            .str.decode('ascii')\
            .fillna(self.treeView.model()._data['custom_Columns'])\
            .apply(lambda x:json.loads(x) if x else {})\
            .apply(lambda x: json.dumps({k: v for k, v in x.items() if k in categories.keys() and v in categories[k]}))
        self.treeView.model()._data['custom_Columns'] = self.treeView.model()._data['custom_Columns'].replace('{}','')

        self.treeView.model()._data['Category'] = self.treeView.model()._data['custom_Columns'].apply( lambda x: "\n".join([f"{k}:{v}" for k, v in json.loads(x).items()]) if x else '')
        self.treeView.model().updateVisibleData()
        if self.libType == 'report':
            treeModel = createTreeModel(self.treeView.model(),'objectID','PARENTID', self.libName,'Name',
                            'Object_Desc')
            self.treeView.setModel(treeModel)
            self.treeView.header().resizeSections(QHeaderView.ResizeToContents)
        cur = s.db.cursor()
        for objectId, custom_Columns in self.treeView.model()._data[['objectID', 'custom_Columns']].values:
            cur.execute(
                f"UPDATE util_obj set custom_Columns = '{custom_Columns}' where PROJECTID = '{self.projectID}' and LIBID = '{self.libID}' and objectID='{objectId}'")

        s.db.commit()

    def updateCategories(self,menu,remove=False,filter=False):
        custom_cols = pd.read_sql_query(
            f"SELECT * from util_custom where PROJECTID = '{self.projectID}' and LIBID = '{self.libID}'", s.db)
        custom_cols['Value_List'] = custom_cols['Value_List'].apply(json.loads)


        if remove:
            for action in menu.actions()[9:-2]:

                    menu.removeAction(action)

        for category, values in custom_cols[['column_Name', 'Value_List']].values:
            categoryMenu = QMenu(category,parent=self)

            for action in ['']+values :
                if filter:
                    action = FilterAction(action, f'{category}|custom_Columns', custom=True, parent=self)
                    categoryMenu.addAction(action)
                else:
                    action = UpdateAction(action,f'{category}|custom_Columns',custom=True,parent=self)
                    categoryMenu.addAction(action)


            menu.insertMenu(menu.actions()[-2],categoryMenu)

    def resetData(self):
        self.filters = {}

        self.treeView.model().updateFilteredData(self.treeView.model()._data)
        # for row in range(self.treeView.model().rowCount()):
        #     self.treeView.closePersistentEditor(self.treeView.model().index(row, 7))
        # self.treeView.setItemDelegateForColumn(7, PrimaryInfoDelegate())
        # for row in range(self.treeView.model().rowCount()):
        #     self.treeView.openPersistentEditor(self.treeView.model().index(row, 7))
        if self.libType == 'report':
            if self.state =='Items':
                treeModel = createTreeModel( self.treeView.model(), 'objectID','PARENTID', self.libName,'Name',
                                'Object_Desc')
                self.treeView.setModel(treeModel)
                self.treeView.header().resizeSections(QHeaderView.ResizeToContents)

            elif self.state == 'Tasks':
                self.populateTasks()


        for group in self.actionGroups:
            for action in group.actions():
                action.setChecked(False)
        for action in self.filterMenu.actions():
            action.setChecked(False)

class TabWidget(QWidget):
    def __init__(self, parent,currentUser):
        super(TabWidget, self).__init__(parent)
        self.currentUser = currentUser.lower()
        self.setupUi()

    def resizeEvent(self, event):
        super(TabWidget, self).resizeEvent(event)
        size = self.tabs.tabBar().size()
        size.setWidth(self.size().width())
        self.tabs.tabBar().resize(size)

        # self.tabs.tabBar().resizeEvent(event)
        # self.tabs.tabBar().setUsesScrollButtons(False)
        # print('sdsd')

    def setupUi(self):
        self.layout = QVBoxLayout(self)

        # Initialize tab screen
        self.tabs = QTabWidget()
        self.tab1 = QWidget()
        self.tab2 = QWidget()
        self.tabs.resize(792, 610)
        self.tabBar = CustomTabBar(self)
        self.tabs.setTabBar(self.tabBar)
        # Add tabs
        self.tabs.addTab(self.tab1, "Project-Library Tracking")
        # self.tabs.addTab(self.tab2, "Misc. Utilities")

        # Create first tab
        self.tab1Layout = QVBoxLayout(self)
        self.tab1Layout.setSpacing(0)
        self.tab1Layout.setContentsMargins(0,0,0,0)
        self.mainUtility = MainUtility(parent=self,currentUser=self.currentUser)
        self.tab1Layout.addWidget(self.mainUtility)
        self.tab1.setLayout(self.tab1Layout)

        # Create second tab
        # self.tab2Layout = QVBoxLayout(self)
        # self.sasWidget = SpecsSAS()
        # self.tab2Layout.addWidget(self.sasWidget)
        # self.tab2.setLayout(self.tab2Layout)
        self.centralWidget = QWidget(self)
        self.tab2Layout = QHBoxLayout(self)
        self.utilityListView = QListView(self)
        self.listModel = ListModel(['Excel Compare', ])
        self.utilityListView.setModel(self.listModel)
        self.utilityListView.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.utilityListView.setMaximumWidth(200)
        self.utilityListView.selectionModel().currentRowChanged.connect(self.stepSelected)
        self.rightWidget = QWidget(self)
        self.rightlayout = QVBoxLayout()

        self.rightWidget.setLayout(self.rightlayout)
        splitter = QSplitter(self)
        splitter.setOrientation(Qt.Horizontal)
        splitter.addWidget(self.utilityListView)
        splitter.addWidget(self.rightWidget)
        splitter.setChildrenCollapsible(False)
        self.tab2Layout.addWidget(splitter)
        self.placeHolderWidget = QLabel("Select a utility from the list", parent=self.rightWidget)
        self.rightlayout.addWidget(self.placeHolderWidget, alignment=Qt.AlignCenter)
        self.tab2.setLayout(self.tab2Layout)

        # Add tabs to widget
        self.layout.addWidget(self.tabs)
        self.setLayout(self.layout)

    def stepSelected(self):
        if len(self.utilityListView.selectedIndexes()) < 1: return
        step = self.listModel.itemData(self.utilityListView.selectedIndexes()[0])[0]

        for i in range(self.rightlayout.count()):
            self.rightlayout.itemAt(i).widget().hide()

        if step == 'Excel Compare':
            self.excelComapre =  ExcelCompare(self)
            self.rightlayout.addWidget(self.excelComapre)


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
        self.includedListView.selectionModel().currentRowChanged.connect(self.populateVarListView)
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

def getPropertyValue(props,col,row):
    if   props.loc[row,col] is None:
        return  ''
    reuse = props.loc[row,col].split(':')
    if reuse[1] == 'true':
        reuseRow = props.loc[row,'REUSEID']
        row = props[props['SYSTEMID'] ==reuseRow].index[0]
        return getPropertyValue(props, col, row)
    else:
        return ":".join(reuse[2:])

if __name__ == '__main__':
    main()
