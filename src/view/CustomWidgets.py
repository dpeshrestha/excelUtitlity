import json
import os
import re
import time
from datetime import datetime
from functools import reduce
from pathlib import Path

import pandas as pd
import numpy as np
from PySide2.QtCore import Qt, QSize, QDateTime, QDate, QUrl, QThread, Signal, Slot
from PySide2.QtGui import QIcon, QStandardItem, QCursor
from PySide2.QtWidgets import QTabBar, QDialog, QVBoxLayout, QLabel, QPushButton, QMenu, QAction, QHBoxLayout, \
    QLineEdit, QPlainTextEdit, QWidget, QListView, QAbstractItemView, QFormLayout, QComboBox, QDateEdit, QSizePolicy, \
    QProgressBar, QFileDialog, QCheckBox
import src.settings as s
from src.delegates.commentDelegate import CommentWidget,AttachmentButton
from src.model.ListModel import ListModel
from src.utils.pdfutil import convertRTFToPDF, merge, createBookmarkDict, genTOC
from src.utils.utils import createTreeModel, orderItems, findChildren
from src.utils.validation import validateUsers, validateAdmin
from src.view.customDialogs import customQMessageBox


class CustomTabBar(QTabBar):
    def __init__(self,parent=True):
        super(CustomTabBar, self).__init__(parent=parent)
        # self.setExpanding(True)

    def tabSizeHint(self, index:int):
        size = super(CustomTabBar, self).tabSizeHint(index)
        w = self.parent().width()/2
        size.setWidth(w)
        return size





class customButtton(QPushButton):
    def __init__(self,parent=None,icon=QIcon(),name=None,tooltipText=None,func=None):
        super(customButtton, self).__init__(parent=parent,icon=icon)
        self.setIconSize(QSize(20,20))
        self.setObjectName(name)
        self.setToolTip(tooltipText)
        self.setMaximumWidth(32)
        self.func = func
        if self.func:
            self.clicked.connect(self.func)
        self.setStyleSheet("""
            QPushButton{
                background-color:white;
                padding: 3px;
            }
            QPushButton:hover:!pressed{background-color:rgba(229,243,255)}
            QPushButton:hover:pressed{background-color:rgba(197,224,247)}
            QPushButton::menu-indicator{image:none}
            """)
    def setMenu(self, menu):
        super(customButtton, self).setMenu(menu)


class EditDialog(QDialog):
    def __init__(self,text='',col='',parent=None):
        super(EditDialog, self).__init__(parent,Qt.WindowSystemMenuHint|Qt.WindowTitleHint)
        self.newValueSaved = False
        self.col = col
        mainLayout = QVBoxLayout()
        self.lineEdit = QPlainTextEdit()
        mainLayout.addWidget(self.lineEdit)
        buttonLayout = QHBoxLayout()
        self.okButton = QPushButton('Ok')
        self.okButton.clicked.connect(self.okClicked)
        self.cancelButton = QPushButton('Cancel')
        self.cancelButton.clicked.connect(self.cancelClicked)
        buttonLayout.addWidget(self.okButton)
        buttonLayout.addWidget(self.cancelButton)

        self.setWindowTitle('Edit')
        self.setSizeGripEnabled(False)
        mainLayout.addLayout(buttonLayout)
        self.setLayout(mainLayout)
        self.setFixedSize(QSize(200,100))
        self.lineEdit.setPlainText(text)

    def okClicked(self):
        self.newValueSaved = True
        self.newValue = self.lineEdit.toPlainText()

        if 'progname' in self.col:
            maxl = 100
        elif 'notes' in self.col:
            maxl = 255
        else:
            maxl = 0
        if maxl:
            msg = customQMessageBox("Text cannot be longer than "+str(maxl)+" characters.")
            msg.exec_()
            return

        self.close()

    def cancelClicked(self):
        self.newValueSaved = False


        self.close()


class AddTaskDialog(QDialog):
    def __init__(self,parent=None):
        #add custom name field
        # validation for Status
        # Status only for tasks not milestones
        # expand/collapse item on single click
        super(AddTaskDialog, self).__init__(parent=parent,f=Qt.WindowSystemMenuHint|Qt.WindowTitleHint)
        self.taskDesign = pd.read_excel('files\TaskDesign.xlsx')
        self.rootTask = self.taskDesign[self.taskDesign['Parent'].isna()]
        print('here add atsk')
        self.setupUI()
        print('created UI add atsk')
        self.setWindowTitle("Add Task")
        self.okButton.clicked.connect(self.okClicked)
        self.cancelButton.clicked.connect(self.cancelClicked)
        self.newTaskAdded = False

    def okClicked(self):
        if self.taskDropDown.currentText() == '':
            msg = customQMessageBox("Please select a predefined milestone")
            msg.exec_()
            return

        if not re.match(r"^[\w+\s+]{1,100}$", self.nameEdit.text().strip()):
            msg = customQMessageBox(
        """Milestone name should only contain alphabets.
           Milestone name should not be blank.
           Milestone name should not be longer than 100 characters. """)
            msg.exec_()
            return False

        cursor = s.db.cursor()
        tasks = self.taskDesign[self.taskDesign['LibraryType'].str.contains(self.parent().libType)]
        tasks.loc[:,'PROJECTID'] = self.parent().projectID
        tasks.loc[:,'LIBID'] = self.parent().libID
        cursor.execute('select max(taskid) as taskid from util_task')
        taskid = cursor.fetchone()[0]
        if not taskid:
            taskid = 1
        else:
            taskid +=1
        tasks.loc[:,'taskid'] = [i for i in range(taskid,taskid+tasks.shape[0])]
        ids = findChildren(tasks, 'Parent', 'Task', 'taskid', self.taskDropDown.currentText())
        tasks = tasks[tasks['taskid'].isin(ids)]
        tasks = tasks.replace(self.taskDropDown.currentText(), self.nameEdit.text().strip())

        sql = "INSERT INTO util_task(taskid,PROJECTID,LIBID,task_Name,Parentid,task_type) values (%s,%s,%s,%s,%s,%s)"
        tasks.loc[:, 'taskid'] = [str(i) for i in range(taskid, taskid + tasks.shape[0])]
        tasks.loc[:,'taskid'] = tasks['taskid'].astype(str)
        tasks.loc[:, 'Parentid'] = tasks['Parent'].apply(
            lambda x: tasks[(tasks['Task'] == x) & (~tasks['Task'].isna())]['taskid'].values.tolist())
        tasks.loc[:, 'Parentid'] = tasks['Parentid'].apply(lambda x: str(int(x[0])) if len(x) else '')
        tasks = tasks.fillna('')

        util_task_values = tasks[['taskid', 'PROJECTID', 'LIBID', 'Task', 'Parentid', 'TaskType']].values.tolist()
        try:
            cursor.executemany(sql,util_task_values)

            s.db.commit()
        except:
            s.db.rollback()

        tasks['temp'] = tasks[tasks['TaskItems'] != '']['TaskItems'].apply(
            lambda x:x.split('\n'))

        tasks = tasks.explode('temp')
        tasks['tOrder'] = ''
        d = {}
        tasks['Task'].apply(lambda x: orderItems(d, x))
        tasks['tOrder'] = [i for k, v in d.items() for i in v]
        itemsList = tasks[~tasks['temp'].isna()][['taskid','tOrder','PROJECTID','LIBID','temp']].values.tolist()



        sql = "INSERT INTO util_task_items(taskid,tOrder,PROJECTID,LIBID,titem_Name) values(%s,%s,%s,%s,%s)"

        try:
            cursor.executemany(sql, itemsList)

            s.db.commit()
        except:
            s.db.rollback()
        self.newTaskAdded = True
        self.close()

    def cancelClicked(self):
        self.close()

    def setupUI(self):
        tasks  = self.rootTask[self.rootTask['LibraryType'].str.contains(self.parent().libType)]['Task'].values.tolist()
        self.mainLayout = QVBoxLayout(self)
        self.taskDropDown = QComboBox()
        self.taskDropDown.addItems(['']+tasks)
        self.mainLayout.addWidget(self.taskDropDown)
        self.nameEdit = QLineEdit()
        self.mainLayout.addWidget(self.nameEdit)
        self.buttonLayout = QHBoxLayout()
        self.okButton = QPushButton('Ok')
        self.buttonLayout.addWidget(self.okButton, alignment=Qt.AlignCenter)
        self.cancelButton = QPushButton('Cancel')
        self.buttonLayout.addWidget(self.cancelButton, alignment=Qt.AlignCenter)
        self.mainLayout.addLayout(self.buttonLayout)
        self.setLayout(self.mainLayout)


class MergeDialog(QDialog):
    def __init__(self,parent=None):
        super(MergeDialog, self).__init__(parent,Qt.WindowSystemMenuHint|Qt.WindowTitleHint)
        self.setupUI()
        self.okButton.clicked.connect(self.okClicked)
        self.cancelButton.clicked.connect(self.cancelClicked)
        self.data = self.parent().treeView.model()._data[~self.parent().treeView.model()._data['outDated']][['object_name','objectID']]
        listModel = ListModel(self.data['object_name'].values.tolist())
        self.listView.setModel(listModel)

    def setupUI(self):
        self.mainLayout = QVBoxLayout()
        self.label = QLabel('Select the new dataset to merge')
        self.mainLayout.addWidget(self.label)
        self.listView = QListView()
        self.mainLayout.addWidget(self.listView)
        self.buttonLayout = QHBoxLayout()
        self.okButton = QPushButton('Ok')
        self.buttonLayout.addWidget(self.okButton,alignment=Qt.AlignCenter)
        self.cancelButton = QPushButton('Cancel')
        self.buttonLayout.addWidget(self.cancelButton,alignment=Qt.AlignCenter)
        self.mainLayout.addLayout(self.buttonLayout)
        self.setLayout(self.mainLayout)
        self.listView.setSelectionMode(QAbstractItemView.SingleSelection)

    @validateAdmin
    def okClicked(self):
        row = self.listView.selectionModel().selectedRows()[0].row()
        objectID = self.data.iloc[row]['objectID']
        outdatedRow = self.parent().treeView.currentIndex().row()
        outdatedID = self.parent().treeView.model().visibleData.iloc[outdatedRow]['objectID']
        cur = s.db.cursor()
        cur.execute(f"delete from util_obj where objectID='{outdatedID}'")
        s.db.commit()
        self.parent().treeView.model().deleteRows()

        cur.execute(f"update util_obj set objectID = '{outdatedID}' where  objectID = '{objectID}' ")
        s.db.commit()
        self.parent().treeView.model()._data.loc[
            self.parent().treeView.model()._data['objectID'] == objectID, 'objectID'] = outdatedID
        self.parent().treeView.model().visibleData.loc[self.parent().treeView.model().visibleData['objectID'] == objectID,'objectID'] = outdatedID
        self.parent().treeView.model().layoutChanged.emit()
        self.close()
        #remove old

    def cancelClicked(self):
        self.close()

class AttachmentItem(QWidget):
    def __init__(self,path,name='',idx=None,parent=None):
        super(AttachmentItem, self).__init__(parent=parent)
        self.path = path
        self.idx = idx
        text = os.path.basename(path) if not name else name
        self.mainLayout = QHBoxLayout(parent)
        self.fileName = QLabel(text)
        self.mainLayout.addWidget(self.fileName)
        self.removeBtn = QPushButton(icon=QIcon("icons/closeOld.png"))
        self.removeBtn.setStyleSheet("*{background-color:transparent}")
        self.removeBtn.setFixedSize(QSize(30,30))
        self.removeBtn.setSizePolicy(QSizePolicy.Minimum,QSizePolicy.Minimum)
        self.mainLayout.addWidget(self.removeBtn)
        self.setLayout(self.mainLayout)


        self.removeBtn.clicked.connect(self.remove)

    def remove(self):
        if self.idx:
            s.cursor.execute(f'DELETE from util_comment_attachments where attachment_id ={self.idx}')
            s.db.commit()

        self.deleteLater()

    def changeFile(self,path):
        self.path = path
        text = os.path.basename(path)
        self.fileName.setText(text)



class CommentDialog(QDialog):
    def __init__(self,row,commentId =None,parent=None):
        super(CommentDialog, self).__init__(parent,Qt.WindowSystemMenuHint|Qt.WindowTitleHint)
        self.issueID = self.parent().treeView.model().visibleData.iloc[row]['issueid']
        self.commentId = commentId


        if commentId is not None:
            self.setWindowTitle('Edit Comment')
            self.attachments = pd.read_sql_query(f"SELECT * from util_comment_attachments where comment_id = {commentId} ",s.db)
        else:
            self.attachments = pd.DataFrame()
            self.setWindowTitle('Add Comment')

        self.setupUI()
        for widget in [AttachmentItem('',name=row['attachment_name'],idx=row['attachment_id'],parent= self) for i,row in self.attachments.iterrows()]:
            self.attachmentLayout.addWidget(widget)
        self.addMoreButton.clicked.connect(self.addAttachments)

        # self.browseButton.clicked.connect(lambda x:QFileDialog.getOpenFileNames(self,"Add Attachments","C:\\","Image files (*.jpg *.png)"))
        self.okButton.clicked.connect(self.okClicked)
        self.cancelButton.clicked.connect(self.cancelClicked)

    def addAttachments(self):
        paths = QFileDialog.getOpenFileNames(self, "Add Attachments", "", "Image files (*.jpg *.png)")[0]
        if not paths:
            return
        savedPaths = [self.attachmentLayout.itemAt(i).widget().path for i in range(self.attachmentLayout.count())]
        fileNames = [self.attachmentLayout.itemAt(i).widget().fileName.text() for i in
                     range(self.attachmentLayout.count())] + [os.path.basename(p) for p in paths]
        fileNamesCount = {}
        for f in fileNames:
            if fileNamesCount.get(f):
                fileNamesCount[f] = fileNamesCount.get(f)+1
            else:
                fileNamesCount[f] = 1
        duplicateNames = []
        for file in fileNamesCount.keys():
            if fileNamesCount[file] > 1:
                duplicateNames.append(file)
        if duplicateNames:
            msg = customQMessageBox("Following files are duplicate please reselect files:\n"+"\n".join(duplicateNames))
            msg.exec_()
            return

        savedPaths += paths
        savedPaths = set(savedPaths)
        # filedata = {}
        # for filename,path in zip(fileNames,savedPaths):
        #     filedata[filename]=path

        for i in range(self.attachmentLayout.count()):
            if self.attachmentLayout.itemAt(i).widget().path:
                self.attachmentLayout.itemAt(i).widget().remove()

        # for name,path in filedata.items():
        for path in savedPaths:
            if path:
                self.attachmentLayout.addWidget(AttachmentItem(path,parent=self))

    def okClicked(self):
        comment = self.comentEdit.toPlainText().strip()

        if  not comment:
            msg = customQMessageBox('Comment cannot be empty')
            msg.exec_()
            return
        if len(comment)>255:
            msg = customQMessageBox("Comment is too long. Comment cannot be more than 255 characters.")
            msg.exec_()
            return


        date = datetime.now().strftime('%Y-%m-%d')
        files = [self.attachmentLayout.itemAt(x).widget().path for x in
                 range(self.attachmentLayout.count())]
        files = [x for x in files if x!=''] #if empty the attachments are old
        fileBlobs = [open(file, 'rb').read() for file in files]


        if self.commentId is not None:

            sql = f"UPDATE util_issue_comment set comment_text = '{comment}' , comment_date= '{date}'   where comment_id = {self.commentId}"
            s.cursor.execute(sql)
            s.db.commit()
            widget = self.parent().detailView.indexWidget(self.parent().detailView.currentIndex())
            savedFilenames = [self.attachmentLayout.itemAt(i).widget().fileName.text() for i in range(self.attachmentLayout.count())]
            attachments = widget.attachments[widget.attachments['attachment_name'].isin(savedFilenames)][['attachment_name','attachment']]
            attachments = attachments.append(pd.DataFrame({'attachment_name':[os.path.basename(f) for f in files],'attachment':fileBlobs})).drop_duplicates()
            widget = CommentWidget(widget.user.text(), comment, widget.date.text(), widget.commentId, widget.item, attachments=attachments, parent=widget.parent())
            # widget.attachmentWidget.hide()
            widget.showAttachments()
            widget.showAttachments()
            self.parent().detailView.setIndexWidget(self.parent().detailView.currentIndex(),widget)

            # self.parent().detailView.update()
        else:
            sql = "INSERT INTO  util_issue_comment(issueid,PROJECTID,LIBID,comment_userid,comment_date,comment_text) values (%s,%s,%s,%s,%s,%s)"
            s.cursor.execute(sql,(str(int(self.issueID)),self.parent().projectID,self.parent().libID,s.currentUser,date,comment))
            s.db.commit()

            item = QStandardItem()
            self.parent().detailView.model().insertRow(0,item)

            index = self.parent().detailView.model().indexFromItem(item)
            userName = s.projectUsers.set_index('userID').loc[s.currentUser]['name']
            userId = s.currentUser
            sql = "SELECT max(comment_id) from util_issue_comment"
            s.cursor.execute(sql)

            self.commentId = s.cursor.fetchone()[0]
            attachments = pd.DataFrame.from_dict({'attachment_name':[os.path.basename(file) for file in files],'attachment':fileBlobs})
            widget = CommentWidget(f"{userName} ({userId})",comment, datetime.now().strftime('%b %d'),self.commentId ,item,attachments,parent=self)

            item.setSizeHint(widget.sizeHint())
            self.parent().detailView.setIndexWidget(index, widget)

        attachSQL = "INSERT INTO util_comment_attachments(comment_id,PROJECTID,LIBID,attachment_name,attachment,attachment_ORDER) values(%s,%s,%s,%s,%s,%s)"
        for i, (file, blob) in enumerate(zip(files, fileBlobs)):
            s.cursor.execute(attachSQL, (self.commentId,self.parent().projectID,self.parent().libID,os.path.basename(file),blob,i))
        # widget.addItems(['asd','asd'])
        # widget.resize(QSize(300,300))
        s.db.commit()


        self.close()

    def cancelClicked(self):
        self.close()

    def setupUI(self):
        self.mainLayout = QVBoxLayout()
        self.comentEdit = QPlainTextEdit()
        self.mainLayout.addWidget(self.comentEdit)
        self.buttonLayout = QHBoxLayout()
        self.okButton = QPushButton('Ok')
        self.mainLayout.addWidget(QLabel("Add attachments:"))
        self.attachmentLayout = QVBoxLayout()
        self.attachmentLayout.setSpacing(1)
        self.attachmentLayout.setContentsMargins(0,0,0,0)
        self.mainLayout.addLayout(self.attachmentLayout)
        self.addMoreButton = QPushButton('Add Files')
        self.mainLayout.addWidget(self.addMoreButton,alignment=Qt.AlignLeft)
        self.addMoreButton.setSizePolicy(QSizePolicy.Minimum,QSizePolicy.Minimum)
        self.buttonLayout.addWidget(self.okButton, alignment=Qt.AlignCenter)
        self.cancelButton = QPushButton('Cancel')
        self.buttonLayout.addWidget(self.cancelButton, alignment=Qt.AlignCenter)
        self.mainLayout.addLayout(self.buttonLayout)
        self.setLayout(self.mainLayout)


class IssueDialog(QDialog):
    def __init__(self,row=None,parent=None):
        #pass row if edit
        super(IssueDialog, self).__init__(parent,Qt.WindowSystemMenuHint|Qt.WindowTitleHint)
        self.setupUI()
        self.row = row

        self.setWindowTitle("Add Issue")
        if self.row is not None:
            self.setWindowTitle("Edit Issue")
            rowData = self.parent().treeView.model().visibleData.iloc[row]
            rowData = self.parent().treeView.model()._data.loc[self.parent().treeView.model()._data['issueid'] == rowData['issueid']].iloc[0]
            self.titleEdit.setText(rowData['issue_Title'])
            self.detailEdit.setPlainText(rowData['issue_Detail'])
            # self.titleEdit.setText(rowData['issue_Title'])
            self.impactEdit.setPlainText(rowData['issue_impact'])
            self.userDropDown.setCurrentText(rowData['assigned_to'])
            self.statusDropDown.setCurrentText(rowData['issue_Status'])

        self.okButton.clicked.connect(self.okClicked)
        self.cancelButton.clicked.connect(self.cancelClicked)


    def okClicked(self):
        formInput = {
            'issue_Title' : self.titleEdit.text().strip(),
            'issue_Detail' : self.detailEdit.toPlainText().strip(),
            'issue_impact' : self.impactEdit.toPlainText().strip(),
            'assigned_to' : self.userDropDown.currentText(),
            'open_By':s.currentUser,
            'open_date':datetime.now().strftime('%Y-%m-%d'),
            'issue_Status': self.statusDropDown.currentText(),
            'LIBID':self.parent().libID,
            'PROJECTID':self.parent().projectID,
        }

        ##validating input


        skip = []
        if formInput['issue_Status'] == 'Closed':
            skip = ['assigned_to']
        for k,v in formInput.items():
            if k not in skip and v == '':
                msg = customQMessageBox(f"All fields are mandatory.")
                msg.exec_()
                return


        if len(formInput['issue_Title']) >   255 :
            msg = customQMessageBox("Title must be less than 255 characters.")
            msg.exec_()
            return
        if len(formInput['issue_Detail']) >   255 :
            msg = customQMessageBox("Detail must be less than 255 characters.")
            msg.exec_()
            return
        if len(formInput['issue_impact']) >   200 :
            msg = customQMessageBox("Impact must be less than 200 characters.")
            msg.exec_()
            return




        tableInput = {k:[v] for k,v in formInput.items()}
        tableInput.update({'Title': [formInput['issue_Title']],
                                       'Detail': [formInput['issue_Detail']],
                                       'Impacts': [formInput['issue_impact']],
                                       'Opened By': [s.currentUser],
                                       'Date': [datetime.now().strftime('%b %d,%Y')],
                                       'Assigned To': [formInput['assigned_to']],
                                       'Status': [formInput['issue_Status']]
                                       })

        # self.parent().treeView.model().updateSourceData(self.parent().treeView.model()._data.merge(pd.DataFrame.from_dict(tableInput),how='outer').fillna(''))
        if self.row is not None:
            del tableInput['Opened By']
            del tableInput['Date']
            del formInput['open_by']
            del tableInput['open_date']
            rowData = self.parent().treeView.model().visibleData.iloc[self.row]
            tableInput['issueid'] = [rowData['issueid']]
            allDataRow = self.parent().treeView.model()._data.loc[self.parent().treeView.model()._data['issueid']==rowData['issueid']]
            for k, v in tableInput.items():

                self.parent().treeView.model()._data.loc[allDataRow.index[0],k] = v[0] if len(v) else ''
            self.parent().treeView.model().updateSourceData(self.parent().treeView.model()._data)

            updateText = ",".join([f"{k}='{v}'" for k,v in formInput.items()])
            sql = f"UPDATE util_issues set {updateText} where issueid='{rowData['issueid']}' and PROJECTID = '{self.parent().projectID}' and LIBID='{self.parent().libID}'"
            s.cursor.execute(sql)

        else:

            newID = pd.read_sql_query('SELECT max(issueid)  as issueid from util_issues;',s.db)
            if not newID['issueid'].values[0]:
                newID = 0
            else:
                newID = newID['issueid'].values[0] + 1
            tableInput['issueid'] = newID
            tableInput['open_date'] = pd.to_datetime(tableInput['open_date'])
            tableInputDF = pd.DataFrame.from_dict(tableInput)
            source =self.parent().treeView.model()._data
            source['open_date'] = pd.to_datetime(source['open_date']) # date time conflict so changed both to datetime
            self.parent().treeView.model().updateSourceData(
                source.merge(tableInputDF, how='outer').fillna(''))

            sql = "INSERT INTO util_issues(issueid,issue_Title,issue_Detail,issue_impact,assigned_to,open_By,open_date,issue_Status,LIBID,PROJECTID)" \
                  " values (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"
            s.cursor.execute(sql,(str(newID),) + tuple(formInput.values()))
        s.db.commit()
        self.close()

    def cancelClicked(self):
        self.close()

    def setupUI(self):
        self.mainLayout = QVBoxLayout()
        self.form = QFormLayout()
        self.titleEdit = QLineEdit()
        self.form.addRow('Title',self.titleEdit)
        self.detailEdit = QPlainTextEdit()
        self.form.addRow('Detail',self.detailEdit)
        self.impactEdit = QPlainTextEdit()
        self.impactEdit.setSizePolicy(QSizePolicy.MinimumExpanding,QSizePolicy.Minimum)
        self.impactEdit.setMinimumHeight(self.detailEdit.height() / 6)
        self.impactEdit.setMaximumHeight(self.detailEdit.height() / 4)
        self.form.addRow('Impacts',self.impactEdit)

        self.userDropDown = QComboBox()
        # projectUsers = pd.read_sql_query(f"select Distinct userID from team where PROJECTID = '{self.parent().projectID}'",s.db)['userID'].values.tolist()
        self.userDropDown.addItems(['']+s.projectUsers['userID'].values.tolist())
        self.form.addRow('Assigned To',self.userDropDown)
        self.statusDropDown = QComboBox()
        self.statusDropDown.addItems(['Open','Deferred','Closed'])
        self.form.addRow('Status',self.statusDropDown)
        self.buttonLayout = QHBoxLayout()
        self.okButton = QPushButton('Ok')
        self.mainLayout.addLayout(self.form)
        self.buttonLayout.addWidget(self.okButton, alignment=Qt.AlignCenter)
        self.cancelButton = QPushButton('Cancel')
        self.buttonLayout.addWidget(self.cancelButton, alignment=Qt.AlignCenter)
        self.mainLayout.addLayout(self.buttonLayout)

        self.setLayout(self.mainLayout)


class CategoryDialog(QDialog):
    def __init__(self,text='',parent=None):
        super(CategoryDialog, self).__init__(parent,Qt.WindowSystemMenuHint|Qt.WindowTitleHint)
        self.valueSaved = False
        self.setupUI()
        self.textBox.setPlainText(text)
        self.okButton.clicked.connect(self.okClicked)
        self.cancelButton.clicked.connect(self.close)
        self.setWindowTitle("Edit Category ")

    def setupUI(self):
        self.mainLayout = QVBoxLayout()
        self.label = QLabel('Please specify categories in following format:  \nTYPE: SAFETY, EFFICACY \nTARGET QUATER: Q1,Q2,Q3')
        self.mainLayout.addWidget(self.label)
        self.textBox = QPlainTextEdit()
        self.mainLayout.addWidget(self.textBox)
        self.iconsLayout = QHBoxLayout()
        self.okButton = QPushButton('Ok')
        self.iconsLayout.addWidget(self.okButton)
        self.cancelButton = QPushButton('Cancel')
        self.iconsLayout.addWidget(self.cancelButton)
        self.mainLayout.addLayout(self.iconsLayout)
        self.setLayout(self.mainLayout)

    def okClicked(self):
        if self.validate(self.textBox.toPlainText()):
            self.valueSaved = True
            self.close()


    def validate(self,text):
        if not text:
            msg = customQMessageBox('Empty value. Enter category in the required format.')
            msg.exec_()
            return False
        for s in text.strip('\n').__add__('\n').split('\n'):
            if s and not re.match("^[\w+\s]*[:]{1}([\w+\s]*[,]*)*$",s):
                msg = customQMessageBox('Invalid text format')
                msg.exec_()
                return False

        return  True


class RunAction(QAction):
    def __init__(self,text='',parent=None):
        super(RunAction, self).__init__(parent=parent,text=text)
        self.triggered.connect(self.run)


    def run(self):
        rows = [row.row() for row in self.parent().treeView.selectionModel().selectedRows()]
        objectIDs = self.parent().treeView.model()._data.iloc[rows]['objectID'].values.tolist()

        if not objectIDs:
            msg = customQMessageBox('Please select a row')
            msg.exec_()
            return
        progCol = 'primaryProgPath' if self.text() == 'Primary' else 'qcProgPath'
        logCol = 'primaryLogPath' if self.text() == 'Primary' else 'qcLogPath'
        errors = []
        for row in rows:
            if self.parent().treeView.model()._data.iloc[row][progCol]:
                if not os.path.exists(self.parent().treeView.model()._data.iloc[row][logCol]):
                    with open(self.parent().treeView.model()._data.iloc[row][logCol],'w') as f:
                        f.write("test log")

                print("Created log file at ",self.parent().treeView.model()._data.iloc[row][logCol])

            else:
                errors.append(row)
        if errors:
            print('No primary program found')
            # msg = customQMessageBox('No primary program found for following items \n'+"\n".join(self.parent().treeView.model()._data.iloc[errors]['object_name'].values.tolist()))
            # msg.exec_()


class ScanAction(QAction):
    def __init__(self,text='',parent=None):
        super(ScanAction, self).__init__(parent=parent,text=text)
        self.triggered.connect(self.run)

    def run(self):
        rows = [row.row() for row in self.parent().treeView.selectionModel().selectedRows()]
        objectIDs = self.parent().treeView.model()._data.iloc[rows]['objectID'].values.tolist()

        if not objectIDs:
            msg = customQMessageBox('Please select a row')
            msg.exec_()
            return
        logCol = 'primaryLogPath' if self.text() == 'Primary' else 'qcLogPath'
        scanCol = 'Scan Primary' if self.text() ==  'Primary' else 'Scan QC'
        dblogCol = 'primary_log' if self.text() == 'Primary' else 'qc_log'
        cursor = s.db.cursor()
        errors = []
        for row in rows:
            errorLines = []
            objectID = self.parent().treeView.model()._data.iloc[row]['objectID']
            if os.path.exists(self.parent().treeView.model()._data.iloc[row][logCol]):
                print("Scanning log file at ",self.parent().treeView.model()._data.iloc[row][logCol])
                from datetime import datetime
                datetime.now().strftime('%d%b%Y %H:%M:%S')
                cursor.execute(f"update  util_obj set {self.text().lower()}_log_dt = '{datetime.now().strftime('%d%b%Y %H:%M:%S')}' where PROJECTID = '{self.parent().projectID}' and LIBID = '{self.parent().libID}' and objectID= {self.parent().treeView.model()._data.iloc[row]['objectID']} ")
                s.db.commit()
                with open(self.parent().treeView.model()._data.iloc[row][logCol],'r') as f:

                    logScan = pd.read_csv('src/validations/logScan.csv').replace(np.nan,'')
                    startExp = logScan[logScan[scanCol]=='Y']['Starts With'].unique()[::-1]

                    errFlag = 0
                    for lines in readChunks(f):
                        if errFlag > 10 :
                            break
                        nErrors = []
                        for i,line in enumerate(lines):
                            line = line.upper()
                            if errFlag > 10 :
                                break
                            for se in startExp:
                                if line.startswith(se):
                                    error = logScan[logScan['Starts With'] == se][
                                        logScan[logScan['Starts With'] == se]['Contains text'].apply(
                                            lambda x: x in line)]
                                    if not error.empty:
                                        # print(line,error['Starts With'].values[0],error['Contains text'].values[0])

                                        errorLines.append(line)
                                        if 'error' in line[len('error')] or  'warning' in line[len('warning')] :
                                            errFlag +=1
                                            if errFlag > 10:
                                                break

                                    break

                    if not nErrors:
                        print('No errors')


                # print(errFlag, errorLines[:10])
                errorLines = json.dumps(errorLines[:10])
                cursor.execute(f"UPDATE util_obj SET {dblogCol} = '{errorLines}' where objectID='{objectID}' AND PROJECTID ='{self.parent().projectID}'  AND LIBID ='{self.parent().libID}' ")
                s.db.commit()

            else:
                errors.append(row)

        if errors :
            print('No primary Log file found for following items\n '+"\n".join(self.parent().treeView.model()._data.iloc[errors]['object_name'].values.tolist()))


            # msg = customQMessageBox('No primary Log file found for following items\n '+"\n".join(self.parent().treeView.model()._data.iloc[errors]['object_name'].values.tolist()))
            # msg.exec_()


class UpdateAction(QAction):
    def __init__(self,text='',col='',editDialog = False,custom=False,parent=None):

        super(UpdateAction, self).__init__(parent=parent,text=text)
        self.col = col
        self.editDialog = editDialog
        self.custom = custom

        self.triggered.connect(self.update)


    def update(self):

        if not self.parent().treeView.selectionModel().selectedRows():
            msg = customQMessageBox("Please select a row")
            msg.exec_()
            return

        rows = [row.row() for row in self.parent().treeView.selectionModel().selectedRows()]
        if self.parent().libType == 'report':
            objectID = [self.parent().treeView.model().itemFromIndex(row).id for row in self.parent().treeView.selectionModel().selectedRows()]

            objectID = self.parent().treeView.model()._data[self.parent().treeView.model()._data['objectID'].isin(objectID)]['objectID']

        else:
            objectID = self.parent().treeView.model().visibleData.iloc[rows]['objectID']


        if self.parent().treeView.model()._data[self.parent().treeView.model()._data['objectID'].isin(objectID)]['outDated'].any():
            outdatedMask = self.parent().treeView.model()._data[self.parent().treeView.model()._data['objectID'].isin(objectID)]
            objectID = outdatedMask[~outdatedMask['outDated']]['objectID']
            if objectID.empty:
                errorText = "Error. Cannot update outdated rows."
                msg = customQMessageBox(errorText)
                msg.exec_()
                return



        if s.isAdmin:
            owners = [s.currentUser]
            errorText = "Access Error. Only admin has access to this functionality."

        else:
            if 'primary' in self.col.lower()  and self.col != 'Primary_Owner':
                ownerMask = self.parent().treeView.model()._data[self.parent().treeView.model()._data['Primary_Owner']==s.currentUser+';']
                owners = ownerMask[ownerMask['objectID'].isin(objectID)]['Primary_Owner'].values.tolist()
                owners = [x[:-1] for x in owners]
                objectID = ownerMask[ownerMask['objectID'].isin(objectID)]['objectID']
                errorText = "Access Error. You  do not have access to this functionality."
            elif 'qc' in self.col.lower() and self.col != 'QC_Type' and self.col != 'QC_Owner':
                ownerMask = self.parent().treeView.model()._data[
                    self.parent().treeView.model()._data['QC_Owner'] == s.currentUser + ';']
                owners = ownerMask[ownerMask['objectID'].isin(objectID)][
                    'QC_Owner'].values.tolist()
                owners = [x[:-1] for x in owners]
                objectID = ownerMask[ownerMask['objectID'].isin(objectID)]['objectID']

                errorText = "Access Error. You  do not have access to this functionality."
            else:
                owners = []
                errorText = "Access Error. Only admin has access to this functionality."


        @validateUsers(owners,errorText=errorText)
        def updateDB(objectID):
            if self.parent().libType == 'report':
                indexes = [row for row in self.parent().treeView.selectionModel().selectedRows() if
                           self.parent().treeView.model().itemFromIndex(row).id in objectID.values.tolist()]
            else:
                indexes = [row for row in self.parent().treeView.selectionModel().selectedRows() if
                           row.row() in objectID.index]

            if self.editDialog:
                if len(self.parent().treeView.selectionModel().selectedRows()) == 1:
                    text = self.parent().treeView.model()._data[self.parent().treeView.model().visibleData['objectID'].isin(objectID)][self.col].values[0]
                    dlg = EditDialog(text=text)
                else:
                    dlg = EditDialog()
                dlg.exec_()
                if dlg.newValueSaved:
                    text = dlg.newValue

                    self.parent().treeView.model().updateData(indexes , self.col, text, Qt.EditRole)
            else:
                text  = self.text()
                self.parent().treeView.model().updateData(indexes, self.col, text, Qt.EditRole)

            # objectID = self.parent().treeView.model().visibleData.iloc[
            #     self.parent().treeView.selectionModel().selectedRows()[0].row()]['objectID']
            # rowData = self.parent().treeView.model()._data[self.parent().treeView.model()._data['objectID'] == objectID]

        if objectID.empty:
            errorText = "Access Error. You do not have permission to update selected rows."
            msg = customQMessageBox(errorText)
            msg.exec_()
            return

        updateDB(objectID)
        notUpdated = len(rows)-objectID.shape[0]

        if notUpdated >0 :
            msg = customQMessageBox(f"{notUpdated} row{'s' if notUpdated>1 else ''} were not updated")
            msg.exec_()

        # self.parent().checkValidations(rowData.to_dict('records')[0])


class TaskEdit(QDialog):
    def __init__(self,parent=None,isMilestone=False):
        super(TaskEdit, self).__init__(parent,Qt.WindowSystemMenuHint|Qt.WindowTitleHint)
        self.setWindowTitle('Edit Tasks')
        self.index = self.parent().treeView.currentIndex()
        self.taskItem = self.parent().treeView.model().itemFromIndex(self.index)
        self.task = self.parent().treeView.model()._data.loc[
            self.parent().treeView.model()._data['taskid'] == self.taskItem.id].iloc[0]

        self.isMilestone = isMilestone
        self.setupUI()
        self.okButton.clicked.connect(self.okClicked)
        self.cancelButton.clicked.connect(self.cancelClicked)

    def validateFields(self):
        if not re.match(r"^[\w+\s+]{1,100}$",self.nameEdit.text().strip()):
            msg = customQMessageBox("Task name should only contain alphabets \n Task name should not be blank \n Task name should not be longer than 100 characters ")
            msg.exec_()
            return  False
        if not len(self.noteTextEdit.toPlainText().strip()) < 255:
            msg = customQMessageBox(
                "Status Notes should be less than 255 characters.")
            msg.exec_()
            return False

        # startDateDelta = abs(datetime.now().date() - self.pStartdateEdit.date().toPython())
        endDateDelta = abs(datetime.now().date() - self.pEndtdateEdit.date().toPython())
        cDateDelta = abs(datetime.now().date() - self.cdateEdit.date().toPython())

        # if startDateDelta.days > 7:
        #     msg = customQMessageBox(
        #         "Start Date cannot be ......")
        #     msg.exec_()
        #     return False

        if endDateDelta.days > 7:
            msg = customQMessageBox(
                "End Date cannot be more than 7 days from "+ datetime.now().date().strftime("%b %d,%Y"))
            msg.exec_()
            return False

        if cDateDelta.days > 3:
            msg = customQMessageBox(
                "Completion Date cannot be more than 3 days from "+ datetime.now().date().strftime("%b %d,%Y"))

            msg.exec_()
            return False

        return True


    def okClicked(self):

        if not self.validateFields():
            return

        formDict = {'task_Name':self.nameEdit.text().strip(),
                    'task_Owner':'' if self.ownerDropDown.currentText().strip() == '_NA_' else self.ownerDropDown.currentText().strip(),
                    'task_status': '' if self.isMilestone else self.status.currentText(),
                    'Planned_Start_Date':self.pStartdateEdit.date().toPython().strftime('%Y-%m-%d'),  #30Oct2021 21:25:10
                    'Planned_End_Date': self.pEndtdateEdit.date().toPython().strftime('%Y-%m-%d'),
                    'Completion_Date':self.cdateEdit.date().toPython().strftime('%Y-%m-%d'), #.strftime("%b %d,%Y")
                    'Status_Notes':self.noteTextEdit.toPlainText(),
                    }
        formToTable = {'Name':formDict['task_Name'],
                       'Owner':formDict['task_Owner'],
                       'Status':formDict['task_status']+', '+self.index.model()._data[self.index.model()._data['taskid']==self.index.model().itemFromIndex(self.index).id]['completion'].astype(str).values[0]+'%' if formDict['task_status'] else ''+self.parent().treeView.model().itemFromIndex(self.index.siblingAtColumn(3)).text().split(',')[-1],
                       'Planned Start Date':self.pStartdateEdit.date().toPython().strftime("%b %d,%Y"),
                       'Planned End Date':self.pEndtdateEdit.date().toPython().strftime("%b %d,%Y"),
                       'Completion Date':self.cdateEdit.date().toPython().strftime("%b %d,%Y"),
                       'Notes':formDict['Status_Notes']}
        # index = self.parent().treeView.currentIndex().siblingAtColumn(0)
        # if  self.index.model().itemFromIndex(self.index).itemType:
        #     formDict['task_status'] = ''
        updateValues = []
        for k, v in formDict.items():
            updateValues.append( f"{k}='{v}'")
            self.task[k] = v

        for k,v in formToTable.items():
            self.task[k] = v
            self.parent().treeView.model().setData(self.index.siblingAtColumn(self.parent().treeView.model()._headerData.index(k)),
                         v,
                         Qt.EditRole)
        self.parent().treeView.model().layoutChanged.emit()

        self.parent().treeView.model()._data.loc[
            self.parent().treeView.model()._data['taskid'] == self.task['taskid']] = self.task.values

        updateValues = ",".join(updateValues)



        sql = f"UPDATE util_task set  {updateValues} where taskid='{self.task['taskid']}' and PROJECTID='{self.task['PROJECTID']}' and LIBID = '{self.task['LIBID']}' "
        s.cursor.execute(sql)
        s.db.commit()
        self.close()

    def cancelClicked(self):
        self.close()


    def setupUI(self):
        self.mainLayout = QVBoxLayout()
        self.formLayout  = QFormLayout()
        self.nameEdit = QLineEdit(self.task['Name'])
        self.formLayout.addRow('Name',self.nameEdit)
        self.mainLayout.addLayout(self.formLayout)
        self.ownerDropDown = QComboBox()
        self.ownerDropDown.addItems(['_NA_']+s.adminUsers)
        self.ownerDropDown.setCurrentText('_NA_' if self.task['task_Owner']=='' else self.task['task_Owner'])
        self.formLayout.addRow('Owner',self.ownerDropDown)
        if self.isMilestone:
            self.status = QLabel()
            self.status.setText(self.task['Status'])

        else:
            self.status = QComboBox()
            self.status.addItems([ '','Completed', 'Working', 'Waiting', 'Deferred', 'Cancelled', 'Not Applicable' ])
            self.status.setCurrentText(self.task['task_status'])
        self.formLayout.addRow('Status',self.status)
        self.pStartdateEdit = QDateEdit(calendarPopup=True)
        self.pStartdateEdit.setDisplayFormat("MMM dd,yyyy")

        self.pEndtdateEdit = QDateEdit(calendarPopup=True)
        self.pEndtdateEdit.setDisplayFormat("MMM dd,yyyy")
        self.cdateEdit = QDateEdit(calendarPopup=True)
        self.cdateEdit.setDisplayFormat("MMM dd,yyyy")

        self.pStartdateEdit.setDate(QDate.currentDate() if self.task['Planned Start Date']=='yyyy/mm/dd' else QDate.fromString(self.task['Planned Start Date'],"MMM dd,yyyy"))
        self.pEndtdateEdit.setDate(QDate.currentDate() if self.task['Planned End Date']=='yyyy/mm/dd' else  QDate.fromString(self.task['Planned End Date'],"MMM dd,yyyy"))
        self.cdateEdit.setDate(QDate.currentDate() if self.task['Completion Date']=='yyyy/mm/dd' else QDate.fromString(self.task['Completion Date'],"MMM dd,yyyy"))
        self.formLayout.addRow('Planned Start Date',self.pStartdateEdit)
        self.formLayout.addRow('Planned End Date',self.pEndtdateEdit)
        self.formLayout.addRow('Completion Date',self.cdateEdit)
        self.noteTextEdit = QPlainTextEdit()
        self.noteTextEdit.setPlainText(self.task['Notes'])
        self.formLayout.addRow('Notes',self.noteTextEdit)
        self.buttonLayout = QHBoxLayout()
        self.okButton = QPushButton('OK')
        self.cancelButton = QPushButton('Cancel')
        self.buttonLayout.addWidget(self.okButton)
        self.buttonLayout.addWidget(self.cancelButton)
        self.mainLayout.addLayout(self.buttonLayout)
        self.setLayout(self.mainLayout)


class CombinePDF(QDialog):
    def __init__(self,parent=None):
        super(CombinePDF, self).__init__(parent=parent)
        self.setWindowTitle("Combibe PDF files")
        self.setupUI()
        self.browseBtn.clicked.connect(lambda: self.pathEdit.setText(
            QFileDialog.getExistingDirectoryUrl(self, 'Select Folder to Store PDF', QUrl('')).path().strip('/')))
        self.obrowseBtn.clicked.connect(lambda: self.opathEdit.setText(
            QFileDialog.getSaveFileUrl(self, 'Select Folder to Store PDF', QUrl(''),"PDF files (*.pdf)")[0].path().strip('/'),))

        self.okBtn.clicked.connect(self.okClicked)
        self.cancelBtn.clicked.connect(self.cancelClicked)

    def okClicked(self):
        path = self.pathEdit.text().strip()
        pdfs = [str(x) for x in Path(path).glob('*.pdf')]
        mergedpath = os.path.normpath(os.path.join(os.path.dirname(self.opathEdit.text().strip()),'merged.pdf'))
        if os.path.exists(mergedpath):
            msg = customQMessageBox("File already exists.")
            msg.exec_()
            return

        names =   [".".join(os.path.basename(x).split('.')[:-1]) for x in pdfs]
        rowData = self.parent().treeView.model()._data[self.parent().treeView.model()._data['object_name'].isin(names)]
        rowData = rowData.sort_values("items_ORDER")
        colmap = {'Title':'Title','Itemid':'ITEMID','Description':'Description','Itemid: Title':'custom'}
        bookmarkCol= colmap[self.bookmarksDropDown.currentText()]

        bookmarkHeaders = rowData[bookmarkCol].values.tolist()
        try:
            merge(pdfs, mergedpath)
        except FileExistsError:
            msg = customQMessageBox("File already exists.")
            msg.exec_()
            return

        bookmarkDict = createBookmarkDict(mergedpath, bookmarkHeaders)
        if self.tableCheck.isChecked():
            tocPath = os.path.normpath(os.path.join(os.path.dirname(self.opathEdit.text().strip()),'toc.pdf'))
            tocCol = colmap[self.tocDropDown.currentText()]
            if tocCol == 'custom':
                tocHeaders = rowData['ITEMID'] + ": " + rowData['Title']
                tocHeaders = tocHeaders.values.tolist()
            else:
                tocHeaders = rowData[tocCol].values.tolist()

            tocDict = createBookmarkDict(mergedpath, tocHeaders)
            tocfile = genTOC(tocPath, tocDict)
            paths = [tocPath, mergedpath]
        else:
            paths = [mergedpath]
        outPath = os.path.normpath(self.opathEdit.text().strip())
        merge(paths, outPath, bookmarkDict)
        os.remove(mergedpath)
        os.remove(tocPath)

        self.close()

    def cancelClicked(self):
        self.close()

    def setupUI(self):
        self.mainlayout = QVBoxLayout()

        self.inputPathLabel = QLabel("Input Folder:")
        self.mainlayout.addWidget(self.inputPathLabel)
        pathSelect = QHBoxLayout()
        self.pathEdit = QLineEdit(parent=self)
        pathSelect.addWidget(self.pathEdit)
        self.browseBtn = QPushButton(icon=QIcon("icons/folder.png"))
        pathSelect.addWidget(self.browseBtn)
        self.mainlayout.addLayout(pathSelect)
        self.form = QFormLayout()
        self.tableCheck = QCheckBox()
        self.form.addRow("Create Table of Contents",self.tableCheck)
        self.bookmarksDropDown = QComboBox()
        self.bookmarksDropDown.addItems(['Title','Itemid', 'Description'])
        self.form.addRow("Bookmarks",self.bookmarksDropDown)
        self.tocDropDown = QComboBox()
        self.tocDropDown.addItems(['Itemid: Title','Title'])
        self.form.addRow("Table of Contents", self.tocDropDown)
        self.mainlayout.addLayout(self.form)
        self.outputPathLabel = QLabel("Output Folder:")
        self.mainlayout.addWidget(self.outputPathLabel)
        opathSelect = QHBoxLayout()
        self.opathEdit = QLineEdit(parent=self)
        opathSelect.addWidget(self.opathEdit)
        self.obrowseBtn = QPushButton(icon=QIcon("icons/folder.png"))
        opathSelect.addWidget(self.obrowseBtn)
        self.mainlayout.addLayout(opathSelect)
        buttonLayout = QHBoxLayout()
        self.okBtn = QPushButton('OK')
        self.cancelBtn = QPushButton('Cancel')
        buttonLayout.addWidget(self.okBtn)
        buttonLayout.addWidget(self.cancelBtn)
        self.mainlayout.addLayout(buttonLayout)
        self.setLayout(self.mainlayout)

class Emitter(QThread):
    fileStarted = Signal(str)
    fileCompleted = Signal(str)
    processSignal = Signal(int)

    def __init__(self):
        super(Emitter, self).__init__()
        # self.rtfs = rtfs

    # def run(self):
    #     convertRTFToPDF(self.rtfs, self)


class RTF2PDF(QDialog):
    def __init__(self,parent=None):
        super(RTF2PDF, self).__init__(parent=parent)
        self.setWindowTitle("RTF to PDF convert")
        self.setupUI()
        self.browseBtn.clicked.connect(lambda: self.pathEdit.setText(
            QFileDialog.getExistingDirectoryUrl(self, 'Select Folder to Store PDF', QUrl('')).path().strip('/')))

        self.okBtn.clicked.connect(self.okClicked)
        self.cancelBtn.clicked.connect(self.cancelClicked)
        # self.pathEdit.setText("C:\\Users\\dipeshs\\PycharmProjects\\excelUtitlity\\pdftest")


    def okClicked(self):
        self.progressBar.show()
        path = os.path.normpath(self.pathEdit.text().strip())

        objectIDs = [index.model().itemFromIndex(index).id for index in
                     self.parent().treeView.selectionModel().selectedRows()]
        rowData = self.parent().treeView.model()._data[self.parent().treeView.model()._data['objectID'].isin(objectIDs)]
        total = rowData.shape[0]
        rtfs = rowData[rowData['extension'] == '.rtf']['outputPath'].values.tolist()
        # rtfs = {k:path for k in rtfs}
        oqueue = []
        self.emitter = Emitter()
        # self.emitter.start()

        self.emitter.processSignal.connect(self.updateProgress)
        self.emitter.fileStarted.connect(self.fileComplete)
        self.emitter.fileCompleted.connect(self.fileComplete)
        sucessful = convertRTFToPDF(rtfs, path, self.emitter)
        msg = customQMessageBox(f"Successfully converted {sucessful} files out of {total}")
        msg.exec_()
        self.close()




        import subprocess
        subprocess.Popen(f"explorer {os.path.normpath(path)}")

    @Slot(str)
    def fileComplete(self,file):
        print(file)

    @Slot(int)
    def updateProgress(self,val):
        if self.progressBar.isHidden():
            self.progressBar.show()

        self.progressBar.setValue(val)
        # if val == 100:
        #     time.sleep(1)
        #     self.close()

    def cancelClicked(self):
        self.close()

    def setupUI(self):
        self.mainlayout = QVBoxLayout()
        self.progressBar = QProgressBar(parent=self)
        self.mainlayout.addWidget(self.progressBar)
        pathSelect = QHBoxLayout()
        self.pathEdit = QLineEdit(parent=self)
        pathSelect.addWidget(self.pathEdit)
        self.browseBtn = QPushButton(icon=QIcon("icons/folder.png"))
        pathSelect.addWidget(self.browseBtn)
        self.mainlayout.addLayout(pathSelect)
        buttonLayout = QHBoxLayout()
        self.okBtn = QPushButton('OK')
        self.cancelBtn = QPushButton('Cancel')
        buttonLayout.addWidget(self.okBtn)
        buttonLayout.addWidget(self.cancelBtn)
        self.mainlayout.addLayout(buttonLayout)
        self.setLayout(self.mainlayout)

        self.progressBar.hide()


class FilterAction(QAction):
    def __init__(self,text='',col='',editDialog = False,custom=False,parent=None):
        super(FilterAction, self).__init__(parent=parent,text=text)
        self.col = col
        self.custom = custom
        self.editDialog = editDialog
        self.triggered.connect(lambda : self.addToFilter(text,col))
        self.setCheckable(True)


    def addToFilter(self,text,col):
        idColMap = {'Items':'objectID','Issues':'issueid'}
        idCol = idColMap[self.parent().state]
        if self.editDialog:
            if len(self.parent().treeView.selectionModel().selectedRows()) == 1:
                row = self.parent().treeView.currentIndex().row()
                objectID = self.parent().treeView.model().visibleData.iloc[row][idCol]
                text = self.parent().treeView.model()._data[
                    self.parent().treeView.model().visibleData[idCol] == objectID][self.col].values[0]
                dlg = EditDialog(text=text)
            else:
                dlg = EditDialog()
            dlg.exec_()

            if dlg.newValueSaved:
                text = dlg.newValue
                self.parent().filters[col] = text
                self.setChecked(True)
            else:
                self.setChecked(False)
                return
        else:

            if self.custom:
                if self.parent().state == "Items":
                    category,col = col.split('|')
                    text = f'"{category}": "{text}"'


        self.parent().filters[col] = text


        conditions = [(self.parent().treeView.model()._data[col].astype(str).str.lower().str.contains(str(self.parent().filters[col]).lower())) for col in self.parent().filters.keys()]
        conditions = reduce(lambda x,y:x & y,conditions)
        if self.parent().state == 'Items':
            if self.parent().libType == 'report':
                data = self.parent().treeView.model()._data
                filteredData = data[data['TEMPLATE'].isin(['Listing', 'Summary', 'Figure'])][conditions].sort_values(
                    'items_ORDER')
                folders = data[data['ITEMID'].isin(set(filteredData['PARENTID'].values.tolist()))]
                while folders.shape[0] != filteredData[filteredData['TEMPLATE'] == 'Folder'].shape[0]:
                    filteredData = filteredData.append(folders)
                    folders = data[data['ITEMID'].isin(set(filteredData['PARENTID'].values.tolist()))]
                    filteredData = filteredData.drop_duplicates()
                treeModel = createTreeModel(filteredData, self.parent().treeView.model(), 'PARENTID', self.parent().libName,
                                            'Object_Desc')
                self.parent().treeView.setModel(treeModel)
            else:
                data = self.parent().treeView.model()._data[conditions].sort_values(idCol)
                self.parent().treeView.model().updateVisibleData(data)
        elif self.parent().state == "Issues":
            if self.col == "hasComment" and self.custom: # special case
                self.parent().resetData()

                # self.parent().filters['assigned_to'] = s.currentUser
                #
                # self.parent().filters['open_By'] = s.currentUser
                # self.parent().filters['issue_Status'] = 'Open'
                # self.parent().filters['hasComment'] = True
                # conditions = [(self.parent().treeView.model()._data[col].astype(str).str.lower().str.contains(
                #     str(self.parent().filters[col]).lower())) for col in self.parent().filters.keys()]
                # conditions = reduce(lambda x, y: x | y, conditions)
                source = self.parent().treeView.model()._data
                data = source[source['issue_Status'] == 'Open'].query(
                    f"assigned_to=='{s.currentUser}'|open_By=='{s.currentUser}'|hasComment==True")
                self.parent().treeView.model().updateFilteredData(data)
                return
            conditions = [(self.parent().treeView.model().filteredData[col].astype(str).str.lower().str.contains(
                str(self.parent().filters[col]).lower())) for col in self.parent().filters.keys()]
            conditions = reduce(lambda x, y: x & y, conditions)
            data = self.parent().treeView.model().filteredData[conditions]
            self.parent().treeView.model().updateFilteredData(data)
            return
        data = self.parent().treeView.model()._data[conditions].sort_values(
                'open_date')
        self.parent().treeView.model().updateFilteredData(data)


        # reduce(lambda x,y:x&y,conditions)





def readChunks(fileObj, chunkSize=1024):
    lines = []
    while True:
        data = fileObj.readline()
        if not data:
            break
        lines.append(data)
        if len(lines) >= chunkSize:
            yield lines
            lines=[]


    yield lines