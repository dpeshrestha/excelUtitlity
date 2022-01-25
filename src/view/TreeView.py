from PySide2.QtCore import QSize
from PySide2.QtGui import QTextDocument, Qt
from PySide2.QtWidgets import QTreeView, QStyledItemDelegate, QLabel, QLineEdit, QTableView


class ItemWordWrap(QStyledItemDelegate):
    def __init__(self,parent):
        super(ItemWordWrap, self).__init__(parent=parent)
    
    def createEditor(self, parent, option, index):
        editor = QLabel(parent=parent)
        return editor

    def setEditorData(self, editor, index):
        text = index.data(Qt.DisplayRole+99)
        editor.setText(text)
        editor.setWordWrap(True)


class MyTreeView(QTreeView):
    def __init__(self,parent=None):
        super(MyTreeView, self).__init__(parent)
        self.setAlternatingRowColors(True)
        # self.header().setStyleSheet("""

        self.header().setStyleSheet("""
            QHeaderView {background-color:white;}
            QHeaderView:section{padding-left:2px;border-right:0.8px solid gray;border-top:0;border-bottom:0.8px solid gray;}
        """)

    def setModel(self, model):
        super(MyTreeView, self).setModel(model)
        # self.setItemDelegate(ItemWordWrap(self))
        # self.openPersistentEditor(self.model().index(4,3))
        self.setWordWrap(True)

        # self.header().setDefaultAlignment(Qt.AlignCenter|Qt.Alignment(Qt.TextWrapAnywhere))

    def updateData(self,id,rowData):
        pass

    def getItemWithid(self,id,root):
        pass