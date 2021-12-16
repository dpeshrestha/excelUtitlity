from PySide2.QtWidgets import QTreeView


class MyTreeView(QTreeView):
    def __init__(self,parent=None):
        super(MyTreeView, self).__init__(parent)
        self.setAlternatingRowColors(True)
        self.header().setStyleSheet("""
            QHeaderView {background-color:white;}
            QHeaderView:section{padding-left:2px;border-right:0.8px solid gray;border-top:0;border-bottom:0.8px solid gray;}
        """)


    def setModel(self, model):
        super(MyTreeView, self).setModel(model)

        # self.expandAll()

    def updateData(self,id,rowData):
        pass

    def getItemWithid(self,id,root):
        pass

