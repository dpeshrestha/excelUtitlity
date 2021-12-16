from PySide2.QtCore import QAbstractListModel, Qt
from PySide2.QtGui import QFont

import src.settings as s
# from src.utils.utils import strike


class ListModel(QAbstractListModel):
    def __init__(self,data=[],checkable=False,dataCol='',checkVal = '',parent=None):
        super(ListModel, self).__init__(parent)
        self._data = data
        self.checkable = checkable
        if checkable:
            self.dataCol = dataCol
            self.checkVal = checkVal

    def flags(self, index):
        defaultFalgs = super(ListModel, self).flags(index)

        if index.isValid():

            f =  Qt.ItemIsDragEnabled | Qt.ItemIsDropEnabled | defaultFalgs
            if self.checkable:
                f = Qt.ItemIsUserCheckable | Qt.ItemIsTristate | Qt.ItemIsEnabled | defaultFalgs
            return  f
        else:
            return  Qt.ItemIsDropEnabled | defaultFalgs

    def data(self, index, role:Qt.DisplayRole):
        if role == Qt.DisplayRole:
            if self.checkable:


                return str(self._data[self.dataCol][index.row()])
            else:
                return str(self._data[index.row()])
        if role == Qt.CheckStateRole:
            if self.checkable:
                if self._data[self.checkVal][index.row()]== 'X':
                    return Qt.Checked
                elif self._data[self.checkVal][index.row()]== 'Z':
                    return Qt.PartiallyChecked
                else:
                    return  Qt.Unchecked

        if role == Qt.FontRole:
            if self.checkable:
                if self._data[self.checkVal][index.row()] == 'Z':
                    font = QFont()
                    font.setStrikeOut(True)
                    return font

            if self.parent() and  self.parent().state == 'Tasks' and self.parent().treeView.model().itemFromIndex(
                self.parent().treeView.currentIndex().parent().siblingAtColumn(3)).text() in ['Cancelled',
                                                                                              'Not Applicable']:
                font = QFont()
                font.setStrikeOut(True)

                return font

    def updateDB(self, index, value):
        self._data.loc[index.row(),self.checkVal] = value
        self.dataChanged.emit(index, index)
        rowData = self._data.loc[index.row()]
        s.cursor.execute(f"UPDATE util_task_items set titem_Status='{value}' where taskid='{rowData['taskid']}' and tOrder='{rowData['torder']}' and PROJECTID='{rowData['PROJECTID']}' and LIBID='{rowData['LIBID']}'")
        s.db.commit()
        #
        # done = self._data[self._data['titem_Status'] == 'X'].shape[0]
        # n = self._data[self._data['titem_Status']!='Z'].shape[0]
        # selectedidx = self.parent().treeView.model().itemFromIndex(self.parent().treeView.selectedIndexes()[0]).id
        # try:
        #     completion = str(round(done/n,2))
        # except ZeroDivisionError:
        #     completion = str(0)
        #
        # self.parent().treeView.model()._data.loc[
        #     self.parent().treeView.model()._data['taskid'] == selectedidx, 'task_Status'] = completion
        # self.parent().treeView.model()._data.loc[
        #     self.parent().treeView.model()._data['taskid'] == selectedidx, 'Status'] = completion
        # self.parent().treeView.model().itemFromIndex(self.parent().treeView.selectedIndexes()[2]).setData(completion,
        #                                                                                                   Qt.EditRole)
        # s.cursor.execute(
        #     f"UPDATE util_task set task_Status='{completion}' where taskid='{rowData['taskid']}' and PROJECTID='{rowData['PROJECTID']}' and LIBID='{rowData['LIBID']}'")
        # s.db.commit()



    def rowCount(self, parent):
        return  len(self._data)