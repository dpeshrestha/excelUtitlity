from collections import deque

import numpy as np
import pandas as pd
from PySide2.QtCore import QAbstractListModel, Qt
from PySide2.QtGui import QFont

import src.settings as s
# from src.utils.utils import strike
from src.utils.utils import getAllChildren


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

        if role == Qt.ToolTipRole:
            if self.checkable:
                return str(self._data[self.dataCol][index.row()])
            else:
                return str(self._data[index.row()])


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

    def rowCount(self,parent=None):
        return len(self._data)

    def updateDB(self, index, value):
        self._data.loc[index.row(),self.checkVal] = value
        self.dataChanged.emit(index, index)
        #calculate row percentage
        currentPercentage = self.parent().treeView.model().itemFromIndex(self.parent().treeView.currentIndex().siblingAtColumn(3)).text()
        currentStatus = currentPercentage.split(',')[0]
        currentPercentage = currentPercentage.split(',')[-1]
        yes = self._data.loc[self._data['titem_Status']=='X','taskid'].count()
        total = self._data.loc[~(self._data['titem_Status']=='Z'),'taskid'].count()
        if int(total) != 0 :
            newPercentage = yes / total * 100
            newPercentage = f'{newPercentage:.1f}%'
        else:
            newPercentage = '0.0%'
        if currentStatus!= currentPercentage:
            newPercentage = ",".join([currentStatus,newPercentage])
        item = self.parent().treeView.model().itemFromIndex(self.parent().treeView.currentIndex().siblingAtColumn(3))
        item.setText(newPercentage)
        item.setToolTip(newPercentage)
        #update parent percentages
        parents = item.getAllParents([])
        for parent in parents:
            children = parent.getImmediateChildrenAtColumn(3)
            siblingStatuses = [child.text() for child in children if not ('Cancelled' in child.text() or 'Not Applicable' in child.text())]
            # siblingPercentages = [s.split(',')[0]for s in siblingStatuses]
            siblingPercentages = [float(sib.split(',')[-1][:-1]) for sib in siblingStatuses]
            newParentPercentage = np.mean(siblingPercentages)
            parentStatus = parent.model().itemFromIndex(parent.index().siblingAtColumn(3)).text()
            oldParentPercentage = parentStatus.split(',')[-1]
            oldParentStatus = parentStatus.split(',')[0]
            if int(newParentPercentage) == 0:
                newParentPercentage = '0.0%'
            else:
                newParentPercentage = f'{newParentPercentage:.1f}%'
            if oldParentStatus != oldParentPercentage:
                newParentPercentage = ",".join([oldParentStatus, newParentPercentage])
            parent.model().itemFromIndex(parent.index().siblingAtColumn(3)).setText(newParentPercentage)
            parent.model().itemFromIndex(parent.index().siblingAtColumn(3)).setToolTip(newParentPercentage)

        rowData = self._data.loc[index.row()]
        s.db.cursor().execute(f"UPDATE util_task_items set titem_Status='{value}' where taskid='{rowData['taskid']}' and tOrder='{rowData['torder']}' and PROJECTID='{rowData['PROJECTID']}' and LIBID='{rowData['LIBID']}'")
        s.db.commit()

    def calculatePercentage(self):
        #only for task
        data = self.parent().treeView.model()._data
        if not data.empty:
            ignoreMask = ~data['task_status'].isin(['Cancelled', 'Not Applicable'])
            for i, row in data[(data['task_type'] == 'Task') & (ignoreMask)].iterrows():
                comp = pd.read_sql_query(
                    f"SELECT x.completed/ y.total as comp from(select count(taskid) as completed from util_task_items where taskid = {row['taskid']} and titem_Status= 'X') x JOIN (SELECT COUNT(taskid) AS total FROM util_task_items WHERE taskid = {row['taskid']} AND (titem_Status <> 'Z' OR titem_Status IS null)) y on 1=1",
                    s.db)
                comp = comp['comp'][0] if comp['comp'][0] is not None else 0.

                data.loc[data['taskid'] == row['taskid'], 'completion'] = comp
                # taskItems = taskItems[~(taskItems['titem_Status'] == 'Z')]
                # ncompleted = taskItems[taskItems['titem_Status'] == 'X'].shape[0]
                # N = taskItems.shape[0]
                # completionPer = int(ncompleted / N)*100
                # print(f"{parent} : {completionPer} % completed")
            # completionPer = {}
            statusValues = ['', 'Completed', 'Working', 'Waiting', 'Deferred', 'Cancelled', 'Not Applicable']
            maxlen = len(sorted(statusValues, key=len)[-1])

            for parent in data[data['Parentid'] == '']['taskid'].values.tolist():
                tree = getAllChildren(data[ignoreMask], str(parent), 'taskid', 'Parentid')
                tree = deque((k, v) for k, v in tree.items())
                while tree:
                    k, v = tree.popleft()
                    # print(k,v)
                    if len(v) > 0:
                        if data[data['taskid'].isin(v)]['completion'].isna().any():
                            tree.append((k, v))
                        else:
                            data.loc[data['taskid'] == int(k), 'completion'] = data[data['taskid'].isin(v)][
                                'completion'].mean()

            data['completion'] = data['completion'].fillna(0)
            # data['Status']  = data['task_status'].replace(np.nan,'').str.strip().apply(lambda x :x+" "*(14 - len(x) +1))+ data['completion'].apply(lambda x:f"{x*100:.1f}% completed") # 14 - len of not applicable
            data['Status'] = data['task_status'].replace(np.nan, '').str.strip().apply(
                lambda x: x + ', ' if len(x) > 0 else x) + data['completion'].apply(
                lambda x: f"{x * 100:.1f}%")  # 14 - len of not applicable
        data = data.fillna('')
        self.parent().treeView.model().setSource(data)