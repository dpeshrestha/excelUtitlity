from _collections import deque

import contextlib
import os
import tempfile

from PySide2 import QtGui
from PySide2.QtGui import QStandardItem, QPixmap, QColor, QIcon
from PySide2.QtWidgets import QTreeView
# tasks['Task'].apply(lambda x:d[x] if d.get(x) else '')
def orderItems(d,x):
    if d.get(x):
        d[x].append(d[x][-1]+1)
    else:
        d[x] = [1]

def findChildren(df,parentCol,childCol,idCol,rootItem):

    children = deque(df[df[parentCol] == rootItem][childCol].values.tolist())
    ids = df[df[childCol] == rootItem][idCol].values.tolist()+ df[df[parentCol] == rootItem][idCol].values.tolist()

    while children:
        child = children.popleft()
        children += df[df[parentCol] == child][childCol].values.tolist()
        ids +=  df[df[parentCol] == child][idCol].values.tolist()

    return ids


class treeItem(QStandardItem):

    def __init__(self, text, recordid='', recordType='', iconPath='', toolTipText=''):
        super().__init__()
        self.setEditable(False)
        self.setText(str(text))
        self.recordid = recordid
        self.recordType = recordType
        self.hover = False
        self.iconPath = iconPath
        if toolTipText: self.setToolTip(str(toolTipText))
        if len(iconPath) == 7 and iconPath[0] == "#":
            pixMap = QPixmap(16, 16)
            try:
                pixMap.fill(QColor(iconPath))
                self.setIcon(QIcon(pixMap))
            except Exception as e:
                print(e)
        elif iconPath:
            self.setIcon(QtGui.QIcon(iconPath))


def createTreeModel(treeModel,idCol,parentID,rootName,itemID,typeCol='',toolTipCol=''):
    """
    :param treeModel: TreeModel whose _data should be shown
    :param idCol: column name which has uniqueID for each item
    :param parentID: column name which contains parentID of each item
    :param rootName: name of rootItem
    :param itemID:
    :param toolTipCol:
    :return:
    """
    dataDF = treeModel._data
    treeModel.setRowCount(0)
    root = treeModel.invisibleRootItem()
    parentFound = {}
    treeRows = deque(dataDF.to_dict('records'))
    # treeRows = []
    c = 0
    while treeRows:
        itemRow = treeRows.popleft()
        if str(itemRow[parentID]) == rootName:
            parent = root
        else:
            rowParent = itemRow[parentID]
            if rowParent not in parentFound:
                treeRows.append(itemRow)

                c+=1
                print("Unwanted step",c)
                continue
            else:
                parent = parentFound[rowParent]
        # parent.appendRow(TreeItem(itemRow[itemName],itemRow[idCol],itemRow[itemName]))
        itemType = itemRow[typeCol] if typeCol in dataDF.columns else ''
        parent.appendRow([TreeItem(itemRow[col],itemRow[idCol],itemType,itemRow[col]) for col in treeModel._headerData])
        # parent.appendRow(QStandardItem(itemRow['object_name']))
        parentFound[str(itemRow[itemID])] = parent.child(parent.rowCount() - 1)

    return  treeModel


class TreeItem(QStandardItem):
    def __init__(self,text,id,type='',tooltip=''):
        super(TreeItem, self).__init__()
        self.setText(str(text))
        self.itemType = type
        if tooltip:
            self.setToolTip(str(tooltip))

        self.id = id

def getAllChildren(data,parentid,idcol,parentcol):
    roots = [str(parentid)]
    tree = {str(parentid): []}
    while roots:
        root = str(roots.pop())
        tree[root] = data[data[parentcol] == root][idcol].values.tolist()
        roots += data[data[parentcol] == root][idcol].values.tolist()

    return  tree

@contextlib.contextmanager
def temp_file(mode,contents,suffix=None):
    try:
        f  =  tempfile.NamedTemporaryFile(mode=mode,suffix=suffix,delete=False)
        f.write(contents)
        tmp_name = f.name
        f.close()
        yield  tmp_name
    finally:
        os.remove(tmp_name)


