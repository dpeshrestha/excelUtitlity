from _collections import deque
import contextlib
import os
import tempfile

import pandas as pd
from PySide2 import QtGui
from PySide2.QtGui import QStandardItem, QPixmap, QColor, QIcon
from pyreadstat import pyreadstat
import src.settings as s
import mysql.connector

from src.view.customDialogs import customQMessageBox
def restartConnection():
    s.db = mysql.connector.connect(host="192.168.100.15", user="read-utility_user", password="NimUtils1",
                                   database="remoteproddb", pool_name='NimUtil',
                                   pool_size=8)

def addSeperator(text,sep=';'):
    if text:
        text+=sep
    return text

def replaceExtension(text,origExt,newExt):
    if origExt == text.split('.')[-1]:
        return ".".join(text.split('.')[:-1]+[newExt])
    else:
        return text


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
    dataDF = treeModel.visibleData
    # dataDF.to_csv('tree.csv',index=False)
    treeModel.setRowCount(0)
    root = treeModel.invisibleRootItem()
    parentFound = {}
    treeRows = deque(dataDF.to_dict('records'))
    # treeRows = []

    c = 0
    while treeRows:
        oldtreeRows = treeRows
        itemRow = treeRows.popleft()
        if str(itemRow[parentID]) == rootName:
            parent = root
        else:
            rowParent = itemRow[parentID]
            if rowParent not in parentFound:
                treeRows.append(itemRow)
                if len(oldtreeRows) == len(treeRows):
                    c+=1
                if c > len(dataDF)*3:
                    print("into inf loop/")
                    msg = customQMessageBox("Corrupt Data:\n Cannot show full data.")
                    msg.exec_()
                    treeModel.empty = True
                    return treeModel
                else:
                    continue
            else:
                parent = parentFound[rowParent]
        # parent.appendRow(TreeItem(itemRow[itemName],itemRow[idCol],itemRow[itemName]))
        itemType = itemRow[typeCol] if typeCol in dataDF.columns else ''
        parent.appendRow([TreeItem(itemRow[col],itemRow[idCol],itemType,itemRow[col]) for col in treeModel._header])
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
        else:
            self.setToolTip(text)
        self.id = id

    def getImmediateChildrenAtColumn(self,col):
        return [self.child(row,col) for row in range(self.rowCount())]

    def getAllParents(self,parents=[]):
        if self.parent() is not None:
            parents.append(self.parent())
            return self.parent().getAllParents(parents)
        else:
            return parents

def getAllChildren(data,parentid,idcol,parentcol):
    roots = [str(parentid)]
    tree = {str(parentid): []}
    while roots:
        root = str(roots.pop())
        tree[root] = data[data[parentcol] == root][idcol].values.tolist()
        roots += data[data[parentcol] == root][idcol].values.tolist()

    return  tree


def read_sas(path):
    try:
        df,meta = pyreadstat.read_sas7bdat(path)
        return df
    except pyreadstat._readstat_parser.PyreadstatError as e:
        return None



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


if __name__ == "__main__":
    print(replaceExtension("ADAE.sas.sas",'sas','log'))
    print(replaceExtension("ADAE.sas.se",'sas','sas7bdat'))
    print(replaceExtension("ADAEsas",'sas','log'))
    print(replaceExtension("AD.A.E.sas",'sas','log'))
