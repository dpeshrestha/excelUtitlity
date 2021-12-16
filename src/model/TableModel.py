import json

import numpy as np
from PySide2.QtCore import QAbstractTableModel, QModelIndex
from PySide2.QtGui import Qt, QColor
import src.settings as s
from src.utils.validation import validateAdmin
from src.view.customDialogs import customQMessageBox


class BaseTableModel(QAbstractTableModel):
    """Base TableModel class meant to be inherited for customization"""

    def __init__(self, data,cols, parent=None):
        super(BaseTableModel, self).__init__(parent)
        self._data = data.reset_index(drop=True)
        self._header = cols
        self.visibleData = self._data[cols]
        self.filteredData = self._data.copy(True)



    @property
    def header(self):
        return self._header

    @header.setter
    def header(self, headerCols):
        if len(headerCols) > 0:
            self._header = headerCols

    def rowCount(self, parent:QModelIndex=...) -> int:
        return self.visibleData.shape[0]

    def columnCount(self, parent:QModelIndex=...) -> int:
        return self.visibleData.shape[1]

    def data(self, index:QModelIndex, role:int=...):
        if index.isValid():
            if role == Qt.DisplayRole:
                return str(self.visibleData.iat[index.row(), index.column()])
            elif  role == Qt.ToolTipRole:
                if index.column() == 2:
                    return self.visibleData.iat[index.row(), 1]
                elif index.column() == 4:
                    return self.visibleData.iat[index.row(), 4]
                elif index.column() == 6:
                    return self.visibleData.iat[index.row(), 6]
            elif role == Qt.TextColorRole:
                if self.parent().state == 'Items':
                    if self._data[self._data['objectID']==self.visibleData.iat[index.row(),0]]['outDated'].any():
                        return  QColor('grey')
            # else:
                # super(BaseTableModel, self).data(index,role)
            
        else:
            return False



    def headerData(self, section:int, orientation:Qt.Orientation, role:int=...):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self._header[section]

    def flags(self, index:QModelIndex) -> Qt.ItemFlags:
        return Qt.ItemIsEnabled | Qt.ItemIsSelectable

    def updateData(self, indexes,col, value, role=Qt.EditRole):
        strVal = ''
        if 'custom_Columns' not in col:
            value = '' if value == '_NA_' or value == 'None' else  value
        if role == Qt.EditRole:
            objectIDs = [int(self.visibleData.iloc[index.row()]['objectID'])  for index in indexes]

            for objectID in objectIDs:
                if 'primary' in col.lower():
                    # colDict = {'Primary_Owner':0,'Primary_Status':1,'Primary_Progname':2,'Primary_Notes':3}
                    self._data.loc[self._data['objectID'] == objectID, col] = value                   # if col == 'Primary_Progname':
                    self._data.loc[(self._data['Primary_Owner'].str.len() > 1) &(~self._data['Primary_Owner'].str.endswith(';')), 'Primary_Owner'] += ";"
                    self._data.loc[(self._data['Primary_Status'].str.len() > 1)  &(~self._data['Primary_Status'].str.endswith('\n')), 'Primary_Status'] += "\n"
                    self._data.loc[(self._data['Primary_Progname'].str.len() > 1)  &(~self._data['Primary_Progname'].str.endswith(';')), 'Primary_Progname'] += ";"
                    self._data['Primary_Notes'] = np.where(self._data['Primary_Notes'].str.decode('ascii').isna(),self._data['Primary_Notes'],self._data['Primary_Notes'].str.decode('ascii'))
                    self._data['Primary Info'] = self._data['Primary_Owner'] + self._data['Primary_Status'] + self._data[
                        'Primary_Progname'] + self._data['Primary_Notes']
                    if 'progname' in col.lower():
                        self._data['primaryProgPath'] = np.where(self._data["Primary_Progname"],
                                                                f"//sas-vm/{self.parent().projectID}/Data/{self.parent().libName.split(' ')[0]}/SAS Programs/" + \
                                                                self._data["Primary_Progname"].str[:-1]+ '.sas',
                                                                self._data["Primary_Progname"])
                        self._data['primaryLogPath'] = np.where(self._data["Primary_Progname"],
                                                               f"//sas-vm/{self.parent().projectID}/Data/{self.parent().libName.split(' ')[0]}/SAS Programs/" + \
                                                               self._data["Primary_Progname"].str[:-1]+ '.log',
                                                               self._data["Primary_Progname"])
                        self._data['outputPath'] = np.where(self._data["Primary_Progname"],
                                                           f"//sas-vm/{self.parent().projectID}/Data/{self.parent().libName.split(' ')[0]}/SAS Programs/" + \
                                                           self._data["object_name"] + '.sas7bdat',
                                                           self._data["Primary_Progname"])
                    self._data['Primary Info'] = self._data['Primary Info'].str.strip(';').str.strip('\n')
                    self.visibleData.loc[:,'Primary Info'] =  self._data['Primary Info']
            #remove SAS for DATA

                    self.layoutChanged.emit()

                elif 'qc' in col.lower():
                    self._data.loc[self._data['objectID'] == objectID, col] = value
                    if col == 'QC_Type':
                        self.visibleData['QC Type'] = self._data['QC_Type']
                    else:
                    # colDict = {'QC_Owner': 0, 'QC_Status': 1, 'QC_Progname': 2,}

                        self._data.loc[(self._data['QC_Owner'].str.len() > 1) &(~self._data['QC_Owner'].str.endswith(';')), 'QC_Owner'] += ";"
                        self._data.loc[(self._data['QC_Status'].str.len() > 1)&(~self._data['QC_Status'].str.endswith('\n')), 'QC_Status'] += "\n"
                        self._data.loc[(self._data['QC_Progname'].str.len() > 1) & ( ~self._data['QC_Progname'].str.endswith(';')), 'QC_Progname'] += ";"
                        self._data['qc_notes'] = np.where(self._data['qc_notes'].str.decode('ascii').isna(),
                                                           self._data['qc_notes'],
                                                           self._data['qc_notes'].str.decode('ascii'))
                        self._data['QC Info'] = self._data['QC_Owner'] + self._data['QC_Status'] +  self._data['QC_Progname'] + self._data['qc_notes']
                        self._data['QC Info'] = self._data['QC Info'].str.strip(';').str.strip('\n')
                        if 'progname' in col.lower():
                            self._data['qcProgPath'] = np.where(self._data["QC_Progname"],
                                                                    f"//sas-vm/{self.parent().projectID}/Data/{self.parent().libName.split(' ')[0]}/SAS Validation/" + \
                                                                    self._data["QC_Progname"].str[:-1] + '.sas',
                                                                    self._data["QC_Progname"])

                            self._data['qcLogPath'] = np.where(self._data["QC_Progname"],
                                                                   f"//sas-vm/{self.parent().projectID}/Data/{self.parent().libName.split(' ')[0]}/SAS Validation/" + \
                                                                    self._data["QC_Progname"].str[:-1]+ '.log',
                                                                   self._data["QC_Progname"])

                        self.visibleData.loc[:,'QC Info'] = self._data['QC Info']
                    self.layoutChanged.emit()
                elif  'custom_Columns' in col:

                    category = col.split('|')[0]
                    colName = col.split('|')[-1]
                    val = self._data.loc[self._data['objectID'] == objectID,colName].values[0]
                    if val:
                        val = json.loads(val)
                        val[category] = value
                    else:
                        val = {}
                        val[category] = value
                    strVal = json.dumps(val)
                    self._data.loc[self._data['objectID'] == objectID, 'Category'] = "\n".join([f"{k}:{v}" for k,v in val.items()])
                    self._data.loc[self._data['objectID'] == objectID, colName] = strVal
                    self.visibleData.loc[:,'Category'] = self._data['Category']
                    self.layoutChanged.emit()

            if strVal:
                value = strVal
                col = colName

            curosr = s.db.cursor()
            curosr.execute(f"UPDATE util_obj SET {col}='{value}' where objectID in {str(objectIDs).replace('[','(').replace(']',')')} and PROJECTID='{self.parent().projectID}' and LIBID='{self.parent().libID}'")
            s.db.commit()
            # self._data.loc[self._data['objectID']==objectID,col] = value
            # update util_obj set {col}='{value}' where objectID='{objectID}' and PROJECTID='{self.parent().projectID}' and LIBID = '{self.parent().libName}'
    @validateAdmin
    def deleteRows(self):
        rows = sorted([row.row() for row in self.parent().treeView.selectionModel().selectedRows()])


        err = []
        objectIDs  = []
        while rows:
            row = rows.pop()
            rowData = self._data[self._data['objectID']==self.visibleData.iloc[row]['objectID']]

            if not rowData['outDated'].any() or rowData.empty:
               err.append(rowData['object_name'].values[0])
               continue
            else:
                objectIDs.append(rowData['objectID'].values[0])


            self._data = self._data.drop(rowData.index).reset_index(drop=True)
            self.visibleData = self.visibleData.drop(self.visibleData.iloc[row].name).reset_index(drop=True)
            self.removeRow(row)
        self.removeRow(row)
        if len(objectIDs)>=1:
            cur = s.db.cursor()
            cur.execute(f"delete from util_obj where objectID in {str(objectIDs).replace('[', '(').replace(']', ')')}")
            s.db.commit()
            self.layoutChanged.emit()
            self.parent().treeView.clearSelection()
        if len(err)>1:
            msg = "Following items are not outdated\n" + "\n".join(err)
            msg = customQMessageBox(msg)
            msg.exec_()
        elif len(err)==1:
            msg = customQMessageBox("Please select an outdated item/")
            msg.exec_()


    def updateSourceData(self, data):
        self._data = data.reset_index(drop=True)
        self.visibleData = self._data[self._header]
        self.layoutChanged.emit()

    def updateFilteredData(self,data):
        self.filteredData = data.reset_index(drop=True)
        self.updateVisibleData(data)

    def updateVisibleData(self,data):
        self.visibleData = data.reset_index(drop=True)[self._header]
        self.layoutChanged.emit()