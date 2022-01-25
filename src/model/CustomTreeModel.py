import json

import numpy as np
from PySide2.QtGui import QStandardItemModel, Qt, QColor
import src.settings as s
from src.utils.utils import addSeperator, replaceExtension
from src.utils.validation import validateAdmin
from src.view.customDialogs import customQMessageBox


class CustomTreeModel(QStandardItemModel):
    def __init__(self,data,cols,parent=None):
        super(CustomTreeModel, self).__init__(parent)
        self._data = data
        self._header = cols
        self.visibleData = self._data

        self.setObjectName('report')
        self.empty = False # added to make compatible with Dataframe


    def data(self, index, role):
        if role == Qt.BackgroundColorRole:
            objectId = self.itemFromIndex(index).id
            if self.parent().state == 'Tasks':
                if self._header[index.column()] == 'Status':
                    for w in ['Cancelled', 'Not Applicable', 'Deferred']:
                        if w in self.itemFromIndex(index).text():
                            return QColor(255, 102, 0)
            if self.parent().state == 'Items':
                color = QColor(250,165,165,50)
                if self._header[index.column()] == 'Primary Info':
                    validation = self._data.set_index('objectID').loc[objectId]['Validation_fail']['Primary']
                    if self._data.set_index('objectID').loc[objectId].copy(True)['primary_log']:
                        validation.append(self._data.set_index('objectID').loc[objectId]['primary_log'])
                    if validation:
                        return color
                if self._header[index.column()] == 'QC Info':
                    validation = self._data.set_index('objectID').loc[objectId]['Validation_fail']['QC']
                    if self._data.set_index('objectID').loc[objectId].copy(True)['qc_log']:
                        validation.append(self._data.set_index('objectID').loc[objectId]['qc_log'])

                    if validation:
                        return color


        elif role == Qt.ForegroundRole:
            if self._header[index.column()] in ['Planned Start Date','Planned End Date','Completion Date']:
                if self.itemFromIndex(index).text() == 'yyyy/mm/dd':
                    return QColor('grey')
            # if self.parent().state == 'Items':
            #     if self._data.set_index('objectID').loc[self.itemFromIndex(index).id]['outDated']:
            #         return QColor('grey')

        elif role == Qt.TextColorRole:
            #     if self._data[self._data['objectID'] == self.visibleData.iat[index.row(), 0]]['outDated'].any():
                return QColor('grey')
            # pass
        else:
            return super(CustomTreeModel, self).data(index,role)



    def setSource(self,data):
        self._data = data
        self.visibleData = data

    def updateData(self, indexes,col, value, role=Qt.EditRole):
        strVal = ''
        value = '' if value == '_NA_' or value == 'None' else value
        if role == Qt.EditRole:
            objectIDs = [self.itemFromIndex(index.siblingAtColumn(0)).id for index in indexes]
            for objectID,index in zip(objectIDs,indexes):

                # self._data.loc[self._data['objectID']==objectID,col] = value
                if 'primary' in col.lower():
                    # colDict = {'Primary_Owner':0,'Primary_Status':1,'Primary_Progname':2,'Primary_Notes':3}
                    self._data.loc[self._data['objectID'] == objectID, col] = value
                    # if col == 'Primary_Progname':


                    # self._data.loc[(self._data['Primary_Owner'].str.len() > 1) &(~self._data['Primary_Owner'].str.endswith(';')), 'Primary_Owner'] += ";"
                    # self._data.loc[(self._data['Primary_Status'].str.len() > 1)  &(~self._data['Primary_Status'].str.endswith('\n')), 'Primary_Status'] += "\n"
                    # self._data.loc[(self._data['Primary_Progname'].str.len() > 1)  &(~self._data['Primary_Progname'].str.endswith(';')), 'Primary_Progname'] += ";"
                    self._data['Primary_Notes'] = np.where(self._data['Primary_Notes'].str.decode('ascii').isna(),
                                                           self._data['Primary_Notes'],
                                                           self._data['Primary_Notes'].str.decode('ascii'))
                    self._data['Primary Owner'] = self._data['Primary_Owner']
                    self._data['Primary Info'] = self._data['Primary_Status'].apply(addSeperator) + \
                                                 self._data['Primary_Progname'].apply(lambda x: addSeperator(x, '\n')) + \
                                                 self._data['Primary_Notes']
                    self._data['Primary Info'] = self._data['Primary Info'].str.strip(';').str.strip('\n')
                    if 'progname' in col.lower():
                        self._data['primaryProgPath'] = np.where(self._data["Primary_Progname"],
                                                                 f"//sas-vm/{self.parent().projectID}/Reports/{self.parent().libName.split(' ')[0]}/SAS Programs/" + \
                                                                 self._data["Primary_Progname"],
                                                                 self._data["Primary_Progname"])

                        self._data['primaryLogPath'] = np.where(self._data["ITEMID"],
                                                                f"//sas-vm/{self.parent().projectID}/Reports/{self.parent().libName.split(' ')[0]}/SAS Programs/" + \
                                                                self._data["ITEMID"]+'.log',
                                                                self._data["ITEMID"])

                    self.setData(index.siblingAtColumn(self._header.index('Primary Info')),self._data.loc[self._data['objectID'] == objectID, 'Primary Info'].values[0],Qt.EditRole)
                    self.setData(index.siblingAtColumn(self._header.index('Primary Owner')),self._data.loc[self._data['objectID'] == objectID, 'Primary Owner'].values[0],Qt.EditRole)
                    self.layoutChanged.emit()

                elif 'qc' in col.lower():
                    self._data.loc[self._data['objectID'] == objectID, col] = value
                    if col == 'QC_Type':
                        self.setData(index.siblingAtColumn(self._header.index('QC Type')),
                                     self._data.loc[self._data['objectID'] == objectID,col].values[0],
                                     Qt.EditRole)
                    else:
                    # colDict = {'QC_Owner': 0, 'QC_Status': 1, 'QC_Progname': 2,}

                        # self._data.loc[(self._data['QC_Owner'].str.len() > 1) &(~self._data['QC_Owner'].str.endswith(';')), 'QC_Owner'] += ";"
                        # self._data.loc[(self._data['QC_Status'].str.len() > 1)&(~self._data['QC_Status'].str.endswith('\n')), 'QC_Status'] += "\n"
                        # self._data.loc[(self._data['QC_Progname'].str.len() > 1) & ( ~self._data['QC_Progname'].str.endswith(';')), 'QC_Progname'] += ";"
                        self._data['QC Owner'] = self._data['QC_Owner']
                        self._data['qc_notes'] = np.where(self._data['qc_notes'].str.decode('ascii').isna(),
                                                      self._data['qc_notes'],
                                                      self._data['qc_notes'].str.decode('ascii'))
                        self._data['QC Info'] = self._data['QC_Status'].apply(addSeperator) + \
                                                self._data['QC_Progname'].apply(lambda x: addSeperator(x, '\n')) + \
                                                self._data['qc_notes']
                        self._data['QC Info'] = self._data['QC Info'].str.strip(';').str.strip('\n')
                        if 'progname' in col.lower():
                            self._data['qcProgPath'] = np.where(self._data["QC_Progname"],
                                                                f"//sas-vm/{self.parent().projectID}/Reports/{self.parent().libName.split(' ')[0]}/SAS Validation/" + \
                                                                self._data["QC_Progname"],
                                                                self._data["QC_Progname"])

                            self._data['qcLogPath'] = np.where(self._data["ITEMID"],
                                                               f"//sas-vm/{self.parent().projectID}/Reports/{self.parent().libName.split(' ')[0]}/SAS Validation/" + \
                                                               'v_'+self._data["ITEMID"]+'.log',
                                                               self._data["ITEMID"])

                    self.setData(index.siblingAtColumn(self._header.index('QC Info')), self._data.loc[self._data['objectID'] == objectID, 'QC Info'].values[0],Qt.EditRole)
                    self.setData(index.siblingAtColumn(self._header.index('QC Owner')), self._data.loc[self._data['objectID'] == objectID, 'QC Owner'].values[0],Qt.EditRole)
                    self.layoutChanged.emit()
                elif 'custom_Columns' in col:

                    category = col.split('|')[0]
                    colName = col.split('|')[-1]
                    val = self._data.loc[self._data['objectID'] == objectID, colName].values[0]
                    if val:
                        val = json.loads(val)
                        val[category] = value
                    else:
                        val = {}
                        val[category] = value
                    strVal = json.dumps(val)
                    self._data.loc[self._data['objectID'] == objectID, 'Category'] = "\n".join(
                        [f"{k}:{v}" for k, v in val.items()])
                    self._data.loc[self._data['objectID'] == objectID, colName] = strVal
                    self.visibleData.loc[:, 'Category'] = self._data['Category']
                    self.setData(index.siblingAtColumn(self._header.index('Category')),
                                 self._data.loc[self._data['objectID'] == objectID, 'Category'].values[0],
                                 Qt.EditRole)
                    self.layoutChanged.emit()

            if strVal:
                value = strVal
                col = colName
            cursor = s.db.cursor()
            cursor.execute(f"UPDATE util_obj SET {col}='{value}' where objectID in {str(objectIDs).replace('[','(').replace(']',')')} and PROJECTID='{self.parent().projectID}' and LIBID='{self.parent().libName}'")
            s.db.commit()
            print(objectID,value,col)

    def updateFilteredData(self,data):
        self.filteredData = data.reset_index(drop=True)
        self.updateVisibleData(data)

    @validateAdmin
    def deleteRows(self):

        if len(self.parent().treeView.selectionModel().selectedRows()) < 1:
            msg = customQMessageBox('Please select a row')
            msg.exec_()
            return
        items = [self.itemFromIndex(row) for row in self.parent().treeView.selectionModel().selectedRows()]
        ids = [item.id for item in items]

        err = []
        objectIDs = []


        while items:
            item = items.pop()
            rowData = self._data.set_index('objectID').loc[item.id]
            if not rowData['outDated'].any() or rowData.empty:
               err.append(item.text())
               continue
            else:
                objectIDs.append(item.id)
            if item.parent():
                self._data = self._data.set_index('objectID').drop(item.id).reset_index()
                self.visibleData = self.visibleData.set_index('objectID').drop(item.id).reset_index()
                item.parent().removeRow(item.index().row())
            else:
                objectIDs.remove(item.id)
                continue


        if objectIDs:
            cur = s.db.cursor()
            cur.execute(f"delete from util_obj where objectID in {str(objectIDs).replace('[', '(').replace(']', ')')}")
            s.db.commit()
            self.parent().treeView.clearSelection()
        if len(err) > 1:
            msg = "Following items are not outdated\n" + "\n".join(err)
            msg = customQMessageBox(msg)
            msg.exec_()
        elif len(err) == 1:
            msg = customQMessageBox("Please select an outdated item.")
            msg.exec_()
        # [self.removeRow(row.row()) for row in self.parent().treeView.selectionModel().selectedRows()]

    def updateVisibleData(self, data=None):
        if data is None:
            self.visibleData = self.visibleData.merge(self._data, on=["objectID" if self.parent().state == 'Items' else "issueID"], suffixes=('_x', ''))
        else:
            self.visibleData = data.reset_index(drop=True)

