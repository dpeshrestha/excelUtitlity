import json

from PySide2.QtGui import QStandardItemModel, Qt, QColor
import src.settings as s

class CustomTreeModel(QStandardItemModel):
    def __init__(self,data,cols,parent=None):
        super(CustomTreeModel, self).__init__(parent)
        self._data = data
        self._headerData = cols
        self.visibleData = self._data[cols]
        self.setObjectName('report')

    def data(self, index, role):
        if role == Qt.BackgroundColorRole:
            if self._headerData[index.column()] == 'Status':
                for w in ['Cancelled', 'Not Applicable', 'Deferred']:
                    if w in self.itemFromIndex(index).text():
                        return QColor(255, 102, 0)
        if role == Qt.ForegroundRole:
            if self._headerData[index.column()] in ['Planned Start Date','Planned End Date','Completion Date']:
                if self.itemFromIndex(index).text() == 'yyyy/mm/dd':
                    return QColor('grey')
            # pass
        else:
            return super(CustomTreeModel, self).data(index,role)


    def setSource(self,data):
        self._data = data
        self.visibleData = [self._headerData]

    def updateData(self, indexes,col, value, role=Qt.EditRole):
        strVal = ''
        value = '' if value == '_NA_' or value == 'None' else value
        if role == Qt.EditRole:
            objectIDs = [self.itemFromIndex(index.siblingAtColumn(0)).id for index in indexes]
            for objectID,index in zip(objectIDs,indexes):

                # self._data.loc[self._data['objectID']==objectID,col] = value
                if 'Primary' in col:
                    # colDict = {'Primary_Owner':0,'Primary_Status':1,'Primary_Progname':2,'Primary_Notes':3}
                    self._data.loc[self._data['objectID'] == objectID, col] = value
                    # if col == 'Primary_Progname':


                    self._data.loc[(self._data['Primary_Owner'].str.len() > 1) &(~self._data['Primary_Owner'].str.endswith(';')), 'Primary_Owner'] += ";"
                    self._data.loc[(self._data['Primary_Status'].str.len() > 1)  &(~self._data['Primary_Status'].str.endswith('\n')), 'Primary_Status'] += "\n"
                    self._data.loc[(self._data['Primary_Progname'].str.len() > 1)  &(~self._data['Primary_Progname'].str.endswith(';')), 'Primary_Progname'] += ";"
                    self._data['Primary Info'] = self._data['Primary_Owner'] + self._data['Primary_Status'] + self._data[
                        'Primary_Progname'] + self._data['Primary_Notes']
                    self._data['Primary Info'] = self._data['Primary Info'].str.strip(';').str.strip('\n')
                    # self.visibleData['Primary Info'] =  self._data['Primary Info']
                    # index = self.index(index.row(), self._headerData.index('Primary Info'))
                    self.setData(index.siblingAtColumn(self._headerData.index('Primary Info')),self._data.loc[self._data['objectID'] == objectID, 'Primary Info'].values[0],Qt.EditRole)
                    self.layoutChanged.emit()

                elif 'QC' in col:
                    self._data.loc[self._data['objectID'] == objectID, col] = value
                    if col == 'QC_Type':
                        self.setData(index.siblingAtColumn(self._headerData.index('QC Type')),
                                     self._data.loc[self._data['objectID'] == objectID,col].values[0],
                                     Qt.EditRole)
                    else:
                    # colDict = {'QC_Owner': 0, 'QC_Status': 1, 'QC_Progname': 2,}

                        self._data.loc[(self._data['QC_Owner'].str.len() > 1) &(~self._data['QC_Owner'].str.endswith(';')), 'QC_Owner'] += ";"
                        self._data.loc[(self._data['QC_Status'].str.len() > 1)&(~self._data['QC_Status'].str.endswith('\n')), 'QC_Status'] += "\n"
                        self._data.loc[(self._data['QC_Progname'].str.len() > 1) & ( ~self._data['QC_Progname'].str.endswith(';')), 'QC_Progname'] += ";"
                        self._data['QC Info'] = self._data['QC_Owner'] + self._data['QC_Status'] + self._data['QC_Progname']
                        self._data['QC Info'] = self._data['QC Info'].str.strip(';').str.strip('\n')
                        # self.visibleData['QC Info'] = self._data['QC Info']
                        # index = self.index(index.row(), self._headerData.index('QC Info'))
                        self.setData(index.siblingAtColumn(self._headerData.index('QC Info')), self._data.loc[self._data['objectID'] == objectID, 'QC Info'].values[0],
                                     Qt.EditRole)
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
                    self.setData(index.siblingAtColumn(self._headerData.index('Category')),
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



    def deleteRows(self):
        print('Delete')

    def updateVisibleData(self,data):
        self.visibleData = data[self._headerData]
