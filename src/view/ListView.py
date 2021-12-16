from PySide2.QtCore import QModelIndex, Qt, QAbstractListModel
from PySide2.QtWidgets import QAbstractItemView, QListView


class DragDropListView(QListView):
    def __init__(self,parent=None):
        super(DragDropListView, self).__init__(parent)

        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.setDragEnabled(True)
        self.setAcceptDrops(True)
        self.showDropIndicator()
        self.setDropIndicatorShown(True)
        self.setDragDropMode(QAbstractItemView.DragDrop)

    def dragEnterEvent(self, event):
        if self.objectName() == 'excludedListView' or self.objectName() == 'includedListView' :
            if event.source().objectName()=='excludedListView' or event.source().objectName()=='includedListView':
                event.setAccepted(True)
                return

        if self.objectName() == 'keysListView' or self.objectName() == 'varListView':
            if event.source().objectName()=='keysListView' or event.source().objectName()=='varListView':
                event.setAccepted(True)
                return

    def dragMoveEvent(self, e):
        super(DragDropListView, self).dragMoveEvent(e)
        e.setAccepted(True)

    def dropEvent(self, event):
        index = self.indexAt(event.pos())
        source = event.source()
        sourceId = event.source().selectedIndexes()[0]
        if not sourceId.isValid():
            return

        tempVal = source.model()._data[sourceId.row()]
        if self==source:
            self.model().beginRemoveRows(QModelIndex(), index.row(), index.row())
            del self.model()._data[sourceId.row()]
            self.model().endRemoveRows()
        else:
            source.model().beginRemoveRows(QModelIndex(), index.row(), index.row())
            del source.model()._data[sourceId.row()]
            source.model().endRemoveRows()

        if self.dropIndicatorPosition() == QAbstractItemView.BelowItem:
            destIndex = index.row() + 1
        else:
            destIndex = index.row()

        self.model().beginInsertRows(QModelIndex(),destIndex,destIndex)

        if self.dropIndicatorPosition() == QAbstractItemView.OnViewport:
            self.model()._data.append(tempVal)
        else:
            self.model()._data.insert(destIndex,tempVal)
        self.model().endInsertRows()

        if self.objectName() == 'keysListView' or source.objectName() == 'keysListView':
            if len(self.parent().includedListView.selectedIndexes())>0:
                sheetName = self.parent().includedListView.selectedIndexes()[0].data()
                self.parent().keys[sheetName] = self.parent().keysListView.model()._data
        if self!=source and self.objectName() == 'includedListView':
            self.parent().keys[tempVal] = []
        if self!=source and source.objectName() == 'includedListView':
            del self.parent().keys[tempVal]
        self.clearSelection()
