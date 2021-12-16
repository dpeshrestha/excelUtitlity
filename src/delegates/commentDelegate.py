import tempfile
import os
import subprocess
import pandas as pd
from PySide2.QtCore import Qt, QSize
from PySide2.QtGui import QFont, QPainter, QIcon, QCursor
from PySide2.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QLabel, QSizePolicy, QStyleOption, QStyle, QPushButton
import src.settings as s
from src.utils.utils import temp_file
         # subprocess.Popen(f'explorer {filename}')


class AttachmentButton(QPushButton):
    def __init__(self,name,file):
        super(AttachmentButton, self).__init__(text=name)
        self.setObjectName('attachment')
        self.name = name
        self.file = file
        self.setStyleSheet("*{background-color:transparent;border:none;text-decoration:none;color:blue;cursor:pointer;}"                             
                            "*:hover{text-decoration:underline;}")
        self.setCursor(QCursor(Qt.PointingHandCursor))
        self.clicked.connect(self.openAttachment)

    def openAttachment(self):
        print('opening: '+self.text())

        self.parent().setDisabled(True)
        with temp_file('wb',self.file,suffix='.'+self.name.split('.')[-1]) as filename:

            os.system(filename)
        self.parent().setEnabled(True)


class CommentWidget(QWidget):
    def __init__(self,user,text,date,commentId,item,attachments=pd.DataFrame(),parent=None):
        super(CommentWidget, self).__init__(parent)
        boldFont = QFont()
        boldFont.setBold(True)
        self.user=QLabel(user.strip())
        self.user.setFont(boldFont)
        self.commentText =text
        self.item = item
        self.text=QLabel(text)
        self.text.setWordWrap(True)
        self.text.setSizePolicy(QSizePolicy.MinimumExpanding,QSizePolicy.MinimumExpanding)
        self.date=QLabel(date)
        self.commentId = commentId
        self.date.setFont(boldFont)
        self.toggleAttachment = False
        # self.attachments = pd.read_sql_query(f"SELECT attachment_name from util_comment_attachments where comment_id= '{self.commentId}'",s.db)['attachment_name'].values.tolist()
        self.attachments = attachments
        self.setupUI()
        if self.attachments.empty:
            self.attachmentBtn.hide()
        self.setObjectName("comment")
        self.attachmentBtn.clicked.connect(self.showAttachments)
        # self.text.setText(self.attachments + "\n" + self.commentText)
    def showAttachments(self):
        #use self.commentId to get attachemnts

        if self.attachmentWidget.isVisible():
            self.attachmentWidget.hide()
            self.setFixedHeight(self.minimumSizeHint().height())
        else:
            self.attachmentWidget.show()
            self.setFixedHeight(self.minimumSizeHint().height()+60*len(self.attachments))
        self.toggleAttachment = not self.toggleAttachment
        self.item.setSizeHint(self.sizeHint())
        # self.resizeWidget()

    def paint(self):
        opt = QStyleOption()
        opt.initForm(self)
        painter = QPainter(self)
        self.style().drawPrimitive(QStyle.PE_Widget,opt,painter,self)

    def setText(self,text):
        self.text.setText(text)

    def mouseDoubleClickEvent(self, event):
        from src.view.customDialogs import customQMessageBox
        s.cursor.execute(f'SELECT comment_userid from util_issue_comment where comment_id = {self.commentId}')
        commentUserId = s.cursor.fetchone()[0]
        if s.currentUser != commentUserId:

            msg = customQMessageBox("You cannot edit this comment because you are not the author of this comment.")
            msg.exec_()
            return

        if self.parent().parent().currentIndex().row() == 0:
            from src.view.CustomWidgets import CommentDialog
            commentDialog = CommentDialog(self.parent().parent().parent().parent().treeView.currentIndex().row(),self.commentId, parent=self.parent().parent().parent().parent())
            commentDialog.comentEdit.setPlainText(self.text.text())
            commentDialog.exec_()
    #
    # def resizeWidget(self):
    #     self.resize(QSize(self.minimumSizeHint().width(),self.minimumHeight()+50*4))
        # self.
        # self.setFixedHeight(216)


    def setupUI(self):

        self.commentLayout = QVBoxLayout(self)
        self.topLayout = QHBoxLayout(self)
        self.topLayout.addWidget(self.user,alignment=Qt.AlignLeft)
        self.attachmentBtn = QPushButton(icon=QIcon("icons/clip.png"))
        self.attachmentBtn.setStyleSheet("*{background-color:white;border:none}"
                                       "QPushButton:hover{background-color:lightgrey;}"
                                       "QPushButton:clicked{background-color:grey;}") # with user id
        self.attachmentBtn.setIconSize(QSize(32,32))
        self.attachmentBtn.setFixedSize(QSize(15,15))
        self.topLayout.addWidget(self.attachmentBtn,alignment=Qt.AlignLeft) # with user id

        self.topLayout.addWidget(self.date,alignment=Qt.AlignRight)
        self.commentLayout.addLayout(self.topLayout)
        # self.bottomLayout = QHBoxLayout(self)
        self.attachmentWidget = QWidget(self)
        self.attachmentLayout = QVBoxLayout(self)
        self.attachmentWidget.setLayout(self.attachmentLayout)
        if not self.attachments.empty:
            atachmentbtns = [AttachmentButton(attachment_name,attachment) for attachment_name,attachment in self.attachments[['attachment_name','attachment']].values.tolist()]
            for btn in atachmentbtns:
                self.attachmentLayout.addWidget(btn,alignment=Qt.AlignLeft)
        self.attachmentWidget.hide()
        self.commentLayout.addWidget(self.attachmentWidget)
        # self.bottomLayout.addWidget(self.text,alignment=Qt.AlignLeft)
        # self.commentLayout.setAlignment(self.text,alignment=Qt.AlignLeft)
        self.commentLayout.addWidget(self.text,alignment=Qt.AlignTop)
        # self.commentLayout.addWidget(QLabel('_'*53),alignment=Qt.AlignBottom)
        # self.text.setSizePolicy(QSizePolicy.Minimum,QSizePolicy.Minimum)

        self.setLayout(self.commentLayout)
        self.commentLayout.addStretch()
        # self.attachmentWidget.setSizePolicy(QSizePolicy.Minimum,QSizePolicy.Minimum)
        self.setSizePolicy(QSizePolicy.Minimum,QSizePolicy.Minimum)
#



