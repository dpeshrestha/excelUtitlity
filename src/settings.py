import mysql.connector
import os
import tempfile

from PySide2.QtWidgets import QPushButton

db = mysql.connector.connect(host = "192.168.100.15",user="read-utility_user",password="NimUtils1",database="remoteproddb",pool_name='NimUtil',
    pool_size = 8)
isAdmin = False
tempPath = os.path.join(tempfile.gettempdir(),"excelUtilityTemp")
if  not os.path.exists(tempPath):
    os.mkdir(tempPath)

processingBtn = QPushButton()