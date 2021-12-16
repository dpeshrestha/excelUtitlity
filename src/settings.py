import mysql.connector
import os
import tempfile
db = mysql.connector.connect(host = "192.168.100.15",user="read-utility_user",password="NimUtils1",database="remoteproddb")
isAdmin = False
cursor = db.cursor()
tempPath = os.path.join(tempfile.gettempdir(),"excelUtilityTemp")
if  not os.path.exists(tempPath):
    os.mkdir(tempPath)
