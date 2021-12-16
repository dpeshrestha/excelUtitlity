import datetime
from time import  time
import mysql.connector
import pandas as pd
import os
import datetime


if __name__ == "__main__":
    begin_time = datetime.datetime.now()
    b = time()
    # pd.read_sql_query('SELECT ProjectID from project',db).to_csv('test.csv',sep=' ')

    db = mysql.connector.connect(host="192.168.100.15", user="read-utility_user", password="NimUtils1",

                                 database="remoteproddb")
    cursor = db.cursor()

    # cursor.execute('SELECt ProjectID from project')
    # values = cursor.fetchall()
    # values = [x[0] for x in values]
    # with open('test.txt','w') as f:
    #     f.writelines(values)
    # print([4,2,3,1,5].sort())
    print(datetime.datetime.now()-begin_time)
    print(time()-b)



