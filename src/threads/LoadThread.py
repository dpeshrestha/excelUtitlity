# from PySide2.QtCore import QThread, Signal, Slot
# import time
#
# class LoadThread(QThread):
#
#     pstart = Signal()
#     pfinished = Signal()
#     current = Signal(int)
#
#     def __init__(self,func,args,kwargs):
#         super(LoadThread, self).__init__()
#         self.func =func
#         self.args =args
#         self.kwargs =kwargs
#     def run(self):
#         self.pstart.emit()
#         self.func(*self.args,**self.kwargs)
#         self.pfinished.emit()
#
# def runThread(button,text):
#
#     def decorator(func):
#         def wrapper(*args,**kwargs):
#             thread = LoadThread(func,args,kwargs)
#             thread.pstart.connect(lambda: changeText(button, text))
#             thread.pfinished.connect(lambda: changeText(button, ''))
#             thread.start()
#         return wrapper
#
#     return decorator
#
# @Slot()
# def changeText(button,text):
#     print(text)
#
# def sleeeper(test):
#     time.sleep(3)
#     print(test+'arg')
#     print("test  sleeper")
#
#
# if __name__ == "__main__":
#     sleeeper('test')