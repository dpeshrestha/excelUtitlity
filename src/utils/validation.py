
import src.settings as s
from src.view.customDialogs import customQMessageBox


def validateAdmin(func):
    def wrapper(*args,**kwargs):
        if s.isAdmin:
            return func(*args,**kwargs)
        else:
            msg = customQMessageBox("Access Error. Only admin has access to this functionality.")
            msg.exec_()

    return wrapper


def validateUsers(usersList=[],errorText="Access Error. Only admin has access to this functionality."):
    def decorator(func):
        def wrapper(*args,**kwargs):
            print(usersList)
            if s.currentUser in usersList:
                return func(*args,**kwargs)
            else:
                msg = customQMessageBox(errorText)
                msg.exec_()
        return wrapper
    return decorator





