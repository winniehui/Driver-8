import os
import datetime
from time import sleep
from pywinauto import *
from pywinauto import keyboard

from base.base_tool import BaseTool

t = 1
current_date = datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S")
current_path = os.path.abspath('.')

class PageSetup(BaseTool):
    def __init__(self, windhandle, appl_2, winhandle_2):
        super().__init__(windhandle, appl_2, winhandle_2)
        self.windhandle = windhandle
        self.appl_2 = appl_2
        self.winhandle_2 = winhandle_2

    def getWinhandle(self, prtname):
        appl = Application(backend="uia").connect(title_re=prtname + "*")
        Winhand = appl.window(best_match = prtname + " Printing Preferences", top_level_only = False)
        return Winhand

    def getWidth(self, Winhand):
        text_width = Winhand.child_window(title="Width:", control_type="Edit").get_value()
        t_width = str(int(float(text_width.strip(" mm")) * 10))
        return t_width

    def getHeight(self, Winhand):
        text_height = Winhand.child_window(title = "Height:", control_type = "Edit").get_value()
        t_height = str(int(float(text_height.strip(" mm")) * 10))
        return t_height

    def getImg(self,Winhand, prtname,testItem):
        a = Winhand.capture_as_image()
        img_path = current_path + "\\reports\\img\\"
        imgName = img_path + prtname + "_" + testItem + "_" + current_date + ".png"
        a.save(imgName)
        sleep(t)

    def enterCancelButton(self, Winhand):
        Winhand.child_window(title = "Cancel", control_type = "Button").click_input()
        sleep(t)

    def enterWidth(self, Winhand, entervalue):
        Winhand.child_window(title="Width:", control_type="Edit").click_input()
        sleep(t)
        for i in range(0, 7):
            keyboard.send_keys("{BACKSPACE}")
            sleep(0.5)
        Winhand.child_window(title="Width:", control_type="Edit").type_keys(entervalue)
        sleep(t)

    def enterHeight(self, Winhand, entervalue):
        Winhand.child_window(title="Height:", control_type="Edit").click_input()
        sleep(t)
        for i in range(0, 7):
            keyboard.send_keys("{BACKSPACE}")
            sleep(0.5)
        Winhand.child_window(title="Height:", control_type="Edit").type_keys(entervalue)
        sleep(t)

    def clickApplyButton(self, Winhand):
        Winhand.child_window(title="Apply", control_type="Button").click()
        sleep(t)

    def pg_setup_default(self, prtname, testItem):
        self.rclick_prtPreferences(self.appl_2, self.winhandle_2, prtname)
        sleep(3)
        Winhand = self.getWinhandle(prtname)
        t_width = self.getWidth(Winhand)
        t_height = self.getHeight(Winhand)
        self.getImg(Winhand, prtname, testItem)
        self.enterCancelButton(Winhand)
        sleep(3)
        return t_width, t_height

    def pg_setup(self, prtname, enterValue, testItem):
        self.rclick_prtPreferences(self.appl_2, self.winhandle_2, prtname)
        sleep(3)
        Winhand = self.getWinhandle(prtname)
        self.enterWidth(Winhand, enterValue)
        self.enterHeight(Winhand, enterValue)
        self.clickApplyButton(Winhand)
        t_width = self.getWidth(Winhand)
        t_height = self.getHeight(Winhand)
        self.getImg(Winhand, prtname, testItem)
        self.enterCancelButton(Winhand)
        sleep(3)
        return t_width, t_height







