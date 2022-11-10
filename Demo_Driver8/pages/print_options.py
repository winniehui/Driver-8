import datetime
from time import sleep
import os
from pywinauto import *

from base.base_tool import BaseTool

t1 =1
current_date = datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S")
current_path = os.path.abspath('.')

class PrintOptions(BaseTool):
    def __init__(self, windhandle, appl_2, winhandle_2):
        super().__init__(windhandle, appl_2, winhandle_2)
        self.windhandle = windhandle
        self.appl_2 = appl_2
        self.winhandle_2 = winhandle_2

    def getWinhandle(self, prtname):
        appl = Application(backend="uia").connect(title_re=prtname + "*")
        Winhand = appl.window(best_match = prtname + " Printing Preferences", top_level_only = False)
        return Winhand

    def clickPrintOptions(self, Winhand, prtname):
        Winhand.child_window(title="Print Options").click_input()
        sleep(t1)

    def getImg(self, Winhand, prtname, testItem):
        # 截图
        a = Winhand.capture_as_image()
        img_path = current_path + "\\reports\\img\\"
        imgName = img_path + prtname + "_" + testItem + "_" + current_date + ".png"
        a.save(imgName)

    def getSpeed(self, Winhand):
        # get speed default value
        d_speed = Winhand.child_window(title="Speed:", control_type="ComboBox").selected_text()
        t_speed = (d_speed.split(' "/s')[0]).split('.')[0]
        return t_speed

    def enterSpeedClick(self, Winhand):
        Winhand.child_window(title="Speed:", control_type="ComboBox").click_input()

    def getSpeedList(self, Winhand):
        i4 = Winhand.child_window(title="Speed:", control_type="List").item_count()
        print("SpeedLimits length: ", i4)

        i5 = Winhand.child_window(title="Speed:", control_type="List").texts()
        t_SpeedLimits = ''
        for i in range(0, i4):
            i_speed = (i5[i][0].split(' "/s')[0]).split('.')[0]
            if i != i4 - 1:
                t_SpeedLimits = t_SpeedLimits + str(i_speed) + ","
            else:
                t_SpeedLimits = t_SpeedLimits + str(i_speed)
        return t_SpeedLimits

    def getDarkness(self, Winhand):
        t_darkness = Winhand.child_window(title="Darkness:", control_type="ComboBox").selected_text()
        return t_darkness

    def enterDarknessClick(self, Winhand):
        Winhand.child_window(title="Darkness:", control_type="ComboBox").click_input()

    def getDarknessList(self, Winhand):
        j4 = Winhand.child_window(title="Darkness:", control_type="List").texts()
        t_DarknessLimits = str(j4[0][0]) + "-" + str(j4[len(j4) - 1][0])
        return t_DarknessLimits

    def enterCancelButton(self, Winhand):
        Winhand.child_window(title="Cancel", control_type="Button").click_input()

    def pg_PrtOption(self, prtname, testItem):
        Winhand = self.getWinhandle(prtname)
        self.clickPrintOptions(Winhand, prtname)
        Winhand = self.getWinhandle(prtname)
        self.getImg(Winhand, prtname, testItem)
        t_speed = self.getSpeed(Winhand)
        self.enterSpeedClick(Winhand)
        self.getWinhandle(prtname)
        t_SpeedLimits = self.getSpeedList(Winhand)
        self.enterSpeedClick(Winhand)
        t_darkness = self.getDarkness(Winhand)
        self.enterDarknessClick(Winhand)
        sleep(t1)
        self.getWinhandle(prtname)
        t_DarknessLimits = self.getDarknessList(Winhand)
        self.enterDarknessClick(Winhand)
        self.getWinhandle(prtname)
        self.enterCancelButton(Winhand)
        return t_speed, t_SpeedLimits, t_darkness, t_DarknessLimits


