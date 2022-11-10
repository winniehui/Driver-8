from time import sleep
from pywinauto import *

from base.base_tool import BaseTool


class PageUnits(BaseTool):
    def __init__(self, windhandle, appl_2, winhandle_2):
        super().__init__(windhandle, appl_2, winhandle_2)
        self.windhandle = windhandle
        self.appl_2 = appl_2
        self.winhandle_2 = winhandle_2

    def getWinhandle(self, prtname):
        appl = Application(backend="uia").connect(title_re=prtname + "*")
        Winhand = appl.window(best_match = prtname + " Printing Preferences", top_level_only = False)
        return Winhand

    def clickUnits(self, Winhand):
        Winhand.child_window(title="Units").click_input()

    def selectUnits(self, Winhand, p_unit):
        Winhand.child_window(title="Default measurement units:", control_type="ComboBox").select(p_unit)
        sleep(2)

    def enterApplyButton(self, Winhand):
        Winhand.child_window(title="Apply", control_type="Button").click_input()
        sleep(1)

    def enterOKButton(self, Winhand):
        Winhand.child_window(title="OK", control_type="Button").click_input()
        sleep(2)


    def pg_changeUnits(self, prtname, p_unit):
        self.rclick_prtPreferences(self.appl_2, self.winhandle_2, prtname)
        sleep(3)
        Winhand = self.getWinhandle(prtname)
        self.clickUnits(Winhand)
        Winhand = self.getWinhandle(prtname)
        #"inch" "millimeter"
        self.selectUnits(Winhand, p_unit)


