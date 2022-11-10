from pywinauto import *
from time import sleep

from base.base_tool import BaseTool

t = 2

class PrintLabel(BaseTool):
    def __init__(self, windhandle, appl_2, winhandle_2):
        super().__init__(windhandle, appl_2, winhandle_2)
        self.windhandle = windhandle
        self.appl_2 = appl_2
        self.winhandle_2 = winhandle_2

    def getWindhand_reference(self,prtname):
        app3 = Application(backend="uia").connect(title_re=prtname + "*")
        dp3 = app3.window(best_match=prtname + " Properties", top_level_only=False)
        # dp3.print_control_identifiers()
        return app3, dp3

    def enter_PrintTestPage(self, dp3, prtname):
        dp3.child_window(title="Print Test Page", control_type="Button").click_input()

    def enter_CloseButton(self, dp3):
        dp3.child_window(title="Close", auto_id="CommandButton_1", control_type="Button").click()

    def enter_CancelButton(self, dp3):
        dp3.child_window(title="Cancel", control_type="Button").click_input()

    def prtlabel(self, prtname):
        try:
            self.rclick_prtProperties(self.appl_2, self.winhandle_2, prtname)
            sleep(t)
            app3, dp3 = self.getWindhand_reference(prtname)
            sleep(t)
            self.enter_PrintTestPage(dp3, prtname)
            sleep(t)
            app3, dp3 = self.getWindhand_reference(prtname)
            sleep(t)
            self.enter_CloseButton(dp3)
            sleep(t)
            self.enter_CancelButton(dp3)
            sleep(t)

        except Exception as e:
            print("control pannel has more than one same printer,just check it")
            print(e)
        return