

class BaseTool:
    def __init__(self, windhandle, appl_2, winhandle_2):
        self.windhandle = windhandle
        self.appl_2 = appl_2
        self.winhandle_2 = winhandle_2

    def rclick_prtPreferences(self, appl_2, winhandle_2, prtname):
        winhandle_2.child_window(title=prtname).click_input(button='right')
        appl_2.Context.Printingpreferences.click_input()

    def rclick_prtProperties(self, appl_2, winhandle_2, prtname):
        winhandle_2.child_window(title=prtname).click_input(button='right')
        appl_2.Context.Printerproperties.click_input()









