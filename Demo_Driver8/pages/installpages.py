from time import sleep

from base.install_tool import InstallTool

t = 0.5
t1 = 2
t2 = 4

class InstallPages(InstallTool):
    def __init__(self, appi, st_win):
        super().__init__(appi, st_win)
        self.appi = appi
        self.st_win = st_win

    def enterNextButton(self):
        self.st_win.child_window(title="Next >", control_type="Button").click()

    def enterNext_Button(self,st_Newwin):
        st_Newwin.child_window(title="Next >", control_type="Button").click()

    def getNewWinhan(self):
        st_Newwin = self.appi.window(best_match="Printer Installation Wizard", top_level_only=False)
        return st_Newwin

    def enterInstallPrtDriverButton(self,st_Newwin):
        st_Newwin.child_window(title="Install Printer Driver", control_type="Button").click_input()

    def clickRadioButton(self,st_Newwin):
        st_Newwin.child_window(title="I accept the terms in the license agreement", control_type="RadioButton").click()

    def clickOtherButton(self,st_Newwin):
        st_Newwin.child_window(title="Other", control_type="Button").click()

    def enterPrtName(self, st_Newwin, prtname):
        # 通过键盘输入打印机名字
        st_Newwin.type_keys(prtname)

    def getSelectingPrtHandle(self):
        st_Selecwin = self.appi.window(best_match="PrnInst-Selecting the printer", top_level_only=False)
        return st_Selecwin


    def chooseFilepath(self,st_win1):
        file_path = "D:\\python for tester\\Demo_Driver8\\testdata\\D8.txt"
        st_win1.child_window(title=file_path, control_type="ListItem").click_input()


    def enterYesButton(self, st_win1):
        try:
            st_win1.child_window(title="Yes", control_type="Button").click()
        except Exception as e:
            print(e)

    def enterInstallButton(self, st_Newwin):
        st_Newwin.child_window(title="Install", control_type="Button").click()

    def enterInstallAnotherPrtButton(self, st_Newwin):
        st_Newwin.child_window(title="Install another printer", control_type="Button").click()

    def enterExitButton(self, st_win):
        st_win.child_window(title="Exit", control_type="Button").click()

    def install_1_Step(self, prtname):
        sleep(t1)
        self.enterNextButton()
        sleep(t1)
        st_win1 = self.getNewWinhan()
        sleep(t1)
        self.enterInstallPrtDriverButton(st_win1)
        sleep(t2)
        st_win1 = self.getNewWinhan()
        sleep(t2)
        self.clickRadioButton(st_win1)
        sleep(t2)
        self.enterNext_Button(st_win1)
        sleep(t2)
        st_win1 = self.getNewWinhan()
        self.clickOtherButton(st_win1)
        self.enterPrtName(st_win1, prtname)
        st_win2 = self.getSelectingPrtHandle()
        sleep(1)
        st_win1 = self.getNewWinhan()
        self.enterNext_Button(st_win1)
        st_win1 = self.getNewWinhan()

        self.chooseFilepath(st_win1)
        sleep(2)

        self.enterNext_Button(st_win1)
        sleep(2)
        self.enterYesButton(st_win1)
        sleep(2)
        st_win1 = self.getNewWinhan()
        self.enterInstallButton(st_win1)
        sleep(7)
        st_win1 = self.getNewWinhan()
        self.enterInstallAnotherPrtButton(st_win1)
        sleep(3)

    def installAnotherPrtStep(self, prtname):
        st_win1 = self.getNewWinhan()
        self.clickOtherButton(st_win1)
        self.enterPrtName(st_win1, prtname)
        st_win2 = self.getSelectingPrtHandle()
        sleep(1)
        st_win1 = self.getNewWinhan()
        self.enterNext_Button(st_win1)
        st_win1 = self.getNewWinhan()

        self.chooseFilepath(st_win1)
        sleep(2)

        self.enterNext_Button(st_win1)
        sleep(2)
        self.enterYesButton(st_win1)
        sleep(2)
        st_win1 = self.getNewWinhan()
        self.enterInstallButton(st_win1)
        sleep(7)
        st_win1 = self.getNewWinhan()
        self.enterInstallAnotherPrtButton(st_win1)
        sleep(3)

    def installStep(self, prtname, flag):
        # 通过y_index的值控制调用函数，当y_index == 3时，安装第一个printer
        if flag == 3:
            self.install_1_Step(prtname)
        else:
            self.installAnotherPrtStep(prtname)

