import pytest
from pywinauto import *
from time import sleep

t1 = 3

@pytest.fixture(scope="class")
def setup(request):
    app = Application(backend="uia").start('control.exe')
    cpanel = Desktop(backend="uia").ControlPanel

    cpanel.child_window(title="Hardware and Sound").click_input()
    appl = Application(backend="uia").connect(title_re=".*Hardware and Sound*")
    windhandle = appl.window(best_match='Hardware and Sound', top_level_only=False)
    windhandle.child_window(title="Devices and Printers").click_input()
    sleep(t1)

    appl_2 = Application(backend="uia").connect(title_re=".*Devices and Printers*")
    winhandle_2 = appl_2.window(best_match="Devices and Printers", top_level_only=False)

    request.cls.windhandle = windhandle
    request.cls.appl_2 = appl_2
    request.cls.winhandle_2 = winhandle_2

    sleep(t1)

    yield
    windhandle.type_keys("%{F4}")


@pytest.fixture(scope="class")
def install_setup(request):
    #packpath = "D:\\DriverTest\\ZD8-6-4-23832"
    packpath = "C:\\ZD8-6-5-26122\\"
    packname = "PrnInst.exe"
    print("enter install setup")

    appi = Application(backend="uia").start(packpath + packname)

    st_win = appi.window(best_match="Printer Installation Wizard", top_level_only=False)

    request.cls.appi = appi
    request.cls.st_win = st_win
    yield
    st_win.child_window(title="Exit", control_type="Button").click()


