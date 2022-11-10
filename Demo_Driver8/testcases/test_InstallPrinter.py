from time import sleep

import pytest
import softest
from ddt import ddt,data,file_data,unpack

from pages.installpages import InstallPages
from utilities.utils import Utils

data_path = "D:\\python for tester\\Demo_Driver8\\testdata\\"
@pytest.mark.usefixtures("install_setup")
@ddt
class InstallPrinter(softest.TestCase):
    @data(*Utils.read_data_from_excel(data_path + "SRSmapping_Forerunner&Stingray_8.xlsx", "Stringray mapping V8"))
    @unpack
    def test_model(self, prtname, d_width, d_height, d_speed, d_darkness, mx_width,
                   mx_height, mn_width, mn_height, speedLimits, darklimits, y_index):
        print("------------------------------------------------------------------------")
        print("\n     prtname: ", prtname)
        print("y_index: ", y_index)
        print("type(y_index): ", type(y_index))
        sleep(3)
        inp = InstallPages(self.appi, self.st_win)
        sleep(3)

        inp.installStep(prtname, y_index)

        sleep(3)


