from ddt import ddt,data,file_data,unpack
import pytest
import softest
from time import sleep
import os

from pages.printlabel import PrintLabel
from utilities.utils import Utils
from pages.page_setup import PageSetup

current_path = os.path.abspath('.')
data_path = current_path + "\\testdata\\"

@pytest.mark.usefixtures("setup")
@ddt
class TestPrinterMinWH(softest.TestCase):
    @data(*Utils.read_data_from_excel(data_path + "SRSmapping_Forerunner&Stingray_8.xlsx", "Stringray mapping V8"))
    @unpack
    def test_model(self, prtname, d_width, d_height, d_speed, d_darkness, mx_width,
                   mx_height, mn_width, mn_height, speedLimits, darklimits, y_index):

        enterValue = "0.0001"
        testItem = "min"
        t1 = 3

        ps = PageSetup(self.windhandle, self.appl_2, self.winhandle_2)
        t_width, t_height = ps.pg_setup(prtname, enterValue, testItem)

        pl = PrintLabel(self.windhandle, self.appl_2, self.winhandle_2)
        pl.prtlabel(prtname)
        sleep(t1)

        ut = Utils()
        v_width, v_height = ut.get_wh(mn_width, mn_height, t_width, t_height)
        x_index = 12
        ut.write_data_to_excel_wh(v_width, v_height, x_index, y_index)
        sleep(t1)