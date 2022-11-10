from ddt import ddt,data,file_data,unpack
import pytest
import softest

from time import sleep
import os

from pages.Units import PageUnits
from pages.print_options import PrintOptions
from pages.printlabel import PrintLabel
from utilities.utils import Utils

current_path = os.path.abspath('.')
data_path = current_path + "\\testdata\\"
t1 = 1

@pytest.mark.usefixtures("setup")
@ddt
class TestPrinterMinWH(softest.TestCase):
    @data(*Utils.read_data_from_excel(data_path + "SRSmapping_Forerunner&Stingray_8.xlsx", "Stringray mapping V8"))
    @unpack
    def test_model(self, prtname, d_width, d_height, d_speed, d_darkness, mx_width,
                   mx_height, mn_width, mn_height, speedLimits, darklimits, y_index):

        p_unit = "inch"
        testItem = "speedDark"
        pu = PageUnits(self.windhandle, self.appl_2, self.winhandle_2)
        pu.pg_changeUnits(prtname, p_unit)
        sleep(t1)

        prtop = PrintOptions(self.windhandle, self.appl_2, self.winhandle_2)
        t_speed, t_SpeedLimits, t_darkness, t_DarknessLimits = prtop.pg_PrtOption(prtname, testItem)

        pl = PrintLabel(self.windhandle, self.appl_2, self.winhandle_2)
        pl.prtlabel(prtname)
        sleep(t1)

        ut = Utils()
        p_dark,p_speed,p_width,p_height = ut.get_value()

        v_speed, v_darkness, v_SpeedLimits, v_DarknessLimits = ut.get_sd(d_speed, speedLimits,
            d_darkness, darklimits, t_speed, t_SpeedLimits, t_darkness, t_DarknessLimits, p_speed, p_dark)
        ut.write_data_to_excel_sd(v_speed, v_darkness, v_SpeedLimits, v_DarknessLimits, y_index)
