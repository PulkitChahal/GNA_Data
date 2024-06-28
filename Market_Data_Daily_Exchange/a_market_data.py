# ------------IEX--------------
import dam_iex_area
import dam_iex_market
import gdam_iex_area
import gdam_iex_market
import hpdam_iex_area
import hpdam_iex_market
import gtam_iex_area
import gtam_iex_area_intraday
import gtam_iex_market
import gtam_iex_market_intraday
import tam_iex_area_all
import tam_iex_dac
import tam_iex_intraday
import tam_iex_market
import tam_iex_market_intraday

# -------------HPX------------
import dam_hpx_area
import dam_hpx_market
import gdam_hpx_area
import gdam_hpx_market
import gtam_hpx_area
import gtam_hpx_area_intraday
import tam_hpx_area
import tam_hpx_area_dac
import tam_hpx_area_intraday

# -------------PXIL-------------
import dam_pxil_area
import dam_pxil_market
import gdam_pxil_area
import gdam_pxil_market
import hpdam_pxil_area
import hpdam_pxil_market
import gtam_pxil_dac
import gtam_pxil_intraday
import gtam_pxil_anyday
import gtam_pxil_market
import gtam_pxil_market_intraday
import tam_pxil_dac
import tam_pxil_intraday
import tam_pxil_anyday
import tam_pxil_daily
import tam_pxil_weekly
import tam_pxil_monthly
import tam_pxil_market
import tam_pxil_market_intraday

import json
from datetime import datetime
import os

status = {
    "IEX": {},
    "HPX": {},
    "PXIL": {}
}

# ------------IEX-------------
status['IEX']['dam_iex_area.run'] = dam_iex_area.run()
status['IEX']['dam_iex_market'] = dam_iex_market.run()
status['IEX']['gdam_iex_area'] = gdam_iex_area.run()
status['IEX']['gdam_iex_market'] = gdam_iex_market.run()
status['IEX']['hpdam_iex_area'] = hpdam_iex_area.run()
status['IEX']['hpdam_iex_market'] = hpdam_iex_market.run()
# status['IEX']['gtam_iex_area'] = gtam_iex_area.run()
# status['IEX']['gtam_iex_area_intraday'] = gtam_iex_area_intraday.run()
# status['IEX']['gtam_iex_market'] = gtam_iex_market.run()
# status['IEX']['gtam_iex_market_intraday'] = gtam_iex_market_intraday.run()
# status['IEX']['tam_iex_area_all'] = tam_iex_area_all.run()
# status['IEX']['tam_iex_dac'] = tam_iex_dac.run()
# status['IEX']['tam_iex_intraday'] = tam_iex_intraday.run()
# status['IEX']['tam_iex_market'] = tam_iex_market.run()
# status['IEX']['tam_iex_market_intraday'] = tam_iex_market_intraday.run()
#
# # -------------HPX--------------
# status['HPX']['dam_hpx_area'] = dam_hpx_area.run()
# status['HPX']['dam_hpx_market'] = dam_hpx_market.run()
# status['HPX']['gdam_hpx_area'] = gdam_hpx_area.run()
# status['HPX']['gdam_hpx_market'] = gdam_hpx_market.run()
# status['HPX']['gtam_hpx_area'] = gtam_hpx_area.run()
# status['HPX']['gtam_hpx_area_intraday'] = gtam_hpx_area_intraday.run()
# status['HPX']['tam_hpx_area'] = tam_hpx_area.run()
# status['HPX']['tam_hpx_area_dac'] = tam_hpx_area_dac.run()
# status['HPX']['tam_hpx_area_intraday'] = tam_hpx_area_intraday.run()
#
# # -------------PXIL-------------
# status['PXIL']['dam_pxil_area'] = dam_pxil_area.run()
# status['PXIL']['dam_pxil_market'] = dam_pxil_market.run()
# status['PXIL']['gdam_pxil_area'] = gdam_pxil_area.run()
# status['PXIL']['gdam_pxil_market'] = gdam_pxil_market.run()
# status['PXIL']['hpdam_pxil_area'] = hpdam_pxil_area.run()
# status['PXIL']['hpdam_pxil_market'] = hpdam_pxil_market.run()
# status['PXIL']['gtam_pxil_dac'] = gtam_pxil_dac.run()
# status['PXIL']['gtam_pxil_intraday'] = gtam_pxil_intraday.run()
# status['PXIL']['gtam_pxil_anyday'] = gtam_pxil_anyday.run()
# status['PXIL']['gtam_pxil_market'] = gtam_pxil_market.run()
# status['PXIL']['gtam_pxil_market_intraday'] = gtam_pxil_market_intraday.run()
# status['PXIL']['tam_pxil_dac'] = tam_pxil_dac.run()
# status['PXIL']['tam_pxil_intraday'] = tam_pxil_intraday.run()
# status['PXIL']['tam_pxil_anyday'] = tam_pxil_anyday.run()
# status['PXIL']['tam_pxil_daily'] = tam_pxil_daily.run()
# status['PXIL']['tam_pxil_weekly'] = tam_pxil_weekly.run()
# status['PXIL']['tam_pxil_monthly'] = tam_pxil_monthly.run()
# status['PXIL']['tam_pxil_market'] = tam_pxil_market.run()
# status['PXIL']['tam_pxil_market_intraday'] = tam_pxil_market_intraday.run()

today = datetime.now().strftime('%d.%m.%y')
output_directory = r"C:/GNA/Market Data"
output_filename = os.path.join(output_directory, f"Market Data Update Report_{today}.json")

with open(output_filename, 'w') as f:
    json.dump(status, f, indent='\n')

for function, result in status.items():
    print(f"{function}: {result}\n")
