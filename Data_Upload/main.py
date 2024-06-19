import pxil_reverse_auction
import hpx_reverse_auction
import iex_reverse_auction
import deep_portal
import hydro_npp
import outage_npp
import coal_data_npp
import tam_revrese_auction
import deep_tam_data


class daily_data_update:
	def __init__(self):
		pass
	
	def get_data(self):
		hydro_npp.get_data()
		outage_npp.get_data()
		coal_data_npp.coal_npp().get_data()
		deep_portal.deep_portal().get_data()
		# pxil_reverse_auction.pxil_reverse_auction().get_data()
		# hpx_reverse_auction.hpx_reverse_auction().get_data()
		# iex_reverse_auction.iex_reverse_auction().get_data()
		# tam_revrese_auction.tam_reverse_auction().merged_data()
		# deep_tam_data.deep_tam_data().get_data()
		pass


if __name__ == '__main__':
	data_update = daily_data_update()
	data_update.get_data()
	pass
