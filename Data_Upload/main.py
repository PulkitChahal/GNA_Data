import Data_Upload.hydro_npp
import Data_Upload.outage_npp
import Data_Upload.coal_data_npp
import Data_Upload.deep_portal
import Data_Upload.reverse_auction
import Data_Upload.deep_tam_data
import Data_Upload.data_mapping
import Data_Upload.plant_plf_monthly


class daily_data_update:
    def __init__(self):
        pass

    def get_data(self):
        Data_Upload.hydro_npp.hydro_npp_daily_data().get_data()
        Data_Upload.outage_npp.get_data()
        Data_Upload.coal_data_npp.coal_npp().get_data()
        Data_Upload.deep_portal.deep_portal().get_data()
        Data_Upload.reverse_auction.tam_reverse_auction().get_data()
        pass


if __name__ == '__main__':
    data_update = daily_data_update()
    data_update.get_data()
    pass













