from config_loader import ConfigLoader
from nyseg_reader import NysegReader


def new_run():
    config = ConfigLoader()
    nyseg_reader = NysegReader(config)
    nyseg_reader.parse_all_files()

new_run()

