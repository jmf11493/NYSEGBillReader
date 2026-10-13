import yaml


class ConfigLoader:

    def __init__(self):
        with open("config.yaml", "r") as config_file:
            config = yaml.safe_load(config_file)
        self.directory = config['directory']
        self.output_csv_directory = config['output_csv_directory']
        self.output_data_directory = config['output_data_directory']
        self.pod_id = config['pod_id']
        self.meter_number = config['meter_number']