import yaml


class ConfigLoader:

    config_path = 'config.yaml'

    def __init__(self):
        with open(self.config_path, "r") as config_file:
            config = yaml.safe_load(config_file)
        self.directory = config['directory']
        self.output_csv_directory = config['output_csv_directory']
        self.output_data_directory = config['output_data_directory']
        self.pod_id = config['pod_id']
        self.meter_number = config['meter_number']
        self.previous_rates = config['previous_rates']

    @classmethod
    def write_data(cls, data: dict):
        with open(cls.config_path, 'r') as file:
            existing_data = yaml.safe_load(file)
        for key, value in data.items():
            existing_data[key] = value
        with open(cls.config_path, "w") as file:
            yaml.safe_dump(existing_data, file, default_flow_style=False)