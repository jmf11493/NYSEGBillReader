import os

from config_loader import ConfigLoader
from electric import ElectricBillDetails
from nyseg_excel_output import NysegExcelOutput
from pdf_reader import PDFReader


class NysegReader:

    def __init__(self, config_loader: ConfigLoader):
        self.config_loader = config_loader

    def parse_all_files(self):
        directory = self.config_loader.directory
        list_bill_details = []
        list_str_details = self._load_data_file()

        for filename in os.listdir(directory):
            if not filename.endswith('.pdf'):
                continue
            if list_str_details and filename in list_str_details.keys():
                print('Skipping file: ' + filename)
                continue
            path = os.path.join(directory, filename)
            self._fix_pdf(path)
            try:
                pdf_reader = PDFReader(path)
            except:
                print('Skipping file: ' + path)
                continue
            contents = self._get_page_contents(pdf_reader, 0, 3)
            bill_details = ElectricBillDetails(
                contents,
                self.config_loader.pod_id,
                self.config_loader.meter_number,
                filename,
                self.config_loader.previous_rates,
            )
            list_bill_details.append(bill_details)
        for details in list_bill_details:
            print('Adding additional data')
            list_str_details[details.file_name] = details.get_str_values_list()
        ConfigLoader.write_data({'previous_rates': ElectricBillDetails.get_previous_rates_dict()})
        self._write_excel_file(list_str_details)
        self._write_csv_file(list_str_details)
        self._write_data_file(list_str_details)

    def _write_excel_file(self, list_bill_details: dict):
        directory = self.config_loader.directory
        excel_file = NysegExcelOutput(directory)

        for bill_details in list_bill_details.values():
            excel_file.write_output(bill_details)
        excel_file.make_chart()
        excel_file.close_file()

    def _write_csv_file(self, list_bill_details: dict):
        directory = self.config_loader.output_csv_directory

        if os.path.exists(directory):
            os.remove(directory)
        csv_file = open(directory, "a")

        for file_name, original_list in list_bill_details.items():
            # TODO don't re-order
            write_list = [file_name, original_list[6], original_list[7]]
            write_list.extend(original_list[0:6])
            write_list.extend(original_list[8:])
            csv_file.write(','.join(write_list) + '\n')

        csv_file.close()

    def _write_data_file(self, list_bill_details: dict):
        # TODO when csv file and this file are in sync can clean up other method
        directory = self.config_loader.output_data_directory

        if os.path.exists(directory):
            os.remove(directory)
        csv_file = open(directory, "a")
        for file_name, bill_details in list_bill_details.items():
            write_list = [file_name]
            write_list.extend(bill_details)
            csv_file.write(','.join(write_list) + '\n')
        csv_file.close()

    def _load_data_file(self) -> dict:
        directory = self.config_loader.output_data_directory
        data = {}
        if os.path.exists(directory):
            with open(directory, "r") as data_file:
                for line in data_file:
                    values = line.split(',')
                    values = [v.replace('\n', '') for v in values]
                    file_name = values.pop(0)
                    data[file_name] = values
            return data
        return data

    def _get_page_contents(self, pdf_reader, page_start: int, page_end: int) -> str:
        contents = ''
        inclusive_end = page_end + 1
        while page_start != inclusive_end:
            contents = contents + pdf_reader.get_page_contents(page_start)
            page_start += 1
        return contents

    def _fix_pdf(self, path: str):
        EOF_MARKER = b'%%EOF'
        with open(path, 'rb') as file:
            contents = file.read()
        # TODO? Just remove everything after the EOF, remove this condition
        if b'<!DOCTYPE html>' in contents:
            remove2 = contents.split(EOF_MARKER)[1]
            contents = contents.replace(remove2, b'')
            with open(path, 'wb') as file:
                file.write(contents)
                file.close()
