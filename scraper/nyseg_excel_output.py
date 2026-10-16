import xlsxwriter

class NysegExcelOutput:

    # TODO move some of these to a different sheet
    columns = [
        'Month',
        'Prior Excess Generation (kwh)',
        'NYSEG Contribution to Grid (kwh)',
        'NYSEG Useage (kwh)',
        'NYSEG Useage Billed',
        'NYSEG Billed Amount',
        'NYSEG Billing Cycle',
        'Billing Days',
        'Avg Daily Tmp',
        'Avg Daily Useage',
        'Basic Service Charge Total',
        'Theoretical Charge',
        'Theoretical Savings',
        'Calculated Price Per KWH',
        'Actual Price Per KWH',
        # Charges
        'Basic Service Charge 1st month',
        'Basic Service Charge 2nd month',
        'Delivery charge KWH 1st month',
        'Delivery charge Rate 1st month',
        'Delivery charge Total 1st month',
        'Delivery charge KWH 2nd month',
        'Delivery charge Rate 2nd month',
        'Delivery charge Total 2nd month',
        'Transition charge KWH 1st month',
        'Transition charge Rate 1st month',
        'Transition charge Total 1st month',
        'Transition charge KWH 2nd month',
        'Transition charge Rate 2nd month',
        'Transition charge Total 2nd month',
        'Revenue decoupling mech 1st month KWH',
        'Revenue decoupling mech 1st month Rate',
        'Revenue decoupling mech 1st month Total',
        'Revenue decoupling mech 2nd month KWH',
        'Revenue decoupling mech 2nd month Rate',
        'Revenue decoupling mech 2nd month Total',
        'SBC charge 1st month KWH',
        'SBC charge 1st month Rate',
        'SBC charge 1st month Total',
        'SBC charge 2nd month KWH',
        'SBC charge 2nd month Rate',
        'SBC charge 2nd month Total',
        'Recovery charge 1st month KWH',
        'Recovery charge 1st month Rate',
        'Recovery charge 1st month Total',
        'Recovery charge 2nd month KWH',
        'Recovery charge 2nd month Rate',
        'Recovery charge 2nd month Total',
        'Subtotal Electricity Delivery',
        'Supply charge KWH',
        'Supply charge Rate',
        'Supply charge Total',
        'Merchant charge 1st month KWH',
        'Merchant charge 1st month Rate',
        'Merchant charge 1st month Total',
        'Merchant charge 2nd month KWH',
        'Merchant charge 2nd month Rate',
        'Merchant charge 2nd month Total',
        'Subtotal Electricity Supply',
        'Taxes on delivery charges Rate',
        'Taxes on delivery charges Total',
        'County sales tax Rate',
        'County sales tax Total',
        'Subtotal Electricity Taxes and Surcharges',
        'Payment & billing svcs charge',
        'Payment & billing svcs charge sales tax',
        'Payment & billing svcs GRT',
        'Sum of All Rates',
    ]

    num_columns = len(columns)
    current_row = 0
    column_lengths = [len(c) for c in columns]
    start_month = None
    start_year = None
    current_column = 0

    def __init__(self, output_directory: str):
        self.output_directory = output_directory
        self.workbook = xlsxwriter.Workbook(self.output_directory + '/solar panels NEW 2.xlsx')

        self.red_format = self.workbook.add_format({'bg_color': 'red'})
        self.green_format = self.workbook.add_format({'bg_color': 'green'})
        self.yellow_format = self.workbook.add_format({'bg_color': 'yellow'})
        self.blue_format = self.workbook.add_format({'bg_color': 'blue', 'font_color': 'white'})

        self.currency_format = self.workbook.add_format({'num_format': '[$$-409]#,##0.00'})
        self.rate_format = self.workbook.add_format({'num_format': '[$$-409]#,##0.00000000'})

        self.column_format = [
            None,
            self.blue_format,
            self.green_format,
            self.red_format,
            self.yellow_format,
            self.currency_format,
            None,
            None,
            None,
            None,
            self.currency_format,
            self.currency_format,
            self.currency_format,
            self.currency_format,
            self.rate_format,
        ]

        self.worksheet = self.workbook.add_worksheet()
        self._write_headers()

    def _write_headers(self):
        self.write_output(self.columns)
        column_index = 0
        for column_length in self.column_lengths:
            self.worksheet.set_column(self.current_row, column_index, column_length)
            column_index += 1

    def write_output(self, data: list):
        column_index = 0
        while len(self.column_format) < len(data):
            self.column_format.append(None)
        for datum, formatting in zip(data, self.column_format):
            if self.current_row == 1 and column_index == 0:
                split = datum.split('-')
                self.start_month = split[0]
                self.start_year = split[1]
            if datum == 'n/a':
                datum = '=NA()'
            try:
                if '$' in datum:
                    datum.replace('$', '')
                datum = float(datum)
            except:
                pass
            if formatting is None:
                self.worksheet.write(self.current_row, column_index, datum)
            else:
                self.worksheet.write(self.current_row, column_index, datum, formatting)
            column_index += 1
        self.current_row += 1

    def add_totals(self):
        pass
        # TODO insert function
        # self.worksheet.write(self.current_row + 1, 11, sum(theoretical_charges), currency_format)
        # self.worksheet.write(self.current_row + 1, 12, sum(theoretical_savings), currency_format)

    def make_chart(self):
        if self.start_month is None:
            raise Exception('Must call write_output prior to make_chart')

        chart = self.workbook.add_chart({'type': 'line'})
        last_data_row = str(self.current_row)

        if chart:
            chart.set_style(48)
            chart.set_x_axis({
                'name': 'KWH per Month',
            })
            x_axis_label = f'=Sheet1!$A2:$A{last_data_row}'
            columns = ['B', 'C', 'D', 'E']
            colors = ['blue', 'green', 'red', 'yellow']
            names = ['Prior Excess Generation', 'Contribution', 'Useage', 'Useage Billed']

            for column, color, name in zip(columns, colors, names):
                chart.add_series(
                    {'categories': x_axis_label, 'values': f'=Sheet1!${column}$2:${column}${last_data_row}',
                     'line': {'color': color},
                     'name': name})
            self.worksheet.insert_chart(self.current_row + 3, 0, chart)
        self.make_comparative_charts()

    def make_comparative_charts(self):
        charts = ['KWH per Month', 'Bill Cost Per Month', 'Sum of All Rates', 'Cost Per Killowatt Hour', 'Delivery Charge Rate', 'Transition Charge Rate', 'Average Daily Temp', 'Average Daily Use KWH']
        data_columns = ['D', 'F', 'BO', 'O', 'S', 'Y', 'I', 'J']
        for chart, column in zip(charts, data_columns):
            self.make_comparative_chart(chart, column)

    def make_comparative_chart(self, label: str, column: str):
        from datetime import datetime

        comparative_chart = self.workbook.add_chart({'type': 'line'})

        if not comparative_chart:
            raise Exception('Error creating chart')

        months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
        if self.start_month not in months:
            raise Exception('Not a valid start month')

        date_obj = datetime.strptime(self.start_month, "%b")

        data_index_start = 2
        data_index_end = 13
        start_month_int = date_obj.month

        did_pad = False
        pad_data_row = self.current_row + 1000
        pad_data_row_start = pad_data_row
        pad_data_row_end = pad_data_row
        pad_data_row = pad_data_row_start

        while start_month_int != 1:
            did_pad = True
            self.worksheet.write(pad_data_row, 0, '=NA()')
            pad_data_row_end = pad_data_row
            data_index_end -= 1
            start_month_int -= 1
            pad_data_row += 1

        comparative_chart.set_style(48)
        comparative_chart.set_x_axis({
            'name': label,
        })

        year = int('20' + self.start_year)
        year_count = int(self.current_row / 12)
        year_mod = self.current_row % 12

        if year_mod != 0:
            year_count += 1
        color_index = 0

        self.worksheet.write_column(0, 67, months)
        colors = self.generate_contrasting_colors(year_count)
        for i in range(0, year_count):
            if color_index + 1 > len(colors):
                color_index = 0

            prefix = ''
            if i == 0 and did_pad:
                prefix = f'Sheet1!$A${pad_data_row_start + 1}:$A${pad_data_row_end + 1},'
            data = f'=({prefix}Sheet1!${column}${data_index_start}:${column}${data_index_end})'

            comparative_chart.add_series(
                {'categories': f'=Sheet1!BP$1:$BP12', 'values': data,
                 'line': {'color': colors[color_index]},
                 'name': f'{year}'})
            data_index_start = data_index_end + 1
            data_index_end = data_index_start + 11
            year += 1
            color_index += 1
        self.worksheet.insert_chart(self.current_row + 20, self.current_column, comparative_chart)
        self.current_column += 5

    def generate_contrasting_colors(self, n: int, saturation: float = 0.7, lightness: float = 0.4) -> list[str]:
        """
        Generates 'n' visually distinct hex colors by evenly spacing hues.

        :param n: Number of colors to generate.
        :param saturation: Color intensity (0.0 to 1.0). Default 0.7 for vibrant colors.
        :param lightness: Color brightness (0.0 to 1.0). Default 0.5 for balanced contrast.
        :return: A list of hex color strings.
        """
        import colorsys
        if n <= 0:
            return []

        hex_colors = []
        for i in range(n):
            # Calculate hue by dividing the 360-degree wheel into 'n' equal parts
            hue = i / n

            # Convert HSL to RGB (values range from 0.0 to 1.0)
            r, g, b = colorsys.hls_to_rgb(hue, lightness, saturation)

            # Convert RGB floats to 0-255 integers and format as Hex
            hex_color = f"#{int(r * 255):02X}{int(g * 255):02X}{int(b * 255):02X}"
            hex_colors.append(hex_color)

        return hex_colors

    def close_file(self):
        self.workbook.close()