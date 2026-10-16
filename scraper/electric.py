import re
from datetime import datetime


class ElectricRegex:

    def __init__(self,
                 name: str,
                 start_regex: str | None = None,
                 end_regex: str | None = None,
                 start_month: str | None = None,
                 end_month: str | None = None
    ):
        self.start_regex = ''
        self.end_regex = ''
        self.start_month = ''
        self.end_month = ''
        self.kwh = ''
        self.rate = ''
        self.total = ''
        self.name = name
        if start_regex is not None:
            self.start_regex = start_regex
        if end_regex is not None:
            self.end_regex = end_regex
        if start_month is not None:
            self.start_month = start_month
        if end_month is not None:
            self.end_month = end_month
        self.values = []
        self.multi_values = []

    def search(self, contents: str):
        names = [self.name]
        if self.start_month and self.end_month:
            names.append(f'{self.name} - {self.start_month}')
            names.append(f'{self.name} - {self.end_month}')
        regex = ''
        name_exists = False
        find_match = False

        keys = ['kwh', 'rate', 'total']

        for index, name in enumerate(names):
            # colon_name = name + ':'
            # definitions end in colon, exclude these
            # highest index, lowest index is the definition
            if name in contents and contents.rfind(name) != contents.find(name):
                name_exists = True
            regex = f'{self.start_regex}{name}{self.end_regex}'
            match = re.search(regex, contents)
            if match:
                find_match = True
                group_count = (re.compile(regex)).groups
                keys = keys[len(keys)-group_count:len(keys)]
                kwargs = dict(zip(keys, match.groups()))
                self.set_values(**kwargs)
                if index > 0:
                    self.multi_values.append(self.values)
                if len(names) == 1:
                    return
            else:
                continue
        # TODO clean up
        if not find_match:
            group_count = (re.compile(regex)).groups
            params = ['n/a' for _ in range(0, group_count)]
            self.set_values(*params)
        if self.start_month and self.end_month:
            if not self.multi_values:
                self.multi_values.append(self.values)
                self.multi_values.append(['n/a' for _ in range(0, group_count)])
        if name_exists and not find_match:
            # name is in the contents but the regex failed to match
            print(regex)
            print(self.name)
            print(contents)
            exit(1)

    def set_values(self, kwh: str | None = None, rate: str | None = None, total: str | None = None):
        self.kwh = kwh.strip() if kwh is not None else ''
        self.rate = rate.strip() if rate is not None else ''

        if self.rate != '' and self.rate != 'n/a' and '.' not in self.rate:
            self.rate = f'0.{self.rate}'
        self.total = total.strip() if total is not None else ''
        if self.total.startswith('-'):
            self.rate = f'-{self.rate}'

        values = [self.kwh, self.rate, self.total]
        self.values = [v for v in values if v != '']


    def get_values_list(self):
        return self.values if not self.multi_values else self.multi_values

    def get_values(self):
        if len(self.values) == 3:
            return f'{self.name}    {self.kwh}kwh    @    {self.rate}    - {self.total}'
        if len(self.values) == 2:
            return f'{self.name}    @    {self.rate}%    - {self.total}'
        if len(self.values) == 1:
            return f'{self.name}    - {self.total}'
        return ''

class ElectricBillDetails:

    previous_dc_rate = 0.0
    previous_t_rate = 0.0
    previous_r_rate = 0.0
    previous_s_rate = 0.0
    previous_rc_rate = 0.0
    previous_sc_rate = 0.0
    previous_mc_rate = 0.0
    previous_delivery_tax_rate = 0.0
    previous_county_tax_rate = 0.0

    def __init__(self, contents: str, pod_id: str, meter_number: str, file_name: str, previous_rates: dict):
        self.basic_charge = None
        self.excess_generation = ''
        self.current_generation = ''
        self.prior_excess_generation = ''
        self.current_usage = ''
        self.net_usage_billed = ''
        self.charge = ''
        self.average_daily_use_current_year = ''
        self.average_daily_temp_current_year = ''
        self.average_daily_use_past_year = ''
        self.average_daily_temp_past_year = ''
        self.current_year = ''
        self.past_year = ''
        self.calculated_ppkwh = 'n/a' # TODO deprecate
        self.price_per_kwh = 'n/a'
        self.charges = []
        self.pod_id = pod_id
        self.meter_number = meter_number
        self.contents = contents
        self.file_name = file_name
        self._set_price_per_kwh()
        self._set_details()
        self.set_previous_rates(**previous_rates)
        self.sum_all_rates = 0.0

    def _set_price_per_kwh(self) -> str:
        supply_str = "providing electricity supply during this billing period was"
        if supply_str in self.contents:
            ppkwh_str = self.contents.split(supply_str)[1]
            ppkwh_str = ppkwh_str.split('/kwh')[0]
            ppkwh_match = re.search(r'\$(\d+\.\d+)', ppkwh_str)
            if ppkwh_match:
                self.price_per_kwh = ppkwh_match.group(1)

    def _set_details(self):
        if 'PoD ID:' not in self.contents:
            print(self.contents)
            exit(1)
        service_from = self.contents.split('Service from:PoD ID:')[1].split(self.pod_id)[0]

        self._set_charges(service_from)
        self._set_generation_and_useage_info()
        self._set_average_info()

    def _get_start_end_month(self, service_from: str):
        start_date_split = service_from.split(' - ')
        start_date = datetime.strptime(start_date_split[0], '%m/%d/%y').date()
        end_date   = datetime.strptime(start_date_split[1], '%m/%d/%y').date()

        return start_date.strftime('%b'), end_date.strftime('%b')

    def _set_charges(self, service_from: str):

        money_regex = '(-?\\d+\\.\\d{2})'
        small_money_regex = '(-?\\d\\.\\d{2})'
        amount_regex = f'\\${money_regex}'

        start_month, end_month = self._get_start_end_month(service_from)
        supply_regex = '(\\d+)kwh(\\d+)\\s*?@-?0\\.' + money_regex

        end_regex = '(?: - )*?(-?\\d{1,3}\\.\\d{2})'

        ec_start_regex = '.\\d{2}(\\d+)kwh(\\d+)\\s*?@-?0\\.'
        ec_end_regex = '(-?\\d{1,3}\\.\\d{2})'

        tax_end_regex = '(-?\\d+\\.\\d+)\\s*?%@(-?\\d+\\.\\d{2})'

        space_regex = '\\s*?'
        space_number_regex = space_regex + '\\d'
        money_space_regex = small_money_regex + space_regex
        date_money_space_regex = '\\/\\d{2}' + money_space_regex

        self.charges = [
            ElectricRegex('Basic service charge', end_regex=end_regex, start_month=start_month, end_month=end_month),
            ElectricRegex('Delivery charge', start_regex=ec_start_regex, end_regex=ec_end_regex, start_month=start_month, end_month=end_month),
            ElectricRegex('Transition charge', start_regex=ec_start_regex, end_regex=ec_end_regex, start_month=start_month, end_month=end_month),
            ElectricRegex('Revenue decoupling mech', start_regex=ec_start_regex, end_regex=ec_end_regex, start_month=start_month, end_month=end_month),
            ElectricRegex('SBC charge', start_regex=ec_start_regex, end_regex=ec_end_regex, start_month=start_month, end_month=end_month),
            ElectricRegex('Recovery charge', start_regex=ec_start_regex, end_regex=ec_end_regex, start_month=start_month, end_month=end_month),
            ElectricRegex('Subtotal Electricity Delivery', end_regex=amount_regex),
            ElectricRegex('Supply charge', end_regex=supply_regex),
            ElectricRegex('Merchant function charge', end_regex=supply_regex, start_month=start_month, end_month=end_month),
            ElectricRegex('Subtotal Electricity Supply', end_regex=amount_regex),

            ElectricRegex('Taxes on delivery charges', end_regex=tax_end_regex),
            ElectricRegex('County sales tax', end_regex=tax_end_regex),
            ElectricRegex('Subtotal Electricity Taxes and Surcharges', end_regex=amount_regex),

            ElectricRegex('Payment & billing svcs charge', start_regex=money_space_regex, end_regex=space_number_regex),
            ElectricRegex('Payment & billing svcs charge sales tax', start_regex=date_money_space_regex, end_regex=space_number_regex),
            ElectricRegex('Payment & billing svcs GRT', start_regex=date_money_space_regex, end_regex=space_number_regex),
        ]

    def _set_generation_and_useage_info(self):
        generation_and_billing_info = self.contents.split('ExcessGenerationCurrent UsageMeter Number')[1].split(self.meter_number)[0].split(' kwh')

        self.excess_generation = generation_and_billing_info[0]
        self.current_generation = generation_and_billing_info[1]
        self.prior_excess_generation = generation_and_billing_info[2]
        self.current_usage = generation_and_billing_info[3]
        self.net_usage_billed = generation_and_billing_info[4]

        self.service_from = self.contents.split('Service from:PoD ID:')[1].split(self.pod_id)[0]
        dates = self.service_from.split(' - ')
        d1 = datetime.strptime(dates[0], "%m/%d/%y")
        d2 = datetime.strptime(dates[1], "%m/%d/%y")
        self.days = abs((d2 - d1).days) + 1

        # Total Energy Charges
        charge_split = self.contents.split('Total Energy Charges')[1].split('BillingPeriodAverageDaily')[0].split('$')[1]
        charge_match = re.search('\\d+\\.\\d+', charge_split)
        self.charge = charge_match.group() if charge_match is not None else '0.0'

    def _set_average_info(self):
        # Average Info
        if 'Miscellaneous Charges' not in self.contents:
            print('Could not find misc charges')
        useage_chart = self.contents.split('Usage Chart Information')[1]
        average_info = useage_chart.split('Miscellaneous Charges-')[0].split('kwh')
        # 14 ,59  F26 ,62  FOct-20Oct-19
        self.average_daily_use_current_year = average_info[0].strip()
        self.average_daily_temp_current_year = average_info[1].split(' ')[0].strip().replace('°', '')
        self.average_daily_use_past_year = average_info[1].split('F')[1].strip()
        self.average_daily_temp_past_year = average_info[2].split(' F')[0]
        self.current_year = average_info[2].split(' F')[1][:6]
        self.past_year = average_info[2].split(' F')[1][6:12]

    def get_values_list(self):
        values_list = []
        for index, c in enumerate(self.charges):
            c.search(self.contents)
            values = c.get_values_list()
            # Currently only charges can return an array of arrays
            if isinstance(values[0], list):
                for value in values:
                    values_list.extend(value)
            else:
                values_list.extend(values)
            if index == 0:
                second_charge = values_list[1]
                second_charge = float(second_charge) if second_charge != 'n/a' and second_charge != '' else 0.0
                basic_charge = values_list[0]
                self.basic_charge = float(basic_charge) + second_charge if basic_charge != 'n/a' else second_charge

        # Note theoretical charge doesn't account for solar that was generated and immediately used
        # Savings would be more than this calculated number
        # if multiple rates, average them
        theoretical_charge = self.calculate_theoretical_charge(values_list)

        values_list_prefix = [
            self.current_year,
            int(self.excess_generation),
            int(self.current_generation),
            int(self.current_usage),
            int(self.net_usage_billed),
            self.charge,
            self.service_from,
            self.days,
            self.average_daily_temp_current_year,
            self.average_daily_use_current_year,
            self.basic_charge,
            theoretical_charge,
            theoretical_charge - float(self.charge),
            self.calculated_ppkwh,
            self.price_per_kwh,
        ]
        values_list.append(self.sum_all_rates)
        return values_list_prefix + values_list

    def get_str_values_list(self):
        values_list = self.get_values_list()
        return [str(v) for v in values_list]

    @classmethod
    def get_previous_rates(cls):
        return (cls.previous_dc_rate, cls.previous_t_rate, cls.previous_r_rate,
                cls.previous_s_rate, cls.previous_rc_rate, cls.previous_sc_rate,
                cls.previous_mc_rate, cls.previous_delivery_tax_rate,
                cls.previous_county_tax_rate)

    @classmethod
    def get_previous_rates_dict(cls):
        return {
            'dc_rate': cls.previous_dc_rate,
            't_rate': cls.previous_t_rate,
            'r_rate': cls.previous_r_rate,
            's_rate': cls.previous_s_rate,
            'rc_rate': cls.previous_rc_rate,
            'sc_rate': cls.previous_sc_rate,
            'mc_rate': cls.previous_mc_rate,
            'delivery_tax_rate': cls.previous_delivery_tax_rate,
            'county_tax_rate': cls.previous_county_tax_rate,
        }

    @classmethod
    def set_previous_rates(cls, dc_rate: float, t_rate: float, r_rate: float,
                           s_rate: float, rc_rate: float, sc_rate: float,
                           mc_rate: float, delivery_tax_rate: float, county_tax_rate: float):
        cls.previous_dc_rate = dc_rate
        cls.previous_t_rate = t_rate
        cls.previous_r_rate = r_rate
        cls.previous_s_rate = s_rate
        cls.previous_rc_rate = rc_rate
        cls.previous_sc_rate = sc_rate
        cls.previous_mc_rate = mc_rate
        cls.previous_delivery_tax_rate = delivery_tax_rate
        cls.previous_county_tax_rate = county_tax_rate

    def calculate_theoretical_charge(self, values_list: list) -> float:
        dc_rate, t_rate, r_rate, s_rate, rc_rate, sc_rate, mc_rate, delivery_tax_rate, county_tax_rate = self.get_previous_rates()

        if self.price_per_kwh != 'n/a':
            def get_value(index: int):
                return float(values_list[index]) if values_list[index] != 'n/a' else 0.0

            def average_rates_or_single(rate_1: float, rate_2: float):
                return rate_1 if rate_2 == 0.0 else (rate_1 + rate_2) / 2.0

            dc_rate = average_rates_or_single(get_value(3), get_value(6))
            t_rate = average_rates_or_single(get_value(9), get_value(12))
            r_rate = average_rates_or_single(get_value(15), get_value(18))
            s_rate = average_rates_or_single(get_value(21), get_value(24))
            rc_rate = average_rates_or_single(get_value(27), get_value(30))
            sc_rate = get_value(34)
            mc_rate = average_rates_or_single(get_value(37), get_value(40))
            delivery_tax_rate = get_value(43) / 100.0
            county_tax_rate = get_value(46) / 100.0
            self.set_previous_rates(
                dc_rate, t_rate, r_rate, s_rate, rc_rate, sc_rate,
                mc_rate, delivery_tax_rate, county_tax_rate,
            )

        total_useage = float(self.current_usage)
        delivery_charge = ((total_useage * dc_rate) + (total_useage * t_rate) +
                           (total_useage * r_rate) + (total_useage * s_rate) + (total_useage * rc_rate))

        supply_charge = (total_useage * sc_rate) + (total_useage * mc_rate)

        delivery_taxes = delivery_charge * delivery_tax_rate
        county_taxes = (self.basic_charge + delivery_taxes + supply_charge + delivery_charge) * county_tax_rate
        tax_total = delivery_taxes + county_taxes

        self.sum_all_rates = dc_rate + t_rate + r_rate + s_rate + rc_rate + sc_rate + mc_rate

        return self.basic_charge + delivery_charge + supply_charge + tax_total

