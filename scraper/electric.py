import re
from datetime import datetime


class ElectricRegex:
    start_regex = ''
    end_regex = ''
    alt_name = ''

    kwh = ''
    rate = ''
    total = ''

    values = []

    def __init__(self,
                 name: str,
                 contents: str,
                 alt_name: str | None = None,
                 start_regex: str | None = None,
                 end_regex: str | None = None
    ):
        self.name = name
        self.contents = contents
        if start_regex is not None:
            self.start_regex = start_regex
        if end_regex is not None:
            self.end_regex = end_regex
        if alt_name is not None:
            self.alt_name = alt_name
        self._search(contents)

    def _search(self, contents: str):
        names = [self.name]
        regex = ''
        name_exists = False

        keys = ['kwh', 'rate', 'total']

        if self.alt_name:
            names.append(self.alt_name)
        for name in names:
            # colon_name = name + ':'
            # definitions end in colon, exclude these
            # highest index, lowest index is the definition
            if name in contents and contents.rfind(name) != contents.find(name):
                name_exists = True
            regex = f'{self.start_regex}{name}{self.end_regex}'
            match = re.search(regex, contents)
            if match:
                group_count = (re.compile(regex)).groups
                keys = keys[len(keys)-group_count:len(keys)]
                kwargs = dict(zip(keys, match.groups()))
                print(kwargs)
                self.set_values(**kwargs)
                return
            else:
                continue
        group_count = (re.compile(regex)).groups
        params = ['n/a' for _ in range(0, group_count)]
        self.set_values(*params)
        if name_exists:
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

        values = [self.kwh, self.rate, self.total]
        self.values = [v for v in values if v != '']

    def get_values_list(self):
        return self.values

    def get_values(self):
        if len(self.values) == 3:
            return f'{self.name}    {self.kwh}kwh    @    {self.rate}    - {self.total}'
        if len(self.values) == 2:
            return f'{self.name}    @    {self.rate}%    - {self.total}'
        if len(self.values) == 1:
            return f'{self.name}    - {self.total}'
        return ''

class ElectricBillDetails:

    basic_charge = None
    excess_generation = ''
    current_generation = ''
    prior_excess_generation = ''
    current_usage = ''
    net_usage_billed = ''
    charge = ''
    average_daily_use_current_year = ''
    average_daily_temp_current_year = ''
    average_daily_use_past_year = ''
    average_daily_temp_past_year = ''
    current_year = ''
    past_year = ''
    calculated_ppkwh = 'n/a'
    price_per_kwh = 'n/a'
    charges = []
    taxes = []
    misc_charges = []

    money_regex = '(-?\\d+\\.\\d{2})'
    small_money_regex = '(-?\\d\\.\\d{2})'
    amount_regex = f'\\${money_regex}'

    def __init__(self, contents: str, pod_id: str, meter_number: str, file_name: str):
        self.pod_id = pod_id
        self.meter_number = meter_number
        self.contents = contents
        self.file_name = file_name
        self._set_price_per_kwh()
        self._set_details()

    def _set_price_per_kwh(self) -> str:
        supply_str = "providing electricity supply during this billing period was"
        if supply_str in self.contents:
            ppkwh_str = self.contents.split(supply_str)[1]
            ppkwh_str = ppkwh_str.split('/kwh')[0]
            ppkwh_match = re.search(r'\$(\d+\.\d+)', ppkwh_str)
            if ppkwh_match:
                self.price_per_kwh = ppkwh_match.group(1)
            print('Actual ppkwh:', self.price_per_kwh)

    def _set_details(self):
        if 'PoD ID:' not in self.contents:
            print(self.contents)
            exit(1)
        service_from = self.contents.split('Service from:PoD ID:')[1].split(self.pod_id)[0]

        self._set_charges(service_from)
        self._set_taxes()
        self._set_misc_charges()
        self._set_generation_and_useage_info()
        self._set_average_info()

    def _get_start_end_month(self, service_from: str):
        start_date_split = service_from.split(' - ')
        start_date = datetime.strptime(start_date_split[0], '%m/%d/%y').date()
        end_date   = datetime.strptime(start_date_split[1], '%m/%d/%y').date()

        return start_date.strftime('%b'), end_date.strftime('%b')

    def _set_charges(self, service_from: str):
        start_month, end_month = self._get_start_end_month(service_from)
        supply_regex = '(\\d+)kwh(\\d+)\\s*?@-?0\\.' + self.money_regex
        revenue = 'Revenue decoupling mech'
        merchant = 'Merchant function charge'
        sbc_charge = 'SBC charge'
        delivery_charge = 'Delivery charge'
        transition_charge = 'Transition charge'

        end_regex = '(?: - )*?(-?\\d{1,3}\\.\\d{2})'
        basic_service_charge = 'Basic service charge'

        ec_start_regex = '.\\d{2}(\\d+)kwh(\\d+)\\s*?@-?0\\.'
        ec_end_regex = '(-?\\d{1,3}\\.\\d{2})'

        self.charges = [
            ElectricRegex(f'{basic_service_charge} - {start_month}', self.contents, alt_name=basic_service_charge, end_regex=end_regex),
            ElectricRegex(f'{basic_service_charge} - {end_month}', self.contents, alt_name=basic_service_charge, end_regex=end_regex),

            ElectricRegex(f'{delivery_charge} - {start_month}', self.contents, alt_name=delivery_charge, start_regex=ec_start_regex, end_regex=ec_end_regex),
            ElectricRegex(f'{delivery_charge} - {end_month}', self.contents, alt_name=delivery_charge, start_regex=ec_start_regex, end_regex=ec_end_regex),

            ElectricRegex(f'{transition_charge} - {start_month}', self.contents, alt_name=transition_charge, start_regex=ec_start_regex, end_regex=ec_end_regex),
            ElectricRegex(f'{transition_charge} - {end_month}', self.contents, alt_name=transition_charge, start_regex=ec_start_regex, end_regex=ec_end_regex),

            ElectricRegex(f'{revenue} - {start_month}', self.contents, alt_name=revenue, start_regex=ec_start_regex, end_regex=ec_end_regex),
            ElectricRegex(f'{revenue} - {end_month}', self.contents, alt_name=revenue, start_regex=ec_start_regex, end_regex=ec_end_regex),

            ElectricRegex(f'{sbc_charge} - {start_month}', self.contents, alt_name=sbc_charge, start_regex=ec_start_regex, end_regex=ec_end_regex),
            ElectricRegex(f'{sbc_charge} - {end_month}', self.contents, alt_name=sbc_charge, start_regex=ec_start_regex, end_regex=ec_end_regex),

            ElectricRegex('Subtotal Electricity Delivery', self.contents, end_regex=self.amount_regex),
            ElectricRegex('Supply charge', self.contents, end_regex=supply_regex),
            ElectricRegex(f'{merchant} - {start_month}', self.contents, end_regex=supply_regex, alt_name=merchant),
            ElectricRegex(f'{merchant} - {end_month}', self.contents, end_regex=supply_regex, alt_name=merchant),
            ElectricRegex('Subtotal Electricity Supply', self.contents, end_regex=self.amount_regex),
        ]
        second_charge = self.charges[1].get_values_list()[0]
        second_charge = float(second_charge) if second_charge != 'n/a' and second_charge != '' else 0.0
        basic_charge = self.charges[0].get_values_list()[0]
        self.basic_charge = float(basic_charge) + second_charge if basic_charge != '' else second_charge

    def _set_taxes(self):
        end_regex = '(-?\\d+\\.\\d+)\\s*?%@(-?\\d+\\.\\d{2})'
        self.taxes = [
            ElectricRegex('Taxes on delivery charges', self.contents, end_regex=end_regex),
            ElectricRegex('County sales tax', self.contents, end_regex=end_regex),
            ElectricRegex('Subtotal Electricity Taxes and Surcharges', self.contents, end_regex=self.amount_regex)
        ]

    def _set_misc_charges(self):
        space_regex = '\\s*?'
        space_number_regex = space_regex + '\\d'
        money_space_regex = self.small_money_regex + space_regex
        date_money_space_regex = '\\/\\d{2}' + money_space_regex

        self.misc_charges = [
            ElectricRegex('Payment & billing svcs charge', self.contents, start_regex=money_space_regex, end_regex=space_number_regex),
            ElectricRegex('Payment & billing svcs charge sales tax', self.contents, start_regex=date_money_space_regex, end_regex=space_number_regex),
            ElectricRegex('Payment & billing svcs GRT', self.contents, start_regex=date_money_space_regex, end_regex=space_number_regex),
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

        # If billed for electric, update the price per kwh charge
        # TODO when net_usage_billed is 0 maybe get previous estimation?
        if self.basic_charge is not None and self.basic_charge not in ['', 'n/a'] and int(self.net_usage_billed) > 0:
            self.calculated_ppkwh = (float(self.charge) - float(self.basic_charge)) / float(self.net_usage_billed)
        else:
            print('Basic Charge: ', self.basic_charge)
        print('Price per kwh: ' + str(self.calculated_ppkwh))

    def get_details(self):
        charge_str = '\\n'.join([c.get_values() for c in self.charges])
        taxes_str = '\\n'.join([t.get_values() for t in self.taxes])
        misc_str = '\\n'.join([m.get_values() for m in self.misc_charges])

        return f'{self.calculated_ppkwh}\\n{self.basic_charge.get_values()}\\n{charge_str}\\n{taxes_str}\\n{misc_str}\\n'

    def get_values_list(self):
        theoretical_charge = ''
        if self.basic_charge != '' and self.calculated_ppkwh != 'n/a':
            theoretical_charge = (float(self.current_usage) * float(self.calculated_ppkwh)) + float(self.basic_charge)
        values_list = [
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
            theoretical_charge - float(self.charge) if theoretical_charge != '' else '',
            self.calculated_ppkwh,
            self.price_per_kwh
        ]
        for c in self.charges:
            values_list.extend(c.get_values_list())
        tax_values = []
        for t in self.taxes:
            tax_values.extend(t.get_values_list())
        misc_values = []
        for m in self.misc_charges:
            misc_values.extend(m.get_values_list())

        values_list.extend(tax_values)
        values_list.extend(misc_values)

        return values_list

    def get_str_values_list(self):
        values_list = self.get_values_list()
        return [str(v) for v in values_list]