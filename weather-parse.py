"""
    This application was written by Avery Uslaner to assist Red Butte
    Garden with downloading weather data collected by a
    wunderground.com compatible home weather station.

    The station code for this RBG weather station is: KAZLITTL3
    It was installed on 10/13/2016
"""
from datetime import date, timedelta
import json
import xml.etree.ElementTree as ET
import tkinter as tk
from urllib.request import urlretrieve

import openpyxl
import pandas as pd

from calendar_widget import Calendar


class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.grid(column=0, row=0, sticky=(tk.N, tk.W, tk.E, tk.S))
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        self.station_code = tk.StringVar()
        self.start_date = {}
        self.end_date = {}
        self.wb = openpyxl.Workbook()

        # Set default end date to today's date
        end_day, end_month, end_year = date.today().strftime('%d %m %Y').split(' ')
        self.end_date['day_selected'] = end_day
        self.end_date['month_selected'] = end_month
        self.end_date['year_selected'] = end_year

        self.day_start = tk.Button(
            self, text='Choose',
            command=lambda: self.datepicker(side='start')).grid(column=2, row=1, sticky=(tk.W, tk.E))
        self.day_end = tk.Button(
            self, text='Choose',
            command=lambda: self.datepicker(side='end')).grid(column=2, row=2, sticky=(tk.W, tk.E))

        self.station_code_entry = tk.Entry(self, width=7, textvariable=self.station_code)
        self.station_code_entry.grid(column=2, row=3, sticky=(tk.W, tk.E))

        self.read_values()

        self.save_defaults = tk.Button(
            self, text='Save selected values', command=self.save_values).grid(column=1, row=4, sticky=tk.E)
        self.download = tk.Button(
            self, text='Download Data', command=self.grab_history).grid(column=3, row=4, sticky=tk.W)
        # self.show_btn = tk.Button(self, text='Dump data', command=self.dump_data)

        tk.Label(self, text='Station Code').grid(column=3, row=3, sticky=tk.W)
        tk.Label(self, text='Start Date').grid(column=1, row=1, sticky=tk.E)
        tk.Label(self, text='End Date').grid(column=1, row=2, sticky=tk.E)
        tk.Label(self, text='(defaults to today)').grid(column=3, row=2, sticky=tk.W)

        for child in self.winfo_children():
            child.grid_configure(padx=5, pady=5)

        self.station_code_entry.focus()

    def alert(self, message):
        child = tk.Toplevel()
        child.wm_title('Alert!')
        msg = tk.Message(child, text=message)
        msg.pack()
        button = tk.Button(child, text='Dismiss', command=child.destroy)
        button.pack()

    def datepicker(self, side):
        """
        Launch a child window with a calendar widget.
        :param side: Specifies if the child window is recording a start
        or end date.
        """
        child = tk.Toplevel()
        child.wm_title('Choose date')
        if side == 'start':
            Calendar(child, values=self.start_date)
        else:
            Calendar(child, values=self.end_date)

    def dump_data(self):
        print('Start date:', self.start_date)
        print('End date:', self.end_date)

    def grab_history(self):
        if len(self.start_date) == 0:
            message = "No start date selected! Can't download data."
            self.alert(message)
            return
        else:
            start_day = str(self.start_date['day_selected'])
            start_month = str(self.start_date['month_selected'])
            start_year = str(self.start_date['year_selected'])

        end_day = str(self.end_date['day_selected'])
        end_month = str(self.end_date['month_selected'])
        end_year = str(self.end_date['year_selected'])

        station_id = self.station_code_entry.get()
        url = 'https://www.wunderground.com/weatherstation/WXDailyHistory.asp?' \
              'ID=' + station_id + '&day=' + start_day + '&month=' + start_month + '&year=' + start_year + \
              '&dayend=' + end_day + '&monthend=' + end_month + '&yearend=' + end_year + \
              '&graphspan=custom&format=1'
        print('Getting daily data from {}/{}/{} to {}/{}/{}.'.format(
            start_month, start_day, start_year, end_month, end_day, end_year
        ))
        df = pd.read_csv(url, header=0, index_col=0)
        df.rename(columns={'PrecipitationSumIn<br>': 'Precipitation Sum (in)'}, inplace=True)
        df = df[df.index != '<br>']
        df.to_csv('updated-' + end_day + '-' + end_month + '-' + end_year + '-daily-' + station_id + '-weather.csv')

        print('Getting 5 minute data from {}/{}/{} to {}/{}/{}.'.format(
            start_month, start_day, start_year, end_month, end_day, end_year
        ))
        start_date = date(self.start_date['year_selected'], self.start_date['month_selected'],
                          self.start_date['day_selected'])
        end_date = date(int(self.end_date['year_selected']), int(self.end_date['month_selected']),
                        int(self.end_date['day_selected']))
        self.set_headers()
        for single_date in daterange(start_date, end_date):
            day, month, year = single_date.strftime("%d %m %Y").split(' ')
            self.parse_day(day, month, year)
        message = 'Downloaded daily data to csv and 5 minute data to xlsx!'
        self.alert(message)

    def parse_day(self, day, month, year):
        station_id = self.station_code_entry.get()
        url = 'https://www.wunderground.com/weatherstation/WXDailyHistory.asp?' \
              'ID=' + station_id + '&day=' + day + '&month=' + month + '&year=' + year + '&graphspan=day&format=XML'
        print(url)
        xmlfile, headers = urlretrieve(url)
        tree = ET.parse(xmlfile)
        tree_root = tree.getroot()

        wb = self.wb
        sheet = wb.active

        for idx, observation in enumerate(tree_root):
            observation_time = observation.find('observation_time').text

            location = observation.find('location')
            location_full = location[0].text
            location_neighborhood = location[1].text
            location_city = location[2].text
            location_state = location[3].text
            location_zip = location[4].text
            location_latitude = location[5].text
            location_longitude = location[6].text
            location_elevation = location[7].text

            temp_f = observation.find('temp_f').text
            temp_c = observation.find('temp_c').text
            relative_humidity = observation.find('relative_humidity').text
            wind = observation.find('wind_string').text
            wind_dir = observation.find('wind_dir').text
            wind_degrees = observation.find('wind_degrees').text
            wind_mph = observation.find('wind_mph').text
            wind_gust_mph = observation.find('wind_gust_mph').text
            pressure_string = observation.find('pressure_string').text
            pressure_mb = observation.find('pressure_mb').text
            pressure_in = observation.find('pressure_in').text
            dewpoint_string = observation.find('dewpoint_string').text
            dewpoint_f = observation.find('dewpoint_f').text
            dewpoint_c = observation.find('dewpoint_c').text
            solar_radiation = observation.find('solar_radiation').text
            uv = observation.find('UV').text
            precip_1hr_string = observation.find('precip_1hr_string').text
            precip_1hr_in = observation.find('precip_1hr_in').text
            precip_1hr_metric = observation.find('precip_1hr_metric').text
            precip_today_string = observation.find('precip_today_string').text
            precip_today_in = observation.find('precip_today_in').text
            precip_today_metric = observation.find('precip_today_metric').text

            sheet.append([observation_time, location_full, location_neighborhood, location_city, location_state,
                         location_zip, location_latitude, location_longitude, location_elevation, temp_f, temp_c,
                         relative_humidity, wind, wind_dir, wind_degrees, wind_mph, wind_gust_mph, pressure_string,
                         pressure_mb, pressure_in, dewpoint_string, dewpoint_f, dewpoint_c, solar_radiation, uv,
                         precip_1hr_string, precip_1hr_in, precip_1hr_metric, precip_today_string, precip_today_in,
                         precip_today_metric])

        wb.save('5-min-' + station_id + '-weather.xlsx')

    def read_values(self):
        """
        Reads default values from file if they exist.
        """
        try:
            with open('default.cfg') as infile:
                data = json.load(infile)
                self.station_code_entry.insert(0, data['station']['station_id'])

                self.start_date['day_selected'] = data['start']['start_day']
                self.start_date['month_selected'] = data['start']['start_month']
                self.start_date['year_selected'] = data['start']['start_year']

                self.end_date['day_selected'] = data['end']['end_day']
                self.end_date['month_selected'] = data['end']['end_month']
                self.end_date['year_selected'] = data['end']['end_year']
        except FileNotFoundError:
            print('No default file found.')

    def save_values(self):
        """
        Saves the currently selected dates and station ID to a config
        file.
        """
        data = {
            'station': {
                'station_id': self.station_code_entry.get()
            }, 'start': {
                'start_day': self.start_date['day_selected'],
                'start_month': self.start_date['month_selected'],
                'start_year': self.start_date['year_selected']
            }, 'end': {
                'end_day': self.end_date['day_selected'],
                'end_month': self.end_date['month_selected'],
                'end_year': self.end_date['year_selected']
            }}
        with open('default.cfg', 'w+') as outfile:
            json.dump(data, outfile)
            message = "Selected values saved to file!"
            self.alert(message)

    def set_headers(self):
        wb = self.wb
        sheet = wb.active

        sheet['A1'] = 'Observation Time'
        sheet['B1'] = 'Location Full'
        sheet['C1'] = 'Neighborhood'
        sheet['D1'] = 'City'
        sheet['E1'] = 'State'
        sheet['F1'] = 'Zip'
        sheet['G1'] = 'Latitude'
        sheet['H1'] = 'Longitude'
        sheet['I1'] = 'Elevation'
        sheet['J1'] = 'Temp (f)'
        sheet['K1'] = 'Temp (c)'
        sheet['L1'] = 'Relative Humidity'
        sheet['M1'] = 'Wind'
        sheet['N1'] = 'Wind Direction'
        sheet['O1'] = 'Wind Degrees'
        sheet['P1'] = 'Wind (mph)'
        sheet['Q1'] = 'Wind Gust (mph)'
        sheet['R1'] = 'Pressure'
        sheet['S1'] = 'Pressure (mb)'
        sheet['T1'] = 'Pressure (in)'
        sheet['U1'] = 'Dewpoint'
        sheet['V1'] = 'Dewpoint (f)'
        sheet['W1'] = 'Dewpoint (c)'
        sheet['X1'] = 'Solar Radiation'
        sheet['Y1'] = 'UV'
        sheet['Z1'] = 'Precipitation Last Hour'
        sheet['AA1'] = 'Precipitation Last Hour (in)'
        sheet['AB1'] = 'Precipitation Last Hour (mm)'
        sheet['AC1'] = 'Precipitation Today'
        sheet['AD1'] = 'Precipitation Today (in)'
        sheet['AE1'] = 'Precipitation Today (mm)'


def daterange(start_date, end_date):
    """
    From SO
    https://stackoverflow.com/questions/1060279/iterating-through-a-range-of-dates-in-python
    :param start_date: starting date
    :param end_date: ending date
    :return: date
    """
    for num in range(int((end_date - start_date).days)):
        yield start_date + timedelta(num)


if __name__ == '__main__':
    root = tk.Tk()
    root.title('Get Weather History')
    app = Application(master=root)
    root.mainloop()
