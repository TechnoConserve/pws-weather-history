from datetime import date
import xml.etree.ElementTree as ET
import tkinter as tk

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

        self.day_start = tk.Button(
            self, text="Choose",
            command=lambda: self.datepicker(side='start')).grid(column=2, row=1, sticky=(tk.W, tk.E))
        self.day_end = tk.Button(
            self, text="Choose",
            command=lambda: self.datepicker(side='end')).grid(column=2, row=2, sticky=(tk.W, tk.E))

        self.station_code_entry = tk.Entry(self, width=7, textvariable=self.station_code)
        self.station_code_entry.grid(column=2, row=3, sticky=(tk.W, tk.E))

        self.download = tk.Button(
            self, text="Download Data", command=self.grab_history).grid(column=3, row=4, sticky=tk.W)
        #self.show_btn = tk.Button(self, text='Dump data', command=self.dump_data)

        tk.Label(self, text="Station Code").grid(column=3, row=3, sticky=tk.W)
        tk.Label(self, text="Start Date").grid(column=1, row=1, sticky=tk.E)
        tk.Label(self, text="End Date").grid(column=1, row=2, sticky=tk.E)

        for child in self.winfo_children():
            child.grid_configure(padx=5, pady=5)

        self.station_code_entry.focus()

    def datepicker(self, side):
        """
        Lauch a child window with a calendar widget.
        :param side: Specifies if the child window is recording a start
        or end date.
        """
        child = tk.Toplevel()
        child.wm_title("Choose date")
        if side == 'start':
            Calendar(child, values=self.start_date)
        else:
            Calendar(child, values=self.end_date)

    def dump_data(self):
        print("Start date:", self.start_date)
        print("End date:", self.end_date)

    def grab_history(self):
        day, month, year = date.today().strftime('%d %m %Y').split(' ')
        station_id = self.station_code_entry.get()
        url = 'https://www.wunderground.com/weatherstation/WXDailyHistory.asp?' \
              'ID=' + station_id + '=13&month=10&year=2016&' \
              'dayend=' + day + '&monthend=' + month + '&yearend=' + year + \
              '&graphspan=custom&format=1'
        df = pd.read_csv(url, header=0, index_col=0)
        df.rename(columns={'PrecipitationSumIn<br>': 'Precipitation Sum (in)'}, inplace=True)
        df = df[df.index != '<br>']
        df.to_csv('updated-' + day + '-' + month + '-' + year + '-blackrock-weather.csv')


def parse_today():
    tree = ET.parse('WXDailyHistory.xml')
    root = tree.getroot()

    wb = openpyxl.Workbook()
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

    for idx, observation in enumerate(root):
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

        sheet['A' + str(idx + 2)] = observation_time
        sheet['B' + str(idx + 2)] = location_full
        sheet['C' + str(idx + 2)] = location_neighborhood
        sheet['D' + str(idx + 2)] = location_city
        sheet['E' + str(idx + 2)] = location_state
        sheet['F' + str(idx + 2)] = location_zip
        sheet['G' + str(idx + 2)] = location_latitude
        sheet['H' + str(idx + 2)] = location_longitude
        sheet['I' + str(idx + 2)] = location_elevation
        sheet['J' + str(idx + 2)] = temp_f
        sheet['K' + str(idx + 2)] = temp_c
        sheet['L' + str(idx + 2)] = relative_humidity
        sheet['M' + str(idx + 2)] = wind
        sheet['N' + str(idx + 2)] = wind_dir
        sheet['O' + str(idx + 2)] = wind_degrees
        sheet['P' + str(idx + 2)] = wind_mph
        sheet['Q' + str(idx + 2)] = wind_gust_mph
        sheet['R' + str(idx + 2)] = pressure_string
        sheet['S' + str(idx + 2)] = pressure_mb
        sheet['T' + str(idx + 2)] = pressure_in
        sheet['U' + str(idx + 2)] = dewpoint_string
        sheet['V' + str(idx + 2)] = dewpoint_f
        sheet['W' + str(idx + 2)] = dewpoint_c
        sheet['X' + str(idx + 2)] = solar_radiation
        sheet['Y' + str(idx + 2)] = uv
        sheet['Z' + str(idx + 2)] = precip_1hr_string
        sheet['AA' + str(idx + 2)] = precip_1hr_in
        sheet['AB' + str(idx + 2)] = precip_1hr_metric
        sheet['AC' + str(idx + 2)] = precip_today_string
        sheet['AD' + str(idx + 2)] = precip_today_in
        sheet['AE' + str(idx + 2)] = precip_today_metric

    todays_date = date.today().strftime('%d-%m-%Y')
    wb.save(todays_date + '-blackrock-weather.xlsx')


if __name__ == '__main__':
    root = tk.Tk()
    root.title('Get Weather History')
    app = Application(master=root)
    root.mainloop()
