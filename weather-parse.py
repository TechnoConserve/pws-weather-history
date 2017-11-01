import calendar
from datetime import date
import xml.etree.ElementTree as ET
from tkinter import *
from tkinter import ttk
import tkinter.font as tkFont

import openpyxl
import pandas as pd


def get_calendar(locale, fwday):
    # instantiate proper calendar class
    if locale is None:
        return calendar.TextCalendar(fwday)
    else:
        return calendar.LocaleTextCalendar(fwday, locale)


class Calendar(ttk.Frame):
    # XXX ToDo: cget and configure

    datetime = calendar.datetime.datetime
    timedelta = calendar.datetime.timedelta

    def __init__(self, master=None, **kw):
        """
        WIDGET-SPECIFIC OPTIONS

            locale, firstweekday, year, month, selectbackground,
            selectforeground
        """
        # remove custom options from kw before initializating ttk.Frame
        fwday = kw.pop('firstweekday', calendar.MONDAY)
        year = kw.pop('year', self.datetime.now().year)
        month = kw.pop('month', self.datetime.now().month)
        locale = kw.pop('locale', None)
        sel_bg = kw.pop('selectbackground', '#ecffc4')
        sel_fg = kw.pop('selectforeground', '#05640e')

        self._date = self.datetime(year, month, 1)
        self._selection = None # no date selected

        ttk.Frame.__init__(self, master, **kw)

        self._cal = get_calendar(locale, fwday)

        self.__setup_styles()       # creates custom styles
        self.__place_widgets()      # pack/grid used widgets
        self.__config_calendar()    # adjust calendar columns and setup tags
        # configure a canvas, and proper bindings, for selecting dates
        self.__setup_selection(sel_bg, sel_fg)

        # store items ids, used for insertion later
        self._items = [self._calendar.insert('', 'end', values='')
                            for _ in range(6)]
        # insert dates in the currently empty calendar
        self._build_calendar()

        # set the minimal size for the widget
        self._calendar.bind('<Map>', self.__minsize)

    def __setitem__(self, item, value):
        if item in ('year', 'month'):
            raise AttributeError("attribute '%s' is not writeable" % item)
        elif item == 'selectbackground':
            self._canvas['background'] = value
        elif item == 'selectforeground':
            self._canvas.itemconfigure(self._canvas.text, item=value)
        else:
            ttk.Frame.__setitem__(self, item, value)

    def __getitem__(self, item):
        if item in ('year', 'month'):
            return getattr(self._date, item)
        elif item == 'selectbackground':
            return self._canvas['background']
        elif item == 'selectforeground':
            return self._canvas.itemcget(self._canvas.text, 'fill')
        else:
            r = ttk.tclobjs_to_py({item: ttk.Frame.__getitem__(self, item)})
            return r[item]

    def __setup_styles(self):
        # custom ttk styles
        style = ttk.Style(self.master)
        arrow_layout = lambda dir: (
            [('Button.focus', {'children': [('Button.%sarrow' % dir, None)]})]
        )
        style.layout('L.TButton', arrow_layout('left'))
        style.layout('R.TButton', arrow_layout('right'))

    def __place_widgets(self):
        # header frame and its widgets
        hframe = ttk.Frame(self)
        lbtn = ttk.Button(hframe, style='L.TButton', command=self._prev_month)
        rbtn = ttk.Button(hframe, style='R.TButton', command=self._next_month)
        self._header = ttk.Label(hframe, width=15, anchor='center')
        # the calendar
        self._calendar = ttk.Treeview(show='', selectmode='none', height=7)

        # pack the widgets
        hframe.pack(in_=self, side='top', pady=4, anchor='center')
        lbtn.grid(in_=hframe)
        self._header.grid(in_=hframe, column=1, row=0, padx=12)
        rbtn.grid(in_=hframe, column=2, row=0)
        self._calendar.pack(in_=self, expand=1, fill='both', side='bottom')

    def __config_calendar(self):
        cols = self._cal.formatweekheader(3).split()
        self._calendar['columns'] = cols
        self._calendar.tag_configure('header', background='grey90')
        self._calendar.insert('', 'end', values=cols, tag='header')
        # adjust its columns width
        font = tkFont.Font()
        maxwidth = max(font.measure(col) for col in cols)
        for col in cols:
            self._calendar.column(col, width=maxwidth, minwidth=maxwidth,
                anchor='e')

    def __setup_selection(self, sel_bg, sel_fg):
        self._font = tkFont.Font()
        self._canvas = canvas = Canvas(self._calendar,
            background=sel_bg, borderwidth=0, highlightthickness=0)
        canvas.text = canvas.create_text(0, 0, fill=sel_fg, anchor='w')

        canvas.bind('<ButtonPress-1>', lambda evt: canvas.place_forget())
        self._calendar.bind('<Configure>', lambda evt: canvas.place_forget())
        self._calendar.bind('<ButtonPress-1>', self._pressed)

    def __minsize(self, evt):
        width, height = self._calendar.master.geometry().split('x')
        height = height[:height.index('+')]
        self._calendar.master.minsize(width, height)

    def _build_calendar(self):
        year, month = self._date.year, self._date.month

        # update header text (Month, YEAR)
        header = self._cal.formatmonthname(year, month, 0)
        self._header['text'] = header.title()

        # update calendar shown dates
        cal = self._cal.monthdayscalendar(year, month)
        for indx, item in enumerate(self._items):
            week = cal[indx] if indx < len(cal) else []
            fmt_week = [('%02d' % day) if day else '' for day in week]
            self._calendar.item(item, values=fmt_week)

    def _show_selection(self, text, bbox):
        """Configure canvas for a new selection."""
        x, y, width, height = bbox

        textw = self._font.measure(text)

        canvas = self._canvas
        canvas.configure(width=width, height=height)
        canvas.coords(canvas.text, width - textw, height / 2 - 1)
        canvas.itemconfigure(canvas.text, text=text)
        canvas.place(in_=self._calendar, x=x, y=y)

    # Callbacks

    def _pressed(self, evt):
        """Clicked somewhere in the calendar."""
        x, y, widget = evt.x, evt.y, evt.widget
        item = widget.identify_row(y)
        column = widget.identify_column(x)

        if not column or not item in self._items:
            # clicked in the weekdays row or just outside the columns
            return

        item_values = widget.item(item)['values']
        if not len(item_values): # row is empty for this month
            return

        text = item_values[int(column[1]) - 1]
        if not text: # date is empty
            return

        bbox = widget.bbox(item, column)
        if not bbox: # calendar not visible yet
            return

        # update and then show selection
        text = '%02d' % text
        self._selection = (text, item, column)
        self._show_selection(text, bbox)

    def _prev_month(self):
        """Updated calendar to show the previous month."""
        self._canvas.place_forget()

        self._date = self._date - self.timedelta(days=1)
        self._date = self.datetime(self._date.year, self._date.month, 1)
        self._build_calendar() # reconstuct calendar

    def _next_month(self):
        """Update calendar to show the next month."""
        self._canvas.place_forget()

        year, month = self._date.year, self._date.month
        self._date = self._date + self.timedelta(
            days=calendar.monthrange(year, month)[1] + 1)
        self._date = self.datetime(self._date.year, self._date.month, 1)
        self._build_calendar() # reconstruct calendar

    # Properties

    @property
    def selection(self):
        """Return a datetime representing the current selected date."""
        if not self._selection:
            return None

        year, month = self._date.year, self._date.month
        return self.datetime(year, month, int(self._selection[0]))


class Application(ttk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.grid(column=0, row=0, sticky=(N, W, E, S))
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        self.station_code = StringVar()
        self.day_start = Calendar(firstweekday=calendar.SUNDAY)
        self.day_start.grid(column=1, row=1, sticky=(W, E))

        self.station_code_entry = ttk.Entry(self, width=7, textvariable=self.station_code)
        self.station_code_entry.grid(column=2, row=1, sticky=(W, E))

        self.download = ttk.Button(self, text="Download Data", command=grab_history).grid(column=3, row=3, sticky=W)

        ttk.Label(self, text="Station Code").grid(column=3, row=1, sticky=W)

        for child in self.winfo_children():
            child.grid_configure(padx=5, pady=5)

        self.station_code_entry.focus()


def grab_history():
    day, month, year = date.today().strftime('%d %m %Y').split(' ')
    url = 'https://www.wunderground.com/weatherstation/WXDailyHistory.asp?' \
          'ID=KAZLITTL3&day=13&month=10&year=2016&' \
          'dayend=' + day + '&monthend=' + month + '&yearend=' + year + \
          '&graphspan=custom&format=1'
    df = pd.read_csv(url, header=0, index_col=0)
    df.rename(columns={'PrecipitationSumIn<br>': 'Precipitation Sum (in)'}, inplace=True)
    df = df[df.index != '<br>']
    df.to_csv('updated-' + day + '-' + month + '-' + year + 'blackrock-weather.csv')


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
    root = Tk()
    root.title('Get Weather History')
    app = Application(master=root)
    root.bind('<Return>', grab_history())
    root.mainloop()
