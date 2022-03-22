from bs4 import BeautifulSoup
import requests, openpyxl, datetime


class Web_scrap:

    def __init__(self):
        self.url = "https://pogoda.interia.pl/prognoza-dlugoterminowa-zukowo,cId,41881"
        self.day_of_week = {0: "poniedziałek", 1: "wtorek", 2: "Środa", 3: "czwartek", 4: "piątek", 5: "weekend sobota",
                            6: "weekend niedziela"}
        self.create_excel()

    @staticmethod
    def interia_soup(s):
        interia_long_term_forecast = s.find("section", class_="weather-forecast-longterm")
        return interia_long_term_forecast

    def interia_scrap(self, day_int, ltf): # ltf = long_term_forecas

        day = self.day_of_week.get(day_int)

        daily_weather_forecast = ltf.find("div", class_=f"weather-forecast-longterm-list-entry {day}")

        day = daily_weather_forecast.find("span", class_="day").text
        data = daily_weather_forecast.find("span", class_="date").text

        cloud = daily_weather_forecast.find("span", class_="weather-forecast-longterm-list-entry-forecast-phrase").text

        temp_h = daily_weather_forecast.find("span", class_="weather-forecast-longterm-list-entry-forecast-temp").text
        temp_l = daily_weather_forecast.find("span", class_="weather-forecast-longterm-list-entry-forecast-lowtemp").text

        wind = daily_weather_forecast.find("span", class_="weather-forecast-longterm-list-entry-wind-value").text
        wind += " km/h"

        try:
            type_of_rainfall = daily_weather_forecast.find("span", class_="weather-forecast-longterm-list-entry-precipitation-type").text
            amount_of_rainfall = daily_weather_forecast.find("span", class_="weather-forecast-longterm-list-entry-precipitation-value").text

        except:
            type_of_rainfall, amount_of_rainfall = "--", "--"

        return day, data, cloud, temp_h, temp_l, wind, type_of_rainfall + amount_of_rainfall

    @staticmethod
    def create_excel():
        excel = openpyxl.Workbook()
        sheet = excel.active
        sheet.title = "5 - Dniowa prognoza pogody"
        sheet.append(["Dzień tygodnia", "Data", "Zachmurzenie", "Max- Temperatura", "Min - Temperatura", "Prędkość wiatru", "Opady"])
        excel.save(filename = 'Prognoza Pogody.xlsx') # zapisanie pliku z nazwą
        excel.close()

    @staticmethod
    def save_to_excel(data):
        excel = openpyxl.open("Prognoza Pogody.xlsx")
        sheet = excel.active
        sheet.append(data)  # dodanie listy elementów do eksela
        excel.save(filename='Prognoza Pogody.xlsx')  # zapisanie pliku z nazwą
        excel.close()

    def data_scrap(self):
        try:
            source = requests.get(self.url)
            source.raise_for_status()
            soup = BeautifulSoup(source.text, "html.parser")
            long_term_forecast = self.interia_soup(soup)

            day_of_week_int = datetime.date.today().weekday()

            for _ in range(6):

                if day_of_week_int > 6:
                    day_of_week_int = 0
                else:
                    data = self.interia_scrap(day_of_week_int, long_term_forecast)
                    self.save_to_excel(data)
                    day_of_week_int += 1
                print(" ")

        except Exception as e:
            print(e)


x = Web_scrap()
x.data_scrap()
