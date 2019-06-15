# -*- coding:utf-8 -*-
"""
项目目的：1、解决老板提的天气预报需求，利用和风天气api获取特定城市的未来三天天气情况
          2、读取城市.txt并生成一个excel文件，文件中记录有相应城市天气预报

"""
import logging
import os
import requests
import xlwings as xl

logging.basicConfig(level = logging.INFO,format = '%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

CURRENT_PATH = os.path.dirname(os.path.abspath(__file__))

class WeatherSpider:
    def __init__(self, city_name):
        self._city_name = city_name
        self._config_init()
        logger.info("WeatherSpider inited")

    def _config_init(self):
        key = ""  # 需要申请， 和风天气api
        params = "location={1}&key={0}".format(key, self._city_name)
        self._url = "https://free-api.heweather.net/s6/weather/forecast?" + params
        self._headers = ""

    def request(self):
        try:
            logger.info("WeatherSpider begin get {0}".format(self._city_name))
            resp = requests.post(url=self._url, headers=self._headers)
        except Exception as e:
            logger.warning("WeatherSpider post fail: {0}".format(e))
        else:
            if resp.status_code == 200:
                return resp.json()
        return False
            


class ExcelController:
    def __init__(self):
        self._app_init()
        self._workbook_init()
        self._sheet_init()
        self._weatherType = ["飓风","龙卷风","阵雨","强阵雨","雷阵雨","强雷阵雨","雷阵雨伴有冰雹",
                             "小雨","中雨","大雨","极端降雨","暴雨","大暴雨","特大暴雨","冻雨",
                             "小到中雨","中到大雨","大到暴雨","暴雨到大暴雨","大暴雨到特大暴雨",
                             "雨","小雪","中雪","大雪","暴雪","雨夹雪","雨雪天气","阵雨夹雪",
                             "阵雪","小到中雪","中到大雪","大到暴雪","雪"]
        logger.info("ExcelController inited")
        

    def _app_init(self):
        self._app = xl.App()
        self._app.visible = True

    def _workbook_init(self):
        self._book = self._app.books[0]

    def _sheet_init(self):
        self._sheet = self._book.sheets[0]

    def addValue(self, rangeName, value):
        if rangeName is None:
            raise ValuError
        if not rangeName[0].isalpha():
            raise ValuError
        if not rangeName[-1].isalnum():
            raise ValuError
        self._sheet.range(rangeName).value = value
        logger.info("ExcelController added {0} to {1}".format(value, rangeName))

    def setStyle(self):
        logger.info("ExcelController setting styles")
        # 合并单元格
        self._sheet.range("A1:E1").api.Merge()
        for i in range(1, 27):
            self._sheet.range("A{0}:A{1}".format(i*2+1, i*2+2)).api.Merge()
        # 单元格宽度和高度
        self._sheet.range("A1").row_height = 24
        self._sheet.range("A1:E1").column_width = 15  # 宽度
        self._sheet.range("A2:A54").row_height = 20  # 高度
        # 单元格颜色标记
        for column in ["C", "D", "E"]:
            for row in range(3, 55):
                if self._sheet.range("{0}{1}".format(column, row)).value in self._weatherType:
                    self._sheet.range("{0}{1}".format(column, row)).color = (255,218,185)

        # 单元格字体及大小设置
        self._sheet.range("A1").api.Font.Name = "微软雅黑"
        self._sheet.range("A1").api.Font.Size = 12
        self._sheet.range("A1").api.Font.Bold = True
        self._sheet.range("A1").api.HorizontalAlignment = -4108
        self._sheet.range("A1").api.VerticalAlignment = -4108
        self._sheet.range("A1:E1").api.Borders(9).LineStyle = 1
        self._sheet.range("A1:E1").api.Borders(8).LineStyle = 1
        self._sheet.range("A1:E1").api.Borders(10).LineStyle = 1
        self._sheet.range("A1:E1").api.Borders(7).LineStyle = 1  # xlContinuous
        for column in ["A", "B", "C", "D", "E"]:
            for row in range(2, 55):
                self._sheet.range("{0}{1}".format(column, row)).api.Font.Name = "微软雅黑"
                self._sheet.range("{0}{1}".format(column, row)).api.Font.Size = 10
                # 设置居中
                self._sheet.range("{0}{1}".format(column, row)).api.HorizontalAlignment = -4108  # xlCenter水平居中
                self._sheet.range("{0}{1}".format(column, row)).api.VerticalAlignment = -4108  # xlCenter垂直居中
                # 设置边框
                self._sheet.range("{0}{1}".format(column, row)).api.Borders(9).LineStyle = 1  # 1 xlContinuous  9 xlEdgeBottom
                self._sheet.range("{0}{1}".format(column, row)).api.Borders(8).LineStyle = 1  # xlContinuous  8 xlEdgeTop
                self._sheet.range("{0}{1}".format(column, row)).api.Borders(10).LineStyle = 1  # xlContinuous  10 xlEdgeRight
                self._sheet.range("{0}{1}".format(column, row)).api.Borders(7).LineStyle = 1  # xlContinuous  7 xlEdgeLeft
        logger.info("ExcelController setted styles")


    def close(self):
        xlsxFilePath = os.path.join(CURRENT_PATH, "各城市天气预报.xlsx")
        self._book.save(xlsxFilePath)
        self._app.visible = False
        self._app.quit()
        logger.info("ExcelController closed")


def readFromTxt():
    txtFilePath = os.path.join(CURRENT_PATH, "城市.txt")
    with open(txtFilePath, "r") as f:
        logger.info("Program readFromTxt")
        return f.readlines()


def run():
    logger.info("Program beginning")
    # 读取城市列表
    excelController= ExcelController()
    cityList = readFromTxt()
    for cityKey, cityName in enumerate(cityList):
        cityName = cityName.replace("\n", "")
        # 获取天气信息
        weather = WeatherSpider(cityName)
        weatherJson = weather.request()
        HeWeather6 = weatherJson.get("HeWeather6")[0]
        update = HeWeather6.get("update").get("loc")
        excelController.addValue("A1", "华南各城市天气预报(更新时间：{0})".format(update))
        excelController.addValue("A2", "城市")
        excelController.addValue("A{0}".format((cityKey+1)*2+1), cityName)  # 城市名称
        excelController.addValue("B2", "昼夜")
        excelController.addValue("B{0}".format((cityKey+1)*2+1), "白天")  # 白天
        excelController.addValue("B{0}".format((cityKey+1)*2+2), "夜间")  # 夜间
        daily_forecast = HeWeather6.get("daily_forecast")
        #
        dateDict = {"0":"C", "1":"D", "2":"E"}
        for key, value in enumerate(daily_forecast):
            date = value.get("date")
            excelController.addValue("{0}{1}".format(dateDict.get(str(key)), 2), date)  # 预报日期
            cond_txt_d = value.get("cond_txt_d")
            cond_txt_n = value.get("cond_txt_n")
            excelController.addValue("{0}{1}".format(dateDict.get(str(key)), 3+cityKey*2), cond_txt_d)  # 白天天气文字
            excelController.addValue("{0}{1}".format(dateDict.get(str(key)), 4+cityKey*2), cond_txt_n)  # 夜间天气文字
    excelController.setStyle()
    excelController.close()
    logger.info("Program finished")
    
def main():
    run()


if __name__ == '__main__':
    main()
