import requests
from bs4 import BeautifulSoup


def getWeather():
    #tenki.jpの目的の地域のページのURL（今回は東京都調布市）
    #url = 'https://tenki.jp/forecast/3/16/4410/13208/'
    url = 'https://tenki.jp/forecast/4/20/5610/17201/'
    #HTTPリクエスト
    r = requests.get(url)

    #プロキシ環境下の場合は以下を記述
    """
    proxies = {
        "http":"http://proxy.xxx.xxx.xxx:8080",
        "https":"http://proxy.xxx.xxx.xxx:8080"
    }
    r = requests.get(url, proxies=proxies)
    """

    #HTMLの解析
    bsObj = BeautifulSoup(r.content, "html.parser")

    #今日の天気を取得
    today = bsObj.find(class_="today-weather")
    weather = today.p.string

    #気温情報のまとまり
    temp=today.div.find(class_="date-value-wrap")

    #気温の取得
    temp=temp.find_all("dd")
    temp_max = temp[0].span.string #最高気温
    temp_max_diff=temp[1].string #最高気温の前日比
    temp_min = temp[2].span.string #最低気温
    temp_min_diff=temp[3].string #最低気温の前日比

    #とりあえず動くのを見たいので天気と気温を繋げて返しています。
    return weather+temp_max+temp_min