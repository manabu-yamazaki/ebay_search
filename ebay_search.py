import datetime
import math
import os
import random
import time
import urllib.parse
from enum import Enum

import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.chrome import service as fs
from selenium.webdriver.common.by import By

from user_agent import USER_AGNETS
from user_setting import CHROME_DRIVER_PATH, GOOGLE_CHROME_PATH, SEARCH_PAGE, SPREADSHEET_NAME, IS_HEADLESS

FIRST_URL = "https://www.ebay.com/"
SEARCH_URL_FORMAT = "{}&_ipg=240&_pgn={}"
# SEARCH_TEXT = r"(.*) 件中"


class SheetIndex(Enum):
    TITLE_ROW = 1
    START_ROW = 2
    URL_COLUMN = 1
    SSN_COLUMN = 1
    ITEM_URL_COLUMN = 2
    IMAGE_COLUMN = 3
    TITLE_COLUMN = 4
    PRICE_COLUMN = 5
    SEND_COLUMN = 6
    WATCHERS_COLUMN = 7
    DATE_COLUMN = 8


"""
概　要：webdriverで指定のURLを開く
引　数：user_agent:UserAgentリスト, url:URL
返り値：webdriver
"""


def open_web_driver(user_agent, url):
    UA = random.choice(user_agent)
    ChromeOptions = webdriver.ChromeOptions()
    ChromeOptions.binary_location = GOOGLE_CHROME_PATH
    ChromeOptions.add_argument("--user-agent=" + UA)
    ChromeOptions.add_experimental_option(
        "excludeSwitches", ["enable-automation", "enable-logging"]
    )
    ChromeOptions.add_argument("start-maximized")
    ChromeOptions.add_argument("enable-automation")
    ChromeOptions.add_argument("--lang=ja-JP")
    ChromeOptions.add_argument("--no-sandbox")
    ChromeOptions.add_argument("--disable-infobars")
    ChromeOptions.add_argument("--disable-extensions")
    ChromeOptions.add_argument("--disable-dev-shm-usage")
    ChromeOptions.add_argument("--disable-browser-side-navigation")
    ChromeOptions.add_argument("--disable-gpu")
    ChromeOptions.add_argument("--ignore-certificate-errors")
    ChromeOptions.add_argument("--ignore-ssl-errors")
    # ChromeOptions.add_argument("--user-data-dir=/Users/yamazakimanabu/Library/Application Support/Google/Chrome Beta")
    # ChromeOptions.add_argument("--profile-directory=Profile 1")
    if IS_HEADLESS:
        ChromeOptions.add_argument("--headless")
    ChromeOptions.add_experimental_option("extensionLoadTimeout", 12000)
    prefs = {"profile.default_content_setting_values.notifications": 2}
    ChromeOptions.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(
        service=fs.Service(os.getcwd() + CHROME_DRIVER_PATH), options=ChromeOptions,
    )
    # ページを開く
    driver.get(url)
    # sleep_time = random.randint(7, 15)
    time.sleep(1)

    # if len(driver.find_elements_by_class_name("bcs_inputBox")) == 0:
    #     time.sleep(10)
    #     user_agent.remove(UA)
    #     print("UAbot判定。UA名: ", UA)
    #     print("残りUA: ", len(user_agent))
    #     logging.debug(UA)
    #     logging.debug(len(user_agent))
    #     driver.close()
    #     driver = open_web_driver(user_agent, url)

    print(f"OPEN URL: {url}")
    return driver


"""
概　要：ログイン処理
引　数：driver
返り値：なし
"""


def init_config(driver):
    element = driver.find_element(by=By.ID, value="gh-eb-Geo-a-default")
    element.click()
    time.sleep(1)
    element = driver.find_element(by=By.ID, value="gh-eb-Geo-a-en")
    element.click()
    time.sleep(1)
    element = driver.find_element(by=By.CSS_SELECTOR, value="#gh-shipto-click > div > button")
    element.click()
    time.sleep(1)
    element = driver.find_element(by=By.CLASS_NAME, value="menu-button__button")
    element.click()
    element = driver.find_element(
        by=By.CSS_SELECTOR, value='.menu-button__item:has(span > span[data-country="USA|US"])'
    )
    element.click()
    time.sleep(1)
    element = driver.find_element(by=By.CLASS_NAME, value="shipto__close-btn")
    element.click()
    time.sleep(3)


"""
概　要：スプレッドシート読み込み
引　数：なし
返り値：スプレッドシート, Googleドライブ
"""


def init_google_tools():
    # gauth = GoogleAuth()
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    # gauth.credentials = ServiceAccountCredentials.from_json_keyfile_name(
    #     "client_secret.json", scope
    # )
    credentials = ServiceAccountCredentials.from_json_keyfile_name("client_secret.json", scope)
    # client = gspread.authorize(gauth.credentials)
    client = gspread.authorize(credentials)
    # return client.open(SPREADSHEET_NAME).sheet1, GoogleDrive(gauth)
    now = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=9), "JST"))
    return (
        client.open(SPREADSHEET_NAME).worksheet("参照URL"),
        client.open(SPREADSHEET_NAME).add_worksheet(
            now.strftime("%Y%m%d%H%M%S"), rows=1000, cols=20
        ),
    )


"""
概　要：情報収集するページの終端を取得する
引　数：sheet:抽出結果シート
返り値：検索結果のページ数、もしくは、限度ページ数
"""


def init_result_page(sheet):
    sheet.update_cell(SheetIndex.TITLE_ROW.value, SheetIndex.SSN_COLUMN.value, "出品者")
    sheet.update_cell(SheetIndex.TITLE_ROW.value, SheetIndex.ITEM_URL_COLUMN.value, "商品ＵＲＬ")
    sheet.update_cell(SheetIndex.TITLE_ROW.value, SheetIndex.IMAGE_COLUMN.value, "画像")
    sheet.update_cell(SheetIndex.TITLE_ROW.value, SheetIndex.TITLE_COLUMN.value, "タイトル")
    sheet.update_cell(SheetIndex.TITLE_ROW.value, SheetIndex.PRICE_COLUMN.value, "金額")
    sheet.update_cell(SheetIndex.TITLE_ROW.value, SheetIndex.SEND_COLUMN.value, "送料")
    sheet.update_cell(SheetIndex.TITLE_ROW.value, SheetIndex.WATCHERS_COLUMN.value, "watchers数")
    sheet.update_cell(SheetIndex.TITLE_ROW.value, SheetIndex.DATE_COLUMN.value, "出品日")


"""
概　要：情報収集するページの終端を取得する
引　数：driver:webdriver, url:URL
返り値：検索結果のページ数、もしくは、限度ページ数
"""


def get_last_page(driver, url):
    driver.get(SEARCH_URL_FORMAT.format(url, 1))
    time.sleep(1)
    init_config(driver)
    if len(driver.find_elements(by=By.CLASS_NAME, value="srp-controls__count-heading")) == 0:
        return 0
    last_page = math.ceil(
        int(
            driver.find_element(by=By.CLASS_NAME, value="srp-controls__count-heading")
            .text.replace(" results", "")
            .replace("+", "")
            .replace(",", "")
        )
        / 240
    )
    return last_page if last_page < SEARCH_PAGE else SEARCH_PAGE


"""
概　要：検索結果1ページ内での処理を実行
引　数：driver:webdriver, sheet:抽出結果シート, row:追加行, ssn_name:出品者名
返り値：検索結果のページ数、もしくは、限度ページ数
"""


def method_one_page(driver, sheet, row, ssn_name):
    print(ssn_name)
    for element in driver.find_elements(by=By.CLASS_NAME, value="s-item__wrapper"):
        time.sleep(1)
        # "いいね"が付いてない商品はスキップ
        if len(element.find_elements(by=By.CLASS_NAME, value="s-item__watchCountTotal")) == 0:
            continue

        # 抽出結果ページに商品情報を追加
        sheet.update_cell(row, SheetIndex.SSN_COLUMN.value, ssn_name)
        sheet.update_cell(
            row,
            SheetIndex.ITEM_URL_COLUMN.value,
            f'=HYPERLINK("{element.find_element(by=By.CLASS_NAME, value="s-item__link").get_attribute("href")}")',
        )
        sheet.update_cell(
            row,
            SheetIndex.IMAGE_COLUMN.value,
            f'=IMAGE("{element.find_element(by=By.CLASS_NAME, value="s-item__image-img").get_attribute("src")}")',
        )
        sheet.update_cell(
            row,
            SheetIndex.TITLE_COLUMN.value,
            element.find_element(by=By.CSS_SELECTOR, value="a.s-item__link > div > span").text,
        )
        sheet.update_cell(
            row,
            SheetIndex.PRICE_COLUMN.value,
            element.find_element(by=By.CLASS_NAME, value="s-item__price")
            .text.replace("JPY ", "")
            .replace("$", ""),
        )
        sheet.update_cell(
            row,
            SheetIndex.SEND_COLUMN.value,
            element.find_element(by=By.CLASS_NAME, value="s-item__logisticsCost").text.replace(
                "Free shipping", "0"
            ),
        )
        sheet.update_cell(
            row,
            SheetIndex.WATCHERS_COLUMN.value,
            element.find_element(
                by=By.CSS_SELECTOR, value="span.s-item__watchCountTotal > span"
            ).text.replace(" watchers", ""),
        )
        sheet.update_cell(
            row,
            SheetIndex.DATE_COLUMN.value,
            datetime.datetime.strptime(
                element.find_element(
                    by=By.CSS_SELECTOR, value="span.s-item__listingDate > span"
                ).text,
                "%b-%d %H:%M",
            ).strftime("%Y/%m/%d"),
        )

        row += 1

        print("item add")


"""
概　要：本処理
引　数：なし
返り値：なし
"""


def start():
    # スプレッドシート読み込み
    url_sheet, result_sheet = init_google_tools()
    print("Google Tools Authorized")

    # 抽出結果ページ初期処理
    init_result_page(result_sheet)

    # 最初のページを開いて初期設定する
    driver = open_web_driver(USER_AGNETS, FIRST_URL)
    # element = driver.find_element(by=By.CSS_SELECTOR, value=".gh-ug-guest > a")
    # element.click()
    # time.sleep(3)
    # element = driver.find_element(by=By.ID, value="textbox__control")
    # element.clear()
    # element.send_keys("god.of.icecream@gmail.com")
    # element = driver.find_element(by=By.ID, value="signin-continue-btn")
    # element.click()

    # # 言語設定&配送先設定
    # init_config(driver)

    # 参照URLリストを検索
    index = SheetIndex.START_ROW.value - 1
    # 抽出結果の追加行
    row = SheetIndex.START_ROW.value
    while True:
        index += 1

        # リスト最終行まできたら終了
        if not url_sheet.cell(index, SheetIndex.URL_COLUMN.value).value:
            print("END ROW !")
            break

        print(f"OPEN URL: {url_sheet.cell(index, SheetIndex.URL_COLUMN.value).value}")

        # ページ数を取得(MAXページ以下)
        last_page = get_last_page(driver, url_sheet.cell(index, SheetIndex.URL_COLUMN.value).value)
        if last_page == 0:
            continue
        # 出品者名取得
        ssn_name = urllib.parse.parse_qs(urllib.parse.urlparse(driver.current_url).query)["_ssn"][
            0
        ]

        # ページ数分繰り返し
        for page in range(1, last_page):
            driver.get(
                SEARCH_URL_FORMAT.format(
                    url_sheet.cell(index, SheetIndex.URL_COLUMN.value).value, page
                )
            )
            time.sleep(1)

            # 検索結果1ページ内での処理を実行
            method_one_page(driver, result_sheet, row, ssn_name)

    print("COMPLETE !")
