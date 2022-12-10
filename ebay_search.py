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
from user_setting import (
    CHROME_DRIVER_PATH,
    GOOGLE_CHROME_PATH,
    SEARCH_PAGE,
    SPREADSHEET_NAME,
    IS_HEADLESS,
)

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
    if driver.find_element(by=By.CLASS_NAME, value="gh-eb-Geo-txt").text != "English":
        element = driver.find_element(by=By.ID, value="gh-eb-Geo-a-default")
        element.click()
        time.sleep(1)
        element = driver.find_element(by=By.ID, value="gh-eb-Geo-a-en")
        element.click()
        time.sleep(2)
    if (
        len(driver.find_elements(by=By.CSS_SELECTOR, value='[aria-label="Ship to United States"]'))
        == 0
    ):
        element = driver.find_element(by=By.CSS_SELECTOR, value="#gh-shipto-click > div > button")
        element.click()
        time.sleep(2)
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
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    credentials = ServiceAccountCredentials.from_json_keyfile_name("client_secret.json", scope)
    client = gspread.authorize(credentials)
    now_str = datetime.datetime.now(
        datetime.timezone(datetime.timedelta(hours=9), "JST")
    ).strftime("%Y%m%d%H%M%S")
    workbook = client.open(SPREADSHEET_NAME)
    workbook.duplicate_sheet(
        source_sheet_id=workbook.worksheet("テンプレート").id,
        new_sheet_name=now_str,
        insert_sheet_index=2,
    )
    return (
        workbook.worksheet("参照URL"),
        workbook.worksheet(now_str),
    )


"""
概　要：情報収集するページの終端を取得する
引　数：driver:webdriver, url:URL
返り値：検索結果のページ数、もしくは、限度ページ数
"""


def get_last_page(driver, url):
    driver.get(SEARCH_URL_FORMAT.format(url, 1))
    time.sleep(2)
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
        # "いいね"が付いてない商品はスキップ
        if len(element.find_elements(by=By.CLASS_NAME, value="s-item__watchCountTotal")) == 0:
            continue

        # time.sleep(1)

        # 抽出結果ページに商品情報を追加
        cell_list = sheet.range(
            row, SheetIndex.SSN_COLUMN.value, row, SheetIndex.DATE_COLUMN.value
        )
        cell_list[SheetIndex.SSN_COLUMN.value - 1].value = ssn_name
        cell_list[
            SheetIndex.ITEM_URL_COLUMN.value - 1
        ].value = f'=HYPERLINK("{element.find_element(by=By.CLASS_NAME, value="s-item__link").get_attribute("href")}")'
        cell_list[
            SheetIndex.IMAGE_COLUMN.value - 1
        ].value = f'=IMAGE("{element.find_element(by=By.CLASS_NAME, value="s-item__image-img").get_attribute("src")}")'
        cell_list[SheetIndex.TITLE_COLUMN.value - 1].value = element.find_element(
            by=By.CSS_SELECTOR, value="a.s-item__link > div > span"
        ).text
        cell_list[SheetIndex.PRICE_COLUMN.value - 1].value = (
            element.find_element(by=By.CLASS_NAME, value="s-item__price")
            .text.replace("JPY ", "")
            .replace("$", "")
        )
        cell_list[SheetIndex.SEND_COLUMN.value - 1].value = element.find_element(
            by=By.CLASS_NAME, value="s-item__logisticsCost"
        ).text.replace("Free shipping", "0")
        cell_list[SheetIndex.WATCHERS_COLUMN.value - 1].value = element.find_element(
            by=By.CSS_SELECTOR, value="span.s-item__watchCountTotal > span"
        ).text.replace(" watchers", "")
        cell_list[SheetIndex.DATE_COLUMN.value - 1].value = datetime.datetime.strptime(
            element.find_element(by=By.CSS_SELECTOR, value="span.s-item__listingDate > span").text,
            "%b-%d %H:%M",
        ).strftime("%m/%d")

        sheet.update_cells(cell_list, value_input_option="USER_ENTERED")

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
        for page in range(1, last_page + 1):
            driver.get(
                SEARCH_URL_FORMAT.format(
                    url_sheet.cell(index, SheetIndex.URL_COLUMN.value).value, page
                )
            )
            time.sleep(1)

            # 検索結果1ページ内での処理を実行
            method_one_page(driver, result_sheet, row, ssn_name)

    print("COMPLETE !")
