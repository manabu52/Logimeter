#!/usr/bin/env python
# coding: utf-8

# In[1]:


# !pip install selenium
# !pip install beautifulsoup4


# In[1]:

# ライブラリインポート
import os
import glob
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
# from selenium.webdriver.common.by import By
import time
# from datetime import date,timedelta
import pandas as pd
import sqlite3

# クロームのオプションを設定する
chrome_options = webdriver.ChromeOptions()
# ヘッドレスモードで起動する
chrome_options.add_argument('--headless')
# enable-automation：ブラウザ起動時のテスト実行警告を非表示
# enable-logging：DevToolsのログを出力しない
chrome_options.add_experimental_option('excludeSwitches', ['enable-automation', 'enable-logging'])
chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
# Chrome起動
browser = webdriver.Chrome(executable_path = 'C:\\Users\\00392352\\Desktop\\MyPandas\\chromedriver.exe')
browser.implicitly_wait(10)

# In[2]:

# サイトアクセス
url_login = "https://logi-auth.kurando.io/saml/auth?SAMLRequest=fVJNb9swDP0rvunkWHKbzRViA0aCAgGyYUjWHnYpCJtphcqSJ9Jb9u8ne0nhHZIr%2BT7IR64IOtvreuA3t8efAxInNREGNt6tvaOhw3DA8Ms0%2BLTfleKNuSedZda%2Fmg4Zw%2BJ9COBavzA%2BGyKRslEyg6gokk3UMw5Gsf%2Bp6difU2ek7aYUL58KtYSlwnRZqCa9L5oiBZU%2FpFIdj%2FBZPSDmKkKJBtw6YnBcilzmeSojpvgupZZ3WsofInmOM03%2B%2BUKK5NRZR3p0K8UQnPZAhrSDDklzow%2F1l52OQA2XEOaU%2FjanD559462oViNaT9OF6mZk096xCi0wrLI5b%2FXvNl%2Bjz3bzzVvT%2FElqa%2F3vdUBgLAWHAUXy6EMHfH0ytVBTxbTpcYJq7MDYum0DEons4nM%2BP7bTM8TbM544Wfuuh2BoDBBP0PBluTlqbWNcezxW5wVu9K74fHTnn1j9BQ%3D%3D"
browser.get(url_login)
# 変数
USER = "db-m.muramatsu@daiwabutsuryu.co.jp"
PASS = "manabu52"
# 値セット
element = browser.find_element_by_id('user_email')
element.clear()
element.send_keys(USER)
element = browser.find_element_by_id('user_password')
element.clear()
element.send_keys(PASS)
browser_button = browser.find_element_by_name('commit')
browser_button.click()
time.sleep(5)

# In[3]:
# 対象センター    三郷：181、浦安支店：204、浦安出張所：278
center_ID = ['181','204','278']
#センター毎に作業時間データをCSV出力
for T_center in center_ID:
    url_login = "https://logimeter.kurando.io/workplace/" + T_center + "/map"
    browser.get(url_login)
    time.sleep(3)
    # 作業データボタン
    browser_button = browser.find_element_by_xpath('/html/body/div/div[2]/div[2]/div/div/div[2]/div[1]/div[1]/div[3]/div[2]/div[1]/div')
    browser_button.click()
    time.sleep(3)
    # from前日セット
    element = browser.find_element_by_xpath('/html/body/div/div[2]/div[2]/div/div/div[2]/div[1]/div[1]/div[1]/div/input')
    element.send_keys(Keys.LEFT)
    element.send_keys(Keys.ENTER)
    # to前日セット
    element = browser.find_element_by_xpath('/html/body/div/div[2]/div[2]/div/div/div[2]/div[1]/div[1]/div[3]/div/input')
    element.send_keys(Keys.LEFT)
    element.send_keys(Keys.ENTER)
    # 表示ボタン
    browser_button = browser.find_element_by_xpath('/html/body/div/div[2]/div[2]/div/div/div[2]/div[3]/button')
    browser_button.click()
    time.sleep(3)
    # 検索結果件数がゼロだったら終わり
    element = browser.find_element_by_xpath('/html/body/div[1]/div[2]/div[2]/div/div/div[3]/div/div/div[1]/div[1]/div[1]/div[2]').text
    if element != 0:
        # CSV出力ボタン
        browser_button = browser.find_element_by_xpath('/html/body/div[1]/div[2]/div[2]/div/div/div[3]/div/div/div[1]/div[3]/div[3]/button')
        browser_button.click()
        time.sleep(5)
# Chrome閉じる
browser.close()
# ダウンロードフォルダパス取得
download_path = os.getenv("HOMEDRIVE") + os.getenv("HOMEPATH") + "\\downloads"
# 対象csvファイルの絶対パス取得
target_pass = download_path + "\\activities_2*csv"
target_files = glob.glob(target_pass)
if target_files != []:
    # DB接続
    dbname = 'C:\\Users\\00392352\Documents\\MyPython\\productivity.db'
    conn = sqlite3.connect(dbname)
    # 対象csvを読んでDB書き込む
    for csv_name in target_files:
        df = pd.read_csv(filepath_or_buffer= csv_name,encoding='shift_jis',index_col=None)
        df.to_sql('logimeter',conn,if_exists='append',index=False)
        conn.commit
        os.remove(csv_name)
    # DBクレンジング
    df_logimeter = pd.read_sql('SELECT * FROM logimeter', conn)
    df = df_logimeter.drop_duplicates()
    df.to_sql('logimeter',conn,if_exists='replace',index=False)
    conn.close()    
    #CSV出力
    output_path = download_path + "\\作業時間.csv"
    df.to_csv(path_or_buf= output_path,index=False)


# %%
