import os
import datetime
import pandas as pd
import openpyxl
from tqdm import tqdm
import streamlit as st
import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.expected_conditions import presence_of_element_located
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.core.utils import ChromeType


def automate_browser(output, output2, output3, output4, output5, output6):

     try:
            
 # ドライバー指定でChromeブラウザを開く
        options = webdriver.ChromeOptions()
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        #options.add_argument('--headless')
        options.add_argument('--disable-gpu')
        options.add_argument('--no-sandbox')
        options.add_argument('start-maximized')
        options.add_argument('disable-infobars')
        options.add_argument('--disable-dev-shm-usage')
        options.add_argument('--remote-debugging-port=0')
        options.add_argument('--disable-extensions')
        options.add_argument('--dns-prefetch-disable')
        options.add_argument('--disable-web-security')





        # CromeTypeクラスを新たにインポート
          

          # webdriver_managerによりドライバーをインストール
        CHROMEDRIVER = ChromeDriverManager(chrome_type=ChromeType.CHROMIUM).install()
        service = fs.Service(CHROMEDRIVER)
        driver = webdriver.Chrome(options=options, service=service)

        

        


        # Webシステムのログインページを開く
        driver.get('http://jrmgsvr11/App01/LogIn.aspx')

        time.sleep(2) # 2秒待つ

        # 日販環境のラジオボタンをクリック
        search_radio = driver.find_element_by_id("MainContent_rblKankyou_2") .click()

        # ID入力項目を入れる
        search_bar_id = driver.find_element_by_id("MainContent_txtUser")
        search_bar_id.send_keys(output3)

        # Password入力項目にパスワードを入れる
        search_bar_pass = driver.find_element_by_id("MainContent_txtPassword")
        search_bar_pass.send_keys(output4)

        # ログインをクリック
        driver.find_element_by_id("MainContent_btnLogIn") .click()

        # 入力・照会タブをクリック
        driver.find_element_by_id("__tab_MainContent_TabContainer1_TabPanel6") .click()

        time.sleep(1) # 1秒待つ

        # JU2220売上照会を開く
        driver.find_element_by_xpath('//a[@href = "JU/JU0410.aspx"]') .click()

            #管理部門を入れる

        element = driver.find_element_by_id("MainContent_txtSKanri")
        driver.execute_script("arguments[0].setAttribute('value' , '')" , element ) 
        Kanribumonbar_id = driver.find_element_by_id("MainContent_txtSKanri")
        Kanribumonbar_id.send_keys(output5)

        element = driver.find_element_by_id("MainContent_txtEKanri")
        driver.execute_script("arguments[0].setAttribute('value' , '')" , element ) 
        Kanribumonbar_id = driver.find_element_by_id("MainContent_txtEKanri")
        Kanribumonbar_id.send_keys(output6)

        #伝票日付を入れる

        element = driver.find_element_by_id("MainContent_txtSDenpyouHiduke")
        driver.execute_script("arguments[0].setAttribute('value' , '')" , element )
        Hidukebar_id = driver.find_element_by_id("MainContent_txtSDenpyouHiduke")
        Hidukebar_id.send_keys(output)

        element = driver.find_element_by_id("MainContent_txtEDenpyouHiduke")
        driver.execute_script("arguments[0].setAttribute('value' , '')" , element )
        Hidukebar2_id = driver.find_element_by_id("MainContent_txtEDenpyouHiduke")
        Hidukebar2_id.send_keys(output2)


        # 表示ボタンをクリック
        driver.find_element_by_id("MainContent_btnHyouji") .click()

        time.sleep(2) # 2秒待つ

        # EXCEL出力をクリック

        element = driver.find_element_by_id("MainContent_lbnExcelOut")
        element.click()

        time.sleep(1) # 1秒待つ

            # すべてボタンをクリック
        element = driver.find_element_by_id("MainContent_cbxAll")
        #driver.execute_script("arguments[0].setAttribute('onclick' , '')" , element ) 

        element.click()

        # 実行ボタンをクリック
        driver.find_element_by_id("MainContent_btnJikkou") .click()

        time.sleep(1) # 1秒待つ

        driver.find_element_by_id("hlkHome") .click()

        # 入力・照会タブをクリック
        driver.find_element_by_id("__tab_MainContent_TabContainer1_TabPanel6") .click()

        time.sleep(1)

        # HA0200:仕入照会を開く
        driver.find_element_by_xpath('//a[@href="HA/HA0410.aspx"]') .click()

        #管理部門を入れる

        element = driver.find_element_by_id("MainContent_txtSKanribumon")
        driver.execute_script("arguments[0].setAttribute('value' , '')" , element ) 
        Kanribumonbar_id = driver.find_element_by_id("MainContent_txtSKanribumon")
        Kanribumonbar_id.send_keys(output5)

        element = driver.find_element_by_id("MainContent_txtEKanribumon")
        driver.execute_script("arguments[0].setAttribute('value' , '')" , element ) 
        Kanribumonbar_id = driver.find_element_by_id("MainContent_txtEKanribumon")
        Kanribumonbar_id.send_keys(output6)

        #伝票日付を入れる

        element = driver.find_element_by_id("MainContent_txtSDenpyouHiduke")
        driver.execute_script("arguments[0].setAttribute('value' , '')" , element )
        Hidukebar_id = driver.find_element_by_id("MainContent_txtSDenpyouHiduke")
        Hidukebar_id.send_keys(output)

        element = driver.find_element_by_id("MainContent_txtEDenpyouHiduke")
        driver.execute_script("arguments[0].setAttribute('value' , '')" , element )
        Hidukebar2_id = driver.find_element_by_id("MainContent_txtEDenpyouHiduke")  
        Hidukebar2_id.send_keys(output2)

        # 表示ボタンをクリック
        driver.find_element_by_id("MainContent_btnHyouji") .click()

        # EXCEL出力をクリック

        element = driver.find_element_by_id("MainContent_lbnExcel")
        driver.execute_script("arguments[0].setAttribute('onclick' , '')" , element ) 

        element.click()

            # すべてボタンをクリック
        element = driver.find_element_by_id("MainContent_chkAll")
        #driver.execute_script("arguments[0].setAttribute('onclick' , '')" , element ) 

        element.click()

        # 実行ボタンをクリック
        driver.find_element_by_id("MainContent_lbnJikkou") .click()

        time.sleep(3)

        #driver.quit()


        return downloaded_file_path
     
     except Exception as e:
        st.write("エラーが発生しました: ", str(e))
        return None
     
st.title("test")

# サイドバーに入力ウィジェットを配置
output_date = st.sidebar.date_input("伝票日付開始:", datetime.date.today())
output = output_date.strftime("%Y%m%d")
output_date2 = st.sidebar.date_input("伝票日付終了:", datetime.date.today())
output2 = output_date2.strftime("%Y%m%d")
output3 = st.sidebar.text_input("ID:", "")
output4 = st.sidebar.text_input("パスワード:", "", type="password")
output5 = st.sidebar.slider("管理部門開始:", 195000, 195300, 195000)
output6 = st.sidebar.slider("管理部門終了:", 195000, 195300, 195300)

if st.sidebar.button("自動操作実行"):
    downloaded_file_path = automate_browser(output, output2, output3, output4, output5, output6)
    if downloaded_file_path:
        st.markdown(f"[ダウンロードファイル]({downloaded_file_path}) をダウンロード")


# メイン関数
def main():
    st.write("Excelファイルの加工アプリ")

    # ファイルアップロード
    st.write("売上Excelファイルのアップロード")
    sales_file = st.file_uploader("売上ファイルを選択してください", type=["xlsx"])

    st.write("仕入れExcelファイルのアップロード")
    purchase_file = st.file_uploader("仕入れファイルを選択してください", type=["xlsx"])

    # 実行ボタン
    if st.button("実行"):
        if sales_file and purchase_file:
            try:
                process_excel_files(sales_file, purchase_file)
                st.success("処理が完了しました。")
            except Exception as e:
                st.error(f"エラーが発生しました: {e}")
        else:
            st.warning("Excelファイルをアップロードしてください。")

# Excelファイルの処理関数
def process_excel_files(sales_file, purchase_file):
    # アップロードされたファイルをPandas DataFrameに読み込む
    sales_df = pd.read_excel(sales_file)
    purchase_df = pd.read_excel(purchase_file)

    # ここに、Excelファイルを処理するコードを書きます。
    # 上記で示したスクリプトを、この関数内で使用するように適応させます。
    # 以下に、変換されたコードの一部を示します。

    # sales_df の処理
    # ...

    # purchase_df の処理
    # ...

    # ピボットテーブルを作成し、Excelファイルに出力する
    # ...

    # 出力されたExcelファイルをダウンロードリンクとして提供する
    output_sales_path = os.path.join("output", "売掛ピボット.xlsx")
    output_purchase_path = os.path.join("output", "買掛ピボット.xlsx")
    st.markdown(f"[売掛ピボット.xlsx]({output_sales_path}) をダウンロード")
    st.markdown(f"[買掛ピボット.xlsx]({output_purchase_path}) をダウンロード")


if __name__ == "__main__":
    main()
