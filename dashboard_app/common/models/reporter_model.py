import datetime as dt
import glob
import os
import pandas as pd
import shutil
import time

# Webスクレイピング関係ライブラリ
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
# Chromeのバージョンを確認して、適切なDriverをダウンロード・インストールするライブラリ
from webdriver_manager.chrome import ChromeDriverManager


class ReporterModel(object):
    """Base scraping model."""
    def __init__(self, url, id, headless_mode=True):
        self.url = url
        self.id = id
        self.headless_mode = headless_mode
        self.today_op_template = 'Today_OP'

        options = Options()

        # ブラウザを表示させない。
        if headless_mode:
            options.add_argument('--headless')
        
        # コマンドプロンプトのログを表示させない。
        options.add_argument('--disable-logging')
        options.add_argument('--log-level=3')
        options.add_experimental_option('excludeSwitches', ['enable-logging'])

        # headlessモードでもダウンロードできるようにする設定
        options.add_experimental_option('prefs', {'download.prompt_for_download': False})

        # Chrome Driver Managerを引数に設定する。
        # headlessモードでもダウンロードできるようにする設定
        self.download_path = 'C:\\Users\\{}\\Downloads'.format(os.getlogin())
        try:
            self.driver = webdriver.Chrome(ChromeDriverManager().install(), options=options)
        except:
            self.driver = webdriver.Chrome('C:/driver/chromedriver.exe', options=options)

        self.driver.command_executor._commands["send_command"] = (
            'POST',
            '/session/$sessionId/chromium/send_command'
        )
        self.driver.execute(
            'send_command',
            params={
                'cmd': 'Page.setDownloadBehavior',
                'params': {'behavior': 'allow', 'downloadPath': self.download_path}
            }
        )

        self.driver.implicitly_wait(5)

        self._login()
    
    def __del__(self):
        """デストラクタ。ドライバーをclose"""
        if hasattr(self, 'driver'):
            self.driver.close()
    
    def _login(self):
        """レポータに接続してログイン"""
        self.driver.get(self.url)
        input_operator_id = self.driver.find_element_by_id('logon-operator-id')
        input_operator_id.send_keys(self.id)
        self.driver.find_element_by_id('logon-btn').click()
    
    def _call_template(self, TEMPLATE, from_date, to_date):
        """テンプレート呼び出し、指定の集計期間を表示"""

        # テンプレート呼び出し
        self.driver.find_element_by_id('template-title-span').click()
        el = self.driver.find_element_by_id('template-download-select')
        s = Select(el)
        s.select_by_value(TEMPLATE)
        self.driver.find_element_by_id('template-creation-btn').click()

        # 集計期間のfromをクリアしてfrom_dateを送信
        self.driver.find_element_by_id('panel-td-input-from-date-0').send_keys(Keys.CONTROL + 'a')
        self.driver.find_element_by_id('panel-td-input-from-date-0').send_keys(Keys.DELETE)
        self.driver.find_element_by_id('panel-td-input-from-date-0').send_keys(from_date.strftime('%Y/%m/%d'))

        # 集計期間のtoをクリアしてto_dateを送信
        self.driver.find_element_by_id('panel-td-input-to-date-0').send_keys(Keys.CONTROL + 'a')
        self.driver.find_element_by_id('panel-td-input-to-date-0').send_keys(Keys.DELETE)
        self.driver.find_element_by_id('panel-td-input-to-date-0').send_keys(to_date.strftime('%Y/%m/%d'))

        # レポート作成
        self.driver.find_element_by_id('panel-td-create-report-0').click()
        time.sleep(5)

    def _get_table_as_array(self, TEMPLATE, from_date, to_date):
        """テンプレートを表示して、tableを2次元配列で取得"""
        self._call_template(TEMPLATE, from_date, to_date)

        html = self.driver.page_source.encode('utf-8')
        soup = BeautifulSoup(html, 'lxml')

        data_table = []

        # headerのリストを作成
        header_table = soup.find(id='normal-list1-dummy-0-table-head-table')
        xmp = header_table.thead.tr.find_all('xmp')
        header_list = [i.string for i in xmp]
        data_table.append(header_list)

        # bodyのリストを作成
        body_table = soup.find(id='normal-list1-dummy-0-table-body-table')
        tr = body_table.tbody.find_all('tr')
        for td in tr:
            xmp = td.find_all('xmp')
            row = [i.string for i in xmp]
            data_table.append(row)
        
        return data_table


    def get_csv_DL(self, TEMPLATE, from_date, to_date, save_path, file_name):
        """テンプレートで表示された内容をダウンロード"""

        self._call_template(TEMPLATE, from_date, to_date)

        # ダウンロードボタンからCSVをダウンロード  
        now = dt.datetime.now()
        self.driver.find_element_by_id('top-download').click()
        num = int(now.strftime('%Y%m%d%H%M%S'))
        
        time.sleep(1)

        # ダウンロードしたファイルを指定の場所へ移動
        while True:
            download_file_name = '{}-{}.csv'.format(file_name, str(num))
            download_file = os.path.join(self.download_path, download_file_name)
            if glob.glob(download_file):
                # ファイルを移動
                shutil.move(download_file, save_path)

                # ファイル名を変更。既に同一ファイル名があれば削除。
                new_file_name = '{}_{}.csv'.format(to_date.strftime('%Y%m%d') ,file_name)
                new_file_path = os.path.join(save_path, new_file_name)
                old_file_path = os.path.join(save_path, download_file_name)
                if os.path.exists(new_file_path):
                    os.remove(new_file_path)
                os.rename(old_file_path, new_file_path)
                break
            num += 1
    
    def get_today_op_as_dataframe(self):
        '''レポータから本日のOP分析情報をDataFrameで取得する'''

        date_obj = dt.date.today()
        data_array = self._get_table_as_array(self.today_op_template, date_obj, date_obj)

        df = pd.DataFrame(data_array[1:], columns=data_array[0])
        df.set_index(df.columns[0], inplace=True)

        def time_to_days(time_str):
            ''' hh:mm:ss 形式の時間を日数に変換する関数'''
            
            t = dt.datetime.strptime(time_str, "%H:%M:%S")
            return (t.hour + t.minute / 60 + t.second / 3600) / 24

        df = df.applymap(time_to_days)

        return df


