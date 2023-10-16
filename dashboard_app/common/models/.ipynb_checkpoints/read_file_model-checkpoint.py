import datetime as dt
import os
import pandas as pd

class ReadExlFileModel(object):
    def __init__(self):
        today_obj = dt.date.today()
        self.today_str = today_obj.strftime("%Y%m%d")
        self.this_month_str = today_obj.strftime("%Y%m")
        self.date_str =today_obj.strftime("%d")

        # 'クローズ_本日.xlsm'ファイルを読み込み
        self.close_file = r'\\mjs.co.jp\datas\CSC共有フォルダ\第47期 東京CSC第二グループ\47期SV共有\ダッシュボードアイテム\クローズ_本日.xlsm'

        # シフトファイル
        shift_path = r"\\mjs.co.jp\datas\CSC共有フォルダ\第47期 東京CSC第二グループ\47期SV共有\ダッシュボードアイテム\シフト表"
        shift_file = f"{self.this_month_str}_Campaign_ScheduleList.csv"
        self.shift_file_path = os.path.join(shift_path, shift_file)
        
        # 当番ファイル
        duty_path = r"\\mjs.co.jp\datas\CSC共有フォルダ\第47期 東京CSC第二グループ\47期SV共有\ダッシュボードアイテム\当番表"
        duty_file = f"{self.this_month_str}_Campaign_ScheduleList.csv"
        self.duty_file_path = os.path.join(duty_path, duty_file)
        
        # 活動ファイル
        activity_path = r"\\mjs.co.jp\datas\CSC共有フォルダ\第47期 東京CSC第二グループ\47期SV共有\ダッシュボードアイテム"
        activity_file = r"todays_activity.xlsx"
        self.activity_file_path = os.path.join(activity_path, activity_file)

    def read_today_close_file(self):

        df = pd.read_excel(self.close_file)

        # 最初の3列をスキップし、5列目をインデックスとして設定します
        df = df.iloc[:, 3:].set_index(df.columns[5])
        
        df.reset_index(inplace=True)
        df.set_index(['完了日時'], inplace=True)

        df = df.loc[self.today_str]
        df.reset_index(inplace=True)
        df.set_index(['所有者'], inplace=True)

        counts = df.index.value_counts()
        df = pd.DataFrame(counts).reset_index()
        df.columns = ['ｵﾍﾟﾚｰﾀ', 'ｸﾛｰｽﾞ']
        df = df.set_index(df.columns[0])

        return df
    
    def read_today_shift_file(self):
        # シフトのCSVファイルからシフトデータを読み込み、名前をインデックスに設定する。
        # CSVファイルを読み込む。ヘッダーは3行目（0-indexed）にある。
        shift_df = pd.read_csv(self.shift_file_path, skiprows=2, header=1, index_col=1, quotechar='"', encoding='shift_jis')
        # 最後の1列を削除
        shift_df = shift_df.iloc[:, :-1]
        # "組織名"、"従業員ID"、"種別" の列を削除
        shift_df = shift_df.drop(columns=["組織名", "従業員ID", "種別"])
        shift_df = shift_df[[self.date_str]]
        shift_df.columns = ['ｼﾌﾄ']

        return shift_df
    
    def read_today_duty_file(self):
        # 当番のCSVファイルから当番データを読み込み、名前をインデックスに設定する。
        # CSVファイルを読み込む。ヘッダーは3行目（0-indexed）にある。
        shift_df = pd.read_csv(self.duty_file_path, skiprows=2, header=1, index_col=1, quotechar='"', encoding='cp932')
        # 最後の1列を削除
        shift_df = shift_df.iloc[:, :-1]
        # "組織名"、"従業員ID"、"種別" の列を削除
        shift_df = shift_df.drop(columns=["組織名", "従業員ID", "種別"])
        shift_df = shift_df[[str(int(self.date_str))]]
        shift_df.columns = ['業務ｽﾃｰﾀｽ']

        return shift_df
    
    def read_today_activity_file(self):
        # 活動ファイルからデータを読込みDataframeを返す。
        df = pd.read_excel(self.activity_file_path)
        
        return df