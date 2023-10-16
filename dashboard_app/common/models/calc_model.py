import datetime as dt
import numpy as np

class Calc(object):
    def __init__(self):
        pass

    def time_to_days(self, time_str):
        ''' hh:mm:ss 形式の時間を日数に変換する関数'''
        
        t = dt.datetime.strptime(time_str, "%H:%M:%S")
        return (t.hour + t.minute / 60 + t.second / 3600) / 24
    
    def add_sum_to_df(self, dataframe):
        pass

    def calc_acw_att_cph(self, dataframe):
        '''スクレイピングした指標から、ACW, ATT, CPHを計算してDataFrameにカラムを追加して返す
        0除算を避けるために、0の場合はいったん1にreplaceしている'''

        columns_to_convert = ['着信通話時間の合計','発信通話時間の合計', 'ワークタイムの合計', '着信後処理時間の合計', '発信後処理時間の合計', '離席時間の合計', '事前準備時間の合計', '一時離席時間の合計', '転送可時間の合計', '着信通話時間の合計(外線)', '発信通話時間の合計(外線)', 'ｸﾛｰｽﾞ']

        # 計算に使用するカラムの値をfloatに変換
        for col in columns_to_convert:
            dataframe[col] = dataframe[col].astype(float)
        
        # 実際の計算
        dataframe['ACW'] = (dataframe['ワークタイムの合計'] + dataframe['着信後処理時間の合計'] + dataframe['発信後処理時間の合計'] + dataframe['事前準備時間の合計'] + dataframe['転送可時間の合計'] + dataframe['一時離席時間の合計']) / dataframe['ｸﾛｰｽﾞ'].replace(0, 1)
        dataframe.loc[dataframe['ｸﾛｰｽﾞ']==0, 'ACW'] = 0
        dataframe['ATT'] = (dataframe['着信通話時間の合計(外線)'] + dataframe['発信通話時間の合計(外線)']) / dataframe['ｸﾛｰｽﾞ'].replace(0, 1)
        dataframe.loc[dataframe['ｸﾛｰｽﾞ']==0, 'ATT'] = 0
        _tmp = (dataframe['着信通話時間の合計'] + dataframe['発信通話時間の合計'] + dataframe['ワークタイムの合計'] + dataframe['着信後処理時間の合計'] + dataframe['発信後処理時間の合計'] + dataframe['離席時間の合計'] + dataframe['事前準備時間の合計'] + dataframe['一時離席時間の合計']) * 24
        dataframe['CPH'] = np.where(
            _tmp == 0,
            0,
            dataframe['ｸﾛｰｽﾞ'] / _tmp
        )
        
        dataframe = dataframe.replace(np.inf, 0)

        return dataframe
    
    def calc_ratio_direct(self, df_arg):
        dataframe = df_arg.copy()
        columns_to_convert = ['電話対応数', '直受け']

        # 計算に使用するカラムの値をfloatに変換
        for col in columns_to_convert:
            dataframe[col] = dataframe[col].astype(float)

        # 実際の計算
        dataframe['直受け'] = dataframe['直受け'] / dataframe['電話対応数'].replace(0, 1)
        dataframe.loc[dataframe['電話対応数']==0, '直受け'] = 0
        
        return dataframe

    def calc_ratio_20_40(self, df_arg):
        dataframe = df_arg.copy()
        columns_to_convert = ['指標集計対象', '20分以内', '40分以内']

        # 計算に使用するカラムの値をfloatに変換
        for col in columns_to_convert:
            dataframe[col] = dataframe[col].astype(float)
        
        # 実際の計算
        dataframe['20分以内'] = dataframe['20分以内'] / dataframe['指標集計対象'].replace(0, 1)
        dataframe.loc[dataframe['指標集計対象']==0, '20分以内'] = 0
        dataframe['40分以内'] = dataframe['40分以内'] / dataframe['指標集計対象'].replace(0, 1)
        dataframe.loc[dataframe['指標集計対象']==0, '40分以内'] = 0
        
        return dataframe
    
    def _float_to_hms(self, value):
        '''1日を1としたfloat型を'hh:mm:ss'形式の文字列に変換'''
        
        # 1日が1なので、24を掛けて時間単位に変換
        hours = value * 24

        # 時間の整数部分
        h = int(hours)

        # 残りの部分を分単位に変換
        minutes = (hours - h) * 60

        # 分の整数部分
        m = int(minutes)

        # 残りの部分を秒単位に変換
        seconds = (minutes - m) * 60

        # 秒の整数部分
        s = int(seconds)

        return f"{h:02}:{m:02}:{s:02}"
    
    def apply_to_acw_att_cph(self, dataframe):
        dataframe['ACW'] = dataframe['ACW'].apply(self._float_to_hms)
        dataframe['ATT'] = dataframe['ATT'].apply(self._float_to_hms)
        dataframe['CPH'] = dataframe['CPH'].apply(lambda x: round(x, 2))
        dataframe['ｸﾛｰｽﾞ'] = dataframe['ｸﾛｰｽﾞ'].apply(lambda x: int(x))

        return dataframe
    
    def apply_to_percent(self, dataframe):
        dataframe['20分以内'] = dataframe['20分以内'].apply(lambda x: round(x * 100, 1))
        dataframe['40分以内'] = dataframe['40分以内'].apply(lambda x: round(x * 100, 1))
        
        return dataframe
