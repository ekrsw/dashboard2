import datetime as dt
import numpy as np
import os
import pandas as pd

from common.models.reporter_model import ReporterModel
from common.models.read_file_model import ReadExlFileModel
from common.models.calc_model import Calc

headless_mode = True
url = 'http://sccts7dxsql/ctreport'
id = 'eisuke_koresawa'

# メンバリストファイル
member_list_path = r"\\mjs.co.jp\datas\CSC共有フォルダ\第47期 東京CSC第二グループ\47期SV共有\ダッシュボードアイテム"
member_list_file = 'メンバーリスト.xlsx'
member_list_file_path = os.path.join(member_list_path, member_list_file)

# 'クローズ_本日.xlsm'ファイル
close_file = r'\\mjs.co.jp\datas\CSC共有フォルダ\第47期 東京CSC第二グループ\47期SV共有\ダッシュボードアイテム\クローズ_本日.xlsm'

# 活動ファイル
activity_path = r"\\mjs.co.jp\datas\CSC共有フォルダ\第47期 東京CSC第二グループ\47期SV共有\ダッシュボードアイテム"
activity_file = r"todays_activity.xlsx"
activity_file_path = os.path.join(activity_path, activity_file)

def get_today_acw_att_cph():
    # レポータ、ファイル読込み、計算インスタンスを作成
    rm = ReporterModel(url, id, headless_mode=headless_mode)
    rf = ReadExlFileModel()
    cl = Calc()

    # レポータから本日の各種情報を読込みDataFrameへ格納
    df = rm.get_today_op_as_dataframe()

    # メンバーリストファイルから、レポータの名前を正式な氏名のフォーマットに変換
    member_list = pd.read_excel(member_list_file_path)
    index_dict_reporter = member_list.set_index('レポータ')['氏名'].to_dict()
    df.index = df.index.map(index_dict_reporter)

    # クローズ_本日.xlsmファイルを読み込み、クローズのDataFrameを作成
    df_close = rf.read_today_close_file()

    # シフトデータを読込み、同時にメンバーリストファイルから、名前を正式な氏名のフォーマットに変換
    df_shift = rf.read_today_shift_file()
    index_dict_shift_table = member_list.set_index('Sweet')['氏名'].to_dict()
    df_shift.index = df_shift.index.map(index_dict_shift_table)

    # 当番データを読込み、同時にメンバーリストファイルから、名前を正式な氏名のフォーマットに変換
    df_duty = rf.read_today_duty_file()
    index_dict_duty_table = member_list.set_index('Sweet')['氏名'].to_dict()
    df_duty.index = df_duty.index.map(index_dict_duty_table)

    # クローズデータとレポータからスクレイピングしたデータをjoin
    df_join = df.join(df_close, how='outer').fillna(0)

    # 各カラムの合計値を計算して新しい行として追加
    df_join.loc['合計'] = df_join.sum()
    
    # ACW, ATT, CPHを計算してカラムを追加
    dataset = cl.calc_acw_att_cph(df_join)

    # シフトデータをjoin
    dataset = dataset.join(df_shift, how='outer').fillna('未設定')
    dataset = dataset.join(df_duty, how='outer').fillna('')
    
    # 必要なカラムのみ抽出
    dataset = dataset[['ｼﾌﾄ', 'ACW', 'ATT', 'CPH', 'ｸﾛｰｽﾞ', '業務ｽﾃｰﾀｽ']]

    # float型を'hh:mm:ss'形式の文字列に変換, CPHを小数点以下2桁でround
    dataset = cl.apply_to_acw_att_cph(dataset)

    # クローズとシフトでソート
    dataset = dataset.sort_values(by=['ｸﾛｰｽﾞ', 'ｼﾌﾄ'], ascending=[False, True])
    
    return dataset

def get_today_20_40():
    
    rf = ReadExlFileModel()
    df = rf.read_today_activity_file()
    df = df.iloc[:, 3:]
    return df

def get_today_acw_att_cph_to_html():
    df = get_today_acw_att_cph()
    df['ｵﾍﾟﾚｰﾀ'] = df.index
    df = df[['ｵﾍﾟﾚｰﾀ', 'ｼﾌﾄ', 'ACW', 'ATT', 'CPH', 'ｸﾛｰｽﾞ', '業務ｽﾃｰﾀｽ']]
    df_sum = df.loc['合計', :]
    df = df.drop('合計')

    dep_acw = df_sum['ACW']
    dep_att = df_sum['ATT']
    dep_cph = df_sum['CPH']

    html_table = df.to_html(index=False, classes="styled-table")
    html_table = html_table.replace(' style="text-align: right;"', '')
    now = dt.datetime.now()
    formatten_datetime = now.strftime('%Y/%m/%d %H:%M:%S')

    # CSSを追加してテーブルを整える
    html = f"""
    <!DOCTYPE html>
    <html lang="ja">
    <head>
        <meta charset="cp932">
        <meta http-equiv="refresh" content="60">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>本日のパフォーマンス指標</title>
        <link rel="stylesheet" href="styles.css"> <!-- CSSファイルを読み込む -->
    </head>
    <body>
        <p class="last-update-time">最終データ取得日時【{formatten_datetime}】</p>
        <h1>本日の部門パフォーマンス【目標: ACW 10, CPH 3】</h1>
        <div class="container">
            <div class="indicator-box">
                <p class="indicator-label">ACW</p>
                <p class="acw">{dep_acw}</p>
            </div>
            <div class="indicator-box">
                <p class="indicator-label">ATT</p>
                <p class="att">{dep_att}</p>
            </div>
            <div class="indicator-box">
                <p class="indicator-label">CPH</p>
                <p class="cph">{dep_cph}</p>
            </div>
        </div>
        <h1>本日の個人別パフォーマンス</h1>
            
        {html_table}
        </body>
        <div class="notes">
            <p>※CPHは暫定値です</p>
            <p>ページは1分ごと, データは3分ごとに更新されます。</p>
        </div>
        </html>
        """
    index_path_sv = r'\\mjs.co.jp\datas\CSC共有フォルダ\第47期 東京CSC第二グループ\47期SV共有\ダッシュボードアイテム'
    index_path_team = r'\\mjs.co.jp\datas\CSC共有フォルダ\第47期 東京CSC第二グループ\是澤チーム\是澤\日次指標'
    index_path_2g = r'\\mjs.co.jp\datas\CSC共有フォルダ\第47期 東京CSC第二グループ\リアルタイム個人別パフォーマンス'
    index_file_sv = os.path.join(index_path_sv, 'index.html')
    index_file_team = os.path.join(index_path_team, 'index.html')
    index_file_2g = os.path.join(index_path_2g, 'index.html')

    with open(index_file_sv, 'w') as f1, open(index_file_team, 'w') as f2, open(index_file_2g, 'w') as f3:
        f1.write(html)
        f2.write(html)
        f3.write(html)

