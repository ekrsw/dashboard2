import datetime as dt
import numpy as np
import os
import pandas as pd
import string

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
close_file = r'\\mjs.co.jp\datas\CSC共有フォルダ\第47期 東京CSC第二グループ\47期SV共有\ダッシュボードアイテム\todays_op.xlsm'

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

def get_today_activity_df():
    '''グループごとのDataframeを取得'''
    
    rf = ReadExlFileModel()
    
    # 本日の活動ファイルを読込み
    df_activity = rf.read_today_activity_file()
    
    # 受付けタイプ「直受け」「折返し」「留守電」のみ残す
    df_activity = df_activity[(df_activity['受付タイプ (関連) (サポート案件)'] == '直受け') | (df_activity['受付タイプ (関連) (サポート案件)'] == '折返し') | (df_activity['受付タイプ (関連) (サポート案件)'] == '留守電')]

    # メンバーリストを読込み、'氏名'、'グループ'のカラムのみにする。
    df_member = pd.read_excel(member_list_file_path)
    df_member = df_member[['氏名', 'グループ']]

    # 活動DataFrameとメンバーリストDataFrameを'所有者'と'氏名'をキーにしてマージ
    df_merge = df_activity.merge(df_member, left_on='所有者 (関連) (サポート案件)', right_on='氏名', how='left')

    # 案件番号、登録日時でソート
    df_merge.sort_values(by=['案件番号 (関連) (サポート案件)', '登録日時'], inplace=True)

    # 同一案件番号の最初の活動のみ残して他は削除  
    df_merge.drop_duplicates(subset='案件番号 (関連) (サポート案件)', keep='first', inplace=True)
    
    # サポート案件の登録日時と、活動の登録日時をPandas Datetime型に変換して、差分を'時間差'カラムに格納、NaNは０変換
    df_merge['登録日時 (関連) (サポート案件)'] = pd.to_datetime(df_merge['登録日時 (関連) (サポート案件)'])
    df_merge['登録日時 (関連) (サポート案件)'] = pd.to_datetime(df_merge['登録日時 (関連) (サポート案件)'])
    df_merge['時間差'] = (df_merge['登録日時'] - df_merge['登録日時 (関連) (サポート案件)']).abs()
    df_merge = df_merge.fillna(0)

    # グループごとのDataFrameに分割
    df_1g = df_merge[df_merge['グループ'] == 1]
    df_2g = df_merge[df_merge['グループ'] == 2]
    df_3g = df_merge[df_merge['グループ'] == 3]
    df_n = df_merge[df_merge['グループ'] == 4]
    df_other = df_merge[df_merge['グループ'] <= 1]

    return df_1g, df_2g, df_3g, df_n, df_other


def convert_to_num_of_cases_by_per_time(df):
    '''DataFrameを以下のルールで振分けて、件数を返す
        c_20: 20分以内
        c_30: 20分超、30分以内
        c_40: 30分超、40分以内
        c_60: 40分超、60分以内 かつ '指標に含める'
        c_60plus: 60分超 かつ '指標に含める'
        not_included: 指標に含めない(40分超、60分以内 かつ 60分超)
    '''

    c_20 = df[df['時間差'] <= pd.Timedelta(minutes=20)].shape[0]
    c_30 = df[(pd.Timedelta(minutes=20) < df['時間差']) & (df['時間差'] <= pd.Timedelta(minutes=30))].shape[0]
    c_40 = df[(pd.Timedelta(minutes=30) < df['時間差']) & (df['時間差'] <= pd.Timedelta(minutes=40))].shape[0]
    c_60 = df[(pd.Timedelta(minutes=40) < df['時間差']) & (df['時間差'] <= pd.Timedelta(minutes=60)) & (df['指標に含めない (関連) (サポート案件)'] == 'いいえ')].shape[0]
    not_included = df[(pd.Timedelta(minutes=40) < df['時間差']) & (df['時間差'] <= pd.Timedelta(minutes=60)) & (df['指標に含めない (関連) (サポート案件)'] == 'はい')].shape[0]
    c_60plus = df[(df['時間差'] > pd.Timedelta(minutes=60)) & (df['指標に含めない (関連) (サポート案件)'] == 'いいえ')].shape[0]
    not_included = not_included + df[(df['時間差'] > pd.Timedelta(minutes=60)) & (df['指標に含めない (関連) (サポート案件)'] == 'はい')].shape[0]
    
    return c_20, c_30, c_40, c_60, c_60plus, not_included


def get_today_direct_df():

    df_1g, df_2g, df_3g, df_n, df_other = get_today_activity_df()
    
    # グループごとの直受け数を抽出
    count_direct_ratio_1g = df_1g[df_1g['受付タイプ (関連) (サポート案件)']=='直受け'].shape[0]
    count_direct_ratio_2g = df_2g[df_2g['受付タイプ (関連) (サポート案件)']=='直受け'].shape[0]
    count_direct_ratio_3g = df_3g[df_3g['受付タイプ (関連) (サポート案件)']=='直受け'].shape[0]
    count_direct_ratio_n = df_n[df_n['受付タイプ (関連) (サポート案件)']=='直受け'].shape[0]
    count_direct_ratio_other = df_other[df_other['受付タイプ (関連) (サポート案件)']=='直受け'].shape[0]

    # データを作成
    data = {
        '電話対応数': [df_2g.shape[0], df_3g.shape[0], df_n.shape[0], df_other.shape[0]],
        '直受け': [count_direct_ratio_2g, count_direct_ratio_3g, count_direct_ratio_ｎ, count_direct_ratio_other]
    }

    # カラム名とインデックスを指定してDataFrameを作成
    df = pd.DataFrame(data, columns=['電話対応数', '直受け'], index=['第2G', '第3G', '長岡', 'その他'])

    return df


def get_today_20_40_df():
    
    df_1g, df_2g, df_3g, df_n, df_other = get_today_activity_df()

    # グループごと、時間差ごとの件数を抽出
    c_20_2g, c_30_2g, c_40_2g, c_60_2g, c_60plus_2g, not_included_2g = convert_to_num_of_cases_by_per_time(df_2g)
    c_20_3g, c_30_3g, c_40_3g, c_60_3g, c_60plus_3g, not_included_3g = convert_to_num_of_cases_by_per_time(df_3g)
    c_20_n, c_30_n, c_40_n, c_60_n, c_60plus_n, not_included_n = convert_to_num_of_cases_by_per_time(df_n)
    c_20_other, c_30_other, c_40_other, c_60_other, c_60plus_other, not_included_other = convert_to_num_of_cases_by_per_time(df_other)

    # データを作成
    data = {
        '指標集計対象': [df_2g.shape[0] - not_included_2g, df_3g.shape[0] - not_included_3g, df_n.shape[0] - not_included_n, df_other.shape[0] - not_included_other],
        '20分以内': [c_20_2g, c_20_3g, c_20_n, c_20_other],
        '40分以内': [c_20_2g + c_30_2g + c_40_2g, c_20_3g + c_30_3g + c_40_3g, c_20_n + c_30_n + c_40_n, c_20_other + c_30_other + c_40_other]
    }

    # カラム名とインデックスを指定してDataFrameを作成
    df = pd.DataFrame(data, columns=['指標集計対象', '20分以内', '40分以内'], index=['第2G', '第3G', '長岡', 'その他'])

    return df


def convert_time_format(time_str):
    try:
        # 時間を時間、分、秒に分割
        hh, mm, ss = map(int, time_str.split(':'))
        
        # 時間を分に変換し、分を合計
        total_minutes = hh * 60 + mm
        
        # 新しい時間フォーマットを作成
        new_time_str = f'{total_minutes:02d}:{ss:02d}'
        return new_time_str
    except ValueError:
        # 不正な時間形式の場合はエラーメッセージを表示
        return "Invalid time format"


def get_today_acw_att_cph_to_html():
    cl = Calc()

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

    # 電話対応数、直受け率のDataFrameを作成
    df_direct_count = get_today_direct_df()
    df_direct_ratio = cl.calc_ratio_direct(df_direct_count)
    df_direct_ratio['直受け'] = df_direct_ratio['直受け'].apply(lambda x: round(x * 100, 1))

    # 指標集計対象、20分以内、40分以内のDataFrameを作成
    df_count = get_today_20_40_df()
    df_ratio = cl.calc_ratio_20_40(df_count)
    df_ratio = cl.apply_to_percent(df_ratio)

    ratio_direct_2g = df_direct_ratio.loc['第2G', '直受け']
    ratio_direct_3g = df_direct_ratio.loc['第3G', '直受け']
    ratio_direct_n = df_direct_ratio.loc['長岡', '直受け']

    ratio_20_2g = df_ratio.loc['第2G', '20分以内']
    ratio_20_3g = df_ratio.loc['第3G', '20分以内']
    ratio_20_n = df_ratio.loc['長岡', '20分以内']
    ratio_40_2g = df_ratio.loc['第2G', '40分以内']
    ratio_40_3g = df_ratio.loc['第3G', '40分以内']
    ratio_40_n = df_ratio.loc['長岡', '40分以内']

    count_20_2g = df_count.loc['第2G', '20分以内']
    count_20_3g = df_count.loc['第3G', '20分以内']
    count_20_n = df_count.loc['長岡', '20分以内']
    count_40_2g = df_count.loc['第2G', '40分以内']
    count_40_3g = df_count.loc['第3G', '40分以内']
    count_40_n = df_count.loc['長岡', '40分以内']

    count_all_2g = df_count.loc['第2G', '指標集計対象']
    count_all_3g = df_count.loc['第3G', '指標集計対象']
    count_all_n = df_count.loc['長岡', '指標集計対象']

    monitor_acw = convert_time_format(dep_acw)
    monitor_att = convert_time_format(dep_att)

    # ダッシュボード用のテンプレート読込み
    with open(r'C:\Users\eisuke_koresawa\Desktop\project\dashboard2\dashboard_app\templates\templates_dashboard.txt', 'r') as template_file:
        t_dashboard = string.Template(template_file.read())
    
    # CSSを追加してテーブルを整える
    html_to_dashboard = t_dashboard.substitute(formatten_datetime=formatten_datetime,
                               dep_acw=dep_acw,
                               dep_att=dep_att,
                               dep_cph=dep_cph,
                               ratio_20_2g=ratio_20_2g,
                               ratio_20_3g=ratio_20_3g,
                               ratio_20_n=ratio_20_n,
                               ratio_40_2g=ratio_40_2g,
                               ratio_40_3g=ratio_40_3g,
                               ratio_40_n=ratio_40_n,
                               count_20_2g=count_20_2g,
                               count_20_3g=count_20_3g,
                               count_20_n=count_20_n,
                               count_40_2g=count_40_2g,
                               count_40_3g=count_40_3g,
                               count_40_n=count_40_n,
                               count_all_2g=count_all_2g,
                               count_all_3g=count_all_3g,
                               count_all_n=count_all_n,
                               html_table=html_table)

    index_path_sv = r'\\mjs.co.jp\datas\CSC共有フォルダ\第47期 東京CSC第二グループ\47期SV共有\ダッシュボードアイテム'
    index_path_team = r'\\mjs.co.jp\datas\CSC共有フォルダ\第47期 東京CSC第二グループ\是澤チーム\是澤\日次指標'
    index_path_2g = r'\\mjs.co.jp\datas\CSC共有フォルダ\第47期 東京CSC第二グループ\リアルタイム個人別パフォーマンス'
    index_file_sv = os.path.join(index_path_sv, 'index.html')
    index_file_team = os.path.join(index_path_team, 'index.html')
    index_file_2g = os.path.join(index_path_2g, 'index.html')

    with open(index_file_sv, 'w') as f1, open(index_file_team, 'w') as f2, open(index_file_2g, 'w') as f3:
        f1.write(html_to_dashboard)
        f2.write(html_to_dashboard)
        f3.write(html_to_dashboard)

    # モニター用テンプレートの読込み
    with open(r'C:\Users\eisuke_koresawa\Desktop\project\dashboard2\dashboard_app\templates\templates_monitor.txt', 'r') as template_file:
        t_monitor = string.Template(template_file.read())
    
    # CSSを追加してテーブルを整える
    html_to_monitor = t_monitor.substitute(formatten_datetime=formatten_datetime,
                               dep_acw=monitor_acw,
                               dep_att=monitor_att,
                               dep_cph=dep_cph,
                               ratio_direct_2g=ratio_direct_2g,
                               ratio_direct_3g=ratio_direct_3g,
                               ratio_direct_n=ratio_direct_n,
                               ratio_20_2g=ratio_20_2g,
                               ratio_20_3g=ratio_20_3g,
                               ratio_20_n=ratio_20_n,
                               ratio_40_2g=ratio_40_2g,
                               ratio_40_3g=ratio_40_3g,
                               ratio_40_n=ratio_40_n)

    index_monitor_file_sv = os.path.join(index_path_sv, 'monitor.html')
    with open(index_monitor_file_sv, 'w') as f_monitor:
        f_monitor.write(html_to_monitor)


