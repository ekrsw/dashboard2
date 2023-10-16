from dashboard_app.common.models.read_file_model import ReadExlFileModel
import pandas as pd
import os


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


# メンバリストファイル
member_list_path = r"\\mjs.co.jp\datas\CSC共有フォルダ\第47期 東京CSC第二グループ\47期SV共有\ダッシュボードアイテム"
member_list_file = 'メンバーリスト.xlsx'
member_list_file_path = os.path.join(member_list_path, member_list_file)

# ReadExlFileModelインスタンスを作成
rm = ReadExlFileModel()

# 本日の活動ファイルを読込み
df_activity = rm.read_today_activity_file()

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
df_2g = df_merge[df_merge['グループ'] == 2]
df_3g = df_merge[df_merge['グループ'] == 3]
df_n = df_merge[df_merge['グループ'] == 4]
df_other = df_merge[df_merge['グループ'] <= 1]

# グループごと、時間差ごとの件数を抽出
c_20_2g, c_30_2g, c_40_2g, c_60_2g, c_60plus_2g, not_included_2g = convert_to_num_of_cases_by_per_time(df_2g)
c_20_3g, c_30_3g, c_40_3g, c_60_3g, c_60plus_3g, not_included_3g = convert_to_num_of_cases_by_per_time(df_3g)
c_20_n, c_30_n, c_40_n, c_60_n, c_60plus_n, not_included_n = convert_to_num_of_cases_by_per_time(df_n)
c_20_other, c_30_other, c_40_other, c_60_other, c_60plus_other, not_included_other = convert_to_num_of_cases_by_per_time(df_other)

# 20分以内、40分以内を計算

# データを作成
data = {
    '指標集計対象': [df_2g.shape[0] - not_included_2g, df_3g.shape[0] - not_included_3g, df_n.shape[0] - not_included_n, df_other.shape[0] - not_included_other],
    '20分以内': [c_20_2g, c_20_3g, c_20_n, c_20_other],
    '40分以内': [c_20_2g + c_30_2g + c_40_2g, c_20_3g + c_30_3g + c_40_3g, c_20_n + c_30_n + c_40_n, c_20_other + c_30_other + c_40_other]
}

# カラム名とインデックスを指定してDataFrameを作成
df = pd.DataFrame(data, columns=['指標集計対象', '20分以内', '40分以内'], index=['第2G', '第3G', '長岡', 'その他'])

# DataFrameを表示
print(df)
