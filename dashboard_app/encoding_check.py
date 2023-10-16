import chardet
import os

duty_path = r"\\mjs.co.jp\datas\CSC共有フォルダ\第47期 東京CSC第二グループ\47期SV共有\ダッシュボードアイテム\当番表"
duty_file = f"202309_Campaign_ScheduleList.csv"
duty_file_path = os.path.join(duty_path, duty_file)

with open(duty_file_path, 'rb') as f:
    result = chardet.detect(f.read())

print(result['encoding'])