from common.models.read_file_model import ReadExlFileModel

rm = ReadExlFileModel()
df = rm.read_today_activity_file()
print(df)