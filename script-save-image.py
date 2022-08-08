import pandas as pd
import requests


def run():
  df = pd.read_excel("./data-api-ekyc.xlsx", "100 CCCD", engine='openpyxl', dtype = str)
  for index, row in df.iterrows():
    # print(row['Ảnh CMND mặt trước'], row['Ảnh CMND mặt sau'], row['Ảnh chân dung'])
    try:
      img_data = requests.get(row['Ảnh chân dung']).content
      print('./image/' + row['Ảnh chân dung'].split('/')[-1]);
      with open('./image/' + row['Ảnh chân dung'].split('/')[-1], 'wb') as handler:
          handler.write(img_data)
    except:
        pass
run()