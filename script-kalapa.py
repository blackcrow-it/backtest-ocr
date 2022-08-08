import pandas as pd
import xlsxwriter
# https://www.educative.io/answers/what-is-sequencematcher-in-python
from difflib import SequenceMatcher
from api import check_ekyc_image_url_kalapa
import base64
import requests
import datetime

"""
Hàm viết dữ liệu vào file excel
"""
def write_info_to_excel():
  df = pd.read_excel("./data-api-ekyc.xlsx", "100 CCCD", engine='openpyxl', dtype = str)

  workbook = xlsxwriter.Workbook('./result/data-result-ekyc-kalapa.xlsx')
  worksheet = workbook.add_worksheet('100 CCCD')
  bold = workbook.add_format({'bold': True})
  error = workbook.add_format({'bg_color': 'red', 'font_color': 'white'})
  warning = workbook.add_format({'bg_color': 'yellow'})
  worksheet.write('A1', 'STT', bold)
  worksheet.write('B1', 'Tên', bold)
  worksheet.write('C1', 'Tên OCR', bold)
  worksheet.write('D1', 'Score Tên', bold)
  worksheet.write('E1', 'Ngày sinh', bold)
  worksheet.write('F1', 'Ngày sinh OCR', bold)
  worksheet.write('G1', 'Score Ngày sinh', bold)
  worksheet.write('H1', 'CCCD', bold)
  worksheet.write('I1', 'CCCD OCR', bold)
  worksheet.write('J1', 'Score CCCD', bold)
  worksheet.write('K1', 'Ngày cấp', bold)
  worksheet.write('L1', 'Ngày cấp OCR', bold)
  worksheet.write('M1', 'Score Ngày cấp', bold)
  worksheet.write('N1', 'Nơi cấp', bold)
  worksheet.write('O1', 'Nơi cấp OCR', bold)
  worksheet.write('P1', 'Score Nơi cấp', bold)
  worksheet.write('Q1', 'Quê quán', bold)
  worksheet.write('R1', 'Quê quán OCR', bold)
  worksheet.write('S1', 'Score Quê quán', bold)
  worksheet.write('T1', 'Thường trú', bold)
  worksheet.write('U1', 'Thường trú OCR', bold)
  worksheet.write('V1', 'Score Thường trú', bold)
  worksheet.write('W1', 'Kiểm tra ảnh mặt trước', bold)
  worksheet.write('X1', 'Kiểm tra ảnh mặt sau', bold)
  worksheet.write('Y1', 'Kiểm tra ảnh selfie', bold)
  worksheet.write('Z1', 'Điểm matching eKyc', bold)
  worksheet.write('AA1', 'Score Trung bình', bold)
  worksheet.write('AB1', 'Thời gian xử lý', bold)
  worksheet.write('AC1', 'Ảnh chân dung', bold)
  worksheet.write('AD1', 'Ảnh CCCD mặt trước', bold)
  worksheet.write('AE1', 'Ảnh CCCD mặt sau', bold)

  row_sheet_result = 1
  col_sheet_result = 0
  for index, row in df.iterrows():
    print(row["STT"], row["Tên"])
    worksheet.write(row_sheet_result, col_sheet_result, row['STT'])

    print("Checking OCR ...")
    # Đọc dữ liệu OCR trên CMND
    try:
      recognition, time_ocr = check_ekyc_image_url_kalapa(row['Ảnh CMND mặt trước'], row['Ảnh CMND mặt sau'], row['Ảnh chân dung'])
      if recognition['status_code'] == 200:
        data = recognition['data']['idCardInfo']
        date_time_obj = datetime.datetime.strptime(row['DOB'], '%Y-%m-%d')
        row['DOB'] = date_time_obj.strftime("%d/%m/%Y")
        date_time_obj = datetime.datetime.strptime(row['Ngày cấp'], '%Y-%m-%d')
        row['Ngày cấp'] = date_time_obj.strftime("%d/%m/%Y")

        list_score = []
        score_name = percent_sequence_matcher(row['Tên'], data['name'])
        score_birthday = percent_sequence_matcher(row['DOB'], data['birthday'])
        score_id = percent_sequence_matcher(str(row['CCCD']), data['id_number'])
        score_issue_date = percent_sequence_matcher(row['Ngày cấp'], data['doi'])
        score_issue_at = percent_sequence_matcher(row['Nơi cấp'], data['poi'])
        # list_score.append(percent_sequence_matcher(row['Quê quán'], data['home_town']))
        score_address = percent_sequence_matcher(row['Thường trú'], data['resident'])

        list_score.append(score_name)
        list_score.append(score_birthday)
        list_score.append(score_id)
        list_score.append(score_issue_date)
        list_score.append(score_issue_at)
        # list_score.append(percent_sequence_matcher(row['Quê quán'], data['home_town']))
        list_score.append(score_address)
        score = round(mean(list_score))

        # Ghi vào bảng
        worksheet.write(row_sheet_result, col_sheet_result + 1, str(row['Tên']))
        worksheet.write(row_sheet_result, col_sheet_result + 2, str(data['name']))
        worksheet.write(row_sheet_result, col_sheet_result + 3, str(score_name))
        worksheet.write(row_sheet_result, col_sheet_result + 4, str(row['DOB']))
        worksheet.write(row_sheet_result, col_sheet_result + 5, str(data['birthday']))
        worksheet.write(row_sheet_result, col_sheet_result + 6, str(score_birthday))
        worksheet.write(row_sheet_result, col_sheet_result + 7, str(row['CCCD']))
        worksheet.write(row_sheet_result, col_sheet_result + 8, str(data['id_number']))
        worksheet.write(row_sheet_result, col_sheet_result + 9, str(score_id))
        worksheet.write(row_sheet_result, col_sheet_result + 10, str(row['Ngày cấp']))
        worksheet.write(row_sheet_result, col_sheet_result + 11, str(data['doi']))
        worksheet.write(row_sheet_result, col_sheet_result + 12, str(score_issue_date))
        worksheet.write(row_sheet_result, col_sheet_result + 13, str(row['Nơi cấp']))
        worksheet.write(row_sheet_result, col_sheet_result + 14, str(data['poi']))
        worksheet.write(row_sheet_result, col_sheet_result + 15, str(score_issue_at))
        worksheet.write(row_sheet_result, col_sheet_result + 16, str(row['Quê quán']))
        worksheet.write(row_sheet_result, col_sheet_result + 17, str(data['home']))
        worksheet.write(row_sheet_result, col_sheet_result + 18, None)
        worksheet.write(row_sheet_result, col_sheet_result + 19, str(row['Thường trú']))
        worksheet.write(row_sheet_result, col_sheet_result + 20, str(data['resident']))
        worksheet.write(row_sheet_result, col_sheet_result + 21, str(score_address))
        worksheet.write(row_sheet_result, col_sheet_result + 22, str(recognition['data']['error']['front']['message']))
        worksheet.write(row_sheet_result, col_sheet_result + 23, str(recognition['data']['error']['back']['message']))
        if 'selfie' in recognition['data']['error']:
          worksheet.write(row_sheet_result, col_sheet_result + 24, str(recognition['data']['error']['selfie']['message']))
        if recognition['data']['selfieCheck'] != None:
          worksheet.write(row_sheet_result, col_sheet_result + 25, str(recognition['data']['selfieCheck']['matching_score']))
        worksheet.write(row_sheet_result, col_sheet_result + 26, str(score))
        worksheet.write(row_sheet_result, col_sheet_result + 27, time_ocr)
        worksheet.write(row_sheet_result, col_sheet_result + 28, row['Ảnh chân dung'])
        worksheet.write(row_sheet_result, col_sheet_result + 29, row['Ảnh CMND mặt trước'])
        worksheet.write(row_sheet_result, col_sheet_result + 30, row['Ảnh CMND mặt sau'])
      print("Time " + str(time_ocr))
      print("Done OCR.")
    except OSError as err:
      worksheet.write(row_sheet_result, col_sheet_result + 1, str(row['Tên']))
      worksheet.write(row_sheet_result, col_sheet_result + 4, str(row['DOB']))
      worksheet.write(row_sheet_result, col_sheet_result + 7, str(row['CCCD']))
      worksheet.write(row_sheet_result, col_sheet_result + 10, str(row['Ngày cấp']))
      worksheet.write(row_sheet_result, col_sheet_result + 13, str(row['Nơi cấp']))
      worksheet.write(row_sheet_result, col_sheet_result + 16, str(row['Quê quán']))
      worksheet.write(row_sheet_result, col_sheet_result + 19, str(row['Thường trú']))
      worksheet.write(row_sheet_result, col_sheet_result + 28, row['Ảnh chân dung'])
      worksheet.write(row_sheet_result, col_sheet_result + 29, row['Ảnh CMND mặt trước'])
      worksheet.write(row_sheet_result, col_sheet_result + 30, row['Ảnh CMND mặt sau'])
      print("Error OCR.")
      print("error: {0}".format(err))
      pass
    
    row_sheet_result += 1
  workbook.close()

# Tính độ chính xác theo phần trăm
def percent_sequence_matcher(str_origin, str_new, is_clean = True):
  if is_clean:
    str_origin = clean_string(str_origin)
    str_new = clean_string(str_new)
  if str_origin == 'nan':
    str_origin = ''
  s = SequenceMatcher(None, str_origin, str_new)
  return round(s.ratio()*100)

def mean(numbers):
  return float(sum(numbers)) / max(len(numbers), 1)

# Xoá khoảng trắng và lower tất cả các text
def clean_string(text):
  text = text.lower().strip()
  return text

def get_as_base64(url):
  return base64.b64encode(requests.get(url).content)

def main():
  write_info_to_excel()

if __name__ == '__main__':
    main()
    