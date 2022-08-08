import pandas as pd
import xlsxwriter
# https://www.educative.io/answers/what-is-sequencematcher-in-python
from difflib import SequenceMatcher
from api import check_recognition_image_url, check_ekyc_image_url_gmo, check_ekyc_image_url_fpt
import base64
import requests
import datetime

"""
Hàm viết dữ liệu vào file excel
"""
def write_info_to_excel():
  df = pd.read_excel("./data-api-ekyc.xlsx", "250 CMND > 10 năm", engine='openpyxl', dtype = str)

  workbook = xlsxwriter.Workbook('./data-result-ekyc-250-gmo.xlsx')
  worksheet = workbook.add_worksheet('250 CMND > 10 năm GMO')
  bold = workbook.add_format({'bold': True})
  error = workbook.add_format({'bg_color': 'red', 'font_color': 'white'})
  warning = workbook.add_format({'bg_color': 'yellow'})
  worksheet.write('A1', 'STT', bold)
  worksheet.write('B1', 'PawnID', bold)
  worksheet.write('C1', 'CustomerID', bold)
  worksheet.write('D1', 'Tên', bold)
  worksheet.write('E1', 'DOB', bold)
  worksheet.write('F1', 'Giới tính', bold)
  worksheet.write('G1', 'CMND', bold)
  worksheet.write('H1', 'Ngày cấp', bold)
  worksheet.write('I1', 'Nơi cấp', bold)
  worksheet.write('J1', 'Quê quán', bold)
  worksheet.write('K1', 'Thường trú', bold)
  worksheet.write('L1', 'Ảnh chân dung', bold)
  worksheet.write('M1', 'Ảnh CMND mặt trước', bold)
  worksheet.write('N1', 'Ảnh CMND mặt sau', bold)
  worksheet.write('O1', 'Điểm so sánh', bold)
  worksheet.write('P1', 'Kiểm tra mờ nhoè', bold)
  worksheet.write('Q1', 'Kiểm tra photocopy', bold)
  worksheet.write('R1', 'Kiểm tra cắt góc', bold)
  worksheet.write('S1', 'Kiểm tra cùng loại', bold)
  worksheet.write('T1', 'Kiểm tra chụp qua màn hình', bold)
  worksheet.write('U1', 'Kiểm tra hết hạn', bold)
  worksheet.write('V1', 'eKyc', bold)
  worksheet.write('W1', 'Độ chính xác', bold)
  worksheet.write('X1', 'Thời gian OCR', bold)
  worksheet.write('Y1', 'Thời gian eKyc', bold)

  row_sheet_result = 1
  col_sheet_result = 0
  for index, row in df.iterrows():
    print(row["STT"], row["Tên"])
    worksheet.write(row_sheet_result, col_sheet_result, row['STT'])
    worksheet.write(row_sheet_result, col_sheet_result + 1, str(row['PawnID']))
    worksheet.write(row_sheet_result, col_sheet_result + 2, str(row['CustomerID']))

    print("Checking OCR ...")
    # Đọc dữ liệu OCR trên CMND
    recognition, time_ocr = check_recognition_image_url(row['Ảnh CMND mặt trước'], row['Ảnh CMND mặt sau'])
    if recognition['status_code'] == 200:
      data = recognition['data']
      list_score = []

      date_time_obj = datetime.datetime.strptime(row['DOB'], '%Y-%m-%d')
      row['DOB'] = date_time_obj.strftime("%d/%m/%Y")
      date_time_obj = datetime.datetime.strptime(row['Ngày cấp'], '%Y-%m-%d')
      row['Ngày cấp'] = date_time_obj.strftime("%d/%m/%Y")

      score_name = percent_sequence_matcher(row['Tên'], data['name'])
      score_birthday = percent_sequence_matcher(row['DOB'], data['birthday'])
      score_id = percent_sequence_matcher(str(row['CMND']), data['id'])
      score_issue_date = percent_sequence_matcher(row['Ngày cấp'], data['issue_date'])
      score_issue_at = percent_sequence_matcher(row['Nơi cấp'], data['issue_at'])
      # list_score.append(percent_sequence_matcher(row['Quê quán'], data['home_town']))
      score_address = percent_sequence_matcher(row['Thường trú'], data['address'])

      if score_name < 50:
        worksheet.write(row_sheet_result, col_sheet_result + 3, str(data['name']), error)
      elif score_name < 100:
        worksheet.write(row_sheet_result, col_sheet_result + 3, str(data['name']), warning)
      else:
        worksheet.write(row_sheet_result, col_sheet_result + 3, str(data['name']))

      if score_birthday < 50:
        worksheet.write(row_sheet_result, col_sheet_result + 4, str(data['birthday']), error)
      elif score_birthday < 100:
        worksheet.write(row_sheet_result, col_sheet_result + 4, str(data['birthday']), warning)
      else:
        worksheet.write(row_sheet_result, col_sheet_result + 4, str(data['birthday']))

      if (data['sex'] == 'Nam' and row['Giới tính'] == "1") or (data['sex'] == 'Nữ' and row['Giới tính'] == "0"):
        worksheet.write(row_sheet_result, col_sheet_result + 5, str(data['sex']))
      else:
        if data['sex'] == "":
          data['sex'] = "Không xác định"
        worksheet.write(row_sheet_result, col_sheet_result + 5, str(data['sex']), error)

      if score_id < 50:
        worksheet.write(row_sheet_result, col_sheet_result + 6, str(data['id']), error)
      elif score_id < 100:
        worksheet.write(row_sheet_result, col_sheet_result + 6, str(data['id']), warning)
      else:
        worksheet.write(row_sheet_result, col_sheet_result + 6, str(data['id']))

      if score_issue_date < 50:
        worksheet.write(row_sheet_result, col_sheet_result + 7, str(data['issue_date']), error)
      elif score_issue_date < 100:
        worksheet.write(row_sheet_result, col_sheet_result + 7, str(data['issue_date']), warning)
      else:
        worksheet.write(row_sheet_result, col_sheet_result + 7, str(data['issue_date']))

      if score_issue_at < 50:
        worksheet.write(row_sheet_result, col_sheet_result + 8, str(data['issue_at']), error)
      elif score_issue_at < 100:
        worksheet.write(row_sheet_result, col_sheet_result + 8, str(data['issue_at']), warning)
      else:
        worksheet.write(row_sheet_result, col_sheet_result + 8, str(data['issue_at']))

      worksheet.write(row_sheet_result, col_sheet_result + 9, str(data['home_town']))

      if score_address < 50:
        worksheet.write(row_sheet_result, col_sheet_result + 10, str(data['address']), error)
      elif score_address < 100:
        worksheet.write(row_sheet_result, col_sheet_result + 10, str(data['address']), warning)
      else:
        worksheet.write(row_sheet_result, col_sheet_result + 10, str(data['address']))
      worksheet.write(row_sheet_result, col_sheet_result + 11, str(row['Ảnh chân dung']))
      worksheet.write(row_sheet_result, col_sheet_result + 12, str(row['Ảnh CMND mặt trước']))
      worksheet.write(row_sheet_result, col_sheet_result + 13, str(row['Ảnh CMND mặt sau']))

      list_score.append(score_name)
      list_score.append(score_birthday)
      list_score.append(score_id)
      list_score.append(score_issue_date)
      list_score.append(score_issue_at)
      # list_score.append(percent_sequence_matcher(row['Quê quán'], data['home_town']))
      list_score.append(score_address)

      score = round(mean(list_score))

      worksheet.write(row_sheet_result, col_sheet_result + 14, str(score) + "%")
      worksheet.write(row_sheet_result, col_sheet_result + 15, str(data['blur_check']))
      worksheet.write(row_sheet_result, col_sheet_result + 16, str(data['color_check']))
      worksheet.write(row_sheet_result, col_sheet_result + 17, str(data['corner_check']))
      worksheet.write(row_sheet_result, col_sheet_result + 18, str(data['front_back_type']))
      worksheet.write(row_sheet_result, col_sheet_result + 19, str(data['throw_screen_check']))
      worksheet.write(row_sheet_result, col_sheet_result + 20, str(data['expire_check']))
      worksheet.write(row_sheet_result, col_sheet_result + 23, time_ocr)
    print("Time " + str(time_ocr))
    print("Done OCR.")

    print("Checking eKyc GMO ...")
    # Kiểm tra CMND với ảnh chân dung
    eKyc, time_ekyc = check_ekyc_image_url_gmo(row['Ảnh chân dung'], row['Ảnh CMND mặt trước'])
    if not (eKyc['status_code'] == 200 and eKyc['data'] and eKyc['data']['result_code'] == 200):
      eKyc, time_ekyc = check_ekyc_image_url_gmo(row['Ảnh chân dung'], row['Ảnh CMND mặt sau'])
      print("Error Param")
    if eKyc['status_code'] == 200 and eKyc['data'] and eKyc['data']['result_code'] == 200:
      worksheet.write(row_sheet_result, col_sheet_result + 21, str(eKyc['data']['face_compare']))
      worksheet.write(row_sheet_result, col_sheet_result + 22, str(eKyc['data']['score']))
      worksheet.write(row_sheet_result, col_sheet_result + 24, time_ekyc)
    print("Time " + str(time_ekyc))
    print("Done eKyc.")
    
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
    