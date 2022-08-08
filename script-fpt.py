import pandas as pd
import xlsxwriter
from api import check_ocr_image_url_fpt, check_ekyc_image_url_fpt
import datetime
import time
from helper import percent_sequence_matcher, mean

def write_info_to_excel():
  df = pd.read_excel("./data-api-ekyc.xlsx", "250 CMND > 10 năm", engine='openpyxl', dtype = str)

  workbook = xlsxwriter.Workbook('./data-result-ekyc-250-fpt-3.xlsx')
  worksheet = workbook.add_worksheet('OCR FPT')
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
  worksheet.write('P1', 'Điểm CMND', bold)
  worksheet.write('Q1', 'Điểm tên', bold)
  worksheet.write('R1', 'Điểm năm sinh', bold)
  worksheet.write('S1', 'Điểm quê quán', bold)
  worksheet.write('T1', 'Điểm nơi trú', bold)
  worksheet.write('U1', 'Điểm ngày cấp', bold)
  worksheet.write('V1', 'Điểm nơi cấp', bold)
  worksheet.write('W1', 'Thời gian OCR', bold)
  worksheet.write('X1', 'eKyc', bold)
  worksheet.write('Y1', 'Thời gian eKyc', bold)

  row_sheet_result = 1
  col_sheet_result = 0
  for index, row in df.iterrows():
    print(row["STT"], row["Tên"])
    worksheet.write(row_sheet_result, col_sheet_result, row['STT'])
    worksheet.write(row_sheet_result, col_sheet_result + 1, str(row['PawnID']))
    worksheet.write(row_sheet_result, col_sheet_result + 2, str(row['CustomerID']))
    worksheet.write(row_sheet_result, col_sheet_result + 11, str(row['Ảnh chân dung']))
    worksheet.write(row_sheet_result, col_sheet_result + 12, str(row['Ảnh CMND mặt trước']))
    worksheet.write(row_sheet_result, col_sheet_result + 13, str(row['Ảnh CMND mặt sau']))
    try:
      recognition, time_ocr = check_ocr_image_url_fpt(row['Ảnh CMND mặt trước'], row['Ảnh CMND mặt sau'], 'cmtnd')
      if recognition['status_code'] != 200:
        worksheet.write(row_sheet_result, col_sheet_result + 3, str(recognition['data']), error)
        worksheet.write(row_sheet_result, col_sheet_result + 22, time_ocr)
        recognition, time_ocr = check_ocr_image_url_fpt(row['Ảnh CMND mặt sau'], row['Ảnh CMND mặt trước'], 'cmtnd')
        if str(recognition['data']) != 'Không đọc được số cmt hoặc cccd':
          worksheet.write(row_sheet_result, col_sheet_result + 3, str(recognition['data']), error)
          worksheet.write(row_sheet_result, col_sheet_result + 22, time_ocr)
      if recognition['status_code'] == 200:
        data = recognition['data']
        list_score = []
        date_time_obj = datetime.datetime.strptime(row['DOB'], '%Y-%m-%d')
        row['DOB'] = date_time_obj.strftime("%d/%m/%Y")
        date_time_obj = datetime.datetime.strptime(row['Ngày cấp'], '%Y-%m-%d')
        row['Ngày cấp'] = date_time_obj.strftime("%d/%m/%Y")

        score_name = percent_sequence_matcher(row['Tên'], data['hoVaTen'])
        score_birthday = percent_sequence_matcher(row['DOB'], data['namSinh'])
        score_id = percent_sequence_matcher(str(row['CMND']), data['soCmt'])
        score_issue_date = percent_sequence_matcher(row['Ngày cấp'], data['ngayCap'])
        score_issue_at = percent_sequence_matcher(row['Nơi cấp'], data['noiCap'])
        # list_score.append(percent_sequence_matcher(row['Quê quán'], data['home_town']))
        score_address = percent_sequence_matcher(row['Thường trú'], data['noiTru'])

        if score_name < 50:
          worksheet.write(row_sheet_result, col_sheet_result + 3, str(data['hoVaTen']), error)
        elif score_name < 100:
          worksheet.write(row_sheet_result, col_sheet_result + 3, str(data['hoVaTen']), warning)
        else:
          worksheet.write(row_sheet_result, col_sheet_result + 3, str(data['hoVaTen']))

        if score_birthday < 50:
          worksheet.write(row_sheet_result, col_sheet_result + 4, str(data['namSinh']), error)
        elif score_birthday < 100:
          worksheet.write(row_sheet_result, col_sheet_result + 4, str(data['namSinh']), warning)
        else:
          worksheet.write(row_sheet_result, col_sheet_result + 4, str(data['namSinh']))

        # if (data['sex'] == 'Nam' and row['Giới tính'] == "1") or (data['sex'] == 'Nữ' and row['Giới tính'] == "0"):
        #   worksheet.write(row_sheet_result, col_sheet_result + 5, str(data['sex']))
        # else:
        #   worksheet.write(row_sheet_result, col_sheet_result + 5, str(data['sex']), error)

        if score_id < 50:
          worksheet.write(row_sheet_result, col_sheet_result + 6, str(data['soCmt']), error)
        elif score_id < 100:
          worksheet.write(row_sheet_result, col_sheet_result + 6, str(data['soCmt']), warning)
        else:
          worksheet.write(row_sheet_result, col_sheet_result + 6, str(data['soCmt']))

        if score_issue_date < 50:
          worksheet.write(row_sheet_result, col_sheet_result + 7, str(data['ngayCap']), error)
        elif score_issue_date < 100:
          worksheet.write(row_sheet_result, col_sheet_result + 7, str(data['ngayCap']), warning)
        else:
          worksheet.write(row_sheet_result, col_sheet_result + 7, str(data['ngayCap']))

        if score_issue_at < 50:
          worksheet.write(row_sheet_result, col_sheet_result + 8, str(data['noiCap']), error)
        elif score_issue_at < 100:
          worksheet.write(row_sheet_result, col_sheet_result + 8, str(data['noiCap']), warning)
        else:
          worksheet.write(row_sheet_result, col_sheet_result + 8, str(data['noiCap']))

        worksheet.write(row_sheet_result, col_sheet_result + 9, str(data['queQuan']))

        if score_address < 50:
          worksheet.write(row_sheet_result, col_sheet_result + 10, str(data['noiTru']), error)
        elif score_address < 100:
          worksheet.write(row_sheet_result, col_sheet_result + 10, str(data['noiTru']), warning)
        else:
          worksheet.write(row_sheet_result, col_sheet_result + 10, str(data['noiTru']))

        list_score.append(score_name)
        list_score.append(score_birthday)
        list_score.append(score_id)
        list_score.append(score_issue_date)
        list_score.append(score_issue_at)
        # list_score.append(percent_sequence_matcher(row['Quê quán'], data['home_town']))
        list_score.append(score_address)

        score = round(mean(list_score))

        worksheet.write(row_sheet_result, col_sheet_result + 14, score)
        worksheet.write(row_sheet_result, col_sheet_result + 15, str(data['soCmtScore']))
        worksheet.write(row_sheet_result, col_sheet_result + 16, str(data['hoVaTenScore']))
        worksheet.write(row_sheet_result, col_sheet_result + 17, str(data['namSinhScore']))
        worksheet.write(row_sheet_result, col_sheet_result + 18, str(data['queQuanScore']))
        worksheet.write(row_sheet_result, col_sheet_result + 19, str(data['noiTruScore']))
        worksheet.write(row_sheet_result, col_sheet_result + 20, str(data['ngayCapScore']))
        worksheet.write(row_sheet_result, col_sheet_result + 21, str(data['noiCapScore']))
        # break
        time.sleep(5)
    except Exception as e:
      print(e)
      pass

    try:
      eKyc, time_ekyc = check_ekyc_image_url_fpt(row['Ảnh chân dung'], row['Ảnh CMND mặt trước'])
      print(row['Ảnh chân dung'])
      if eKyc['status_code'] == 200 and eKyc["data"]["message"] == "Không tìm thấy gương mặt trong ảnh mặt trước":
        eKyc, time_ekyc = check_ekyc_image_url_fpt(row['Ảnh chân dung'], row['Ảnh CMND mặt sau'])
        print("- Error Param")
      if eKyc['status_code'] == 200 and eKyc['data']:
        if eKyc['data']['message'] == 'Thành công':
          worksheet.write(row_sheet_result, col_sheet_result + 23, str(eKyc['data']['message']))
        else:
          worksheet.write(row_sheet_result, col_sheet_result + 23, str(eKyc['data']['message']), warning)
        worksheet.write(row_sheet_result, col_sheet_result + 24, time_ekyc)
      print(eKyc['data'])
      print("- Done eKyc FPT.")
      time.sleep(5)
    except:
      print("- Error eKyc FPT.")
    if (row_sheet_result == 54):
      break
    row_sheet_result += 1
  workbook.close()

def main():
  write_info_to_excel()

if __name__ == '__main__':
    main()
    