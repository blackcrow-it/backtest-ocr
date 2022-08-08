import pandas as pd
import xlsxwriter
from api import check_ocr_image_url_savis, check_ekyc_image_url_savis
import datetime
import time
from helper import percent_sequence_matcher, mean

def write_info_to_excel():
  df = pd.read_excel("./data-api-ekyc.xlsx", "250 CMND > 10 năm", engine='openpyxl', dtype = str)

  workbook = xlsxwriter.Workbook('./data-result-ekyc-250-savis.xlsx')
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
  worksheet.write('V1', 'eKyc', bold)
  worksheet.write('W1', 'Thời gian OCR', bold)
  worksheet.write('X1', 'Thời gian eKyc', bold)

  row_sheet_result = 1
  col_sheet_result = 0
  for index, row in df.iterrows():
    print(row["STT"], row["Tên"])
    worksheet.write(row_sheet_result, col_sheet_result, row['STT'])
    worksheet.write(row_sheet_result, col_sheet_result + 1, str(row['PawnID']))
    worksheet.write(row_sheet_result, col_sheet_result + 2, str(row['CustomerID']))
    try:
      recognition, time_ocr = check_ocr_image_url_savis(row['Ảnh CMND mặt trước'], row['Ảnh CMND mặt sau'])
      # print(recognition['status_code'])
      if recognition['status_code'] == 200:
        data = recognition['data']
        list_score = []

        date_time_obj = datetime.datetime.strptime(row['DOB'], '%Y-%m-%d')
        row['DOB'] = date_time_obj.strftime("%d/%m/%Y")
        date_time_obj = datetime.datetime.strptime(row['Ngày cấp'], '%Y-%m-%d')
        row['Ngày cấp'] = date_time_obj.strftime("%d/%m/%Y")

        score_name = percent_sequence_matcher(row['Tên'], data['ho_ten']['value'])
        score_birthday = percent_sequence_matcher(row['DOB'], data['ngay_sinh']['normalized']['value'])
        score_id = percent_sequence_matcher(str(row['CMND']), data['id']['value'])
        score_issue_date = percent_sequence_matcher(row['Ngày cấp'], data['ngay_cap']['normalized']['value'])
        score_issue_at = percent_sequence_matcher(row['Nơi cấp'], data['noi_cap']['value'])
        # list_score.append(percent_sequence_matcher(row['Quê quán'], data['home_town']))
        score_address = percent_sequence_matcher(row['Thường trú'], data['ho_khau_thuong_tru']['value'])

        if score_name < 50:
          worksheet.write(row_sheet_result, col_sheet_result + 3, str(data['ho_ten']['value']), error)
        elif score_name < 100:
          worksheet.write(row_sheet_result, col_sheet_result + 3, str(data['ho_ten']['value']), warning)
        else:
          worksheet.write(row_sheet_result, col_sheet_result + 3, str(data['ho_ten']['value']))

        if score_birthday < 50:
          worksheet.write(row_sheet_result, col_sheet_result + 4, str(data['ngay_sinh']['normalized']['value']), error)
        elif score_birthday < 100:
          worksheet.write(row_sheet_result, col_sheet_result + 4, str(data['ngay_sinh']['normalized']['value']), warning)
        else:
          worksheet.write(row_sheet_result, col_sheet_result + 4, str(data['ngay_sinh']['normalized']['value']))

        # if (data['sex'] == 'Nam' and row['Giới tính'] == "1") or (data['sex'] == 'Nữ' and row['Giới tính'] == "0"):
        #   worksheet.write(row_sheet_result, col_sheet_result + 5, str(data['sex']))
        # else:
        #   worksheet.write(row_sheet_result, col_sheet_result + 5, str(data['sex']), error)

        if score_id < 50:
          worksheet.write(row_sheet_result, col_sheet_result + 6, str(data['id']['value']), error)
        elif score_id < 100:
          worksheet.write(row_sheet_result, col_sheet_result + 6, str(data['id']['value']), warning)
        else:
          worksheet.write(row_sheet_result, col_sheet_result + 6, str(data['id']['value']))

        if score_issue_date < 50:
          worksheet.write(row_sheet_result, col_sheet_result + 7, str(data['ngay_cap']['normalized']['value']), error)
        elif score_issue_date < 100:
          worksheet.write(row_sheet_result, col_sheet_result + 7, str(data['ngay_cap']['normalized']['value']), warning)
        else:
          worksheet.write(row_sheet_result, col_sheet_result + 7, str(data['ngay_cap']['normalized']['value']))

        if score_issue_at < 50:
          worksheet.write(row_sheet_result, col_sheet_result + 8, str(data['noi_cap']['value']), error)
        elif score_issue_at < 100:
          worksheet.write(row_sheet_result, col_sheet_result + 8, str(data['noi_cap']['value']), warning)
        else:
          worksheet.write(row_sheet_result, col_sheet_result + 8, str(data['noi_cap']['value']))

        worksheet.write(row_sheet_result, col_sheet_result + 9, str(data['nguyen_quan']))

        if score_address < 50:
          worksheet.write(row_sheet_result, col_sheet_result + 10, str(data['ho_khau_thuong_tru']['value']), error)
        elif score_address < 100:
          worksheet.write(row_sheet_result, col_sheet_result + 10, str(data['ho_khau_thuong_tru']['value']), warning)
        else:
          worksheet.write(row_sheet_result, col_sheet_result + 10, str(data['ho_khau_thuong_tru']['value']))
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

        worksheet.write(row_sheet_result, col_sheet_result + 14, score)
        worksheet.write(row_sheet_result, col_sheet_result + 15, str(data['id']['confidence']))
        worksheet.write(row_sheet_result, col_sheet_result + 16, str(data['ho_ten']['confidence']))
        worksheet.write(row_sheet_result, col_sheet_result + 17, str(data['ngay_sinh']['confidence']))
        worksheet.write(row_sheet_result, col_sheet_result + 18, str(data['nguyen_quan']['confidence']))
        worksheet.write(row_sheet_result, col_sheet_result + 19, str(data['ho_khau_thuong_tru']['confidence']))
        worksheet.write(row_sheet_result, col_sheet_result + 20, str(data['ngay_cap']['confidence']))
        worksheet.write(row_sheet_result, col_sheet_result + 21, str(data['noi_cap']['confidence']))
        worksheet.write(row_sheet_result, col_sheet_result + 23, time_ocr)
        # break
        # time.sleep(5)
    except:
      pass

    print("- Checking eKyc Savis ...")
    eKyc, time_ekyc = check_ekyc_image_url_savis(row['Ảnh chân dung'], row['Ảnh CMND mặt trước'])
    print(row['Ảnh chân dung'])
    if eKyc['status_code'] == 200 and eKyc["data"]["is_matched"]["value"] == "False" and "error" in eKyc["data"]:
      eKyc, time_ekyc = check_ekyc_image_url_savis(row['Ảnh chân dung'], row['Ảnh CMND mặt sau'])
      print("- Error Param")
    if eKyc['status_code'] == 200 and eKyc['data']:
      print(eKyc)
      worksheet.write(row_sheet_result, col_sheet_result + 22, str(eKyc["data"]["is_matched"]["value"]))
      worksheet.write(row_sheet_result, col_sheet_result + 24, time_ekyc)
    print(eKyc['data'])
    print("- Done eKyc Savis.")
    
    row_sheet_result += 1
  workbook.close()

def main():
  write_info_to_excel()

if __name__ == '__main__':
    main()
    