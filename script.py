# Reading an excel file using Python
import re
from typing import Counter
from xlrd import open_workbook
import pandas as pd
import xlsxwriter
import datetime

import string
from sklearn.metrics.pairwise import cosine_similarity
from sklearn.feature_extraction.text import CountVectorizer
from nltk.corpus import stopwords
from difflib import SequenceMatcher

from api import check_recognition
# stopwords = stopwords.words('english')


def check_data():
  # Đọc file dữ liệu mẫu
  # xl = pd.ExcelFile("./data-check.xlsx")
  # df = xl.parse("data_sheet")
  df = pd.read_excel("./data-check.xlsx", engine='openpyxl')

  # Truy cập vào file cần tạo dữ liệu
  workbook = xlsxwriter.Workbook('./data-result.xlsx')
  # Tạo bảng rong file excel
  worksheet = workbook.add_worksheet('result')
  bold = workbook.add_format({'bold': True})
  worksheet.write('A1', 'STT', bold)
  worksheet.write('B1', 'id', bold)
  worksheet.write('C1', 'id_new', bold)
  worksheet.write('D1', 'score_id', bold)
  worksheet.write('E1', 'fullname', bold)
  worksheet.write('F1', 'fullname_new', bold)
  worksheet.write('G1', 'score_fullname', bold)
  worksheet.write('H1', 'address', bold)
  worksheet.write('I1', 'address_new', bold)
  worksheet.write('J1', 'score_address', bold)
  worksheet.write('K1', 'birthday', bold)
  worksheet.write('L1', 'birthday_new', bold)
  worksheet.write('M1', 'score_birthday', bold)
  worksheet.write('N1', 'expiry', bold)
  worksheet.write('O1', 'expiry_new', bold)
  worksheet.write('P1', 'score_expiry', bold)
  worksheet.write('Q1', 'issue_by', bold)
  worksheet.write('R1', 'issue_by_new', bold)
  worksheet.write('S1', 'score_issue_by', bold)
  worksheet.write('T1', 'issue_date', bold)
  worksheet.write('U1', 'issue_date_new', bold)
  worksheet.write('V1', 'score_issue_date', bold)
  worksheet.write('W1', 'status_old', bold)
  worksheet.write('X1', 'score', bold)
  worksheet.write('Y1', 'note', bold)
  worksheet.write('Z1', 'front', bold)
  worksheet.write('AA1', 'back', bold)


  row_sheet_result = 1
  col_sheet_result = 0
  limit = 0
  for index, row in df.iterrows():
    # if (pd.isna(row["OK?"])):
    #     print(row)
    # if (row["OK?"] == "OK"):
    # if (row["OK?"] == "NG"):
      path_image1 = './data-image-test/'+str(row.front)+'.jpg' if row.front != 'Không có' else None
      path_image2 = './data-image-test/'+str(row.back)+'.jpg' if row.back != 'Không có' else None
      if (row.back == 'Không có'):
        row.back = None
      recognition = check_recognition(path_image1, path_image2)
      if recognition['status_code'] == 200:
        data = recognition['data']
        if isinstance(row.birthday, datetime.datetime):
          row.birthday = row.birthday.strftime("%m/%d/%Y")
        if isinstance(row.expiry, datetime.datetime):
          row.expiry = row.expiry.strftime("%m/%d/%Y")
        if isinstance(row.issue_by, datetime.datetime):
          row.issue_by = row.issue_by.strftime("%m/%d/%Y")
        if isinstance(row.issue_date, datetime.datetime):
          row.issue_date = row.issue_date.strftime("%m/%d/%Y")


        worksheet.write(row_sheet_result, col_sheet_result, row.STT)

        score_id = percent_sequence_matcher(str(row.id), data['id'], is_clean=False)
        worksheet.write(row_sheet_result, col_sheet_result + 1, str(row.id))
        worksheet.write(row_sheet_result, col_sheet_result + 2, data['id'])
        worksheet.write(row_sheet_result, col_sheet_result + 3, score_id)

        score_fullname = percent_sequence_matcher(str(row.fullname), data['name'])
        worksheet.write(row_sheet_result, col_sheet_result + 4, str(row.fullname))
        worksheet.write(row_sheet_result, col_sheet_result + 5, data['name'])
        worksheet.write(row_sheet_result, col_sheet_result + 6, score_fullname)

        score_address = percent_sequence_matcher(str(row.address), data['address'])
        worksheet.write(row_sheet_result, col_sheet_result + 7, str(row.address))
        worksheet.write(row_sheet_result, col_sheet_result + 8, data['address'])
        worksheet.write(row_sheet_result, col_sheet_result + 9, score_address)

        score_birthday = percent_sequence_matcher(str(row.birthday), data['birthday'])
        worksheet.write(row_sheet_result, col_sheet_result + 10, str(row.birthday))
        worksheet.write(row_sheet_result, col_sheet_result + 11, data['birthday'])
        worksheet.write(row_sheet_result, col_sheet_result + 12, score_birthday)

        score_expiry = percent_sequence_matcher(str(row.expiry), data['expiry'])
        worksheet.write(row_sheet_result, col_sheet_result + 13, str(row.expiry))
        worksheet.write(row_sheet_result, col_sheet_result + 14, data['expiry'])
        worksheet.write(row_sheet_result, col_sheet_result + 15, score_expiry)

        score_issue_at = percent_sequence_matcher(str(row.issue_by), data['issue_at'])
        worksheet.write(row_sheet_result, col_sheet_result + 16, str(row.issue_by))
        worksheet.write(row_sheet_result, col_sheet_result + 17, data['issue_at'])
        worksheet.write(row_sheet_result, col_sheet_result + 18, score_issue_at)

        score_issue_date = percent_sequence_matcher(str(row.issue_date), data['issue_date'])
        worksheet.write(row_sheet_result, col_sheet_result + 19, str(row.issue_date))
        worksheet.write(row_sheet_result, col_sheet_result + 20, data['issue_date'])
        worksheet.write(row_sheet_result, col_sheet_result + 21, score_issue_date)

        worksheet.write(row_sheet_result, col_sheet_result + 22, str(row["OK?"]))

        worksheet.write(
          row_sheet_result,
          col_sheet_result + 23,
          round(mean([score_id, score_fullname, score_address, score_birthday, score_expiry, score_issue_at, score_issue_date]))
        )

        worksheet.write(row_sheet_result, col_sheet_result + 24, str(row["Note"]))
        worksheet.write(row_sheet_result, col_sheet_result + 25, str(row["front"]))
        worksheet.write(row_sheet_result, col_sheet_result + 26, str(row["back"]))



        row_sheet_result += 1
        print("done: " + str(row.fullname))
      else:
        break
      # if (limit == 9):
      #   break
      # limit += 1
  workbook.close()
  

def percent_compare_char_in_string(str_origin, str_new):
  sentences = [
    str_origin,
    str_new
  ]
  cleaned = list(map(clean_string, sentences))
  vectors = get_vectors(cleaned)
  result = cosine_sim_vectors(vectors[0], vectors[1])
  return round(result * 100)

def clean_string(text):
  text = text.lower().strip()
  return text

def get_vectors(cleaned):
  vectorizer = CountVectorizer().fit_transform(cleaned)
  vectors = vectorizer.toarray()
  return vectors

def cosine_sim_vectors(vec1, vec2):
  vec1 = vec1.reshape(1, -1)
  vec2 = vec2.reshape(1, -1)
  return cosine_similarity(vec1, vec2)[0][0]

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

if __name__ == '__main__':
  print(percent_compare_char_in_string(
    'Ấp Phú Lợi, Phú Mỹ Hưng, Củ Chi, Hồ Chí Minh',
    'Ấp Phú Lợi Phú Mỹ Hưng Củ Chi, TP. Hồ Chí Minh'
  ))
  print(percent_sequence_matcher(
    'Ấp Phú Lợi, Phú Mỹ Hưng, Củ Chi, Hồ Chí Minh',
    'Ấp Phú Lợi Phú Mỹ Hưng Củ Chi, TP. Hồ Chí Minh'
  ))
  # check_data()
  # xl = pd.read_excel("./data-check.xlsx", engine='openpyxl')
  # print(xl.iterrows())
  # df = xl.parse("data_sheet")
  # print(df.headline())
  # print(mean([1,2,3]))