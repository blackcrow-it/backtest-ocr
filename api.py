import requests
import base64
import json
from PIL import Image 
import PIL 


url_gmo = "https://f88.smartocr.vn"
api_key = "a08eb42a-4a57-449b-84f4-1f67219f2679"
token_gmo = "eyJhbGciOiJSU0ExXzUiLCJlbmMiOiJBMjU2R0NNIiwia2lkIjoiMSJ9.XUH4hTiekg8168K5Xc-kVmMgQysYvDwUlWfDjH0n0FXX192IPsZdyFycey4yMe3ePwpaBIAerhu91fQFZ5NTiRifZ_SID8I-U7A4vtiBXEQ7llKdzlhUcpYgNHJhNQj7cb_mhvL6T-wh_UGjmk-R5AIS1z6WRCx9pBpgNYvyXE9NblQKfolNHso68liM_P7Bh2ZXAU2gznNzK4bWuiswnNbLj0LMKbo8XpvZr3givm77uUEokEOV2DX3RQbXs6bVt9RZcUo3HyWccHBJkO43Tfrbd2Zkt5SivY4hKXqrSiKQgfDXSAfZo2wDhaloRU43HOiCyvkF7XPZ-UhNfHnUzA.w30TBM506Ssfi8G4zofQJg.jpjJ-96HofQbMuQbbhOBxWByoHbqmMXqM01c9xupq_Eow9HJkXGIcF4xcxQd1ARGLJGXZ6RYLn0ajVoGDHyYKTmKufYwEBJG.cd3WMtg5RRBgRqn_Kp3duw"

url_fpt = "http://api-poc-eid.paas.xplat.fpt.com.vn/api"
url_ocr_fpt = "http://ocrhub.vn/api"
token = "2875b647-0a7c-4954-b1ec-a6e32f5138fb"
code = "F88TEST"

url_savis = "https://uat-gateway.digital-id.vn"
apikey = "eyJ4NXQiOiJOVGRtWmpNNFpEazNOalkwWXpjNU1tWm1PRGd3TVRFM01XWXdOREU1TVdSbFpEZzROemM0WkE9PSIsImtpZCI6ImdhdGV3YXlfY2VydGlmaWNhdGVfYWxpYXMiLCJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJzdWIiOiJhZG1pbkBjYXJib24uc3VwZXIiLCJhcHBsaWNhdGlvbiI6eyJvd25lciI6ImFkbWluIiwidGllclF1b3RhVHlwZSI6bnVsbCwidGllciI6IjUwUGVyTWluIiwibmFtZSI6IkY4OC1FS1lDIiwiaWQiOjEwLCJ1dWlkIjoiYmQwMjdlM2MtNzllYy00YzY4LThkYmEtZDEwODcwMWI3YjRkIn0sImlzcyI6Imh0dHBzOlwvXC91YXQtYXBpcG9ydGFsLmRpZ2l0YWwtaWQudm46NDQzXC9vYXV0aDJcL3Rva2VuIiwidGllckluZm8iOnsiQnJvbnplIjp7InRpZXJRdW90YVR5cGUiOiJyZXF1ZXN0Q291bnQiLCJncmFwaFFMTWF4Q29tcGxleGl0eSI6MCwiZ3JhcGhRTE1heERlcHRoIjowLCJzdG9wT25RdW90YVJlYWNoIjp0cnVlLCJzcGlrZUFycmVzdExpbWl0IjowLCJzcGlrZUFycmVzdFVuaXQiOm51bGx9fSwia2V5dHlwZSI6IlBST0RVQ1RJT04iLCJwZXJtaXR0ZWRSZWZlcmVyIjoiIiwic3Vic2NyaWJlZEFQSXMiOlt7InN1YnNjcmliZXJUZW5hbnREb21haW4iOiJjYXJib24uc3VwZXIiLCJuYW1lIjoiRjg4LUVLWUMiLCJjb250ZXh0IjoiXC9lLWt5Y1wvMS4wIiwicHVibGlzaGVyIjoiYWRtaW4iLCJ2ZXJzaW9uIjoiMS4wIiwic3Vic2NyaXB0aW9uVGllciI6IkJyb256ZSJ9LHsic3Vic2NyaWJlclRlbmFudERvbWFpbiI6ImNhcmJvbi5zdXBlciIsIm5hbWUiOiJGODgtT0NSLUNBUkQiLCJjb250ZXh0IjoiXC9vY3JcLzEuMCIsInB1Ymxpc2hlciI6ImFkbWluIiwidmVyc2lvbiI6IjEuMCIsInN1YnNjcmlwdGlvblRpZXIiOiJCcm9uemUifV0sInBlcm1pdHRlZElQIjoiIiwiaWF0IjoxNjIzNDEzMTY1LCJqdGkiOiI4MGZmOWIwMy1hNTRjLTRjYzUtYWE2My00MmZjYjAyZWRlYjAifQ==.HC3Pi-6Jkwc6nZvcUbzkS4E3BxAPr_IUxreP7cA_lJul8l35A9sLpp_A8WpogOtczpk3kWAz-7JHt8HEx1KJyQbN2403Bv0ByDyI_n0ntD1PCI9ca7SbZknOAwEUnVEHNnDL9_R4Car9hFvpLgA2dYX67Y2DV_g5bmxBn1ztQGODSUS3b6mcopGw3Dq2_UhPppyVQMvi0StgmGEo87F4RLboNvqd27DeMnYbG3RRoU7Z3fk82TW48rm2WdJg6A0HF6hjWOsWGhT9HORn_418ZDSyQ-jYSOjB5zf7m37G6fIt4D4dJ7oFbccxU3nLUW519NEc7a4RAT36fDbW-LA4DA=="

url_kalapa = "https://api.kalapa.vn/user-profile"
token_kalapa = "5bb42ea331ee010001a0b7d7eb309480283540b0b087d26c748c8e4e"

url_kalapa_epay = "https://ekyc-api.kalapa.vn/api"
token_kalapa_epay = "5bb42ea331ee010001a0b7d76cfc1e45a65141038389808b9ff3562a"

def check_recognition(path_image_front, path_image_back):
  url = url_gmo + '/idfull/v1/recognition'
  files = {}
  if path_image_front:
    files['image1'] = open(path_image_front, 'rb')
  if path_image_back:
    files['image2'] = open(path_image_back, 'rb')
  r = requests.post(
    url,
    files=files,
    data = {'encode': 1},
    headers= {'api-key': api_key}
  )
  if r.status_code == 200:
    return {'status_code': 200, 'data': r.json()}
  else:
    return {'status_code': 500, 'data': None}

def check_recognition_image_url(url_image_front, url_image_back):
  url = url_gmo + '/idfull/v1/recognition'
  files = {}
  if url_image_front:
    files['image1'] = requests.get(url_image_front).content
  if url_image_back:
    files['image2'] = requests.get(url_image_back).content
  r = requests.post(
    url,
    files=files,
    data = {'encode': 1},
    headers= {'api-key': api_key}
  )
  if r.status_code == 200:
    return {'status_code': 200, 'data': r.json()}, r.elapsed.total_seconds()
  else:
    return {'status_code': 500, 'data': None}, r.elapsed.total_seconds()

def check_ekyc_image_url_gmo(url_image_face, url_image_front):
  url = url_gmo + '/face/v1/recognition'
  files = {}
  if url_image_face:
    files['image1'] = requests.get(url_image_face).content
  if url_image_front:
    files['image2'] = requests.get(url_image_front).content
  r = requests.post(
    url,
    files=files,
    headers= {'api-key': api_key}
  )
  if r.status_code == 200:
    return {'status_code': 200, 'data': r.json()}, r.elapsed.total_seconds()
  else:
    return {'status_code': 500, 'data': None}, r.elapsed.total_seconds()

def check_ekyc_image_url_fpt(url_image_face, url_image_front):
  url = url_fpt + '/public/all/so-sanh-anh'
  data = {
    'anhKhachHang': base64.b64encode(requests.get(url_image_face).content).decode("utf-8"),
    'anhMatTruoc': base64.b64encode(requests.get(url_image_front).content).decode("utf-8")
  }
  r = requests.post(
    url,
    json=data,
    headers= {'token': token, 'code': code, 'Accept-Language': 'vi'}
  )
  print(r.content)
  print(r.elapsed.total_seconds())
  if r.status_code == 200:
    return {'status_code': 200, 'data': r.json()}, r.elapsed.total_seconds()
  else:
    return {'status_code': 500, 'data': None}, r.elapsed.total_seconds()

def check_ocr_image_url_fpt(url_image_front, url_image_back, type_image):
  url = url_fpt + '/public/all/doc-noi-dung-ocr'
  data = {
    'anhMatTruoc': base64.b64encode(requests.get(url_image_front).content).decode("utf-8"),
    'anhMatSau': base64.b64encode(requests.get(url_image_back).content).decode("utf-8"),
    'maGiayTo': type_image
  }
  r = requests.post(
    url,
    json=data,
    headers= {'token': token, 'code': code, 'Accept-Language': 'vi'}
  )
  print(r.json())
  if r.status_code == 200 and r.json()['status'] == 200:
    return {'status_code': 200, 'data': r.json()['data']}, r.elapsed.total_seconds()
  else:
    try:
      return {'status_code': 500, 'data': r.json()['message']}, r.elapsed.total_seconds()
    except:
      return {'status_code': 500, 'data': None}, r.elapsed.total_seconds()

def check_ocr_image_url_savis(url_image_1, url_image_2):
  url = url_savis + '/ocr/1.0/predict'
  data_result = {}
  status_code = 0

  r = requests.post(
    url,
    data = {
      'input': [url_image_1, url_image_2],
      'check_liveness': True
    },
    headers = {'apikey': apikey}
  )
  if r.status_code == 200:
    data_result['class_name'] = []
    data_result['liveness'] = []
    for item in r.json()['output']:
      for key, value in item.items():
        if key == 'class_name' or key == 'liveness':
          data_result[key].append(value)
          if key == 'class_name' and value['normalized']['value'] != (-1):
            status_code = 200
        else:
          data_result[key] = value
  else:
    status_code = r.status_code
    data_result = r.json()
  return {'status_code': int(status_code), 'data': data_result}, r.elapsed.total_seconds()

def check_ekyc_image_url_savis(url_image_face, url_image_front):
  url = url_savis + '/e-kyc/1.0/check_match_image_card_image_general'
  data = {
    'image_general': url_image_face,
    'image_card': url_image_front,
    'threshold': 0.8
  }
  r = requests.post(
    url,
    data=data,
    headers= {'apikey': apikey}
  )
  if r.status_code == 200:
    return {'status_code': 200, 'data': r.json()['output']}, r.elapsed.total_seconds()
  else:
    return {'status_code': 500, 'data': None}, r.elapsed.total_seconds()

def check_ocr_dkx_image_url_gmo(url_image_front, url_image_back):
  url = url_gmo + '/dkx/v1/recognition'
  files = {}
  if url_image_front:
    files['image1'] = requests.get(url_image_front).content
  if url_image_back:
    files['image2'] = requests.get(url_image_back).content
  r = requests.post(
    url,
    files = files,
    data = {'encode': 1},
    headers= {'Authorization': 'Bearer ' + token_gmo}
  )
  print(r.json())
  if r.status_code == 200:
    return {'status_code': 200, 'data': r.json()}, r.elapsed.total_seconds()
  else:
    return {'status_code': 500, 'data': None}, r.elapsed.total_seconds()

def check_ocr_image_url_kalapa(url_image_back):
  url = url_kalapa + '/registration/ocr'
  files = {}
  files['image'] = requests.get(url_image_back).content
  r = requests.post(
    url,
    files=files,
    headers= { 'Authorization': token_kalapa }
  )
  print(r.status_code)
  print(r.json())
  if r.status_code == 200:
    return {'status_code': 200, 'data': r.json()}, r.elapsed.total_seconds()
  else:
    try:
      return {'status_code': 500, 'data': r.json()['message']}, r.elapsed.total_seconds()
    except:
      return {'status_code': 500, 'data': None}, r.elapsed.total_seconds()

def check_ekyc_image_url_kalapa(url_image_front, url_image_back, url_image_self):
  url = url_kalapa_epay + '/kyc/get-all'
  files = {}
  if url_image_front:
    files['front_image'] = open('./image/' + url_image_front.split('/')[-1], 'rb')
  if url_image_back:
    files['back_image'] = open('./image/' + url_image_back.split('/')[-1], 'rb')
  if url_image_self:
    try:
      files['selfie_image'] = open('./image/' + url_image_self.split('/')[-1], 'rb')
    except:
      files['selfie_image'] = open('./image/' + url_image_front.split('/')[-1], 'rb')
      pass
  r = requests.post(
    url,
    files = files,
    data = {'app_token': token_kalapa_epay}
  )
  if r.status_code == 200:
    return {'status_code': 200, 'data': r.json()}, r.elapsed.total_seconds()
  else:
    return {'status_code': 500, 'data': None}, r.elapsed.total_seconds()

print(check_ekyc_image_url_kalapa(
  'https://f88.s3-ap-southeast-1.amazonaws.com/POS/Uploads/Customer/CCCD/z2536306464562_245ed48e740e4197e0a74b54a7bda86d1202106061732135134438.jpg',
  'https://f88.s3-ap-southeast-1.amazonaws.com/POS/Uploads/Customer/CCCD/z2536306463466_5ace44a3025c485af39b295ae093bece1202106061733134837525.jpg',
  'https://f88.s3-ap-southeast-1.amazonaws.com/POS/Uploads/HDCC/HPG0233/2106/58/Avatar/z2536311968290_2fbc71ed5bf7511bb690866d4b7b80a41202106061807376357636.jpg'
))