
from difflib import SequenceMatcher
import base64
import requests

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

def clean_string(text):
  text = text.lower().strip()
  return text

def get_as_base64(url):
  return base64.b64encode(requests.get(url).content)