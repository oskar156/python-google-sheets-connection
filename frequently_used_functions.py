import os.path
import time 
from datetime import datetime
import re

from selenium import webdriver
from selenium.webdriver import Firefox
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select

import google.auth
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

def column_to_number(column_letter):
  column_number = 0
  for i, char in enumerate(column_letter[::-1]):  # Iterate through the letters in reverse
    column_number += (ord(char.upper()) - ord('A') + 1) * (26 ** i)
  return column_number

def google_sheet_read(service, google_sheet_id, sheet_name, range):
  sheet = service.spreadsheets()
  combined_range = "'" + str(sheet_name) + "'" + "!" + str(range)
  result = sheet.values().get(spreadsheetId = str(google_sheet_id), range=str(combined_range)).execute()
  return result.get("values", [])

def google_sheet_write(service, google_sheet_id, sheet_name, range, values_to_write):
  sheet = service.spreadsheets()
  body = {"values": values_to_write}
  combined_range = "'" + str(sheet_name) + "'" + "!" + str(range)
  sheet.values().update(spreadsheetId = str(google_sheet_id), range=str(combined_range), valueInputOption="USER_ENTERED", body=body,).execute()

def google_sheet_append_rows(service, google_sheet_id, sheet_name, values_to_append):
  sheet = service.spreadsheets()
  body = {"values": values_to_append}
  combined_range = "'" + str(sheet_name) + "'" + "!A1" #appends to the bottom of the sheet no matter what is here
  sheet.values().get(spreadsheetId = str(google_sheet_id), range=str(combined_range), valueInputOption="USER_ENTERED", insertDataOption='INSERT_ROWS', body=body,).execute()

def google_sheet_insert_blank_rows(service, google_sheet_id, sheet_name, insert_index, rows_to_insert):
  spreadsheet = service.spreadsheets().get(spreadsheetId=google_sheet_id).execute()
  sheet = service.spreadsheets()
  sheet_id = -1
  for sheet in spreadsheet['sheets']:
    if sheet['properties']['title'] == sheet_name:
      sheet_id = sheet['properties']['sheetId']
      break
  request_body = {"requests": [{ "insertDimension": {
    "range": {
      "sheetId": sheet_id,
      "dimension": "ROWS",
      "startIndex": insert_index,
      "endIndex": insert_index + rows_to_insert
    },
    "inheritFromBefore": True
  }}]}
  service.spreadsheets().batchUpdate(spreadsheetId=google_sheet_id, body=request_body).execute()

#WIP BELOW
def google_sheet_copy_paste(service, google_sheet_id, copy_sheet_name, copy_range, paste_sheet_name, paste_range):
  spreadsheet = service.spreadsheets().get(spreadsheetId=google_sheet_id).execute()
  sheet = service.spreadsheets()

  combined_copy_range = "'" + str(copy_sheet_name) + "'" + "!" + str(copy_range)
  result = service.spreadsheets().values().get(spreadsheetId=google_sheet_id, range=combined_copy_range).execute()
  values = result.get('values', [])

  paste_sheet_id = -1
  for sheet in spreadsheet['sheets']:
    if sheet['properties']['title'] == paste_sheet_name:
        paste_sheet_id = sheet['properties']['sheetId']
        break

  #combined_paste_range = "'" + str(paste_sheet_name) + "'" + "!" + str(paste_range)   
  start_row_index = int(re.search("[0-9]+", paste_range).group())
  end_row_index = re.search("[0-9]+", paste_range[::-1]).group()[::-1]
  start_col_index = int(column_to_number(re.search(r"[A-z]+", paste_range).group()))
  end_col_index = column_to_number(re.search(r"\:[A-z]+", paste_range).group()[1:])
  
  requests = []

  for i, row in enumerate(values):
    for j, cell in enumerate(row):
      # For each formula cell, create a request to paste it
      requests.append({"updateCells": {
        "range": {
            "sheetId": 0,  # Use the sheet ID or get it from a separate API call
            "startRowIndex": start_row_index + i,
            "endRowIndex": start_row_index + i,
            "startColumnIndex": start_col_index + i,
            "endColumnIndex": start_col_index + i
        },
        "rows": [{"values": [{"userEnteredValue": {"formulaValue": cell}}]}],
        "fields": "userEnteredValue"
      }})
  body = {'requests': requests}
  service.spreadsheets().batchUpdate(spreadsheetId=google_sheet_id, body=body).execute()
