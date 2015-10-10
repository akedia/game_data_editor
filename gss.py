__author__ = 'yaoyuanchao'
import json
import gspread
import os.path
from oauth2client.client import OAuth2WebServerFlow
from oauth2client.tools import run_flow,argparser
from oauth2client.file import Storage
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive

credential_data_path="creds.data"
def authorize_from_web():

    CLIENT_ID = '879811213950-qtlqq74tfclt8co4799f3kbq7hf83qua.apps.googleusercontent.com'
    CLIENT_SECRET = 'enYf8pgBWxNGZyoxsXwxDFPA'

    flow = OAuth2WebServerFlow(
              client_id = CLIENT_ID,
              client_secret = CLIENT_SECRET,
              scope = 'https://spreadsheets.google.com/feeds https://docs.google.com/feeds',
              redirect_uri = 'http://example.com/auth_return'
           )
    flags = argparser.parse_args(args=[])
    storage = Storage(credential_data_path)
    credentials = run_flow(flow, storage,flags)
    print "access_token: %s" % credentials.access_token
    return credentials

def authorize_from_local():
    if os.path.isfile(credential_data_path):
        return Storage(credential_data_path).get()
    else:
        return None

def credentials_from_auth():
    credentials=authorize_from_local()
    if credentials is None:
        return authorize_from_web()
    else:
        return credentials


credentials=credentials_from_auth()
gc= gspread.authorize(credentials)
wks=gc.open("test for python").sheet1
wks.update_acell('B2','heasallo')

