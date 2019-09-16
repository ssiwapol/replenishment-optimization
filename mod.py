import shutil
import os
import io

import yaml
import numpy as np
import xlrd
from google.cloud import storage


class PathFile:
    def __init__(self):
        with open("config.yaml") as f:
            self.config = yaml.load(f, Loader=yaml.Loader)
            # gcp client
            if self.config['app']['run'] == "gcp":
                if self.config['path']['gcp']['authfile'] == 'None':
                    client = storage.Client()
                else:
                    client = storage.Client.from_service_account_json(
                        self.config['path']['gcp']['authfile'])
                self.gcpbucket = client.get_bucket(self.config['path']['gcp']['bucket'])
            else:
                pass

    def getuser(self):
        filename = self.config['file']['user']
        # load file from GCP
        if self.config['app']['run'] == "gcp":
            fullpath = '/'.join((self.config['path']['gcp']['path'], filename))
            blob = storage.Blob(fullpath, self.gcpbucket)
            byte_stream = io.BytesIO()
            blob.download_to_file(byte_stream)
            byte_stream.seek(0)
            with byte_stream as f:
                user = yaml.load(f, Loader=yaml.Loader)
        # load file from local
        else:
            path = os.path.join(self.config['path']['local']['path'], filename)
            with open(path) as f:
                user = yaml.load(f, Loader=yaml.Loader)
        return user

    def setuser(self, user):
        self.user = user

    def loadfile(self, filename):
        # load file from GCP
        if self.config['app']['run'] == "gcp":
            fullpath = '/'.join((self.config['path']['gcp']['path'], self.user, filename))
            blob = storage.Blob(fullpath, self.gcpbucket)
            byte_stream = io.BytesIO()
            blob.download_to_file(byte_stream)
            byte_stream.seek(0)
            return byte_stream
        # load file from local
        else:
            path = os.path.join(self.config['path']['local']['path'], self.user, filename)
            return path

    def savefile(self, file, filename):
        # save file to GCP
        if self.config['app']['run'] == "gcp":
            fullpath = '/'.join((self.config['path']['gcp']['path'], self.user, filename))
            blob = storage.Blob(fullpath, self.gcpbucket)
            blob.upload_from_file(file)
        # save file to local
        else:
            path = os.path.join(self.config['path']['local']['path'], self.user, filename)
            with open(path, 'wb') as f:
                shutil.copyfileobj(file, f)


def converttofloat(x):
    try:
        X = float(x)
    except Exception:
        X = np.nan
    return X


def write_dict_to_worksheet(x, sheetname, workbook):
    worksheet = workbook.add_worksheet(sheetname)
    row = 0
    for key, val in x.items():
        worksheet.write(row, 0, key)
        worksheet.write(row, 1, val)
        row += 1


def read_dict_from_worksheet(file, sheet, stream=False):
    if stream:
        ws = xlrd.open_workbook(file_contents=file.read()).sheet_by_name(sheet)
    else:
        ws = xlrd.open_workbook(filename=file).sheet_by_name(sheet)
    status = {}
    row = 0
    while True:
        try:
            status[ws.cell_value(row, 0)] = ws.cell_value(row, 1)
        except Exception:
            break
        row += 1
    return status
