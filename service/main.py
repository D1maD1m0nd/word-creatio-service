import json
import os

import requests
from flask import Flask, request
from flask_cors import CORS
from flask_restful import Resource, Api
from requests.cookies import RequestsCookieJar
import uuid


def readJson():
    f = open('authCookie.json')

    data = json.load(f)

    # Iterating through the json
    # list
    result = {}
    for i in data:
        result[i] = data[i]

    # Closing file
    f.close()

    return result


def sendFile(file, size, name):
    print("PRE_READ")
    BPMCSRF = mCookie["BPMCSRF"]
    headers = {
        "Content-Disposition": f"attachment;filename =${name}",
        "Content-Length": f'{size}',
        "BPMCSRF": BPMCSRF,
        "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "Content-Range": f"bytes 0-{str(size)}/{size + 1}",
    }
    base_url = 'http://bakdev.lexiasoft.ru/0/rest/FileApiService/UploadFile'

    params = {
        'fileapi16721259223571': '',
        'totalFileLength': size,
        'fileId': str(uuid.uuid4()),
        'mimeType': 'application%2Fvnd.openxmlformats-officedocument.wordprocessingml.document',
        'fileName': name,
        'parentColumnName': 'Contact',
        'parentColumnValue': 'c58cbd8d-015c-4597-a43e-00a313eb3033',
        'entitySchemaName': 'ContactFile'
    }
    r = requests.post(base_url,
                      params=params,
                      headers=headers,
                      cookies=mCookie,
                      data=file)
    print("SEND_FILE")
    print(r.json())
    return r


class FileSample(Resource):
    def post(self):
        data = request.files['file']  # pass the form field name as key
        print(data.content_length)

        data.save(data.filename)
        file_stats = os.stat(data.filename)
        in_file = open(data.filename, "rb")# opening for [r]eading as [b]inary
        fileData = in_file.read()  # if you only wanted to read 512 bytes, do .read(512)
        in_file.close()

        resp = sendFile(fileData, file_stats.st_size, data.filename)
        return resp.json()


class AuthCreatio(Resource):
    def post(self):
        url = 'http://bakdev.lexiasoft.ru/ServiceModel/AuthService.svc/Login'
        myobj = {
            "UserName": "Supervisor",
            "UserPassword": "Supervisor"
        }

        x = requests.post(url, json=myobj)
        for key, morsel in x.cookies.items():
            mCookie[key] = morsel
        with open('authCookie.json', 'w') as convert_file:
            convert_file.write(json.dumps(mCookie))
        return x.json()


app = Flask(__name__)
CORS(app)
api = Api(app)
mCookie = readJson()

api.add_resource(AuthCreatio, '/auth')
api.add_resource(FileSample, '/file')
if __name__ == '__main__':
    app.run(debug=True)
