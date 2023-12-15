"""
!/usr/bin/env python
-*- coding: utf-8 -*-
@File  : read_data.py
@Author: GinaChou
@Date  : 2023/12/12
"""
import requests
from pprint import pprint
import json
import os
import openpyxl

app_id = 'ginachou-1524620d-b8ae-4796'
app_key = 'e0c769a4-0796-4540-b504-d66668428ce1'

auth_url="https://tdx.transportdata.tw/auth/realms/TDXConnect/protocol/openid-connect/token"
# url = "https://tdx.transportdata.tw/api/basic/v2/Road/Link/Shape/Geometry/WKT/0000300140000T?%24format=JSON"
url_geo = 'https://tdx.transportdata.tw/api/basic/v2/Road/Link/Shape/Geometry/WKT/'
url_len = 'https://tdx.transportdata.tw/api/basic/v2/Road/Link/LinkID/'

class Auth():

    def __init__(self, app_id, app_key):
        self.app_id = app_id
        self.app_key = app_key

    def get_auth_header(self):
        content_type = 'application/x-www-form-urlencoded'
        grant_type = 'client_credentials'

        return{
            'content-type' : content_type,
            'grant_type' : grant_type,
            'client_id' : self.app_id,
            'client_secret' : self.app_key
        }

class data():

    def __init__(self, app_id, app_key, auth_response):
        self.app_id = app_id
        self.app_key = app_key
        self.auth_response = auth_response

    def get_data_header(self):
        auth_JSON = json.loads(self.auth_response.text)
        access_token = auth_JSON.get('access_token')

        return{
            'authorization': 'Bearer '+access_token
        }

if __name__ == '__main__':

    params = {
        '$format': 'JSON',
    }
    json_data = [
        '0000101006000F',
        '0000101007000F',
        '0000101007010F',
    ]
    print(f'json_data: {type(json_data)}')
    try:
        d = data(app_id, app_key, auth_response)
        data_response = requests.get(url_geo, headers=d.get_data_header())  #2023/12/12
    except:
        a = Auth(app_id, app_key)
        auth_response = requests.post(auth_url, a.get_auth_header())
        d = data(app_id, app_key, auth_response)
        # data_response = requests.get(url, headers=d.get_data_header())
        res_geo = requests.post(url_geo, params=params,headers=d.get_data_header(),json=json_data)
        res_dist = requests.post(url_len, params=params,headers=d.get_data_header(),json=json_data)
    print(f'auth:{auth_response}') #AUTH標頭回應
    pprint(f'auth_text:{auth_response.text}')

    ##POST產出經緯度、路段長度
    geo_list = json.loads(res_geo.text)
    dist_list = json.loads(res_dist.text)
    # pprint(f'res_geo_text:{res_geo.text}')
    # print(f'data_type: {type(geo_list)}')
    print(f'geometry: {geo_list}')
    print(f'res_dist: {dist_list}')

    ##切換到.py檔目前的絕對路徑，並於底下新增output資料夾
    result_path = os.path.join(os.path.dirname(os.path.abspath(__file__)),'output')
    ##檢查是否存在output資料夾，若無則新增
    if not os.path.exists(result_path):
        os.makedirs(result_path)
    # # os.chdir 是 python 切換到電腦指定路徑的方法
    os.chdir(result_path)

    wb = openpyxl.Workbook()
    sheet = wb.create_sheet("POST", 0) # 新增工作表 1
    result = []
    ##為EXCEL增加標頭
    titles = ("LinkID","Geometry","Version","UpdateDate","Distance")
    sheet.append(titles)
    LinkID =""
    res_dist =""
    Geometry =""
    Version =""
    UpdateDate =""

    # 輸出link資訊(用分號合併)
    for link in geo_list:
        ls = []
        res_dist = 0
        # print("link:", link,"link_type:",type(link))

        if LinkID =="":
            LinkID = link["LinkID"]
        else:
            LinkID = LinkID +";" + link["LinkID"]
        Geo = link["Geometry"]
        ##只保留經緯度資料
        Geo = Geo.replace("LINESTRING", "")
        Geo = Geo.replace("(", "")
        Geo = Geo.replace(")", "")
        Geo = Geo.replace(",", ";")  ##逗號改為分號
        if Geometry =="":
            Geometry = Geo
        else:
            Geometry = Geometry +";" + Geo
        if Version =="":
            Version = link["Version"]
        else:
            Version = Version +";" + link["Version"]
        if UpdateDate =="":
            UpdateDate = link["UpdateDate"]
        else:
            UpdateDate = UpdateDate +";" + link["UpdateDate"]

        ls.append(LinkID)
        ls.append(Geometry)
        ls.append(Version)
        ls.append(UpdateDate)
        for link in dist_list:
            res_dist += link["Length"]
            print("res_dist:", res_dist)
        ls.append(res_dist)

        result.append(ls)
        print("result:", result)

    print("FINAL_result[-1]:", result[-1])
    sheet.append(result[-1])
    wb.save("Geometry.xlsx")