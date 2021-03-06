# -*- coding: utf-8 -*-
import json
import requests
import os
import pandas as pd
import openpyxl
import re
import sys

# github issuesをエクセルへ

def main():
    # os.environ["no_proxy"] = "localhost"
    # トークンを設定している場合は以下のtokenをurlの末尾に + toeknで付与する。
    # トークンの検証をしていないので動作しないかもしれない。トークンを使う場合は別途検証が必要。
    user = "Git login user"
    reps = "repository"
    outputfile = "issues01.xlsx"
    pswd = ""    # Git login user password 認証が必要な場合は記述する
    token = ""   # トークンが必要な場合は記述する
    page = 1
    issue_max_page = 1
    json_dict = None

    # issues の max page を取得する
    url = "https://api.github.com/repos/" + user + "/" + reps + "/issues?pre_page=1000&page=" + str(page) + "&filter=all&state=all" + token

    if not pswd :
        response = requests.head(url)
    else :
        response = requests.head(url, auth=(user, pswd))
    
    if response.status_code != 200 :
        print('response error status_code : ' + str(response.status_code))
        sys.exit()

    if ('link' in response.headers._store) :
        headerLinks = response.headers._store['link'][1].split(',')
        for link in headerLinks :
            if re.search(r'rel=\"last\"', link) :
                lastPage = re.findall(r'\&page=\d+\&', link)
                issue_max_page = int(re.findall(r'\d+', lastPage[0])[0])
                break

    # max page　分 issues を取得する
    while page <= issue_max_page :   
        url = "https://api.github.com/repos/" + user + "/" + reps + "/issues?pre_page=1000&page=" + str(page) + "&filter=all&state=all"
        page += 1
        
        if not pswd :
            response = requests.get(url)
        else :
            response = requests.get(url, auth=(user, pswd))

        if response.status_code != 200 :
            print('response error status_code : ' + str(response.status_code))
            sys.exit()

        # 辞書
        if not json_dict :
            json_dict = json.loads(response.text)
        else :
            for jd in json.loads(response.text) :
                json_dict.append(jd)

    # csv
    csv_header = ["No", "issuesNo", "title", "内容", "コメント", "ステータス", "ラベル"]
    csv_body = []

    cnt = 1
    for issue_items in json_dict:
        csv_line_body = []
        csv_line_body.append(str(cnt))
        csv_line_body.append(issue_items["number"])
        csv_line_body.append(issue_items["title"])
        if not issue_items["body"] and ("pull_request" in issue_items) :
            csv_line_body.append("pull_request処理")
        else :
            csv_line_body.append(issue_items["body"])

        if issue_items["comments"] > 0:
            url = issue_items["comments_url"] + token

            if not pswd :
                comments_response = requests.get(url)
            else :
                comments_response = requests.get(url, auth=(user, pswd))

            comments_json = json.loads(comments_response.text)
            comment_str = ""
            for commnet in comments_json:
                comment_str = commnet["updated_at"] + "\n"+ commnet["body"] +  "\n\n" + comment_str

            csv_line_body.append(comment_str)
        else:
            csv_line_body.append("")

        csv_line_body.append(issue_items["state"])
        if issue_items["labels"] :
            csv_line_body.append(issue_items["labels"][0]['description'])

        csv_body.append(csv_line_body)

        cnt = cnt + 1

    # 当初CSVで作っていたが見ずらいのでエクセルにした。
    df = pd.DataFrame(csv_body, columns=csv_header)
    # with open("./sjis.csv", mode="w", encoding="cp932", errors="ignore") as f:
        # df.to_csv(f, index=False)
    
    # excel
    df.to_excel(outputfile, index=False)

if __name__ == '__main__':
    main()
