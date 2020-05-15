# -*- coding: utf-8 -*-
import json
import requests
import os
import pandas as pd
import openpyxl
import re
import time
from tqdm import tqdm
import sys

# github issuesをエクセルへ

def main():
    # os.environ["no_proxy"] = "localhost"
    # token = ""
    # トークンを設定している場合は以下のtokenをurlの末尾に + toeknで付与する。
    # トークンの検証をしていないので動作しないかもしれない。トークンを使う場合は別途検証が必要。
    # token = "?access_token=生成されたtokentoken"
    # user = "github login user"
    # pswd = "github login user password"
    # reps = "github login user repository"
    # outputfile = "issues.xlsx"
    user = args[1]
    reps = args[2]
    outputfile = args[3]
    pswd = args[4] if len(args) == 5 else ""
    token = args[5] if len(args) == 6 else ""
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
    # while page <= issue_max_page :   
    max_pages = list(range(issue_max_page))  # プログレスバーを表示するため配列にする
    for page in tqdm(max_pages) :
        page += 1
        url = "https://api.github.com/repos/" + user + "/" + reps + "/issues?pre_page=1000&page=" + str(page) + "&filter=all&state=all"
        
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
    # for issue_items in json_dict:
    for issue_items in tqdm(json_dict):
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

    book = openpyxl.load_workbook(outputfile)
    sheet = book['Sheet1']

    counter = 1
    while counter <= sheet.max_row:
        # 行の高さ変更
        sheet.row_dimensions[counter].height = 30
        if counter != 1 :
            # 上付、左揃え
            sheet['A' + str(counter)].alignment = openpyxl.styles.Alignment(horizontal = 'left', vertical = 'top')
            sheet['B' + str(counter)].alignment = openpyxl.styles.Alignment(horizontal = 'left', vertical = 'top')

            # セル内折り返し
            sheet['C' + str(counter)].alignment = openpyxl.styles.Alignment(horizontal = 'left', vertical = 'top', wrapText=True)
            sheet['D' + str(counter)].alignment = openpyxl.styles.Alignment(horizontal = 'left', vertical = 'top', wrapText=True)
            sheet['E' + str(counter)].alignment = openpyxl.styles.Alignment(horizontal = 'left', vertical = 'top', wrapText=True)

        counter+=1
    
    # 列の幅を変更
    sheet.column_dimensions['C'].width = 30
    sheet.column_dimensions['D'].width = 50
    sheet.column_dimensions['E'].width = 50

    book.save(outputfile)

if __name__ == '__main__':
    args = sys.argv
    if len(args) < 4 :
        print("Usages $ python GetGitIssues.ph $1 $2 $3 [$4 $5] \n"
              "第一引数 $1 : Git login user (required) \n" +
              "第二引数 $2 : Git repository (required) \n" +
              "第三引数 $3 : Git outputfile_name (required) \n" +
              "第四引数 $4 : Git login user password \n" +
              "第五引数 $5 : token \n")
        sys.exit()

    main()
