# GetGitIssues

## 概要
Github の Issues をエクセルに出力する  
Python 環境で実行する  
出力したエクセルのサンプル issues01.xlsx  

## 実行
```
python GetGitIssues.py "login user" "repository" "outputfilename"
```

GetGitIssues.py  
実行中にプログレスバー表示あり。処理に時間がかかる場合ことらの方がよい。引数あり。    

GetGitIssues_lite.py  
実行中プログレスバーなし。引数なし。コード上に必要なパラメータを記載する。  

## パラメータ
必須
- 第一引数    login user
- 第二引数    repository
- 第三引数    outputfilename

非必須  
private なリポジトリの場合に必要になる。
- 第四引数    login user password
- 第五引数    token


## token の取得方法

個人用アクセストークンは、通常のOAuthアクセストークンと同様に機能します。 HTTPSを介したGitのパスワードの代わりに使用したり、基本認証を介したAPIの認証に使用したりできます。

GUI  
https://qiita.com/kz800/items/497ec70bff3e555dacd0  
https://help.github.com/en/github/authenticating-to-github/creating-a-personal-access-token-for-the-command-line

CUI  
https://qiita.com/rentalname@github/items/9a185d445e45b8b4857c

