# twstockspider

這個小爬蟲是為了更新excel市價的欄位而製作的，可在設定變數的區間設定目標網頁、目標檔案、目標工作表、起始讀取位置及寫入位置。
# Requirements
* Python 3
* openpyxl
* requests
* BeautifulSoup

# 設定變數、import
這邊一開始看不是很懂沒關係，只是之後會用到必須先設定。
```
url = "https://tw.stock.yahoo.com/quote/" # 奇摩股市url
targetfile = '投資總表.xlsx'              # 目標檔案
targetsheet = '庫存股票'                  # 目標工作表名稱
targetnumber = ''                        # 目標股號
column = 1                               # 工作表目標 欄
targetrow = 3                            # 工作表目標 列
pricecol = 7                             # 目標股價 欄
```

```
    import openpyxl
    # 必須import
    import requests
    # 整理 html 必須的套件
    from bs4 import BeautifulSoup
```
需要用的的python套件是openpyxl，這個套件可以建立、開啟、讀寫、儲存excel檔案
> 注意 老的excel版本openpyxl可能無法支援，但具體要多新我並沒有去查。

```
    # 開檔案
    workbook = openpyxl.load_workbook(targetfile)
```
workbook指的是一個excel檔案，targetfile需要填入欲讀取的檔案名稱，不在同一資料夾的話要給予[相對路徑](https://ithelp.ithome.com.tw/articles/10268186)。

```
    # 抓到所有資料表
    sheet = workbook.worksheets

    sheetlength = len(sheet)
    # 搜尋所有資料表的名字假如與targetsheet相同就跳出去執行搜尋股價的程式
    for targetsheetnumber in range(0,sheetlength):
        if sheet[targetsheetnumber].title==targetsheet:
            break
```
sheet為工作表，在這邊因為我不確定未來會不會再增加其他的表在目標的前後，所以先讀取所有工作表再去比對我要的是哪一個。
搜尋是由大而小搜尋抓到檔案之後先找到目標工作表 ( targetsheet ) ，當找到想要的 title 之後就立刻跳出 ( break ) 此迴圈往下走。

```
# 台股總共971支(應該)不會紀錄超過這個range
for row in range(targetrow,900):
    # 如果為空單位即停止搜尋
    if sheet[targetsheetnumber].cell(row,column).value==None:
        break
```
從 row = 3 開始搜尋，儲存格的內容讀取方式是 .value 。

```
    # 讀取儲存格資料強制字串化
    targetnumber = str(sheet[targetsheetnumber].cell(row,column).value)

    # strip()去除前後空白 split('ETF')分割資料為純str(數字) 並且只留下股號的部分
    targetnumber = targetnumber.strip().split('ETF')
    targetnumber = targetnumber[len(targetnumber)-1]
```
需要強制字串化是因為excel會有預設數字的問題，可能在讀到資料後把向 0050 變成 50 從而導致之後的爬蟲失敗。
為確保爬蟲時的url正確，首先先把空白去除乾淨，並且在紀錄時可能會有ETF紀錄的出現，必須先行濾除只留下股號，最後因為 split( ) 會讓原本一個字串變成兩個陣列所以只選取最後一個陣列資料存進 targetnumber。

# 要求、整理、儲存資料

```
    # 要求資料requests.get(url , headers , timeout) ， headers timeout逾期時長為選填
    r = requests.get(str(url)+str(targetnumber), timeout=5)
```
把前置作業做完之後就來到了爬蟲時刻。
```
    # 如果成功要求到資料則執行
    if r.status_code == 200:
        # 整理
        soup = BeautifulSoup(r.text,"html.parser").select("#atomic .Fz\(32px\)")
        # 取出數字輸出
        for price in soup:
            price = price.text
        print(targetnumber , price)
        # 目標股價儲存格更改
        sheet[targetsheetnumber].cell(row,pricecol).value=float(price)
    else:
        print(targetnumber , "未讀取到資料")
        sheet[targetsheetnumber].cell(row,pricecol).value="-"
```
Beautifulsoup 想必有再用爬蟲的人都不陌生，整理一個漂亮的 html 不香嗎。
else 的部分是如果沒有 get 成功會先在 cmd 視窗裡顯現並且在儲存格理儲存一個 " - "。

# 儲存
```
workbook.save(targetfile)
```
**!!!絕對要記得!!!**，這是最重要的一步，沒有做儲存，前面的爬蟲阿處理阿全部都沒用。

## 參考
[openpyxl](https://openpyxl.readthedocs.io/en/stable/)

[openpyxl實作](https://hackmd.io/@howkii-studio/python_autoporcessing_xl)

[twstock](https://hackmd.io/@s02260441/HJcMcnds8)
