#! python3

# モジュールのインポート

from bs4 import BeautifulSoup
import requests, time, re, openpyxl, os, sys


# 関数定義

def get_number_of_papers(query_word):
    
    """
    検索キーワードの対象論文数を表示(引数：検索キーワード)
    """
    
    URL = "https://saemobilus.sae.org/search/?op=navigatePage&pageNumber=1\
        &conditions%5B0%5D.keyword={0}".format(query_word)
    headers = {"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.116 Safari/537.36"}
    resp = requests.get(URL, timeout=5, headers=headers)
    time.sleep(2)
    soup = BeautifulSoup(resp.text, "html.parser")

    result = soup.find(class_ = "filter-number").text
    number = re.sub(r"\D", "", result)

    print("対象件数は" + number + "件です．")
    
    # 前回取得した文献数の入力
    print("前回までに取得した件数を入力してください．")
    num_past = input(">> ")
    page_start = int(-(-int(num_past)/10))

    #取得したい論文数の入力
    print("今回取得する文献数を入力してください．\
        \n前回取得した文献と重複がないように取得します．")
    num_str = input(">> ")
    page_end = int(-(-int(num_str)/10)) + int(-(-int(num_past)/10))

    return page_start, page_end


def webscraping(query_word, num):
    
    """
    検索キーワードの対象文献数を表示して，ページのソースを返す（引数：検索キーワード）
    """
    
    URL = "https://saemobilus.sae.org/search/?op=navigatePage&pageNumber={0}\
        &conditions%5B0%5D.keyword={1}".format(num+1, query_word)
    headers = {"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.116 Safari/537.36"}
    try:
        resp = requests.get(URL, timeout=10, headers=headers)
        time.sleep(1)
        soup = BeautifulSoup(resp.text, "html.parser")
    except:
        soup = None
    
    return soup


def get_paper_infomation(src, list_empty):

    """
    論文番号，タイトル，アブストラクトを返す(引数：ソース，検索キーワード，論文数)
    """
    
    # 論文番号の取得
    paper_numbers = src.find_all(class_="paper-number")
    
    # 各論文のページのソースを取得し，文献番号，タイトル，アブストラクトをリストとして返す
    for papernumber in paper_numbers:
        papernumber = papernumber.text
        URL = "https://saemobilus.sae.org/content/{}".format(papernumber)
        headers = {"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.116 Safari/537.36"}
        
        try:
            resp = requests.get(URL, timeout=10, headers=headers)
            time.sleep(1)
            soup = BeautifulSoup(resp.text, "html.parser")

            # 各種情報の取得
            # タイトル
            try:
                    
                title = soup.find("h1") or "NONE"
                title = title.text
            
            except:
                title = "NONE"

            # アブストラクト
            try:

                abstract = soup.find(class_="htmlview paragraph") or "NONE"
                abstract = abstract.text
            
            except:
                abstract = "NONE"

            # 発行年
            try:

                year = soup.find(class_="published") or "NONE"
                year = year.text
                year = re.sub(".*,", "", year)

            except:
                year = "NONE"

            # 著者
            try:

                authors = soup.find(class_="authors")
                authors = authors.find_all("a")

                authors_list = []

                for author in authors:
                    author = author.text
                    author = author.strip()
                    authors_list.append(author)
                
                authors = ", ".join(authors_list)
            
            except:
                authors = "NONE"

            # セクター
            try:

                sector = soup.find(id="sector")
                sector = sector.a.text

            except:
                sector = "NONE"


            # トピック
            try:

                topics = soup.find(class_="group") or "NONE"
                topics = topics.find_all("a") or "NONE"
            
                topics_list = []

                for v in topics:
                    v = v.text
                    v = v.strip()
                    topics_list.append(v)
                
                topics = ", ".join(topics_list)

            except:
                topics = "NONE"
        
        except:
            title = "NONE"
            abstract = "NONE"
            year = "NONE"
            authors = "NONE"
            sector = "NONE"
            topics = "NONE"

        list_empty.append([papernumber, title, abstract, year, authors, sector, topics])




def datawrite_excel(query_word, list_paper_inforamtion, folder_pass):

    """
    Excelに論文番号，タイトル，アブストラクトを出力
    """

    # ExcelBookの作成
    wb = openpyxl.Workbook()
    sheets = wb.sheetnames
    sheet = wb._sheets[0]

    # Excelsheetの見出しの設定
    title0 = "number of results"
    title1 = "Paper number"
    title2 = "Paper title"
    title3 = "Paper Abstract"
    title4 = "Year"
    title5 = "Authors"
    title6 = "Sector"
    title7 = "Topics"
    
    row_base = 1 
    col_base = 1

    sheet.cell(row = row_base, column= col_base).value = title0
    sheet.cell(row = row_base, column= col_base + 1).value = title1
    sheet.cell(row = row_base, column= col_base + 2).value = title2
    sheet.cell(row = row_base, column= col_base + 3).value = title3
    sheet.cell(row = row_base, column= col_base + 4).value = title4
    sheet.cell(row = row_base, column= col_base + 5).value = title5
    sheet.cell(row = row_base, column= col_base + 6).value = title6
    sheet.cell(row = row_base, column= col_base + 7).value = title7


    # データの書き込み
    for i in range(len(list_paper_inforamtion)):

        sheet.cell(row = row_base + i + 1, column= col_base + 1).value = list_paper_inforamtion[i][0]
        sheet.cell(row = row_base + i + 1, column= col_base + 2).value = list_paper_inforamtion[i][1]
        sheet.cell(row = row_base + i + 1, column= col_base + 3).value = list_paper_inforamtion[i][2]
        sheet.cell(row = row_base + i + 1, column= col_base + 4).value = list_paper_inforamtion[i][3]
        sheet.cell(row = row_base + i + 1, column= col_base + 5).value = list_paper_inforamtion[i][4]
        sheet.cell(row = row_base + i + 1, column= col_base + 6).value = list_paper_inforamtion[i][5]
        sheet.cell(row = row_base + i + 1, column= col_base + 7).value = list_paper_inforamtion[i][6]

    
    #フォルダの確認(無ければ作成)
    if not os.path.exists(folder_pass):
        os.makedirs(folder_pass)
    wb_svpass = folder_pass + "/{}.xlsx".format(query_word)
    
    # ExcelBookの保存，終了
    wb.save(wb_svpass)
    wb.close()


def main():
    """
    1. 変数の設定(キーワード入力，Excelシートのパス指定)
    2. スクリプトの実行
    """

    # 変数の設定
    print("キーワードを入力してください")
    query_word = input(">> ")

    folder_pass = "./data/{}".format(query_word)

    # リストの作成
    list_empty = []

    # 対象文献数の取得
    page_start, page_end = get_number_of_papers(query_word)

    # 論文情報(論文番号，タイトル，アブストラクト)を取得してリストに追加

    number_of_processing = 0

    for num in range(page_start, page_end):
        
        number_of_processing += 1
        
        try:
            src = webscraping(query_word, num)
        
            # 論文情報(論文番号，タイトル，アブストラクト)をリストへ追加
            get_paper_infomation(src, list_empty)
            
            print(str((number_of_processing)*10) + "件処理しました")

        except:
            print(str(number_of_processing*10+1)+"件目から" + str((number_of_processing+1)*10) + "件目までの取得に失敗しました")

    
    # リストの再定義
    list_paper_information = list_empty

    # Excelへの出力
    datawrite_excel(query_word, list_paper_information, folder_pass)

    print("終了しました")

#実行
if __name__ == "__main__":
    main()