import requests
import logging
import os
import sys
import gspread
import time
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
from bs4 import BeautifulSoup
from datetime import datetime
from dotenv import load_dotenv
from pathlib import Path
from const import *


################################################################
#環境変数のLOAD

load_dotenv(verbose=True)
JSON_KEYFILE_NAME = DEBUG =os.environ.get("JSON_KEYFILE_NAME")
SPREADSHEET_KEY = DEBUG =os.environ.get("SPREADSHEET_KEY")
CLIENT_ID = DEBUG =os.environ.get("CLIENT_ID")
CLIENT_SECRET = DEBUG =os.environ.get("CLIENT_SECRET")

################################################################
# 呼び出し関数部品
################################################################


def get_cagegory_to_retrieve():
    """
    ATEGORY_TO_RETRIEVEで指定されたスプレッドシートの全データを取得
    引数：なし
    戻り値：読み込んだシートのデータ
    """
    try:
        #ファイルオープン
        ws = open_google_spread()
        #シート「取得するカテゴリ」をアクティブ
        ws = ws.worksheet(CATEGORY_TO_RETRIEVE)
        #シート全体の情報を取得
        splead_datas = ws.get_all_values()
        return splead_datas
    except Exception as e:
        message = '[ERROR]' + 'スプレッドシートの読み込みで予期せぬエラーが発生しました。'
        write_error_log(message, e)


def init_variables():
    """
    変数を初期する関数
    引数：なし
    戻り値：初期化した変数
    """
    page_count = 1    #スクレイピング対象のページ数
    new_item_count = 0    #新規の製品情報の数
    next_page_flag = True    #次のページが存在するかを判定するフラグ
    return page_count, new_item_count, next_page_flag


def start_logging():
    """
    ログのスタート
    カレントディレクトリ直下に以下の形式でログファイルを生成する。
    yodobashi_yyyymmdd-hhmmss.log
    ex) yodobashi_20210604-201132.log
    引数：なし
    戻り値：なし
    """
    logdir= os.path.join(os.getcwd(), "logs")
    if not os.path.exists(logdir):
        # ディレクトリが存在しない場合、ディレクトリを作成する
        os.makedirs(logdir)
    #現在時刻の取得
    date_time = datetime.now().strftime("%Y%m%d-%H%M%S")  
    #ファイル名の生成
    file_name = os.path.join(logdir, "yodobashi_"  + date_time + ".log")
    logging.basicConfig(filename=file_name,level=logging.INFO,format='%(asctime)s %(message)s', 
    datefmt='%m/%d/%Y %I:%M:%S %p')


def write_error_log(message, e):
    """
    エラーログをログファイルに書き込みプログラムを強制終了する。
    引数1：ログに書き込むメッセージ
    引数2：エラー情報
    戻り値：なし
    """
    logging.info(message)
    logging.info('[ERROR]エラー内容:%s', str(e.args))
    logging.exception('Detail: %s', e)
    sys.exit(1)


def r_get(url, category):
    """
    引数のURLに対してHTMLを解析する。
    引数1：アクセス先のURL
    戻り値：beautifulsoupの解析結果(データ型：bs4.BeautifulSoup)
    """
    try:        
        time.sleep(SLEEP_TIME)  #待機
        headers = {'User-Agent': 'Sample Header'}
        response = requests.get(url,  headers=headers)
        response.encoding = response.apparent_encoding
        bs = BeautifulSoup(response.text, 'html.parser')
        if not is_valid_page(bs):
            logging.info('[ERROR]指定したURL：' + url + 'が存在しません。')
            sys.exit(1)
    except Exception as e:
        message = '[ERROR]' + 'カテゴリ' + '【' + category + '】の商品情報の取得処理で予期せぬエラーが発生しました。'
        write_error_log(message, e)
    return bs


def create_product_url(href):
    """
    商品ページのURLを生成する。
    
    引数1：URLの一部(product/100000001004143210/)
    戻り値：URL(https://www.yodobashi.com/product/100000001004143210/)
    """
    url = "https://www.yodobashi.com/" + href
    return url



def is_exists_class_name(bs, class_name):
    """
    引数で受け取ったクラス名がbs内に存在するかをチェックする関数
    class="next"が存在するかどうかで判定する。
    引数1：beautifulsoupの解析結果(データ型：bs4.BeautifulSoup)
    引数2：存在チェックするclassの名称
    戻り値：True(存在する) or False(存在しない)
    """
    is_exists_flag = False
    try:
        # class有無を判定する
        if len(bs.select('.' + class_name)) > 0 or class_name in bs["class"]:
            is_exists_flag = True
    except (KeyError, AttributeError):
        # 存在しない場合（＝エラーの場合）はfalseを返却する
        pass
    return is_exists_flag



def open_google_spread():
    """
    グーグルスプレッドシートにアクセスするためのオブジェクトを生成する
    引数：なし
    戻り値：共有設定したスプレッドシートをオープンしたオブジェクト
    """    
    #2つのAPIを記述しないとリフレッシュトークンを3600秒毎に発行し続けなければならない
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    #認証情報設定
    #ダウンロードしたjsonファイル名をクレデンシャル変数に設定（秘密鍵、Pythonファイルから読み込みしやすい位置に置く）
    credentials = ServiceAccountCredentials.from_json_keyfile_name(JSON_KEYFILE_NAME, scope)
    #OAuth2の資格情報を使用してGoogle APIにログインします。
    gc = gspread.authorize(credentials)
    #共有設定したスプレッドシートキーを変数[SPREADSHEET_KEY]に格納する。
    #SPREADSHEET_KEY = '1wJwtfPjnapc2pkT9Vhg734TGdtZloZRcj6OjcrrJn_A'
    #共有設定したスプレッドシートを開く
    ws = gc.open_by_key(SPREADSHEET_KEY)
    return ws


def add_product_info_to_spread(category, datas):
    """
    引数で指定されたカテゴリ名のシートに製品情報を書き込む。
    引数1：カテゴリ名
    引数2：スクレイピングで収集した製品情報（複数ある場合はすべてのデータを含む）
    戻り値：なし
    """
    try:
         #スプレッドシートを開く
        ws = open_google_spread()
        #追加したカテゴリーシートを選択
        ws = ws.worksheet(category)
        #取得した製品データを最終行の下に追記
        #A列のデータを配列として取得
        A_COL_ARRAY = ws.col_values(1)
        #最下行インデックスを取得
        LAST_ROW_IDX = len(A_COL_ARRAY)
        ws.append_rows(datas, table_range="A" + str(LAST_ROW_IDX + 1))
    except Exception as e:
        message = '[ERROR]' + 'カテゴリ' + '【' + category + '】の製品情報の登録またはダウンロード処理で予期せぬエラーが発生しました。'
        write_error_log(message, e)  
    
    
def add_header_info_to_spread(category, header,rows, cols): 
    """
    引数で指定されたカテゴリ名のシートにヘッダ情報を書き込む。
    第1引数で指定したシートが存在しない場合はシートを生成する。
    引数1：カテゴリ名
    引数2：ヘッダ情報
    引数3：生成するシートの行数
    引数4：生成するシートの列数
    戻り値：なし
    """
    try:
        #スプレッドシートを開く
        ws = open_google_spread()
        #カテゴリ名のシートを作成
        try:
            ws.add_worksheet(title=category,rows=rows, cols=cols)
        except gspread.exceptions.APIError:
            #既にシートが存在する場合AIPErrorが発生するのでスキップして処理を続行させる。
            pass
            
        #追加したカテゴリーシートを選択
        ws = ws.worksheet(category)
        #Header情報を追記
        ws.append_row(header)
    except Exception as e:
        message = '[ERROR]'  + category + 'シートへのヘッダ情報書込みで予期せぬエラーが発生しました。'
        write_error_log(message, e)


def get_last_index(category):
    """
    引数で指定されたカテゴリ名のシートの最終行のインデックス番号を取得する。
    引数1：カテゴリ名
    戻り値：シートの最終行インデックス番号
    """
    try:
        #ファイルオープン
        ws = open_google_spread()
        #シート「取得するカテゴリ」をアクティブ
        ws = ws.worksheet(category)
        #A列のデータを配列として取得 
        A_COL_ARRAY = ws.col_values(1)
        #最下行インデックスを取得
        LAST_ROW_IDX = len(A_COL_ARRAY)
    except Exception as e:
        message = '[ERROR]'  + category + 'シートの最終行INDEX番号の取得処理で予期せぬエラーが発生しました。'
        write_error_log(message, e)

    return LAST_ROW_IDX




def is_product_id(category, product_id):
    """
    引数で指定されたカテゴリ名の製品IDがスプレッドシート内に存在するかチェックする。
    引数1：カテゴリ名
    引数2：チェック対象の製品ID
    戻り値：True(存在する） or False(存在しない）
    """
    try:
        #ファイルオープン
        ws = open_google_spread()
        #シートをアクティブ
        try:
            ws = ws.worksheet(category)
        except gspread.exceptions.APIError as e:
            #既にシートが存在する場合AIPErrorが発生するのでスキップして処理を続行させる。
            message = '[ERROR]' + 'gspred関連のエラーが発生しました。エラー内容を確認してください。'
            write_error_log(message, e)
        #A列(商品ID)のデータを配列として取得
        product_array = ws.col_values(1)
    except Exception as e:
        message = '[ERROR]' + 'カテゴリ' + '【' + category + '】の製品ID存在チェック処理で予期せぬエラーが発生しました。'
        write_error_log(message, e)  
    return (product_id in product_array)


def get_item_info(bs_target_item, category):
    """
    以下の製品情報を取得または生成する。
    
    製品名/製品価格/製品のイメージURL/製品の説明/製品の概要/

    引数1：製品ページのbeautifulsoup解析結果(データ型：bs4.BeautifulSoup))
  
    戻り値：product_name（製品名）, product_price（製品価格）, product_spec
    　　　　description（製品説明）, main_img_url（製品のイメージURL)
    """
    try:
        #製品名の取得
        product_name = bs_target_item.find("p", class_= "js_ppPrdName").text
        #製品価格の取得
        product_price =bs_target_item.find("span", class_="price js_ppSalesPrice").text.split("￥")[-1]
        #メインイメージ
        main_img_url = bs_target_item.find(id="mainImg")["src"]
        #製品説明の取得
        product_description = (bs_target_item.find(id="pinfo_productSummury")).get_text('\n')
        #製品概要と正しく一致比較させるために改行コードを削除
        product_description = product_description.replace( '\n' , '' )

        #製品概要と製品説明を比較するために改行を除外したデータを生成
        #product_description_excluded = [elem.replace("\n","") for elem in product_description]
        #product_description_excluded = '\n'.join(product_description_excluded)

        #製品概要一式の取得
        bs_product_overview = bs_target_item.select("ul.pDescription li div.pDesBody")
        
        if bs_product_overview: #製品概要が空の場合がある。
            product_overview = [ overview.get_text('\n') for overview in bs_product_overview]
            #連結して１つの文字列にする。
            product_overview = ''.join(product_overview)
            #製品説明と正しく一致比較させるために改行コードを削除
            product_overview = product_overview.replace( '\n' , '' )
        else:
            #製品概要が空の場合はエラーを回避するため「概要情報なし」を入れる。
            product_overview = "製品の概要情報なし"

        #製品のスペック情報を取得
        bs_product_spec = bs_target_item.select('table[class$="specTbl"]')
        product_spec = [ product_spec.get_text('\n') for product_spec in bs_product_spec]
        product_spec = '\n'.join(product_spec)

        if bs_product_overview:
            #if product_description_excluded == bs_product_overview[0].text:
            if product_description == product_overview:
                description = main_img_url + "\n■商品説明\n" + product_description + "■\n商品スペック\n" + product_spec
            else:
                description = main_img_url +"\n■商品説明\n" + product_description +"\n■商品概要\n" + product_overview + "\n■商品スペック\n" +product_spec
        else:
            #製品概要がない場合は商品説明だけ記載する。
            description = main_img_url +"\n■商品説明\n" + product_description  + "\n■商品スペック\n" +product_spec
    except Exception as e:
        message = '[ERROR]' + 'カテゴリ' + '【' + category + '】の製品情報の取得処理で予期せぬエラーが発生しました。'
        write_error_log(message, e)
    return product_name, product_price, product_spec, description, main_img_url
           

def is_valid_page(bs):
    """
    アクセス先のURLページが存在するかチェックする
    引数1：製品ページのbeautifulsoup解析結果(データ型：bs4.BeautifulSoup))
  
    戻り値：True（ページが存在する）,False（ページが存在しない
    """
    try:
        if bs.select("#contents > div > div.nfTtl")[0].string == 'ご指定のページが見つかりません':
            return False
    except IndexError:
        return True


def write_to_excel(category):
    """
    Googleスプレッドシート上の指定したカテゴリシートのデータをExcelに出力する。
    生成されるファイル名：yodobashi_<カテゴリ名>_yyyymmdd_hhmmss.xlsx
    例：yodobashi_グミ_20210604_214342.xlsx
    引数1：取得したスプレッドシートのデータ
    引数2：カテゴリ名    
    戻り値：なし
    """
    #ファイルオープン
    ws = open_google_spread()
    #指定カテゴリのシートを選択
    ws = ws.worksheet(category)
    #指定したカテゴリシートの全データを取得
    spread_datas = ws.get_all_values()
    #翻訳処理を開始
    for data in spread_datas[1:]:
        #製品名(B列）を日本語→韓国語に翻訳
        data[1] = translate_ja_to_ko(data[1],category)
        #説明(AC列）を日本語→韓国語に翻訳
        data[28] = translate_ja_to_ko(data[28],category)
        
    # Excelファイルのカラムを定義
    Coulum = spread_datas[0]
    # データフレームを作成
    df = pd.DataFrame(spread_datas[1:], columns=Coulum)
    #ファイルパスを生成
    file_dir= os.path.join(os.getcwd(), "data")
    if not os.path.exists(file_dir):
        # ディレクトリが存在しない場合、ディレクトリを作成する
        os.makedirs(file_dir)
    #現在日時を取得
    date_time = datetime.now().strftime("%Y%m%d_%H%M%S")  
    file_path = os.path.join(file_dir, "yodobashi_" + category + "_"  + date_time + ".xlsx")
    df.to_excel( file_path, sheet_name=category, index=False)


def translate_ja_to_ko(text,category):
    """
    短文翻訳API(Papago Text Translation API)を利用して引数で受け取った
    日本語文章を韓国語に翻訳する。
   
    サイトURL：https://www.ncloud.com/product/aiService/papagoNmt
    引数1：文章（日本語）
    戻り値：文章（韓国語）
    """
    POST_URL = "https://naveropenapi.apigw.ntruss.com/nmt/v1/translation"
    # ボディーの定義
    request_body = {'source': 'ja',
                    'target': 'ko',
                    'text': text}
    # ヘッダ情報の定義
    headers = {"X-NCP-APIGW-API-KEY-ID":CLIENT_ID,
               "X-NCP-APIGW-API-KEY": CLIENT_SECRET,
               "Content-Type Content-Type" : "application/json"}
    try:
        # POSTリクエスト送信
        response = requests.post(POST_URL, json=request_body, headers=headers)
        response = response.json()
        translatedText = response['message']['result']['translatedText']
        return translatedText

    except Exception as e:
        message = '[ERROR]' + 'カテゴリ' + '【' + category + '】の翻訳処理で予期せぬエラーが発生しました。'
        write_error_log(message, e)


def get_sheet_list():
    """
    Googleスプレッドシート上に存在する全シート名を取得する。
    引数：なし    
    戻り値：シート名のリスト
    """
    ws = open_google_spread()
    #すべてのシート名を取得
    sheet_list = ws.worksheets()
    #シート名を取り出す
    sheet_list = [sheet.title for sheet in sheet_list]
    return sheet_list


def get_translate_categorys():
    """
    Googleスプレッドシートの「翻訳するカテゴリ」シートに記載されているカテゴリ名を取得する。
    引数：なし    
    戻り値：翻訳をするカテゴリのA列のデータ
    """
    #ファイルオープン
    ws = open_google_spread()
    #シートをアクティブ
    ws = ws.worksheet(SHEET_TO_TRANSLATE)
    #A列(商品ID)のデータを配列として取得
    category_array = ws.col_values(1)
    return category_array