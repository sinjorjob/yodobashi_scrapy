#########################################
#ツール2
#翻訳ツール
#########################################


from utils import *
from const import *
from colorama import init
import time

init()

#######################
# main処理
#######################

#ロギングの開始
start_logging()
print("[INFO]翻訳処理を開始します。")
logging.info('[INFO]翻訳処理を開始します。')

#A列(翻訳するカテゴリ名)のデータを配列として取得
category_array = get_translate_categorys()

#Googleスプレッドシート上に存在する全シート名を取得する。
sheet_list = get_sheet_list()

#１カテゴリずつ翻訳処理→Excelダウンロードを実施
for category in category_array[1:]:
    
    #シート名が存在していたら処理を続行
    if category in sheet_list:
        #製品（B列）、説明（AC列）を翻訳した後Excelファイルを出力。
        write_to_excel(category)
    else:
        print(Fore.YELLOW + f"[WARNING]カテゴリ名:【{category} 】のシートが存在していません。")
        logging.info('[WARNING]カテゴリ名:' + category +'のシートが存在していません。')

print(Fore.WHITE + "[INFO]すべての処理が完了しました。")
logging.info('[INFO]すべての処理が完了しました。')
time.sleep(30)
