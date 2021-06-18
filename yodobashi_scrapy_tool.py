#########################################
#ツール1
#ヨドバシサイトのスプレイピングツール
#########################################

from utils import *
from const import *

#########################
# main処理
#########################

#ロギングの開始
start_logging()
print("[INFO]製品情報の収集処理を開始します。")
logging.info('[INFO]製品情報の収集処理を開始します。')

#ATEGORY_TO_RETRIEVEで指定されたスプレッドシートの全データを取得
splead_datas = get_cagegory_to_retrieve()

#全体の予期せぬエラーをキャッチ
try:
    
    #シート「取得するカテゴリ」に記載されている情報を１つずつ取り出してスクレイピング処理を投げる。
    for index, data in enumerate(splead_datas):
        
        #ループ内で使う変数の初期化
        page_count, new_item_count, next_page_flag = init_variables() 
        if index > 0:   # Header行を飛ばす。
            
            ##カテゴリ名,モデル番号,スクレイピング対象カテゴリURLを取得
            category, model_number, category_url = data[0], data[1], data[2]
            print(f"[INFO]【{category}】の収集処理を開始します。")
            #指定されたカテゴリURLを解析
            bs = r_get(category_url, category)
            
            #Googleスプレッドシートにヘッダ情報を書き込む（既にシートがあればスキップ）
            add_header_info_to_spread(category, HEADER, 1000, 324)
            datas = [] #スプレッドシートに書き込む全データ
            
            #モデル番号の枝番を取得
            model_count = get_model_count(category)

            #次のページが存在する間ループ
            while next_page_flag:
                print(f"[INFO]【カテゴリ名：{category}】{page_count}ページ目の情報を取得中....")
                logging.info('[INFO]【カテゴリ名：' + category + '】' + str(page_count) + 'ページ目の情報を取得中....')
                try:

                    #Activeなページ内の製品情報だけを取得
                    bs_items = bs.select('div[data-salesinformationcode]')
                except Exception as e:
                    message = '[ERROR]' + 'カテゴリ' + '【' + category + '】の全体の商品情報の取得処理で予期せぬエラーが発生しました。'
                    write_error_log(message, e)              

                #Hitした製品情報に対するループ処理
                for bs_item in bs_items:

                    #在庫情報を取得(li要素が3以下の場合は在庫情報がない
                    if len(bs_item.select(".js_addLatestSalesOrder li" )) >= 4:   
                        
                        #在庫ステータスを取得
                        try:
                            zaiko_status = bs_item.select(".js_addLatestSalesOrder li:nth-of-type(4)" )[0].string
                        except Exception as e:
                            message = '[ERROR]' + 'カテゴリ' + '【' + category + '】の在庫ステータスの取得処理で予期せぬエラーが発生しました。'
                            write_error_log(message, e) 
                        
                        #在庫がある場合に後続処理を実行（なければスキップして次の製品へ進む）
                        if zaiko_status in ZAIKO_STATUS:
                            try:

                                #該当製品のURL情報を取得(例：product/100000001004143210/)
                                part_of_url = bs_item.find(class_="js_productListPostTag js-clicklog js-clicklog_OPT_CALLBACK_POST js-taglog-schRlt js_smpClickableFor cImg").attrs['href']
                            except Exception as e:
                                message = '[ERROR]' + 'カテゴリ' + '【' + category + '】の製品のURL情報の取得処理で予期せぬエラーが発生しました。'
                                write_error_log(message, e) 
                            
                            #製品のURLを生成
                            item_url = create_product_url(part_of_url)  
                            #print(item_url)
                            
                            #製品IDの取得
                            product_id = part_of_url.split("/")[-2]
                            
                            #対象製品ページからHTMLソースを取得
                            logging.info('[INFO]アクセス先URL：' + item_url)
                            bs_target_item = r_get(item_url, category)
                            
                            #製品IDがスプレッドシートに存在するかチェック
                            if is_product_id(category, product_id):
                                
                                #存在する場合は報収取得処理をスキップして次の商品ループへ進む。
                                continue
                            
                            #製品のモデル番号の生成
                            model_number_data = model_number + str(model_count) 
                            
                            #製品情報の取得(製品名,価格,スペック,説明,モデル番号,イメージURL)
                            product_name, product_price, product_spec, description, main_img_url = get_item_info(bs_target_item, category)
                           
                            #製品情報をリストに追加する。
                            datas.append([product_id, product_name,"",zaiko_status, model_number_data,"", "", product_price,"", "", "",
                                      "", "",COMPANY_NAME," "," "," "," "," "," "," "," "," "," ", main_img_url,"","","",description,description])
                            
                            #モデル番号を生成する枝番を＋１する
                            model_count += 1
                            
                            #登録対象製品数のカウントを＋１する。
                            new_item_count += 1
                
                    else:
                        
                        #在庫情報がない製品は情報取得処理をスキップ
                        continue
                
                #次のページが存在するかチェック   
                if is_exists_class_name(bs, "next"):
                    print("次のページ情報を取得します...")
                    
                    #次のページのURLを取得
                    next_page_url = create_product_url(bs.find(class_= "next")["href"][1:])
                    
                    #次のページのHTMLソースを取得
                    bs = r_get(next_page_url, category)
                    page_count += 1
                else:
                    
                    #次のページが存在しない場合はWhileループを抜ける
                    next_page_flag = False

            #最後に対象カテゴリに対して新規に収集した製品情報を一括で書き込む。
            #追加する新規データが１件以上ある場合のみ
            if datas:
                add_product_info_to_spread(category, datas)
                print(f"[INFO]【カテゴリ名：{category}】{str(new_item_count)} 件のデータを登録しました。")
                logging.info('[INFO]' + '【カテゴリ名：' + category + '】' + str(new_item_count) + '件のデータを登録しました。')
            else:
                print(f"[INFO]【カテゴリ名：{category}】に対する新規商品データは１件もありませんでした。")
                logging.info('[INFO]' + '【カテゴリ名：' + category +'】に対する新規商品データは１件もありませんでした。')
    print("[INFO]すべての処理が完了しました。")
    logging.info('[INFO]すべての処理が完了しました。')
    

except Exception as e:
    message = '[ERROR]' + '予期せぬエラーが発生しました。'
    write_error_log(message, e)