########################################################################################
# クエリ条件ファイル読み込み処理
#[概要]クエリ条件ファイルを読みこみクエリ条件を辞書のリストとして返却します。
# [引数1]ツールルートディレクトリパス
# [戻り値]./setting/query_condition.txtの内容を「,」でスプリットした辞書オブジェクトのリスト
########################################################################################

# モジュール読み込み
import os

def read_query_conditions(root_dir_path):
        
        # クエリ条件ファイル絶対パス
        root_file_path = os.path.join(root_dir_path, "setting", "query_condition.txt")
        
        # クエリ条件リスト（{スキーマ, テーブル物理名, データ取得条件}を各行に格納）
        query_condition_list = []
        
       # クエリ条件ファイル読み込んで辞書型のオブジェクトに変換
        with open(root_file_path, 'r', encoding='utf-8') as query_file:
            for query_condition in query_file:
                
                #始まりが「#」もしくは空の行はスキップする。
                if query_condition.startswith("#"):
                    continue
                if query_condition == "\n":
                    continue
                    
                splitted_query_condition = query_condition .replace("\n","").split(",", 2)
                
                query_condition_list.append({"schema":splitted_query_condition[0].strip(), "table":splitted_query_condition[1].strip(), "where":splitted_query_condition[2].strip()})
                
        return query_condition_list