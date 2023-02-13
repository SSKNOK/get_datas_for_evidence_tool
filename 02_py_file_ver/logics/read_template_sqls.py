########################################################################################
# テンプレートSQLファイル読み込み処理
#[概要]テーブル論物名取得SQL、カラム論物名取得SQL、データ取得SQLを読み込みます。
# [引数1]ツールルートディレクトリパス
# [戻り値]./sql_templates配下のSQLファイルを読みこんだ辞書オブジェクト
########################################################################################

# モジュール読み込み
import os

def read_template_sqls(root_dir_path):
    
    # テンプレートSQLディレクトリ絶対パス
    template_sql_dir_path = os.path.join(root_dir_path, "sql_templates")
    
    # テンプレートSQL辞書
    template_sql_dict = {}
    
    # テンプレートSQLディレクトリ配下のファイル一覧を取得（再帰的には取得しない）
    files = os.listdir(template_sql_dir_path)
    template_sql_files = [f for f in files if os.path.isfile(os.path.join(template_sql_dir_path, f))]
    
    # テンプレートSQLディレクトリ配下のファイル内容をそれぞれ読み取り
    for template_sql_file in template_sql_files:
        with open(os.path.join(template_sql_dir_path, template_sql_file), 'r', encoding='utf-8') as query_file:
            template_sql_dict[template_sql_file.replace(".sql","")] = query_file.read()
    
    return template_sql_dict
            