########################################################################################
# DB検索およびワークブックアウトプット処理
#[概要]DBに接続しクエリ条件ファイルに設定された条件をもとにテーブル論理名、カラム論理名、データの取得を行います。
# [引数1]ツール設定辞書
# [引数2]テンプレートSQL辞書
# [引数3]クエリ条件リスト
# [引数4]エクセル操縦クラス
# [戻り値]./sql_templates配下のSQLファイルを読みこんだ辞書オブジェクト        
########################################################################################

# モジュール読み込み
import psycopg2
import pandas
from sshtunnel import SSHTunnelForwarder

def query_and_output_excel(setting_dict, template_sql_dict, query_condition_list, operate_excel):
    
    # SSHトンネルオブジェクト
    ssh_tonnel = None
    # DB接続オブジェクト
    connection = None
    # カーソル
    cursor = None
    
    try:
        # SSHトンネリング
        if setting_dict["use_ssh_tonnel"] == "yes":
            ssh_tonnel = SSHTunnelForwarder((setting_dict['ssh_server_address'], int(setting_dict['ssh_server_port'])),
                                                ssh_host_key = setting_dict['ssh_host_key'] if setting_dict['ssh_host_key'] else None,
                                                ssh_username = setting_dict['ssh_user'],
                                                ssh_password = setting_dict['ssh_password'],
                                                ssh_pkey = setting_dict['ssh_pkey_file_path'] if setting_dict['ssh_pkey_file_path'] else None,
                                                remote_bind_address=(setting_dict['ssh_remote_bind_address'], int(setting_dict['ssh_remote_bind_port']))
                                            )
            # トンネリングスタート
            ssh_tonnel.start()
        
        # DB接続オブジェクト生成
        connection = psycopg2.connect(host = setting_dict['db_host'],
                                      port = setting_dict['db_port'],
                                      database = setting_dict['db_name'],
                                      user = setting_dict['db_user'],
                                      password = setting_dict['db_password'])
        cursor = connection.cursor()
        
        # クエリ条件リスト1要素ごとにクエリとエクセルファイルへの出力を実行
        for query_condition_dict in query_condition_list:
            # 1.テーブル論理名取得およびエクセルファイル出力
            cursor.execute(template_sql_dict['get_table_logical_name'], [query_condition_dict['schema'], query_condition_dict['table']])
            table_names_df = pandas.DataFrame(cursor.fetchall())
            operate_excel.write_table_names(table_names_df)
            
            # 2.テーブルカラム名取得およびエクセルファイル出力
            cursor.execute(template_sql_dict['get_column_logical_names'], [query_condition_dict['schema'], query_condition_dict['table']])
            column_names_df = pandas.DataFrame(cursor.fetchall())
            operate_excel.write_column_names(column_names_df, query_condition_dict)
            
            # 3.データ取得およびエクセルファイル出力
                # カラム名をカンマ区切りの文字列に変形してデータ取得用SQLのカラム名を置換
            get_datas_sql = template_sql_dict['get_datas']
            get_datas_sql = get_datas_sql.replace(":column_names", ", ".join(column_names_df.iloc[:,1].tolist()))
            
                # データ取得用SQLのテーブル名を検索対象のテーブル名に置換
            get_datas_sql = get_datas_sql.replace(":table_name", query_condition_dict['table'])
            
                # データ取得用SQLのwhere句を指定された検索条件に置換
            if query_condition_dict['where']:
                get_datas_sql = get_datas_sql.replace(":condition", "and " + query_condition_dict['where'])
            else:
                get_datas_sql = get_datas_sql.replace(":condition", "")
                
                #データ取得
            cursor.execute(get_datas_sql)
            datas_df = pandas.DataFrame(cursor.fetchall())
                # エクセルファイル出力
            operate_excel.write_datas(datas_df)
            
            
    finally:
        # DB切断
        cursor.close()
        connection.close()
        
        # トンネルクローズ
        if setting_dict["use_ssh_tonnel"] == "yes":
            ssh_tonnel.stop()