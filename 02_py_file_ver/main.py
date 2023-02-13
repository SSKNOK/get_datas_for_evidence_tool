########################################################################################
# main
########################################################################################

# モジュール
import PySimpleGUI
import traceback
import os
# 自作のモジュール
from logics.read_tool_settings import read_tool_settings
from logics.read_query_conditions import read_query_conditions
from logics.read_template_sqls import read_template_sqls
from logics.query_and_output_excel import query_and_output_excel
from classes.operate_excel import operate_excel


# Jupyter用の設定
root_dir_path = os.getcwd()

try:
    # ツール設定ファイル読み込み
    setting_dict = read_tool_settings(root_dir_path)

    # クエリ条件ファイル読み込み
    query_condition_list = read_query_conditions(root_dir_path)

    # エクセルオブジェクト生成
    operate_excel = operate_excel(root_dir_path, setting_dict)

    # テンプレートSQLファイル読み込み
    template_sql_dict = read_template_sqls(root_dir_path)

    # クエリ実行
    query_and_output_excel(setting_dict, template_sql_dict, query_condition_list ,operate_excel)

    # エクセル保存
    operate_excel.save_excel_file()

    
# エラッた場合はGUIで表示
except Exception as e:
    
    # PySimpleGUIウィンドウ設定
    PySimpleGUI.theme('Dark Blue 3')
    
    # 表示内容設定
    layout = [
        [PySimpleGUI.Text('type:' + str(type(e)))],
        [PySimpleGUI.Text('args:' + str(e.args))],
        [PySimpleGUI.Text('trace:' + traceback.format_exc())],
        [PySimpleGUI.Submit(button_text='OK')]
    ]

    # PySimpleGUIウィンドウ表示
    window = PySimpleGUI.Window('ERROR', layout)

    while True:
        event, values = window.read()

        if event is None:
            break
        if event == 'OK':
            break

    # PySimpleGUIウィンドウの破棄と終了
    window.close()