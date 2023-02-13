########################################################################################
# ツール設定ファイル読み込み処理
# [概要]ツール設定ファイルを読み込んで辞書として返却します。
# [引数1]ツールルートディレクトリパス
# [戻り値]./setting/tool_setting.txtの内容を「=」でスプリットした辞書オブジェクト
########################################################################################

# モジュール読み込み
import os

def read_tool_settings(root_dir_path):

    # ツール設定ファイル絶対パス
    root_file_path = os.path.join(root_dir_path, "setting", "tool_setting.txt")

    # ツール設定ルール辞書
    setting_dict = {}

    # ツール設定ファイル読み込んで辞書型のオブジェクトに変換
    with open(root_file_path, 'r', encoding='utf-8') as setting_file:
        for setting_rule in setting_file:

            #始まりが「#」もしくは空の行はスキップする。
            if setting_rule.startswith("#"):
                continue
            if setting_rule == "\n":
                continue

            setting_item = setting_rule.replace("\n","").split("=")[0].strip()
            setting_value = setting_rule.replace("\n","").split("=")[1].strip()

            setting_dict[setting_item] = setting_value

    return setting_dict