# 指定した .pptx ファイルからインデックス用コードを抽出し、index ファイルを出力する

from pptx import Presentation
import os
import sys
from pathlib import Path
import argparse
import re
import json

import logging

import PpIndexConfig as pic
from PpIndexCommon import remove_slides

version='0.9'

# ロガーの作成
logger = logging.getLogger(__name__)

def setLogger(loglevel, logoutput):
    # ログレベルの設定
    if loglevel.upper() == 'DEBUG':
        logger.setLevel(logging.DEBUG)
    else:
        logger.setLevel(logging.INFO)

    # ログファイルの設定
    if logoutput.upper() != 'STDOUT':
        file_handler = logging.FileHandler(logoutput, 'a')
        file_handler.setLevel(logging.DEBUG)
        #file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        file_handler.setFormatter(logging.Formatter('%(message)s'))
        logger.addHandler(file_handler)

# 対象ファイル読み込みエラー、書き込みエラーなど
class ProcessError(Exception):
    def __init__(self, message):
        self.message = message
        super().__init__(message)

    def __str__(self):
        return f"ProcessError: {self.message}"


# ファイル/フォルダーが存在しない場合にエラーを発生させる
def raiseProcessError_if_not_exists(file_path):
    if not os.path.exists(file_path):
        raise ProcessError(f"File/Folder not found: {file_path}")


# コマンドライン引数を解析して返す
def parse_commandargs():
    # コマンドライン引数のパーサーを作成
    parser = argparse.ArgumentParser(description="Read and process a YAML file.")
    # 'file' 引数の定義
    parser.add_argument("file", help="The YAML file to read")
    # オプション引数 --ll 、略称 -l,文字型、デフォルトは info を指定
    parser.add_argument('--ll', '-l', type=str, default='info', help='log level [debug|INFO]')
    # オプション引数 --lf, 略称 -f,文字型、デフォルトは STDOUT を指定
    parser.add_argument('--lf', '-f', type=str, default='STDOUT', help='log file')
    # オプション引数 --dump 、略称 -d,文字型、デフォルトは 空文字列 を指定
    parser.add_argument('--dump', '-d', type=str, default='', help='dump file name')
    # バージョン番号を表示
    parser.add_argument('--version', '-v', action='version', version=f'%(prog)s {version}')
    
    # 引数を解析
    _args = parser.parse_args()
    return _args


# indexファイルをもとに .pptx ファイルのラベル置換を行い、新しい .pptx ファイルを出力する
def generate_target(genparams,folderobj):
    _sourcepptx = genparams['SOURCE']
    _sourcepptxpath = folderobj / _sourcepptx
    _generatepptx = genparams['GENERATE']
    _generatepptxpath = folderobj / _generatepptx
    logger.debug(f"source:{_sourcepptx} generate:{_generatepptx}")

    # _originpptx を読み込む
    prs = Presentation(_sourcepptxpath)

    # 削除対象スライドを削除して、スライド番号マッピング用ハッシュを受け取る key:元のスライド番号 value:新しいスライド番号
    snummap = remove_slides(prs, genparams['CSL'])  

    firstlevelreplacements = []
    secondlevelreplacements = []
    mokujireplacements = []
    # 置換パターンを作成
    for indexfile in genparams['INDEX']:
        _indexpath = folderobj / indexfile
        logger.debug(f"indexfile:{indexfile} ")
        # jsonファイルを読み込む
        with open(_indexpath, 'r', encoding='utf-8') as file:
            indexdata = json.load(file)

        # 置換パターン
        # #DT# タイトル部のインデックス
        # #RI# 参照部　インデックス
        # #RT# 参照部　タイトル
        # #RN# 参照部　スライド番号
        # #RIT# 参照部　インデックスとタイトル
        # firstlevel の置換パターンを作成
        for key in indexdata:
            logger.debug(f"key:{key}")
            logger.debug(f"index:{indexdata[key]['index']}")
            logger.debug(f"slidenumber:{indexdata[key]['slidenumber']}")
            logger.debug(f"slidenumber:{snummap[indexdata[key]['slidenumber']]}")   # スライド削除後の番号へのマッピング対応
            logger.debug(f"title:{indexdata[key]['title']}")
            # #DT#FA1 を indexdata[key]['index'] で置換するためのデータを作成
            replacements = {
                make_replacetargetSingleKey_reg("DT",key): rf"{indexdata[key]['index']}\g<1>",
                make_replacetargetSingleKey_reg("RI",key): rf"{indexdata[key]['index']}\g<1>",
                make_replacetargetSingleKey_reg("RT",key): rf"{indexdata[key]['title']}\g<1>",
                make_replacetargetSingleKey_reg("RN",key): rf"{snummap[indexdata[key]['slidenumber']]}\g<1>",   # スライド削除後の番号へのマッピング対応
                make_replacetargetSingleKey_reg("RIT",key): rf"{indexdata[key]['index']} {indexdata[key]['title']}\g<1>",
            }
            # print(replacements)
            firstlevelreplacements.append(replacements)

            # secondlevel の置換パターンを作成
            secondlevelassoc = indexdata[key]['secondlevelassoc']
            secondlevelidxtitlelist = ""
            secondlevelslidenumberlist = ""
            for skey in secondlevelassoc:
                logger.debug(f"skey:{skey}")
                logger.debug(f"index:{secondlevelassoc[skey]['index']}")
                logger.debug(f"slidenumber:{snummap[secondlevelassoc[skey]['slidenumber']]}")
                logger.debug(f"title:{secondlevelassoc[skey]['title']}")
                # #T#FA1 を indexdata[key]['index'] で置換するためのデータを作成
                replacements = {
                    make_replacetargetDualKey_reg("DT",key, skey): rf"{indexdata[key]['index']}\g<1>{secondlevelassoc[skey]['index']}\g<2>",
                    make_replacetargetDualKey_reg("RI",key, skey): rf"{indexdata[key]['index']}\g<1>{secondlevelassoc[skey]['index']}\g<2>",
                    make_replacetargetDualKey_reg("RT",key, skey): rf"{secondlevelassoc[skey]['title']}\g<2>",
                    make_replacetargetDualKey_reg("RN",key, skey): rf"{snummap[secondlevelassoc[skey]['slidenumber']]}\g<2>",
                    make_replacetargetDualKey_reg("RIT",key, skey): rf"{indexdata[key]['index']}\g<1>{secondlevelassoc[skey]['index']}\g<2> {secondlevelassoc[skey]['title']} ",
                }
                # print(replacements)
                secondlevelreplacements.append(replacements)
                secondlevelidxtitlelist += f"{indexdata[key]['index']}.{secondlevelassoc[skey]['index']}) {secondlevelassoc[skey]['title']}\n"
                secondlevelslidenumberlist += f"{snummap[secondlevelassoc[skey]['slidenumber']]}\n"

            replacements = {
                make_replacetargetSingleKey_reg("RLISTIT",key): rf"{secondlevelidxtitlelist}\g<1>",
                make_replacetargetSingleKey_reg("RLISTSN",key): rf"{secondlevelslidenumberlist}\g<1>",
            }
            mokujireplacements.append(replacements)

    logger.debug('-----------------')
    logger.debug(firstlevelreplacements)
    logger.debug('-----------------')
    logger.debug(secondlevelreplacements)
    logger.debug('-----------------')
    logger.debug(mokujireplacements)


    # slidenumber はスライド削除後のスライド番号を保持
    slidenumber=0

    for slide in prs.slides:
        logger.debug(f"slidenumber:{slidenumber}")
        slidenumber += 1
        tfnumber = 0
        for shape in slide.shapes:
            if shape.has_text_frame:
                tfnumber += 1
                _shapedeleteflag = False
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        # run.text に #CSP# が含まれていたら、そのshapeを削除する
                        if genparams['CSP'] and re.search(r'#CSP#', run.text):
                            _shapedeleteflag = True
                            logger.debug(f"deleting shape:{shape}")
                            break

                        # run.text に #CSP# が含まれていない場合は、置換を行う
                        replacedtext = replace_text(run.text, firstlevelreplacements, secondlevelreplacements,mokujireplacements)
                        run.text = replacedtext
                if _shapedeleteflag:
                    shape.element.getparent().remove(shape.element)                    
                    continue

        # logger.debug(f"tfnumber:{tfnumber}")

    # 変更を保存
    prs.save(_generatepptxpath)


def make_replacetargetDualKey_reg(tag,key, skey):
    # 「ピリオドまたはハイフンで区切ったキー文字列」の後ろに「非英数文字または空白または文末」が続く場合にマッチ
    return rf"#{tag}#{key}([-.]){skey}([^0-9a-zA-Z_.]|\s|$)"
    #return rf"#{tag}#{key}([-.]){skey}([^0-9a-zA-Z_.]+|\s|$)"
    #return rf"#{tag}#{key}([-.]){skey}([^0-9a-zA-Z_.]*\s?$)"

def make_replacetargetSingleKey_reg(tag,key):
    # 「キー文字列」の後ろに「非英数文字または空白または文末」が続く場合にマッチ
    return rf"#{tag}#{key}([^0-9a-zA-Z_.]|\s|$)"




# text に対して firstlevel と secondlevel の置換を行う
def replace_text(text, first, second,mokuji):
    # text に #\w+# が含まれない場合はそのまま返す
    if not re.search(r'#\w+#', text):
        return text
    
    for replacements in second:
        for target, replacement in replacements.items():
            #logger.debug(f"searching text:{text} target:{target} replacement:{replacement}")
            if re.search(target, text):
                replacedtext = textreplacewrapper_reg(text, target, replacement)
                logger.debug(f"replacing text:{text} target:{target} replacement:{replacement}")
                logger.debug(f"  replacedtext:{replacedtext}")
                text = replacedtext

    for replacements in first:
        for target, replacement in replacements.items():
            #logger.debug(f"searching text:{text} target:{target} replacement:{replacement}")
            if re.search(target, text):
                replacedtext = textreplacewrapper_reg(text, target, replacement)
                logger.debug(f"replacing text:{text} target:{target} replacement:{replacement}")
                logger.debug(f"  replacedtext:{replacedtext}")
                text = replacedtext

    for replacements in mokuji:
        for target, replacement in replacements.items():
            #logger.debug(f"searching text:{text} target:{target} replacement:{replacement}")
            if re.search(target, text):
                replacedtext = textreplacewrapper_reg(text, target, replacement)
                logger.debug(f"replacing text:{text} target:{target} replacement:{replacement}")
                logger.debug(f"  replacedtext:{replacedtext}")
                text = replacedtext

    return text


# text に対して target を replacement に置換する
def textreplacewrapper(text, target, replacement):
    logger.debug(f"text:{text} target:{target} replacement:{replacement}")
    return text.replace(target, replacement)


# text に対して target を replacement に置換する、正規表現バージョン
def textreplacewrapper_reg(text, target, replacement):
    logger.debug(f"text:{text} target:{target} replacement:{replacement}")
    replaced_text = re.sub(target, replacement, text)
    logger.debug(f"replaced_text:{replaced_text}")
    return replaced_text


# メイン処理
def main():
    args = parse_commandargs()

    # ロガーの設定
    setLogger(args.ll, args.lf)

    # YAMLファイルを読み込む
    logger.info( "--------- configuration ---------")
    configs = pic.verify_parameter_formats(pic.load_yaml(args.file,logger=logger),logger=logger)
    logger.debug(f"configs: {configs}")
    raiseProcessError_if_not_exists(configs['FOLDER'])

    _folder = configs['FOLDER']
    _folderobj = Path(_folder)

    logger.info( "--------- processing---------")
    _generating=configs['GENERATING']   
    for i in range(len(_generating)):
        # _generating[i] のパス名に拡張子を付与、ファイルの存在確認
        logger.debug(f"generating: {_generating[i]}")
        _generate = pic.add_extension(pic.remove_extension(_generating[i]['GENERATE']), 'pptx')
        _generating[i]['GENERATE'] = _generate
        _generatepathobj = _folderobj / _generate
        if os.path.exists(_generatepathobj):
            # _generatepathobj のファイルが既に存在したら警告
            logger.info(f"Warning: overwriting existing pptx file: {_generatepathobj}")

        _source = pic.add_extension(pic.remove_extension(_generating[i]['SOURCE']), 'pptx')
        _generating[i]['SOURCE'] = _source
        _sourcepathobj = _folderobj / _source
        # _sourcepath のファイルが存在しなかったらエラー
        raiseProcessError_if_not_exists(_sourcepathobj)

        for j in range(len(_generating[i]['INDEX'])):
            _index = pic.add_extension(pic.remove_extension(_generating[i]['INDEX'][j]), 'json')
            _generating[i]['INDEX'][j] = _index
            _indexpathobj = _folderobj / _index
            # _indexpath のファイルが存在しなかったらエラー
            raiseProcessError_if_not_exists(_indexpathobj)

        # ラベリング処理をする
        logger.debug(f"generating:\n  sourcepath:{_sourcepathobj}\n  generatepath:{_generatepathobj}")
        generate_target(_generating[i],_folderobj)

    logger.info(f"{sys.argv[1]} done.")



if __name__ == "__main__":
    try:
        main()
        exit(0)
    except pic.ConfigError as e:
        logger.info(f"ConfigError: {e}")
        exit(1)
    except ProcessError as e:
        logger.info(f"ProcessError: {e}")
        exit(2)

