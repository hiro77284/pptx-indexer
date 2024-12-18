# 指定した .pptx ファイルからインデックス用コードを抽出し、index ファイルを出力する

import os
import yaml
import re


# 設定ファイル内のファイル名置換用正規表現
__pattern_with_key_suffix = re.compile(r'<([A-Za-z][0-9A-Za-z_]*):([0-9A-Za-z_]+)>')
__pattern_key_only = re.compile(r'<([A-Za-z][0-9A-Za-z_]*)>')
__pattern_key_notacceptable = re.compile(r'.*[<>].*')   # 変換処理後に angle bracket が残っていたらエラー
# バージョンチェック用（これより小さいバージョンのファイルは処理しない）
__MINIMUMVERSION__ = 1.0


# 設定ファイル読込失敗、フォーマット不正、指定されたファイルが存在しない、などのエラー
class ConfigError(Exception):
    def __init__(self, message):
        self.message = message
        super().__init__(message)

    def __str__(self):
        return f"ConfigError: {self.message}"

# VERSION が存在しないか、最小値以下の場合はエラー
def verify_version(configs, logger=None):
    if logger is None:
        import logging
        logger = logging.getLogger(__name__)

    if 'VERSION' not in configs:
        raise ConfigError("VERSION not found in the YAML file.")
    if configs['VERSION'] < __MINIMUMVERSION__:
        raise ConfigError(f"Config's VERSION is too old: {configs['VERSION']}")

    logger.info(f"VERSION: {configs['VERSION']}")
    return configs

# 指定されたYAMLファイルを読み込む
def load_yaml(file_path, dump='', logger=None):
    if logger is None:
        import logging
        logger = logging.getLogger(__name__)

    try:
        with open(file_path, 'r', encoding='utf-8') as file:
            _data = yaml.safe_load(file)
        # 読み込んだデータを表示
        logger.debug(f"load_yaml: {_data}")
        # ダンプファイルが指定されていたら、そのファイルにダンプする
        if dump:
            with open(dump, 'w', encoding='utf-8') as file:
                yaml.dump(_data, file,sort_keys=False)
        return _data
    except FileNotFoundError:
        raise ConfigError(f"File not found: {file_path}")
    except yaml.YAMLError as e:
        raise ConfigError(f"YAML format error: {e}")
    except Exception as e:
        raise ConfigError(f"Unexpected error: {e}")


# 拡張子を削除
def remove_extension(pathname):
    return os.path.splitext(pathname)[0]

# 拡張子を付加
def add_extension(pathname, ext):
    return f"{pathname}.{ext}"


# パラメータを置換する
# <A> は targets['A'] に置換する
# <A:add> は targets['A'] + 'add' に置換する
# 置換処理後に angle bracket <> が残っていたらエラー
# いずれにも該当しない場合は pathname をそのまま返す
def modify_filenames(pathname, targets):
    # 書式が <英字の1文字以上> にマッチする正規表現
    global __pattern_with_key_suffix, __pattern_key_only, __pattern_key_notacceptable
    match = __pattern_with_key_suffix.match(pathname)
    try:
        if match:
            return f"{targets[match.group(1)]}{match.group(2)}"
    except KeyError:
        raise ConfigError(f"unmatch TARGETS key on <KEY:SUFFIX> {match.group(1)}")

    match = __pattern_key_only.match(pathname)
    try:
        if match:
            return targets[match.group(1)]
    except KeyError:
        raise ConfigError(f"unmatch TARGETS key on <KEY> {match.group(1)}")

    match = __pattern_key_notacceptable.match(pathname)
    if match:
        raise ConfigError(f"unconvertible angle brackets <...> format on: {pathname}")

    return pathname


# 設定値を検証する。 logger を指定されたらそれを使う
def verify_parameter_formats(configs, dump='', logger=None):
    if logger is None:
        import logging
        logger = logging.getLogger(__name__)

    # VERSION が存在しないか、最小値以下の場合はエラー
    verify_version(configs, logger)

    # FOLDER がない場合はカレントディレクトリを設定する
    if 'FOLDER' not in configs:
        logger.info("FOLDER not found in the YAML file, so set the current directory.")
        configs['FOLDER'] = os.getcwd()
    logger.info(f"FOLDER: {configs['FOLDER']}")

    # TARGETS がない場合は警告
    if 'TARGETS' not in configs:
        logger.info("'TARGETS' not specified in the YAML file.")

    targets = configs['TARGETS']
    # TARGETS の拡張子を除去し、表示する
    logger.info("TARGETS files:")
    for key in targets:
        targets[key] = remove_extension(targets[key])
        logger.info(f"  {key}:{targets[key]}")

    # INDEXING がないか、あっても空の場合は警告
    if 'INDEXING' not in configs or not configs['INDEXING']:
        logger.info("'INDEXING' not specified or empty in the YAML file.")

    indexing = configs['INDEXING']
    # INDEXING の配列をすべて走査して TARGET と SOURCE を表示する
    logger.info("INDEXING files:")
    for idxfile in indexing:
        logger.debug(f"  INDEX   :{idxfile['INDEX']} SOURCE:{idxfile['SOURCE']}")
        # TARGET の書式が <A> ならば、対応する SOURCE に置換する
        # 書式が <英字の1文字以上> にマッチする正規表現
        pattern_apd = re.compile(r'<(\w+):(\S+)>')

        # idxfile['CSL'] がない場合は False に設定する
        if 'CSL' not in idxfile:
            idxfile['CSL'] = False
        # CSL はOFF/ON、Trur/False で指定されるので、それ以外の値はエラー
        if idxfile['CSL'] not in ['OFF', 'ON', True, False]:
            raise ConfigError(f"CSL value is not valid: {idxfile['CSL']}")

        idxfile['INDEX'] = modify_filenames(idxfile['INDEX'], targets)
        idxfile['SOURCE'] = modify_filenames(idxfile['SOURCE'], targets)
        logger.info(f"  INDEX:{idxfile['INDEX']} SOURCE:{idxfile['SOURCE']}")

#    configs['INDEXING'] = indexing 
    # genarating の配列をすべて走査して GENERATE, SOURCE, INDEX, CSL を表示する
    if 'GENERATING' not in configs or not configs['GENERATING']:
        logger.debug("'GENERATING' not specified or empty in the YAML file.")

    generating = configs['GENERATING']
    logger.debug("GENERATING files:")
    for genfile in generating:
        # genfile['CSL'] がない場合は False に設定する
        if 'CSL' not in genfile:
            genfile['CSL'] = False
        # CSL はOFF/ON、Trur/False で指定されるので、それ以外の値はエラー
        if genfile['CSL'] not in ['OFF', 'ON', True, False]:
            raise ConfigError(f"CSL value is not valid: {genfile['CSL']}")

        # genfile['CSP'] がない場合は False に設定する
        if 'CSP' not in genfile:
            genfile['CSP'] = False
        # CSP はOFF/ON、Trur/False で指定されるので、それ以外の値はエラー
        if genfile['CSP'] not in ['OFF', 'ON', True, False]:
            raise ConfigError(f"CSP value is not valid: {genfile['CSP']}")

        # GENERATE, SOURCE, INDEX のファイル名を置換する
        logger.debug(f"  GENERATE:{genfile['GENERATE']} SOURCE:{genfile['SOURCE']} INDEX:{genfile['INDEX']} CSL:{genfile['CSL']}")
        genfile['GENERATE'] = modify_filenames(genfile['GENERATE'], targets)
        genfile['SOURCE'] = modify_filenames(genfile['SOURCE'], targets)
        logger.debug(f"  GENERATE:{genfile['GENERATE']} SOURCE:{genfile['SOURCE']} CSL:{genfile['CSL']}")
        # genfile['INDEX'] は配列なので、すべての要素を置換する
        for i in range(len(genfile['INDEX'])):
            genfile['INDEX'][i] = modify_filenames(genfile['INDEX'][i], targets)
        logger.debug(f"  GENERATE INDEX files:")
        for i in range(len(genfile['INDEX'])):
            logger.debug(f"    {genfile['INDEX'][i]}")

    # ダンプファイルが指定されていたら、そのファイルにダンプする
    if dump:
        with open(dump, 'w', encoding='utf-8') as file:
            yaml.dump(configs, file,sort_keys=False)
    return configs

