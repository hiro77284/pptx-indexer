# コンフィグファイルの読み込みと検証のテスト

import pytest
import PpIndexConfig
import yaml
import logging

logger = logging.getLogger('testlogger')
logger.setLevel(logging.DEBUG)
filehandler = logging.FileHandler('test/test.log')
logger.addHandler(filehandler)

remove_extension_parameters = [
    ('PowerPointFile.pptx',         # standard target
     'PowerPointFile'),
    ('日本語のファイル名.pptx',         # Japanese characters
     '日本語のファイル名'),
    (r'D:\somefolder\PowerPointFile.pptx',  # with a folder and a drive letter, Windows style
     r'D:\somefolder\PowerPointFile'),
    ('/somefolder/PowerPointFile.pptx',  # with a folder and a drive letter, Unix style
     '/somefolder/PowerPointFile'),
    ('PowerPointFile.sample.pptx',  # two periods
     'PowerPointFile.sample'),      # delete just one extension 
    ('PowerPointFile.ext',          # extensions other than pptx
     'PowerPointFile'),             # also deleted
    ('PowerPointFile',              # no extension
     'PowerPointFile'),             # no change
    ('invalidcharacters +!<>file.ext',              # invalid characters
     'invalidcharacters +!<>file'),                 # don't care whether or not the characters' validity, just remove the extension
]



# 拡張子を除去する
def test_remove_extension():
    for target, expected in remove_extension_parameters:
        assert PpIndexConfig.remove_extension(target) == expected

add_extension_parameters = [
    ('PowerPointFile', 'pptx', 'PowerPointFile.pptx'),
    ('日本語のファイル名', 'pptx', '日本語のファイル名.pptx'),
]

def test_add_extension():
    for target, ext, expected in add_extension_parameters:
        assert PpIndexConfig.add_extension(target, ext) == expected

simple_config={'VERSION': 1.0, 
               'FOLDER': 'test/testdata', 
               'INDEXING': [{'INDEX': 'index1.json', 
                             'SOURCE': 'source1.pptx'}, 
                            {'INDEX': 'index2.json', 
                             'SOURCE': 'source2.pptx'}]}

testdatafolder = 'test/testdata'


# yaml ファイルの単純な読込み。load_yamlで読み込んだデータとの照合用
def justreadyaml(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        _data = yaml.safe_load(file)
    return _data


def test_load_yaml():
    # 存在しない設定ファイル
    with pytest.raises(PpIndexConfig.ConfigError):
        PpIndexConfig.load_yaml(f'{testdatafolder}/loadyaml/notexist.yaml')
    # yamlフォーマットに準拠していない
    with pytest.raises(PpIndexConfig.ConfigError):
        PpIndexConfig.load_yaml(f'{testdatafolder}/loadyaml/yamlformaterror.yaml')
    # ↓これは一旦読んだデータをそのままダンプしたものと比較しているだけなので、テスト不要な気がする
    assert PpIndexConfig.load_yaml(f'{testdatafolder}/loadyaml/simpleconfig.yaml') == justreadyaml(f'{testdatafolder}/loadyaml/simpleconfig_dump.yaml')
    

def test_verify_parameter_formats():
    # TARGETS に存在しないキーがINDEXINGで使われている
    configs = PpIndexConfig.load_yaml(f'{testdatafolder}/vpf/simpleconfig_targetunmatch.yaml')
    with pytest.raises(PpIndexConfig.ConfigError):
        PpIndexConfig.verify_parameter_formats(configs) 

    # SUFFIX の書式がおかしい
    configs = PpIndexConfig.load_yaml(f'{testdatafolder}/vpf/simpleconfig_suffixformaterror.yaml')
    with pytest.raises(PpIndexConfig.ConfigError):
        PpIndexConfig.verify_parameter_formats(configs, logger=logger) 

    # 複数ファイルの記述を正しく変換する

    vpfparams = justreadyaml(f'{testdatafolder}/vpf_parameters.yaml')
    for i in range(len(vpfparams["DATA"])):
        configs = PpIndexConfig.load_yaml(f'{testdatafolder}/{vpfparams["DATA"][i]["SOURCE"]}')
        print('tesging: ', f'{testdatafolder}/{vpfparams["DATA"][i]["PURPOSE"]}')
        assert PpIndexConfig.verify_parameter_formats(configs, logger=logger) == justreadyaml(f'{testdatafolder}/{vpfparams["DATA"][i]["VERIFIER"]}')

  
