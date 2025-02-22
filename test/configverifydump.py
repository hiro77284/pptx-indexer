import yaml
import logging

import sys
import os

# 親ディレクトリへのパスをsys.pathに追加
current_dir = os.path.dirname(__file__)
print('current_dir:', current_dir)
parent_dir = os.path.dirname(current_dir)
print('parent_dir:', parent_dir)
sys.path.append(parent_dir)

import PpIndexConfig

logfile='configverifydump.log'
testdatafolder = 'testdata'

logger = logging.getLogger( '__name__' )
logger.setLevel( logging.DEBUG )
filehandler = logging.FileHandler( logfile )
logger.addHandler(filehandler)


configs = PpIndexConfig.load_yaml(f'{testdatafolder}/multifileconfig.yaml')
PpIndexConfig.verify_parameter_formats(configs, dump=f'{testdatafolder}/multifileconfig_verifieddump.yaml', logger=logger) 

