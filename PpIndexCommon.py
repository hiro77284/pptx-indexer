# PowerPoint(.pptx) ファイル処理の共通関数

from pptx import Presentation
import logging
import re

# ロガーの作成
logger = logging.getLogger(__name__)


def remove_slides(prs, deletion):
    snummap = {}    

    _snums_to_remove = ()   # 削除するスライドID
    _newsnum=0                 # スライド番号(削除後)
    _sourcesnum=0           # スライド番号(削除前)
    for slide in prs.slides:
        _sdelflag=False     # このスライドは削除するよフラグ
        snummap[_sourcesnum] = _newsnum
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        # run.text に #CSL# が含まれていたら、そのslideをまるごと削除する
                        if deletion :
                            if re.search(r'#CSL#', run.text):
                                # 削除するスライド番号を _sid_to_remove に追加
                                _snums_to_remove += (_sourcesnum,)
                                logger.debug(f"deleting slide:{_sourcesnum}")
                                _sdelflag=True
                                break
        if not _sdelflag:
            _newsnum += 1
        _sourcesnum += 1

    logger.debug(f"snummap:{snummap}")

    logger.debug(f"_snums_to_remove:{_snums_to_remove}")
    xml_slides = prs.slides._sldIdLst  # スライドのXMLリストへのアクセス
    logger.debug(f"xml_slides:{xml_slides}")
    # _sid_to_remove に含まれるスライドIDを逆順に削除する。まあ逆順じゃなくてもかまわんみたいだが
    for _sn in reversed(_snums_to_remove):
          # 削除したいスライドID
        logger.debug(f"removing slide:{xml_slides[_sn]}")
        xml_slides.remove(xml_slides[_sn])

    return snummap

