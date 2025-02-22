from pptx import Presentation
import logging
import re

# PowerPoint(.pptx) ファイル処理の共通関数

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


def replace_CSLandCSP(run, deletecsl, deletecsp):
    # run.text に #CSP# が含まれていたら shape を削除 または #CSP# の文字列のみを削除
    tobreak = False
    needshapedeletetion = False
    if  re.search(r'#CSP#', run.text):
        if deletecsp:
            needshapedeletetion = True
            logger.debug(f"need shape deletion:{run.text}")
            # このシェイプは後にまるごと削除されるので、
            # 呼び出し元に戻った後の replace_text が実行不要、
            # run の残りの処理も不要なので tobreak=True にしておく
            tobreak = True
        else:
            # 正規表現を使って '#CSP#\s?' を削除
            replacedtext = re.sub(r'#CSP#\s?', '', run.text)
            run.text = replacedtext
            #continue
    
    # run.text に #CSL# が含まれていて、genparams['CSL'] がFalseの場合は #CSL# とそれに続く文字列のみを削除
    if  re.search(r'#CSL#', run.text):
        if deletecsl:
            # ページ削除は既に済んでいるのでここでは特にやることはない
            # というか実際にはここには来ないので下の tobreak=True は dead code だが、
            # 呼び出し元に戻った後の replace_text が実行不要、
            # run の残りの処理も不要なことを明示するために残しておく
            tobreak = True
        else:
            # ページ削除をしない場合は #CSL# とそれに続く文字列のみを削除
            # 正規表現を使って '#CSL#\s?.*' を削除
            replacedtext = re.sub(r'#CSL#\s?.*', '', run.text)
            run.text = replacedtext
            #continue

    return tobreak, needshapedeletetion