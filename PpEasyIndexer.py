#!/usr/bin/env python   # PowerPoint Easy Numbering tool
# -*- coding: utf-8 -*- 

from pptx import Presentation
import os
import re
import argparse
import win32com.client

from PpIndexCommon import remove_slides, replace_CSLandCSP

#指定したPowerPointファイルを読み込み、章番号と節番号を付与して新しいファイルを生成します。
#章番号は #CHAPT# 、節番号は #SECTION# という文字列をスライドのタイトルに含むことで判別します。

matchchapt=r'#CHAPT#'
matchsect=r'#SECTION#'
chaptnumpat=r'#CHAPTNUM##SEPA##CONTENT#'
sectnumpat=r'#CHAPTNUM##DELM##SECTNUM##SEPA##CONTENT#'

separator=') '
delimitter='.'
deletecsl=False
deletecsp=False

postfix='_indexed'

# chaptnumber=0
# sectnumber=0


# コマンドライン引数を解析して返す
def parse_commandargs():
    # コマンドライン引数のパーサーを作成
    parser = argparse.ArgumentParser(description="A simple chapter and section numbering processor.")
    # 'pptx' 引数の定義　処理対象のpptxファイルを指定する
    parser.add_argument("pptx", help="PowerPoint file to process")
    parser.add_argument('--skipsildes', '-ss', type=int, default=0, help='Number of slides to skip at the beginning')
    parser.add_argument('--separator', '-sp', type=str, default=separator, help='Separator between chapter.section and title content')
    parser.add_argument('--delimitter', '-dl', type=str, default=delimitter, help='Delimitter between chapter and section numbers')
    parser.add_argument('--deletecsl', '-csl', action='store_true', help='Delete slides containing #CSL#')
    parser.add_argument('--deletecsp', '-csp', action='store_true', help='Delete shapes containing #CSP#')
    parser.add_argument('--postfix',  '-p', type=str, default=postfix, help='Postfix for the generated filename')

    # 引数を解析
    _args = parser.parse_args()
    return _args


def modify_slide_titles(file_path,sldskip=0):
    global matchchapt,matchsect,chaptnumpat,sectnumpat,separator,delimitter,deletecsl,deletecsp,postfix

    chaptnumber=0  # 章番号をインクリメント
    sectnumber=0    # 節番号をリセット（次スライドから1になるのでここでは0にする）

    prs = Presentation(file_path)  # PowerPointファイルを読み込む

    # #CSL# ページ削除オプションが有効な場合はここで行う
    if deletecsl:
        prs = Presentation(file_path)
        remove_slides(prs, deletecsl)  

    titles = []
    startchapter=False
    startsection=False
    insection=False
    slidenumber=0

    # 全スライドについてループ        
    for slide in prs.slides:
        slidenumber+=1
        if slidenumber <= sldskip:
            print(f"[{slidenumber}] スキップ")
            continue

        title = ''
        # タイトルプレースホルダーを探す
        titleshape = None

        if hasattr(slide.shapes.title, 'has_text_frame'):
            print(f"[{slidenumber}] 置換前タイトル:{slide.shapes.title.text}")
            title = slide.shapes.title.text
            titleshape = slide.shapes.title
        else:
            print(f"[{slidenumber}] 置換処理なし")



        # 章節番号処理のための shapes ループ（#CSP#と#CSL#は別ループなことに注意）
        # スライド中に matchchapt または matchsect があるかどうかを確認
        startchapter=False
        startsection=False
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    text = run.text
                    # matchchat または matchsect にマッチするtextを探す
                    if re.search(matchchapt, text):
                        startchapter=True
                        text = text.replace(matchchapt,'')
                        run.text = text
                    if re.search(matchsect, text):
                        startsection=True
                        text = text.replace(matchsect,'')
                        run.text = text
            # 両方が見つかったら終了。片方だけなら継続
            if startchapter and startsection:
                break
        # ここでは startchapter,startsection の判別がついている

        # 章・節番号の切り替わり処理.

        if startchapter and startsection:    
            # 両方発見されたので、章・節番号切り替わりの処理をする
            chaptnumber+=1  # 章番号をインクリメント
            sectnumber=0    # 節番号をリセット（このスライドから1にする）
            insection=True  # 次スライド以降はstartchapter,startsection未該当でも節が継続する
            newtitle = chaptnumpat
            newtitle = newtitle.replace('#CHAPTNUM#',str(chaptnumber))
            newtitle = newtitle.replace('#SEPA#',separator)
            newtitle = newtitle.replace('#CONTENT#',title)

        elif startchapter:    
            # matchchapt のみなので章番号切り替わりの処理のみ行う
            chaptnumber+=1  # 章番号をインクリメント
            sectnumber=0    # 節番号をリセット（次スライドから1になるのでここでは0にする）
            insection=True  # 次スライド以降はstartchapter,startsection未該当でも節が継続する
            newtitle = chaptnumpat
            newtitle = newtitle.replace('#CHAPTNUM#',str(chaptnumber))
            newtitle = newtitle.replace('#SEPA#',separator)
            newtitle = newtitle.replace('#CONTENT#',title)

        elif startsection or insection:    # 節番号切り替わりの処理のみ行う
            # matchsect のみなので節番号切り替わりの処理のみ行う
            if startsection:
                sectnumber=1    # 節番号をリセット（このスライドから1にする）
                insection=True  # 次スライド以降はstartchapter,startsection未該当でも節が継続する
            else:
                sectnumber+=1   # insection に該当するので節番号をインクリメント
            newtitle = sectnumpat
            newtitle = newtitle.replace('#CHAPTNUM#',str(chaptnumber))
            newtitle = newtitle.replace('#DELM#',delimitter)
            newtitle = newtitle.replace('#SECTNUM#',str(sectnumber))
            newtitle = newtitle.replace('#SEPA#',separator)
            newtitle = newtitle.replace('#CONTENT#',title)

        else: 
            # 最初の Chapter より前のスライドはどこにも該当しない
            newtitle = title
            pass

        if titleshape:
            titleshape.text = newtitle
        titles.append(newtitle)

        print(f"[{slidenumber}] 置換後タイトル:{newtitle}")


        # #CSP# と #CSL# の処理のための shapes ループ（章節番号処理とは別ループなことに注意）
        needshapedeletetion = False
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        tobreak, _needshapedeletetion = replace_CSLandCSP(run, deletecsl, deletecsp)
                        if _needshapedeletetion:
                            needshapedeletetion = True
                        if tobreak: # このシェイプの残りの run は処理不要
                            break

    # 新しいファイル名を生成
    base_name, ext = os.path.splitext(file_path)
    _generatepptxpath = f"{base_name}{postfix}{ext}"
    
    # 変更を保存
    prs.save(_generatepptxpath)


    # 絶対パスでないと win32 アプリでファイルを開けなかったので絶対パス取得
    newpptx_path_abs = os.path.abspath(_generatepptxpath)
    # print(f"BEFORE win32 _generatepptxpath:{_generatepptxpath}")
    # print(f"BEFORE win32 newpptx_path_abs :{newpptx_path_abs}")

    # シェイプ削除がある場合は pywin32 で削除処理を行う
    # (python-pptx でXMLエレメントをremoveする方法で削除しようとすると、保存後のpptxを開いたときに
    # 構成エラーが発生するので、pywin32 を使って削除します。ただしこの方法は遅い)
    if needshapedeletetion:
        win32pptapp = win32com.client.Dispatch("Powerpoint.Application")
        win32prs = win32pptapp.Presentations.Open(str(newpptx_path_abs), WithWindow=False)
        for slide in win32prs.Slides:
            # 逆順でシェイプをループ（削除中にコレクションを変更しないように）
            for shape in reversed(list(slide.Shapes)):
                if shape.HasTextFrame == -1 and shape.TextFrame.HasText:
                    # テキストの中に #CSP# または #CSL# の文字があれば削除
                    if "#CSP#" in shape.TextFrame.TextRange.Text or "#CSL#" in shape.TextFrame.TextRange.Text:
                        print(f"Found #CSP# in text box {shape.TextFrame.TextRange.Text}")
                        shape.Delete()

        print(f"saving as newpptx_path_abs:{newpptx_path_abs}")
        win32prs.SaveAs(newpptx_path_abs)

        # プレゼンテーションを閉じる
        win32prs.Close()

        # PowerPointを閉じる
        win32pptapp.Quit()


    print(f"変換結果: {newpptx_path_abs}")



arg = parse_commandargs()
separator=arg.separator
delimitter=arg.delimitter
deletecsl=arg.deletecsl
deletecsp=arg.deletecsp
postfix=arg.postfix

modify_slide_titles(arg.pptx,arg.skipsildes)

