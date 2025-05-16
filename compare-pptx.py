import argparse
import os
from datetime import datetime
import imagehash
from itertools import combinations
import numpy as np
import sys

VERSION = "0.6.0"
programstr = f"PowerPoint比較解析ツール %(prog)s {VERSION}"

print(programstr)
if sys.argv[1] == "--version":
    sys.exit(0)

# 重いライブラリをインポートする前に、PowerPoint 起動の注意と操作禁止の確認を求めておく
user_input = input("PowerPointを2回起動して画像を出力します。その間、キー操作を行わないでください。\nInvoking PowerPoint app 2 times. Please do not perform any actions during this time. (y/n): ")
if user_input.lower() != 'y':
    print("処理を中止します。")
    sys.exit(1)

print("please wait for a while...")

# 操作禁止の確認が終わったところで、重いライブラリをインポート
from sklearn.metrics.pairwise import cosine_similarity
import comtypes.client
from PIL import Image, ImageChops
from sentence_transformers import SentenceTransformer

#-------------------------------------------------------------------------
# 動作パラメータ定数
#-------------------------------------------------------------------------
defaultsourcedir = "."  # ソースpptxの探索基点ディレクトリ
defaultexportroot = "./export/analyzed#DT#"  # 出力ルートディレクトリ, #DT# は日時に変換
defaultderiveddir = "derived"  # 新ファイル解析結果ディレクトリ
defaultbasedir = "base"  # 旧ファイル解析結果ディレクトリ
defaultdiffdir = "diff"  # 旧ファイル解析結果ディレクトリ
defaultderivedexportname = "derived"  # 新ファイルexport画像のファイル名（番号、拡張子なし）
defaultbaseexportname = "base"  # 旧ファイルexport画像のファイル名（番号、拡張子なし）
defaultmatch = 0  # 完全一致とみなす閾値
defaulthigh = 4  # 類似とみなす閾値
defaultlow = 10  # 類似かもしれない閾値
defaulttextmatch = 0.95  # 完全一致とみなす閾値
defaulttexthigh = 0.90  # 類似とみなす閾値
defaulttextlow = 0.80  # 類似かもしれない閾値
defaultoutput = "analyzed"  # 出力ファイル名（拡張子なし）


#-------------------------------------------------------------------------
# グローバル変数
#-------------------------------------------------------------------------

#テキスト解析モデル
model = SentenceTransformer('all-MiniLM-L6-v2') 


#-------------------------------------------------------------------------
# 例外の定義
#-------------------------------------------------------------------------
# 設定ファイル読込失敗、フォーマット不正、指定されたファイルが存在しない、などのエラー
class ConfigError(Exception):
    def __init__(self, message):
        self.message = message
        super().__init__(message)

    def __str__(self):
        return f"ConfigError: {self.message}"

# 対象ファイル読み込みエラー、書き込みエラーなど
class ProcessError(Exception):
    def __init__(self, message):
        self.message = message
        super().__init__(message)

    def __str__(self):
        return f"ProcessError: {self.message}"


#-------------------------------------------------------------------------
# 引数のパース
#-------------------------------------------------------------------------
def parse_args():
    parser = argparse.ArgumentParser(description='PowerPoint比較解析ツール')

    # 必須引数（位置引数）
    parser.add_argument("derivedfile", help="比較対象のpptxファイル（新しい方）")
    parser.add_argument("basefile", help="比較元のpptxファイル（元ファイル）")

    # オプション引数
    parser.add_argument("--sourcedir", type=str, default=f"{defaultsourcedir}", help="ファイル探索の基点ディレクトリ")
    parser.add_argument("--exportroot", type=str, default=f"{defaultexportroot}", help="出力ルートディレクトリ")
    parser.add_argument("--deriveddir", type=str, default=f"{defaultderiveddir}", help="derivedfileの解析結果ディレクトリ")
    parser.add_argument("--basedir", type=str, default=f"{defaultbasedir}", help="basefileの解析結果ディレクトリ")
    parser.add_argument("--diffdir", type=str, default=f"{defaultdiffdir}", help="差分画像の保存ディレクトリ")
    parser.add_argument("--derivedexportname", type=str, default=f"{defaultderivedexportname}", help="derived export画像のファイル名")
    parser.add_argument("--baseexportname", type=str, default=f"{defaultbaseexportname}", help="base export画像の保存ディレクトリ")

    parser.add_argument("--match", type=int, default=f"{defaultmatch}", help="完全一致とみなす閾値（デフォルト: 0）")
    parser.add_argument("--high", type=int, default=f"{defaulthigh}", help="類似とみなす閾値（デフォルト: 10）")
    parser.add_argument("--low", type=int, default=f"{defaultlow}", help="類似かもしれない閾値（デフォルト: 20）")
    parser.add_argument("--textmatch", type=float, default=f"{defaulttextmatch}", help="テキストを完全一致とみなす閾値（デフォルト: 0.95）")
    parser.add_argument("--texthigh", type=float, default=f"{defaulttexthigh}", help="テキストを類似とみなす閾値（デフォルト: 0.90）")
    parser.add_argument("--textlow", type=float, default=f"{defaulttextlow}", help="テキストを類似かもしれないとみなす閾値（デフォルト: 0.80）")
    parser.add_argument("--output", type=str, default=f"{defaultoutput}", help="解析結果のファイル名（拡張子なし）")

    args = parser.parse_args()

    # args.exportroot に #DT# が含まれている場合は、現在の日時に置き換える
    if "#DT#" in args.exportroot:
        now = datetime.now()
        args.exportroot = args.exportroot.replace("#DT#", now.strftime("%Y%m%d%H%M%S"))

    # ==== 入力確認 ====
    print(f"新pptxファイル       : {args.derivedfile}")
    print(f"旧pptxファイル       : {args.basefile}")
    print(f"pptx探索パス         : {args.sourcedir}")
    print(f"解析結果出力ルート   : {args.exportroot}")
    print(f"新ファイル解析出力先 : {args.deriveddir}")
    print(f"旧ファイル解析出力先 : {args.basedir}")
    print(f"完全一致閾値(画像)      : {args.match}")
    print(f"高い類似閾値(画像)      : {args.high}")
    print(f"低い類似閾値(画像)      : {args.low}")
    print(f"完全一致閾値(テキスト)   : {args.textmatch}")
    print(f"高い類似閾値(テキスト)   : {args.texthigh}")
    print(f"低い類似閾値(テキスト)   : {args.textlow}")
    print(f"比較結果出力ファイル名   : {args.output}.json")

    return args


#-------------------------------------------------------------------------
# 解析結果ディレクトリを作る
# argsの deriveddir, basedir を絶対パスに変換して返す
#-------------------------------------------------------------------------
def create_directory(args):
    # 出力ルートディレクトリの作成
    if not os.path.exists(args.exportroot):
        os.makedirs(args.exportroot)

    # deriveddirとbasedirのディレクトリを作成
    derived_dir = os.path.abspath(os.path.join(args.exportroot, args.deriveddir))
    base_dir = os.path.abspath(os.path.join(args.exportroot, args.basedir))
    diff_dir = os.path.abspath(os.path.join(args.exportroot, args.diffdir))

    if not os.path.exists(derived_dir):
        os.makedirs(derived_dir)
    if not os.path.exists(base_dir):
        os.makedirs(base_dir)
    if not os.path.exists(diff_dir):
        os.makedirs(diff_dir)

    print(f"解析結果出力ルートディレクトリ: {args.exportroot}")
    print(f"新ファイル解析結果ディレクトリ: {derived_dir}")
    print(f"旧ファイル解析結果ディレクトリ: {base_dir}")
    print(f"差分画像保存ディレクトリ      : {diff_dir}")

    args.deriveddir = derived_dir   #絶対パスにして返す
    args.basedir = base_dir         #絶対パスにして返す
    args.diffdir = diff_dir         #絶対パスにして返す

    return args



#-------------------------------------------------------------------------
# 出力ディレクトリにpptxファイルのスライド画像をexportし、hash値を計算して保存する
#-------------------------------------------------------------------------
def export_pptx_images(pptxpath, exportdir, exportfilename):
    print(f"PowerPointファイルを開きます: {pptxpath}")
    print(f"出力先ディレクトリ: {exportdir}")
    # pptxpath をディレクトリとファイル名に分離
    pptdir, pptxfile = os.path.split(pptxpath)
    #pptdir を絶対パスに変換
    pptdir = os.path.abspath(pptdir)

    analyzed = {
        "sourcedir": pptdir,
        "pptxfile": pptxfile,
        "exportdir": exportdir,
        "slides": []
    }

    ppt = comtypes.client.CreateObject("PowerPoint.Application")
    ppt.Visible = True

    presentation = ppt.Presentations.Open(pptxpath)

    # スライドの数を取得
    slide_count = len(presentation.Slides)

    for i in range(slide_count):
        slide = presentation.Slides[i + 1]  # スライドは1から始まるので、i+1で取得

        slide_text = ""
        # スライド内のすべてのシェイプのテキストを結合してベクター化
        for shape in slide.Shapes:
            if shape.HasTextFrame:  # テキストフレームがある場合
                text_frame = shape.TextFrame
                if text_frame.HasText:  # テキストが存在する場合
                    slide_text += text_frame.TextRange.Text + " "

        textvector = model.encode(slide_text)

        # 各スライドを PNG で出力
        imagefile = f"{exportfilename}_{i}.png"
        imaagepath = os.path.join(exportdir, imagefile)
        slide.Export(imaagepath, "PNG")
        with Image.open(imaagepath) as img:
            hash = imagehash.phash(img)
        analyzed["slides"].append({
            "slideimage": imagefile,
            "imagehash": hash,
            "textvector": textvector,
        })


    presentation.Close()
    #ppt.Quit()

    return analyzed


def output_html(derived_analyzed, args):
    # 入力ファイル
    output_htmlfile = "comparison_report.html"
    output_htmlpath = os.path.join(args.exportroot, output_htmlfile)
    derived_imagedir = os.path.join(args.exportroot,args.deriveddir)
    base_imagedir = os.path.join(args.exportroot,args.basedir)

    slides = derived_analyzed["slides"]

    html = [
        "<!DOCTYPE html>",
        "<html lang='ja'>",
        "<head>",
        "<meta charset='UTF-8'>",
        "<title>スライド比較レポート</title>",
        "<style>",
        "body { font-family: sans-serif; }",
        "table { border-collapse: collapse; width: 100%; margin-bottom: 40px; }",
        "th, td { border: 1px solid #ccc; padding: 10px; text-align: center; vertical-align: top; }",
        "th { background-color: #f0f0f0; }",
        "img.thumb { width: 240px; height: auto; cursor: zoom-in; border: 2px solid #aaa; }",
        "img.thumb:hover { border-color: #2196f3; }",
        "details summary { cursor: pointer; font-weight: bold; margin: 10px 0; }",
        "</style>",
        "</head>",
        "<body>",
        f"<h1>📊 スライド比較レポート：{os.path.basename(derived_analyzed['pptxfile'])}</h1>"
    ]

    for di,slide in enumerate(slides):
        derived_image = slide["slideimage"]
        derived_path = os.path.join(args.deriveddir, derived_image).replace("\\", "/")

        graded = {"match": [], "high": [], "low": []}
        for sim in slide.get("similars", []):
            graded[sim["grade"]].append(sim)

        html.append(f"<details open><summary>🖼️ {derived_image}</summary>")
        html.append("<table>")
        html.append("<tr><th>Original</th><th>Match</th><th>High</th><th>Low</th></tr>")
        html.append("<tr>")

        # Original cell
        html.append(f"<td><a href='{derived_path}' target='_blank'><img src='{derived_path}' class='thumb'></a><br>{derived_image}</td>")

        # Grade cells
        for grade in ["match", "high", "low"]:
            cell = ""
            for sim in graded[grade]:
                sim_path = os.path.join(args.basedir, sim["slideimage"]).replace("\\", "/")
                label = f"NewSlide: {sim['slideimage']}<br>ImageScore: {sim['imagescore']} pt<br>TextScore: {sim['textscore']} pt<br>OldPptx: {sim['pptxfile']}"
                cell += f"<a href='{sim_path}' target='_blank'><img src='{sim_path}' class='thumb'></a><br>{label}<br><br>"
                diff_path = os.path.join(args.diffdir, f"diff_{di}_{sim['slideindex']}.png")
                difflabel = f"ImageDifference"
                cell += f"<a href='{diff_path}' target='_blank'><img src='{diff_path}' class='thumb'></a><br>{difflabel}<br><br>"

            html.append(f"<td>{cell if cell else '-'}</td>")

        html.append("</tr></table>")
        html.append("</details>")

    html.append("</body></html>")

    with open(output_htmlpath, "w", encoding="utf-8") as f:
        f.write("\n".join(html))

    print(f"✅ 比較結果出力完了(HTML): {output_htmlpath}")



def main():
    args = parse_args()

    # ディレクトリの作成
    print("出力ディレクトリを作成します")
    args = create_directory(args)

    # ファイルの絶対パスを取得 args.sourcepath + args.basefile
    basepptxpath = os.path.abspath(os.path.join(args.sourcedir, args.basefile))
    derivedpptxpath = os.path.abspath(os.path.join(args.sourcedir, args.derivedfile))
    print(f"新ファイルの絶対パス: {derivedpptxpath}")
    print(f"旧ファイルの絶対パス: {basepptxpath}")

    derived_analyzed = export_pptx_images(derivedpptxpath, args.deriveddir, args.derivedexportname)
    base_analyzed = export_pptx_images(basepptxpath, args.basedir, args.baseexportname)

    # derived_analyzed["slides"] と base_analyzed["slides"] のハッシュ値を比較する
    for di , derived_slide in enumerate(derived_analyzed["slides"]):
        derived_slide["similars"] = []
        # スライド画像のファイル名を取得
        derived_hash = derived_slide["imagehash"]
        derived_vector = derived_slide["textvector"]
        # print(f"derived textvector {derived_vector}")
        for bi, base_slide in enumerate(base_analyzed["slides"]):
            base_hash = base_slide["imagehash"]
            base_vector = base_slide["textvector"]
            # print(f"base textvector {base_vector}")

            similarity = cosine_similarity(derived_vector.reshape(1,-1), base_vector.reshape(1,-1))
            #print( f"similarity {similarity}")
            vector_similarity = similarity[0][0]

            hash_diff = abs(derived_hash - base_hash)

            grade = "different"

            # ハッシュ値を比較
            if hash_diff <= args.match or vector_similarity >= args.textmatch:
                grade = "match"
                print(f"(完全)一致: derived:{di} base:{bi}")
            elif hash_diff <= args.high or vector_similarity >= args.texthigh:
                grade = "high"
                print(f"高い類似性: derived:{di} base:{bi}")
            elif hash_diff <= args.low or vector_similarity >= args.textlow:
                grade = "low"
                print(f"低い類似性: derived:{di} base:{bi}")
            else:
                grade = "different"
                #print(f"相違: derived:{di} base:{bi}")

            if grade != "different":
                derived_slide["similars"].append({
                    "slideimage": base_slide["slideimage"],
                    "grade": grade,
                    "imagescore": hash_diff,
                    "textscore":  format(vector_similarity, '.2f'),
                    "pptxfile": base_analyzed["pptxfile"],
                    "slideindex": bi,
                })
                # ここで差分画像を作る
                # 元画像は derived_analyzed["slides"][di]["slideimage"]
                # 旧画像は base_analyzed["slides"][bi]["slideimage"]
                # 画像のパスを取得
                derived_image_path = os.path.join(args.deriveddir, derived_slide["slideimage"])
                base_image_path = os.path.join(args.basedir, base_slide["slideimage"])

                img1 = Image.open(derived_image_path).convert('RGB')
                img2 = Image.open(base_image_path).convert('RGB')

                # 差分画像を作る
                diff = ImageChops.difference(img1, img2)
                # diff_di_bi.png というファイル名で差分画像を保存する
                diff_filename = f"diff_{di}_{bi}.png"
                diff_path = os.path.join(args.diffdir, diff_filename)

                diff.save(diff_path)
                #diff.show()                

    # derived_analyzed の imagehash を文字列に変換
    for slide in derived_analyzed["slides"]:
        slide["imagehash"] = str(slide["imagehash"])
        slide["textvector"] = np.array_str(slide["textvector"])

    # base_analyzed の imagehash を文字列に変換
    for slide in base_analyzed["slides"]:
        slide["imagehash"] = str(slide["imagehash"])
        slide["textvector"] = np.array_str(slide["textvector"])

    # ★ ★ ★ textvector もここで文字列に変換する方がいいかもしれない ★ ★ ★ 

    # derived_analyzed を JSON形式で保存
    jsonfile = os.path.join(args.exportroot, "derived_" + args.output + ".json")
    with open(jsonfile, "w", encoding="utf-8") as f:
        import json
        json.dump(derived_analyzed, f, ensure_ascii=False, indent=4)

    # base_analyzed を JSON形式で保存
    jsonfile = os.path.join(args.exportroot, "base_" + args.output + ".json")
    with open(jsonfile, "w", encoding="utf-8") as f:
        import json
        json.dump(base_analyzed, f, ensure_ascii=False, indent=4)

    # HTML出力
    output_html(derived_analyzed, args)

if __name__ == "__main__":
    main()
