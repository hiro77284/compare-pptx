from difflib import SequenceMatcher
from math import sqrt
from itertools import combinations

#前提：各スライドは以下の形式で格納されていると仮定
# 座標値はスライドのサイズに対する比率で格納されている
slidedataformat = """
slides = [
    {
        "slidetitle": "タイトル1",
        "shapes": [ 
            { "type": "text", 
            "text": ...,
            "left": ..., 
            "top": ..., 
            "width": ..., 
            "height": ...
            }, ... 
        ]
    },
    ...
]
"""

#--------------------------------------------
# スライドの類似度を計算するための設定
#--------------------------------------------

# レイアウトとテキストの類似度を計算するための重み
# 合計で1.0 になるように設定
layout_weight = 0.2
text_weight = 0.8

# スライドの類似度の閾値
slide_threshold = 0.8  

# シェイプの類似度の閾値
shape_threshold = 0.75  

# テキストの完全一致を要求する場合は True にする
text_strictmatch = False    

"""
類似度計算の設定を変更する関数
:param layout_weight: レイアウトの重み
:param text_weight: テキストの重み
:param slide_threshold: スライドの類似度の閾値
:param shape_threshold: シェイプの類似度の閾値
:param text_completematch: テキストの完全一致を要求する場合は True にする
"""
def set_similarity_settings(layout=0.2, text=0.8, slide=0.8, shape=0.75, text_strict=False):
    global layout_weight, text_weight, slide_threshold, shape_threshold, text_strictmatch
    layout_weight = layout
    text_weight = text
    slide_threshold = slide
    shape_threshold = shape
    text_strictmatch = text_strict


# レイアウトの類似性を計算
# zip はペア反復子。v1,v2の各要素をペアにして a,b に渡し、その差の2乗和の平方根を取る
def layout_similarity(s1, s2):
    v1 = [s1["left"], s1["top"], s1["width"], s1["height"]]
    v2 = [s2["left"], s2["top"], s2["width"], s2["height"]]
    dist = sqrt(sum((a - b)**2 for a, b in zip(v1, v2)))
    return 1 - min(dist, 1.0)  # normalize to [0,1]

# テキストの類似性を比較、completematch=True の場合は完全一致を要求
def text_similarity(t1, t2, strictmatch=False):
    if strictmatch:
        return 1.0 if t1 == t2 else 0.0
    return SequenceMatcher(None, t1, t2).ratio()

# シェイプのレイアウトとテキストの類似度を計算して返す
# alpha はレイアウトの重み、beta はテキストの重み
def shape_similarity(shape1, shape2, alpha=0.7, beta=0.3):
    if shape1["type"] != shape2["type"]:
        return 0.0
    layout_sim = layout_similarity(shape1, shape2)

    text1 = shape1.get("text", "")
    text2 = shape2.get("text", "")
    text_sim = text_similarity(text1, text2, strictmatch=text_strictmatch)
    # text1 に「具体的なものでは」が含まれている場合、print する
    # if "具体的なものでは" in text1 and "具体的なものでは" in text2:
    #     print(f"text_sim: {text_sim}")
    #     print(f"text1: {text1}")
    #     print(f"text2: {text2}")
    return alpha * layout_sim + beta * text_sim

# slide1 と slide2 の類似度を計算し、shape_threshold以上でマッチするshapeの比率をカウントする
# 戻り値は shape_threshold 以上 1.0 以下の値、有効数字2桁
def slide_similarity(slide1, slide2, shape_threshold=0.75):
    shapes1 = slide1["shapes"]
    shapes2 = slide2["shapes"]
    if not shapes1 or not shapes2:
        return 0.0

    matched = 0
    total = max(len(shapes1), len(shapes2))

    for s1 in shapes1:
        best = max((shape_similarity(s1, s2, alpha=layout_weight, beta=text_weight) for s2 in shapes2), default=0.0)
        if best >= shape_threshold:
            matched += 1

    return round(matched / total,2)


# slides のリストを総当たりで比較し類似度がthreshold以上のペアのリストを返す
def find_similar_slide_pairs(slides, slide_threshold=slide_threshold, shape_threshold=shape_threshold):
    similar_pairs = []
    for i, j in combinations(range(len(slides)), 2):
        score = slide_similarity(slides[i], slides[j], shape_threshold)
        if score >= slide_threshold:
            similar_pairs.append((i, j, score))
    return similar_pairs
