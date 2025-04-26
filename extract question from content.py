import pandas as pd

# 1. excelを識別
file_path = 'C:/Users/XXX.xlsx'
data = pd.read_excel(file_path)

# 2. キーワード定義
keywords = [
    "でしょうか", "ですか", "でしょう", "しますか", "知りたい", "しりたい", "聞きたい", "ききたい",
    "どうしたら", "相談したい", "教えて", "おしえて", "わかりません", "分かりません", "わからない",
    "分からない", "不安", "心配", "怖い", "アドバイス", "意見"
]
endings = ["か", "か。", "？"]

# 3. 条件に合う文を抽出する関数を定義する

def extract_sentences(cell):
    """
    セル内の文から、以下の条件に合うものを抽出する：
    - 指定したキーワードを含む文
    - 指定した語尾（"か", "か。", "？"）で終わる文
    """
    if pd.isna(cell):  # セルが空かどうかを確認
        return ""
    sentences = str(cell).split("。")  # 句点で文を分割
    extracted = [
        sentence + "。" for sentence in sentences
        if any(keyword in sentence for keyword in keywords) or
        any(sentence.endswith(ending) for ending in endings)
    ]
    return " ".join(extracted)  # 抽出した文を結合して返す

# 4. A列に関数を適用し、結果をB列に保存する
data["B列"] = data["A列"].apply(extract_sentences)

# 5. 結果を新しいファイルに保存する
output_path = 'C:/Users/XXXXXX.xlsx'  # 替换为保存路径
data.to_excel(output_path, index=False)

