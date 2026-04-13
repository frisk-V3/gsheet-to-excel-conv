import os
import sys
import codecs
from datetime import datetime

def fix_csv_for_excel(file_path):
    """UTF-8のCSVにBOMを付与してExcelでの文字化けを防ぐ"""
    try:
        if not file_path.lower().endswith('.csv'):
            return f"[SKIP] CSVファイルではありません: {os.path.basename(file_path)}"

        # 1. 元のデータをバイナリで読み込む
        with open(file_path, 'rb') as f:
            content = f.read()

        # 2. すでにBOM(EF BB BF)があるかチェック
        BOM = codecs.BOM_UTF8
        if content.startswith(BOM):
            return f"[SKIP] すでに修正済みです: {os.path.basename(file_path)}"

        # 3. 出力ファイル名の生成 (fixed_ 元ファイル名)
        dir_name = os.path.dirname(file_path)
        base_name = os.path.basename(file_path)
        output_path = os.path.join(dir_name, f"fixed_{base_name}")

        # 4. BOMを付与して書き出し
        with open(output_path, 'wb') as f:
            f.write(BOM)
            f.write(content)

        return f"[SUCCESS] 修正完了: {output_path}"

    except Exception as e:
        return f"[ERROR] 失敗しました ({os.path.basename(file_path)}): {str(e)}"

if __name__ == "__main__":
    print("="*50)
    print(" Excel CSV 文字化け修正ツール (BOM Adder)")
    print("="*50)

    # ドラッグ＆ドロップされたファイルを取得
    targets = sys.argv[1:]

    if not targets:
        print("\n[!] 使い方: CSVファイルをこのEXEにドラッグ＆ドロップしてください。")
    else:
        print(f"\n{len(targets)} 個のファイルを処理中...\n")
        for t in targets:
            result = fix_csv_for_excel(t)
            print(result)

    print("\n" + "="*50)
    print(f"完了時刻: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("何かキーを押すと終了します...")
    input()
