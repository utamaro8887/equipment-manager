import os
import csv
import segno
import argparse

# デフォルトのGAS WebアプリURL (後で実際のURLに差し替えて運用します)
DEFAULT_BASE_URL = "https://script.google.com/macros/s/DUMMY_URL/exec?id="
OUTPUT_DIR = "qr_codes"

def generate_qr_codes(csv_filepath, base_url=DEFAULT_BASE_URL):
    """
    CSVファイルから管理IDを読み込み、QRコードを一括生成する
    CSVフォーマット想定: ヘッダーあり、1列目が「管理ID」
    """
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)
        print(f"ディレクトリを作成しました: {OUTPUT_DIR}")

    target_count = 0
    with open(csv_filepath, mode='r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        for row in reader:
            # 必須列「管理ID」を取得（列名は実際のCSVに合わせて調整可能）
            item_id = row.get("管理ID") or row.get("id") or row.get("ID")
            if not item_id:
                continue

            # QRコード化するURLを生成
            target_url = f"{base_url}{item_id}"
            
            # QRコードの生成 (Segnoを使用。マイクロQRなども対応可能ですが標準を使用)
            qr = segno.make(target_url)
            
            # 画像として保存 (PNG, 拡張性が高くラベル印刷にも適している)
            output_path = os.path.join(OUTPUT_DIR, f"{item_id}.png")
            qr.save(output_path, scale=5)
            print(f"[{item_id}] のQRコードを生成しました -> {output_path}")
            target_count += 1
            
    print(f"\n生成完了: 合計 {target_count} 件のQRコードを出力しました。")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="備品管理用QRコード一括生成ツール")
    parser.add_argument("csv_file", help="管理IDが記載されたCSVファイルのパス")
    parser.add_argument("--url", help="ベースとなるURL", default=DEFAULT_BASE_URL)
    
    args = parser.parse_args()
    
    # 実行前に要件であるSegnoライブラリのチェックを行う
    try:
        import segno
    except ImportError:
        print("エラー: segnoライブラリがインストールされていません。")
        print("実行してインストールしてください: pip install segno")
        exit(1)

    generate_qr_codes(args.csv_file, args.url)
