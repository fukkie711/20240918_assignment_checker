import sys
import os
import difflib
import openpyxl
from openpyxl.styles import PatternFill

def check_assignments(answer_file, submission_dir):
    # 答えのファイルを読み込む
    with open(answer_file, 'r', encoding='utf-8') as f:
        answer_content = f.read().replace("\t", "").replace(" ", "").splitlines()

    # Excelワークブックを作成
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "チェック結果"
    ws.append(["学籍番号", "氏名", "類似度", "判定"])

    # 提出ディレクトリ内のファイルをチェック
    for filename in os.listdir(submission_dir):
        if filename.endswith('.py'):
            filepath = os.path.join(submission_dir, filename)
            with open(filepath, 'r', encoding='utf-8') as f:
                submission_content = f.read().replace("\t", "").replace(" ", "").splitlines()

            # difflibを使用して内容を比較
            # matcher = difflib.SequenceMatcher(None, answer_content, submission_content)
            matcher = difflib.SequenceMatcher(None, a=answer_content, b=submission_content)
            similarity = matcher.ratio()

            # 判定
            if similarity == 1.0:
                result = "〇"
            elif similarity == 0:
                result = "×"
            else:
                result = "△"

            # 学籍番号を取得（ファイル名から）
            student_id = filename.split(' ')[0]
            student_name = filename.split('_')[0]

            # 結果をExcelに追加
            # ws.append([student_id, filename, result])
            ws.append([student_id, student_name, similarity, result])

            # コンソールに出力
            # print(f"学籍番号: {student_id}, ファイル名: {filename}, , 類似度：{similarity}, 判定: {result}")
            print(f"学籍番号: {student_id}, 類似度：{similarity}, 判定: {result}")

    # Excelファイルを保存
    excel_filename = str(answer_file) + "チェック結果.xlsx"
    wb.save(excel_filename)
    print(f"チェック結果を{excel_filename}に保存しました。")

def main():
    if len(sys.argv) != 3:
        print("エラーが発生しました。引数を確認してください：[ファイル名, ディレクトリ名]")
        sys.exit(1)

    answer_file = sys.argv[1]
    submission_dir = sys.argv[2]

    if not os.path.isfile(answer_file) or not os.path.isdir(submission_dir):
        print("エラーが発生しました。引数を確認してください：[ファイル名, ディレクトリ名]")
        sys.exit(1)

    check_assignments(answer_file, submission_dir)

if __name__ == "__main__":
    main()