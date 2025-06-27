import os
import re
import ast
import difflib
import argparse
import openpyxl
import subprocess

# --- 設定項目 ---
# 実行タイムアウト（秒）
EXECUTION_TIMEOUT = 5
# 評価の閾値
GRADE_THRESHOLDS = {
    'A': 0.95,
    'B': 0.80,
    'C': 0.50,
}
# チェック対象の拡張子
TARGET_EXTENSION = '.py'
# ファイル名から学籍番号と氏名を抽出するための正規表現
FILENAME_PATTERN = re.compile(r'([a-zA-Z0-9]+)[_ ]?([\w-]+)?')

def normalize_code_by_ast(source_code: str) -> str:
    """
    ソースコードをAST(抽象構文木)に変換し、正規化された文字列表現を返す。
    """
    try:
        tree = ast.parse(source_code)
        return ast.dump(tree, annotate_fields=False, include_attributes=False)
    except SyntaxError:
        return ""

def check_execution(filepath: str) -> str:
    """
    Pythonスクリプトを実行し、その結果（成功、エラー、タイムアウト）を返す。
    """
    try:
        # タイムアウトとエラー出力をキャプチャしてサブプロセスで実行
        result = subprocess.run(
            ['python', filepath],
            capture_output=True,
            text=True,
            encoding='utf-8',
            timeout=EXECUTION_TIMEOUT
        )
        # 実行時エラーがあればそれを返す
        if result.returncode != 0:
            # エラーメッセージの最後の行を抽出
            error_line = result.stderr.strip().splitlines()[-1] if result.stderr.strip() else "実行時エラー"
            return error_line
        return "成功"
    except subprocess.TimeoutExpired:
        return f"{EXECUTION_TIMEOUT}秒タイムアウト"
    except Exception as e:
        return f"実行不可: {e}"

def check_assignments(answer_file: str, submission_dir: str):
    """
    提出物と模範解答を比較し、結果をExcelファイルに出力する。
    """
    try:
        with open(answer_file, 'r', encoding='utf-8') as f:
            answer_code = f.read()
        normalized_answer = normalize_code_by_ast(answer_code)
        if not normalized_answer:
            print(f"エラー: 模範解答ファイル '{answer_file}' に構文エラーがあります。")
            return
    except Exception as e:
        print(f"エラー: 模範解答ファイルの読み込み中にエラー: {e}")
        return

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "チェック結果"
    ws.append(["学籍番号", "氏名", "実行可否", "類似度", "判定"])

    for filename in sorted(os.listdir(submission_dir)):
        if not filename.endswith(TARGET_EXTENSION):
            continue

        filepath = os.path.join(submission_dir, filename)
        student_id, student_name = "-", "-"
        exec_status, similarity, grade = "-", 0.0, "-"

        try:
            match = FILENAME_PATTERN.match(os.path.splitext(filename)[0])
            if match:
                student_id, student_name = match.group(1) or "不明", match.group(2) or ""

            with open(filepath, 'r', encoding='utf-8') as f:
                submission_code = f.read()
            
            normalized_submission = normalize_code_by_ast(submission_code)
            
            if not normalized_submission:
                exec_status = "構文エラー"
                grade = "エラー"
            else:
                # 実行評価
                exec_status = check_execution(filepath)
                
                # 類似度評価
                matcher = difflib.SequenceMatcher(None, normalized_answer, normalized_submission)
                similarity = round(matcher.ratio(), 4)
                if similarity >= GRADE_THRESHOLDS['A']:
                    grade = 'A'
                elif similarity >= GRADE_THRESHOLDS['B']:
                    grade = 'B'
                elif similarity >= GRADE_THRESHOLDS['C']:
                    grade = 'C'
                else:
                    grade = 'D'

        except UnicodeDecodeError:
            exec_status = "文字コードエラー"
            grade = "エラー"
        except Exception as e:
            exec_status = f"予期せぬエラー: {e}"
            grade = "エラー"
        
        ws.append([student_id, student_name, exec_status, similarity, grade])
        print(f"ファイル: {filename:<25} | 実行: {exec_status:<15} | 類似度: {similarity:.2%} | 判定: {grade}")

    dir_name = os.path.basename(os.path.normpath(submission_dir))
    excel_filename = f"チェック結果_{dir_name}.xlsx"
    try:
        wb.save(excel_filename)
        print(f"\nチェック結果を '{excel_filename}' に保存しました。")
    except Exception as e:
        print(f"\nエラー: Excelファイルの保存に失敗しました: {e}")

def main():
    parser = argparse.ArgumentParser(description='Python課題の類似度と実行可否をチェックし、結果をExcelに出力します。')
    parser.add_argument('answer_file', type=str, help='模範解答のPythonファイルへのパス')
    parser.add_argument('submission_dir', type=str, help='学生の提出物が格納されたディレクトリへのパス')
    
    args = parser.parse_args()

    if not os.path.isfile(args.answer_file):
        print(f"エラー: 模範解答ファイルが見つかりません: {args.answer_file}")
        return
    if not os.path.isdir(args.submission_dir):
        print(f"エラー: 提出物ディレクトリが見つかりません: {args.submission_dir}")
        return

    check_assignments(args.answer_file, args.submission_dir)

if __name__ == "__main__":
    main()

