#coding:utf-8

import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog
import os
import comtypes.client
from urllib.parse import unquote
import re
import sys
import zipfile

def extract_zip_file():
    global file_path
    global zip_number
    
    # ファイル選択ダイアログを開く
    root = tk.Tk()
    root.withdraw()

    # ファイルの種類を制限する
    filetypes = [("zipファイル", "*.zip")]
    file_path = filedialog.askopenfilename(filetypes=filetypes, title="zipファイルを選択してください")

    print(file_path)

    # キャンセルが押された場合
    if not file_path:
        print('ファイルが選択されませんでした。処理を終了します。')
        sys.exit()

    # 選択されたファイルがzipファイルであるかを確認する
    if file_path.endswith('.zip'):
        # zipファイルの展開先を設定する
        zip_file_path = file_path     
        destination_path = os.path.splitext(file_path)[0]

        # 展開先フォルダが存在しない場合は作成
    if not os.path.exists(destination_path):
        os.makedirs(destination_path)

    # Python 標準ライブラリで展開
        try:
            with zipfile.ZipFile(zip_file_path, 'r') as zip_ref:
                zip_ref.extractall(destination_path)
            print("ファイルの展開が完了しました。")
        except zipfile.BadZipFile as e:
            print(f"エラー: ZIPファイルの展開に失敗しました。詳細: {e}")

        # zipファイルの数字を抽出する
        match = re.search('第(\d+)回', os.path.basename(file_path))
        if match:
            zip_number = match.group(1)
            print(f'報告会の通し番号: {zip_number}')
        else:
            print('zipファイル名に番号(報告会の実施回)が含まれていません。')

    else:
        print('選択されたファイルがzipファイルではありません。')
        sys.exit()

def find_word_files(root_folder):
    word_files = []
    for foldername, _, filenames in os.walk(root_folder):
        for filename in filenames:
            if filename.endswith('.docx') or filename.endswith('.doc'):
                filename = os.path.normpath(filename)
                filename = filename.replace('\u3000', '　')  # 全角スペースを半角スペースに置き換え
                word_files.append((foldername, filename))
    return word_files

def convert_word_to_pdf(root_folder, pdf_folder, callback=None):
    word_files = find_word_files(root_folder)

    if not os.path.exists(pdf_folder):
        os.mkdir(pdf_folder)

    completed_files = 0

    for foldername, filename in word_files:
        group = foldername.split("\\")[-1]
        word_path = os.path.join(foldername, filename)
        word_path = os.path.normpath(word_path)
        print(word_path)
        pdf_name = "【" + group + "】 " + os.path.splitext(filename)[0] + ".pdf"
        pdf_path = os.path.join(pdf_folder, pdf_name)
        pdf_path = os.path.normpath(pdf_path)

        word = comtypes.client.CreateObject('Word.Application', dynamic=True)
        word.Visible = False
        doc = word.Documents.Open(word_path)
        doc.SaveAs(pdf_path, 17)
        doc.Close()
        word.Quit()

        # ループ内で変換が完了したファイル数をインクリメントする
        completed_files += 1

        if callback:
            callback(total_files, completed_files)

    print('PDFへの変換が完了しました。')

def update_progress(total_files, completed_files):
    progress["value"] = completed_files
    label_completed.config(text="Completed Files: {}".format(completed_files))
    root.update()

def rename_pdf_files(pdf_folder):
    for filename in os.listdir(pdf_folder):
        if filename.endswith('.pdf'):
            # ファイル名に含まれる文字を置き換える
            new_name = re.sub('[,，、]', '・', filename)
            new_name = re.sub('[₋_＿　]', ' ', new_name)
            new_name = re.sub('報告会', '報告書', new_name)

             # "回"の後に"報告書"を追加する
            if "回" in new_name:
                pos = new_name.find("回")
                if new_name[pos+1] == " ":
                    new_name = new_name[:pos] + "回報告書 " + new_name[pos+2:]

            # ファイル名を変更する
            old_path = os.path.join(pdf_folder, filename)
            new_path = os.path.join(pdf_folder, new_name)
            os.rename(old_path, new_path)

    print('PDFファイルのファイル名を変更しました。')



# 関数を呼び出してzipファイルを展開する
extract_zip_file()
extract_path = os.path.splitext(file_path)[0]
print(extract_path)
#extract_path = r"C:\Users\Haruki\Desktop\第18回"
#zip_number = 18

# ファイルを検索してwordファイルを抽出する
word_files = find_word_files(extract_path)
total_files = len(word_files)

print(str(len(word_files)) + "個のwordファイルが見つかりました")
print(word_files)

# GUIの部分
root = tk.Tk()
root.protocol("WM_DELETE_WINDOW", lambda: None)
root.attributes('-toolwindow', True)
root.title("Word to PDF Converter")
frame = tk.Frame(root, padx=10, pady=10)
frame.pack()
progress = ttk.Progressbar(frame, orient="horizontal", length=200, mode="determinate")
progress["maximum"] = total_files
progress["value"] = 0
progress.grid(row=0, column=0, columnspan=2, pady=10)
label_total = tk.Label(frame, text=f"Total Files: {total_files}")
label_total.grid(row=1, column=0)
label_completed = tk.Label(frame, text="Completed Files: 0")
label_completed.grid(row=1, column=1)

# ワードファイルをPDFに変換し、指定のフォルダに保存する
pdf_folder = os.path.join(extract_path, "第" + str(zip_number) + "回報告書PDF")
# 変換処理の開始
convert_word_to_pdf(extract_path, pdf_folder, callback=update_progress)

rename_pdf_files(pdf_folder)