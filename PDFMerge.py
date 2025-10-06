#coding:utf-8

import os
import sys
import tkinter as tk
from tkinter import *
from tkinter import filedialog
from PyPDF2 import PdfMerger

class PdfMergerApp(Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.master.title("PDF Merger")
        self.pack()
        self.create_widgets()
        self.pdf_merger = PdfMerger()
        self.pdf_files = []

    def create_widgets(self):
        # Select PDF Filesボタンとメッセージを配置するフレーム
        select_frame = Frame(self)
        select_frame.pack(side="top", padx=10, pady=5)

        select_label = Label(select_frame, text="PDFが保存されているフォルダを選択", width=30)
        select_label.pack(side="left")

        self.select_button = Button(select_frame, text="選択", command=self.select_files)
        self.select_button.pack(side="left", padx=5)

        # PDFファイルのリストボックスとスクロールバーを配置するフレーム
        list_frame = Frame(self)
        list_frame.pack(side="top", padx=10, pady=0, fill="both", expand=True)

        self.listbox = Listbox(list_frame, width=45, height=25, selectmode="extended") # 幅を50に指定、複数選択可能にする
        self.listbox.pack(side="left", fill="both", expand=True)

        scrollbar = Scrollbar(list_frame, command=self.listbox.yview)
        scrollbar.pack(side="right", fill="y")
        self.listbox.config(yscrollcommand=scrollbar.set)

        # Shiftキーが押された時の動作を定義
        self.listbox.bind("<Shift-Button-1>", self.on_shift_click)

        # 上下移動ボタンのフレーム
        move_frame = Frame(self)
        move_frame.pack(side="top", padx=20, pady=10)

        self.move_up_button = Button(move_frame, text="選択アイテムを上へ", command=self.move_up)
        self.move_up_button.pack(side="left", padx=5)
        self.move_up_button.configure(borderwidth=2, highlightthickness=5, bg="LightCoral")

        self.move_down_button = Button(move_frame, text="選択アイテムを下へ", command=self.move_down)
        self.move_down_button.pack(side="left", padx=5)
        self.move_down_button.configure(borderwidth=2, highlightthickness=5, bg="LightSteelBlue") 

        self.sort_button = Button(move_frame, text="ソート", command=self.sort_files)
        self.sort_button.pack(side="right", padx=(5,0))  
        self.sort_button.configure(bg="pale green") 

        # Shift選択のラベルを表示するフレーム
        label_frame = Frame(self)
        label_frame.pack(side="top", pady=(0, 10))

        shift_label = Label(label_frame, text="Shiftを押しながら複数選択できます")
        shift_label.pack(side="left", padx=(0,25)) 


        # Merge PDF Filesボタンとメッセージを配置するフレーム
        # Delete PDF Filesボタンとメッセージを配置するフレーム
        merge_delete_frame = Frame(self)
        merge_delete_frame.pack(side="top", padx=20, pady=(0,5))

        self.merge_button = Button(merge_delete_frame, text="結合して保存", command=self.merge_files)
        self.merge_button.pack(side="left", padx=(90,50))
        self.merge_button.configure(borderwidth=2, highlightthickness=5) 

        self.delete_button = Button(merge_delete_frame, text="削除", command=self.delete_files)
        self.delete_button.pack(side="right", padx=(30,0))
        #self.delete_button.configure(bg="snow") 


    def select_files(self):
        default_folder = os.path.expanduser("~/Downloads") # ダウンロードフォルダのパスを取得
        folder_path = filedialog.askdirectory(title="Select Folder", initialdir=default_folder)
        if folder_path:
            self.pdf_files = [os.path.join(folder_path, file) for file in os.listdir(folder_path) if file.endswith(".pdf")]
            self.pdf_files = sorted(self.pdf_files, key=lambda x: ("【S班】" in x, "【N班】" in x, "【R班】" in x, "【AI班】" in x))
            for pdf_file in self.pdf_files:
                self.listbox.insert(END, os.path.basename(pdf_file))

    def on_shift_click(self, event):
        # シフトボタンが押された場合、複数選択を許可する
        if event.state == 1:
            # クリックされたアイテムのインデックスを取得
            index = self.listbox.nearest(event.y)
            # 選択されたアイテムの前後のアイテムを含む範囲を計算
            selected_range = [index, index]
            try:
                while True:
                    selected_range[0] -= 1
                    if selected_range[0] < 0:
                        selected_range[0] = 0
                        break
                    if self.listbox.selection_includes(selected_range[0]):
                        break
            except:
                pass
            try:
                while True:
                    selected_range[1] += 1
                    if self.listbox.selection_includes(selected_range[1]):
                        break
            except:
                pass
            # 選択範囲を設定
            self.listbox.selection_clear(0, END)
            self.listbox.selection_set(selected_range[0], selected_range[1])

    def move_up(self):
        selected_indices = self.listbox.curselection()
        if selected_indices:
            for index in selected_indices:
                if index > 0:
                    text = self.listbox.get(index)
                    self.listbox.delete(index)
                    self.listbox.insert(index-1, text)
                    self.pdf_files[index], self.pdf_files[index-1] = self.pdf_files[index-1], self.pdf_files[index]
            self.listbox.selection_clear(0, END)
            for index in selected_indices:
                self.listbox.selection_set(index-1)

    def move_down(self):
        selected_indices = self.listbox.curselection()
        if selected_indices:
            max_index = self.listbox.size() - 1
            for index in sorted(selected_indices, reverse=True):
                if index < max_index:
                    text = self.listbox.get(index)
                    self.listbox.delete(index)
                    self.listbox.insert(index+1, text)
                    self.pdf_files[index], self.pdf_files[index+1] = self.pdf_files[index+1], self.pdf_files[index]
            self.listbox.selection_clear(0, END)
            for index in selected_indices:
                self.listbox.selection_set(index+1)
                

    def delete_files(self):
        selected_items = self.listbox.curselection()
        for item in reversed(selected_items):
            file_path = self.pdf_files[item]
            self.listbox.delete(item)
            self.pdf_files.remove(file_path)


    def merge_files(self):
        output_file = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")], title="Save Merged PDF File")
        if output_file:
            for file_path in self.pdf_files:
                if os.path.exists(file_path):
                    with open(file_path, "rb") as pdf_file:
                        self.pdf_merger.append(pdf_file)
            with open(output_file, "wb") as output_stream:
                self.pdf_merger.write(output_stream)
            self.master.destroy() # ウィンドウを閉じる
    """
    def select_text_file(self):
        text_file_path = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt")], title="ソート用のテキストファイルを選択")
        if text_file_path:
            print("Selected file:", text_file_path) 
        return text_file_path
    """

    def sort_name(self, team, group_list):
        sorted_group_list = []

        for name in team:
            for item in group_list:
                if name in item:
                    if item not in sorted_group_list:
                        sorted_group_list.append(item)
                    break
        
        for item in group_list:
            if item not in sorted_group_list:
                sorted_group_list.append(item)


        return sorted_group_list

    def sort_files(self):
        ai_list = []
        r_list = []
        n_list = []
        s_list = []

        for pdf_file in self.pdf_files:
            if "【AI班】" in pdf_file:
                ai_list.append(pdf_file)
            elif "【R班】" in pdf_file:
                r_list.append(pdf_file)
            elif "【N班】" in pdf_file:
                n_list.append(pdf_file)
            elif "【S班】" in pdf_file:
                s_list.append(pdf_file)
        

        # 実行ファイルのパスを取得
        exe_path = sys.executable
        # 実行ファイルのディレクトリパスを取得
        exe_dir = os.path.dirname(exe_path)
        # テキストファイルのパスを生成
        text_file_path = os.path.join(exe_dir, '発表順ソートリスト.txt')

        if os.path.exists(text_file_path):            
            # ファイルの読み込み
            with open(text_file_path, 'r', encoding='utf-8') as file:
                content = file.read()

            # 発表順のリストを初期化
            ai_team = []
            r_team = []
            n_team = []
            s_team = []

            # 行ごとに処理
            for line in content.split('\n'):
                if any(x in line for x in ["【AI班】", "【R班】", "【N班】", "【S班】"]):
                    # 行をスペースで分割して班名とメンバーリストに分ける
                    team, members = line.split(' ')

                    # 班名に応じて発表順のリストにメンバーを追加
                    if team == '【AI班】':
                        ai_team = eval(members)
                    elif team == '【R班】':
                        r_team = eval(members)
                    elif team == '【N班】':
                        n_team = eval(members)
                    elif team == '【S班】':
                        s_team = eval(members)

            ai_list = self.sort_name(ai_team, ai_list)
            r_list = self.sort_name(r_team, r_list)
            n_list = self.sort_name(n_team, n_list)
            s_list = self.sort_name(s_team, s_list)

        self.pdf_files = ai_list + r_list + n_list + s_list

        self.listbox.delete(0, END)
        for pdf_file in self.pdf_files:
            self.listbox.insert(END, os.path.basename(pdf_file))

if __name__ == "__main__":
    # アプリケーションを開始する
    root = Tk()
    root.resizable(False, False)  # ウィンドウのサイズを固定する
    app = PdfMergerApp(master=root)
    app.mainloop()
