"""Desktop GUI application for the robotics report PDF converter.

This module provides a Tkinter-based interface that mirrors the workflow of
the Flask web application defined in :mod:`app`.  Users can choose a ZIP file
that contains the Word documents, adjust the presentation order, edit the
default ordering preferences, and monitor the conversion progress in a
dedicated status view.  The conversion pipeline – extraction, Word to PDF
conversion, merging, and optional email delivery – reuses the existing
utility functions so that the desktop edition remains compatible with the web
edition.
"""

from __future__ import annotations

import os
import queue
import shutil
import subprocess
import sys
import threading
import uuid
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Optional

import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

from email_service import send_email_with_attachment
from pdf_merge import merge_pdfs

from word_to_pdf_converter import ConversionError, convert_word_to_pdf

# Import the shared logic from the Flask application.  The imported objects are
# intentionally prefixed with underscores in the original module, but they are
# re-used here to guarantee that ordering rules and data directories are kept
# in sync with the web version.
from app import (  # type: ignore  # noqa: E402
    EMAIL_CONFIG,
    DEFAULT_RECIPIENT_EMAIL,
    UNGROUPED_TEAM_KEY,
    UPLOAD_DIR,
    WORK_DIR,
    ZipEntry,
    _apply_team_prefixes,
    _cleanup_data_directories,
    _determine_report_number,
    _extract_entries,
    _normalize_team_key,
    _team_display_label,
    order_manager,
)


@dataclass
class TeamBlock:
    """Represents a logical block of ordered entries for a single team."""

    key: str
    label: str
    entries: List[ZipEntry]

    def to_names(self) -> List[str]:
        return [entry.display_name for entry in self.entries]


class ScrollableFrame(ttk.Frame):
    """A vertically scrollable frame used for the ordering screen."""

    def __init__(self, master: tk.Widget) -> None:
        super().__init__(master)
        canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self._content = ttk.Frame(canvas)
        self._content.bind(
            "<Configure>",
            lambda event: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        canvas.create_window((0, 0), window=self._content, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        self._canvas = canvas

    @property
    def content(self) -> ttk.Frame:
        return self._content

    def scroll_to_top(self) -> None:
        self._canvas.yview_moveto(0.0)


class BaseView(ttk.Frame):
    """Base class for all views."""

    def __init__(self, master: tk.Widget, controller: "DesktopApp") -> None:
        super().__init__(master)
        self.controller = controller

    def on_show(self) -> None:  # pragma: no cover - UI hook
        """Hook that is executed every time the view becomes visible."""


class WelcomeView(BaseView):
    """Initial screen allowing users to choose the ZIP archive."""

    def __init__(self, master: tk.Widget, controller: "DesktopApp") -> None:
        super().__init__(master, controller)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        wrapper = ttk.Frame(self, padding=24, style="WelcomeWrapper.TFrame")
        wrapper.grid(row=0, column=0, sticky="nsew")
        wrapper.columnconfigure(0, weight=1)

        ttk.Label(wrapper, text="ロボ研報告書作成ツール", style="Heading.TLabel").grid(
            row=0, column=0, sticky="w"
        )

        ttk.Label(
            wrapper,
            text=(
                "Word形式の報告書をまとめたZIPファイルをアップロードしてください。\n"
                "PDFに変換して結合します。"
            ),
            style="BodyText.TLabel",
            justify="left",
        ).grid(row=1, column=0, pady=(12, 16), sticky="w")

        hint = ttk.Frame(wrapper, padding=(16, 12), style="Hint.TFrame")
        hint.grid(row=2, column=0, sticky="ew")
        ttk.Label(
            hint,
            text=(
                "ZIP内にR班やN班などのフォルダーがある場合は班名として自動認識します。\n"
                "サブフォルダーが無い場合はアップロードしたZIPファイル名を班名として表示します。"
            ),
            style="HintText.TLabel",
            justify="left",
        ).grid(row=0, column=0, sticky="w")

        controls = ttk.Frame(wrapper, style="WelcomeWrapper.TFrame")
        controls.grid(row=3, column=0, pady=(24, 0))

        self._selected_path = tk.StringVar(value="ファイルが選択されていません")

        select_button = ttk.Button(
            controls,
            text="ZIPファイルを選択",
            command=self._choose_file,
            style="AccentWide.TButton",
        )
        select_button.grid(row=0, column=0, padx=6)

        ttk.Label(
            controls,
            textvariable=self._selected_path,
            style="Path.TLabel",
        ).grid(row=1, column=0, pady=(8, 16))

        self._next_button = ttk.Button(
            controls,
            text="順番を編集",
            command=self._proceed,
            state=tk.DISABLED,
            style="SecondaryWide.TButton",
        )
        self._next_button.grid(row=2, column=0)

    def _choose_file(self) -> None:
        path = filedialog.askopenfilename(
            title="ZIPファイルを選択",
            filetypes=[("ZIP Archives", "*.zip"), ("All Files", "*.*")],
        )
        if not path:
            return
        self._selected_path.set(path)
        self._next_button.state(["!disabled"])

    def _proceed(self) -> None:
        zip_path = self._selected_path.get().strip()
        if not zip_path:
            return
        try:
            self.controller.prepare_upload(Path(zip_path))
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("エラー", f"ZIPファイルの読み込みに失敗しました: {exc}")


class TeamBlockWidget(ttk.Frame):
    """Widget used inside :class:`OrderView` to represent a team block."""

    def __init__(
        self,
        master: tk.Widget,
        controller: "DesktopApp",
        block: TeamBlock,
        index: int,
    ) -> None:
        super().__init__(master, padding=16, style="TeamBlock.TFrame")
        self.controller = controller
        self.block = block
        self.index = index

        self.columnconfigure(0, weight=1)

        header = ttk.Frame(self, style="TeamHeader.TFrame")
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)

        summary = ttk.Frame(header, style="TeamSummary.TFrame")
        summary.grid(row=0, column=0, sticky="w")
        ttk.Label(
            summary,
            text=f"{block.label}",
            style="TeamTitle.TLabel",
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            summary,
            text=f"{len(block.entries)} 件",
            style="TeamCount.TLabel",
        ).grid(row=0, column=1, padx=(10, 0))

        control_frame = ttk.Frame(header, style="TeamActions.TFrame")
        control_frame.grid(row=0, column=1, sticky="e")
        ttk.Button(
            control_frame,
            text="▲",
            width=3,
            command=lambda: controller.move_team(self.index, -1),
            style="Secondary.TButton",
        ).grid(row=0, column=0, padx=2)
        ttk.Button(
            control_frame,
            text="▼",
            width=3,
            command=lambda: controller.move_team(self.index, 1),
            style="Secondary.TButton",
        ).grid(row=0, column=1, padx=2)

        entries_container = ttk.Frame(self, style="DocumentList.TFrame")
        entries_container.grid(row=1, column=0, sticky="ew", pady=(16, 0))
        entries_container.columnconfigure(0, weight=1)

        for entry_index, entry in enumerate(block.entries):
            self._create_entry_row(entries_container, entry, entry_index)

    def _create_entry_row(
        self,
        container: ttk.Frame,
        entry: ZipEntry,
        entry_index: int,
    ) -> None:
        row = ttk.Frame(container, padding=(12, 10), style="DocumentItem.TFrame")
        row.grid(row=entry_index, column=0, sticky="ew", pady=(0, 10))
        row.columnconfigure(0, weight=1)

        info = ttk.Frame(row, style="DocumentInfo.TFrame")
        info.grid(row=0, column=0, sticky="w")
        ttk.Label(info, text=entry.display_name, style="DocTitle.TLabel").grid(
            row=0,
            column=0,
            sticky="w",
        )

        actions = ttk.Frame(row, style="DocActions.TFrame")
        actions.grid(row=0, column=1, sticky="e")
        ttk.Button(
            actions,
            text="▲",
            width=3,
            style="Secondary.TButton",
            command=lambda: self.controller.move_entry(self.index, entry_index, -1),
        ).grid(row=0, column=0, padx=2)
        ttk.Button(
            actions,
            text="▼",
            width=3,
            style="Secondary.TButton",
            command=lambda: self.controller.move_entry(self.index, entry_index, 1),
        ).grid(row=0, column=1, padx=2)
        ttk.Button(
            actions,
            text="削除",
            style="Danger.TButton",
            command=lambda: self.controller.remove_entry(self.index, entry_index),
        ).grid(row=0, column=2, padx=(8, 0))


class OrderView(BaseView):
    """Screen that allows the user to tweak the document order."""

    def __init__(self, master: tk.Widget, controller: "DesktopApp") -> None:
        super().__init__(master, controller)

        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        wrapper = ttk.Frame(self, padding=24, style="OrderWrapper.TFrame")
        wrapper.grid(row=0, column=0, sticky="nsew")
        wrapper.columnconfigure(0, weight=1)
        wrapper.rowconfigure(2, weight=1)

        ttk.Label(wrapper, text="報告書の並び替え", style="Heading.TLabel").grid(
            row=0, column=0, sticky="w"
        )

        instructions = ttk.Frame(wrapper, padding=(16, 12), style="Instruction.TFrame")
        instructions.grid(row=1, column=0, sticky="ew", pady=(16, 12))
        ttk.Label(
            instructions,
            text=(
                "上下ボタンで班単位・個人単位の並び替えができます。発表の順番に合わせて調整してください。\n"
                "デフォルト順でソートされています。不要なファイルは「削除」で除外できます。"
            ),
            style="InstructionText.TLabel",
            justify="left",
            wraplength=640,
        ).grid(row=0, column=0, sticky="w")

        self.scrollable = ScrollableFrame(wrapper)
        self.scrollable.grid(row=2, column=0, sticky="nsew")
        self.scrollable.content.columnconfigure(0, weight=1)

        controls = ttk.Frame(wrapper, style="Controls.TFrame")
        controls.grid(row=3, column=0, pady=(18, 0))

        ttk.Button(
            controls,
            text="ソート",
            style="SecondaryWide.TButton",
            command=controller.apply_default_order,
        ).grid(row=0, column=0, padx=6)
        ttk.Button(
            controls,
            text="デフォルト順を編集",
            style="SecondaryWide.TButton",
            command=controller.open_default_order_editor,
        ).grid(row=0, column=1, padx=6)
        ttk.Button(
            controls,
            text="最初の状態に戻す",
            style="SecondaryWide.TButton",
            command=controller.reset_order,
        ).grid(row=0, column=2, padx=6)
        ttk.Button(
            controls,
            text="PDF 変換を開始",
            style="AccentWide.TButton",
            command=self._start_processing,
        ).grid(row=0, column=3, padx=12)

    def on_show(self) -> None:
        self.refresh_blocks()

    def refresh_blocks(self) -> None:
        for child in self.scrollable.content.winfo_children():
            child.destroy()
        blocks = self.controller.team_blocks
        for index, block in enumerate(blocks):
            widget = TeamBlockWidget(self.scrollable.content, self.controller, block, index)
            widget.grid(row=index, column=0, sticky="ew", pady=8, padx=12)
        if not blocks:
            ttk.Label(
                self.scrollable.content,
                text="対象のファイルがありません。",
                style="Muted.TLabel",
            ).grid(row=0, column=0, pady=24)
        self.scrollable.content.columnconfigure(0, weight=1)

    def _start_processing(self) -> None:
        self.controller.start_processing()


class StatusView(BaseView):
    """Displays progress updates and completion details."""

    def __init__(self, master: tk.Widget, controller: "DesktopApp") -> None:
        super().__init__(master, controller)
        self.columnconfigure(0, weight=1)

        title = ttk.Label(
            self,
            text="処理状況",
            font=("Helvetica", 20, "bold"),
            foreground="#0b7285",
        )
        title.grid(row=0, column=0, pady=(16, 12))

        self.status_label = ttk.Label(
            self,
            text="待機中",
            font=("Helvetica", 16, "bold"),
        )
        self.status_label.grid(row=1, column=0, pady=(0, 8))

        self.message_label = ttk.Label(self, text="", foreground="#475569")
        self.message_label.grid(row=2, column=0)

        progress_frame = ttk.Frame(self, padding=12, style="Card.TFrame")
        progress_frame.grid(row=3, column=0, pady=12, sticky="ew")
        progress_frame.columnconfigure(0, weight=1)

        self.progress = ttk.Progressbar(progress_frame, maximum=100)
        self.progress.grid(row=0, column=0, sticky="ew")

        self.progress_text = ttk.Label(progress_frame, text="")
        self.progress_text.grid(row=1, column=0, sticky="w", pady=(8, 0))

        completion_frame = ttk.Frame(self, padding=12)
        completion_frame.grid(row=4, column=0, sticky="ew")
        completion_frame.columnconfigure(0, weight=1)

        self.pdf_path_var = tk.StringVar(value="")
        ttk.Label(
            completion_frame,
            textvariable=self.pdf_path_var,
            wraplength=600,
            foreground="#0f172a",
        ).grid(row=0, column=0, sticky="w")

        button_frame = ttk.Frame(completion_frame)
        button_frame.grid(row=1, column=0, pady=(8, 0), sticky="w")

        self.open_pdf_button = ttk.Button(
            button_frame,
            text="PDFを開く",
            state=tk.DISABLED,
            command=lambda: self.controller.open_output_file(self._pdf_path),
        )
        self.open_pdf_button.grid(row=0, column=0, padx=(0, 8))

        self.reveal_folder_button = ttk.Button(
            button_frame,
            text="フォルダーを開く",
            state=tk.DISABLED,
            command=lambda: self.controller.reveal_output_folder(self._pdf_path),
        )
        self.reveal_folder_button.grid(row=0, column=1)

        self._pdf_path: Optional[Path] = None
        self._report_number: str = ""

    def update_status(self, payload: Dict[str, object]) -> None:
        status = payload.get("status", "queued")
        message = str(payload.get("message", ""))
        self.status_label.config(text=self._status_label(status))
        self.message_label.config(text=message)

        progress_current = int(payload.get("progress_current", 0))
        progress_total = int(payload.get("progress_total", 0))
        percent = int(payload.get("progress_percent", 0))
        if progress_total > 0:
            if percent <= 0:
                percent = round(progress_current / progress_total * 100)
            self.progress.config(value=percent)
            self.progress_text.config(
                text=f"進捗: {progress_current}/{progress_total} ({percent}%)"
            )
        else:
            self.progress.config(value=0)
            self.progress_text.config(text="")

    def show_completion(self, pdf_path: Path, report_number: str) -> None:
        self._pdf_path = pdf_path
        self._report_number = report_number
        self.status_label.config(text="完了")
        if report_number:
            self.message_label.config(
                text=f"第{report_number}回報告書のPDFを生成しました。"
            )
        else:
            self.message_label.config(text="PDFの結合が完了しました。")
        self.progress.config(value=100)
        self.progress_text.config(text="処理が完了しました。")
        self.pdf_path_var.set(f"生成されたPDF: {pdf_path}")
        self.open_pdf_button.state(["!disabled"])
        self.reveal_folder_button.state(["!disabled"])

    def show_failure(self, message: str) -> None:
        self.status_label.config(text="失敗")
        self.message_label.config(text=message)
        self.progress_text.config(text="エラーが発生しました。")
        self.progress.config(value=0)
        self.open_pdf_button.state(["disabled"])
        self.reveal_folder_button.state(["disabled"])

    def _status_label(self, status: str) -> str:
        mapping = {
            "queued": "待機中",
            "running": "処理中",
            "completed": "完了",
            "failed": "失敗",
        }
        return mapping.get(status, status)


class DefaultOrderEditor(tk.Toplevel):
    """Modal window that lets the user edit default member sequences."""

    def __init__(self, controller: "DesktopApp") -> None:
        super().__init__(controller.root)
        self.controller = controller
        self.title("デフォルト順の編集")
        self.geometry("640x520")
        self.transient(controller.root)
        self.grab_set()

        container = ttk.Frame(self, padding=16)
        container.pack(fill=tk.BOTH, expand=True)
        container.columnconfigure(0, weight=1)

        selector_frame = ttk.Frame(container)
        selector_frame.grid(row=0, column=0, sticky="ew")
        selector_frame.columnconfigure(1, weight=1)

        ttk.Label(selector_frame, text="対象の班").grid(row=0, column=0, sticky="w")
        self.team_combo = ttk.Combobox(selector_frame, state="readonly")
        self.team_combo.grid(row=0, column=1, sticky="ew", padx=(12, 0))

        ttk.Button(
            selector_frame,
            text="この班の並びを削除",
            command=self._delete_sequence,
        ).grid(row=0, column=2, padx=(12, 0))

        self.new_team_entry = ttk.Entry(selector_frame)
        self.new_team_entry.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(12, 0))
        ttk.Button(
            selector_frame,
            text="新しい班を追加",
            command=self._create_team,
        ).grid(row=1, column=2, padx=(12, 0), pady=(12, 0))

        self.member_list = tk.Listbox(container, activestyle="none")
        self.member_list.grid(row=1, column=0, sticky="nsew", pady=(16, 0))
        container.rowconfigure(1, weight=1)

        controls = ttk.Frame(container)
        controls.grid(row=2, column=0, sticky="ew", pady=(12, 0))
        ttk.Button(controls, text="▲", width=3, command=lambda: self._move(-1)).grid(
            row=0, column=0
        )
        ttk.Button(controls, text="▼", width=3, command=lambda: self._move(1)).grid(
            row=0, column=1, padx=8
        )
        ttk.Button(controls, text="削除", command=self._remove_member).grid(
            row=0, column=2
        )

        add_frame = ttk.Frame(container)
        add_frame.grid(row=3, column=0, sticky="ew", pady=(16, 0))
        add_frame.columnconfigure(0, weight=1)
        self.member_entry = ttk.Entry(add_frame)
        self.member_entry.grid(row=0, column=0, sticky="ew")
        ttk.Button(add_frame, text="追加", command=self._add_member).grid(
            row=0, column=1, padx=(12, 0)
        )

        footer = ttk.Frame(container)
        footer.grid(row=4, column=0, sticky="e", pady=(16, 0))
        ttk.Button(footer, text="キャンセル", command=self.destroy).grid(
            row=0, column=0, padx=(0, 12)
        )
        ttk.Button(footer, text="保存", command=self._save).grid(row=0, column=1)

        self.team_combo.bind("<<ComboboxSelected>>", lambda event: self._load_members())

        self._load_initial_data()

    def _load_initial_data(self) -> None:
        preferences = order_manager.load_preferences()
        teams = self.controller.collect_preference_teams(preferences)
        if not teams:
            teams = [
                {
                    "key": UNGROUPED_TEAM_KEY,
                    "label": _team_display_label(UNGROUPED_TEAM_KEY),
                    "members": [],
                }
            ]
        self._teams = teams
        self.team_combo["values"] = [team["label"] for team in teams]
        self.team_combo.current(0)
        self._load_members()

    def _current_team(self) -> Dict[str, object]:
        index = self.team_combo.current()
        return self._teams[index]

    def _create_team(self) -> None:
        name = self.new_team_entry.get().strip()
        if not name:
            messagebox.showwarning("警告", "班名を入力してください。")
            return
        normalized = _normalize_team_key(name)
        existing_keys = {team["key"] for team in self._teams}
        if normalized in existing_keys:
            messagebox.showinfo("情報", "既に存在する班です。")
            index = next(
                (i for i, team in enumerate(self._teams) if team["key"] == normalized),
                0,
            )
            self.team_combo.current(index)
            self._load_members()
            return
        display_label = _team_display_label(normalized)
        self._teams.append({"key": normalized, "label": display_label, "members": []})
        self.team_combo["values"] = [team["label"] for team in self._teams]
        self.team_combo.current(len(self._teams) - 1)
        self._load_members()
        self.new_team_entry.delete(0, tk.END)

    def _load_members(self) -> None:
        team = self._current_team()
        self.member_list.delete(0, tk.END)
        for member in team.get("members", []):
            self.member_list.insert(tk.END, member)

    def _move(self, direction: int) -> None:
        selection = self.member_list.curselection()
        if not selection:
            return
        index = selection[0]
        target = index + direction
        if target < 0 or target >= self.member_list.size():
            return
        value = self.member_list.get(index)
        self.member_list.delete(index)
        self.member_list.insert(target, value)
        self.member_list.selection_set(target)

    def _remove_member(self) -> None:
        selection = self.member_list.curselection()
        if not selection:
            return
        self.member_list.delete(selection[0])

    def _add_member(self) -> None:
        value = self.member_entry.get().strip()
        if not value:
            return
        existing = {self.member_list.get(i) for i in range(self.member_list.size())}
        if value not in existing:
            self.member_list.insert(tk.END, value)
        self.member_entry.delete(0, tk.END)

    def _delete_sequence(self) -> None:
        team = self._current_team()
        key = str(team.get("key", ""))
        if messagebox.askyesno("確認", "この班のデフォルト順を削除しますか？"):
            order_manager.delete_member_sequence(key)
            self.destroy()
            messagebox.showinfo("完了", "デフォルト順を削除しました。")
            self.controller.reload_preferences()

    def _save(self) -> None:
        team = self._current_team()
        key = _normalize_team_key(str(team.get("key", "")))
        members = [self.member_list.get(i) for i in range(self.member_list.size())]
        order_manager.save_member_sequence(key, members)
        messagebox.showinfo("保存", "デフォルト順を保存しました。")
        self.controller.reload_preferences()
        self.destroy()


class DesktopApp:
    """Main application controller."""

    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("ロボ研報告書作成ツール (デスクトップ版)")
        self.root.geometry("960x720")
        self._configure_styles()

        self.container = ttk.Frame(self.root)
        self.container.pack(fill=tk.BOTH, expand=True)

        self.views: Dict[str, BaseView] = {}
        self.current_view: Optional[BaseView] = None

        self.selected_zip: Optional[Path] = None
        self.zip_original_name: str = ""
        self.entry_map: Dict[str, ZipEntry] = {}
        self.team_blocks: List[TeamBlock] = []
        self.initial_layout_snapshot: List[Dict[str, object]] = []

        self._preferences_cache = order_manager.load_preferences()

        self.status_queue: "queue.Queue[Dict[str, object]]" = queue.Queue()
        self._processing_thread: Optional[threading.Thread] = None

        self.show_view("welcome")

    def _configure_styles(self) -> None:
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:  # pragma: no cover - depends on platform
            pass
        style.configure(
            "Accent.TButton",
            foreground="#ffffff",
            background="#0b7285",
            padding=(8, 6),
        )
        style.map("Accent.TButton", background=[("active", "#0c8599")])
        style.configure(
            "AccentWide.TButton",
            foreground="#ffffff",
            background="#0b7285",
            padding=(10, 8),
        )
        style.map("AccentWide.TButton", background=[("active", "#0c8599")])
        style.configure(
            "Secondary.TButton",
            foreground="#ffffff",
            background="#495057",
            padding=(6, 4),
        )
        style.map("Secondary.TButton", background=[("active", "#343a40")])
        style.configure(
            "SecondaryWide.TButton",
            foreground="#ffffff",
            background="#495057",
            padding=(10, 8),
        )
        style.map("SecondaryWide.TButton", background=[("active", "#343a40")])
        style.configure(
            "Danger.TButton",
            foreground="#ffffff",
            background="#fa5252",
            padding=(6, 4),
        )
        style.map("Danger.TButton", background=[("active", "#e03131")])

        style.configure("WelcomeWrapper.TFrame", background="#f8fafc")
        style.configure("OrderWrapper.TFrame", background="#f8fafc")
        style.configure("Controls.TFrame", background="#f8fafc")
        style.configure("Instruction.TFrame", background="#e7f5ff", borderwidth=0)
        style.configure("InstructionText.TLabel", background="#e7f5ff", foreground="#1f2937")
        style.configure("BodyText.TLabel", background="#f8fafc", foreground="#1f2933")
        style.configure("Hint.TFrame", background="#e7f5ff", borderwidth=0)
        style.configure("HintText.TLabel", background="#e7f5ff", foreground="#0b7285")
        style.configure("Path.TLabel", background="#f8fafc", foreground="#334155")
        style.configure("Heading.TLabel", font=("Helvetica", 20, "bold"), foreground="#0b7285")
        style.configure("Muted.TLabel", foreground="#64748b")
        style.configure("Card.TFrame", background="#f8fafc", relief="groove", borderwidth=1)

        style.configure("TeamBlock.TFrame", background="#ffffff", relief="ridge", borderwidth=1)
        style.configure("TeamHeader.TFrame", background="#ffffff")
        style.configure("TeamSummary.TFrame", background="#ffffff")
        style.configure("TeamActions.TFrame", background="#ffffff")
        style.configure("TeamTitle.TLabel", font=("Helvetica", 14, "bold"), foreground="#0b7285", background="#ffffff")
        style.configure(
            "TeamCount.TLabel",
            background="#0b7285",
            foreground="#ffffff",
            padding=(8, 2),
        )
        style.configure("DocumentList.TFrame", background="#ffffff")
        style.configure(
            "DocumentItem.TFrame",
            background="#fdfdfd",
            relief="solid",
            borderwidth=1,
        )
        style.configure("DocumentInfo.TFrame", background="#fdfdfd")
        style.configure("DocActions.TFrame", background="#fdfdfd")
        style.configure("DocTitle.TLabel", background="#fdfdfd", font=("Helvetica", 11, "bold"))

    def show_view(self, name: str) -> None:
        if name not in self.views:
            if name == "welcome":
                self.views[name] = WelcomeView(self.container, self)
            elif name == "order":
                self.views[name] = OrderView(self.container, self)
            elif name == "status":
                self.views[name] = StatusView(self.container, self)
            else:
                raise ValueError(f"Unknown view: {name}")
        if self.current_view is not None:
            self.current_view.pack_forget()
        view = self.views[name]
        view.pack(fill=tk.BOTH, expand=True)
        self.current_view = view
        view.on_show()

    def prepare_upload(self, zip_path: Path) -> None:
        entries = _extract_entries(zip_path, original_name=zip_path.name)
        if not entries:
            raise RuntimeError("アップロードされたZIPにWordファイルが見つかりませんでした。")

        self.selected_zip = zip_path
        self.zip_original_name = zip_path.name
        self.entry_map = {entry.display_name: entry for entry in entries}

        team_sequence, team_entries = order_manager.initial_layout(entries)
        blocks: List[TeamBlock] = []
        initial_snapshot: List[Dict[str, object]] = []

        for key in team_sequence:
            items = list(team_entries.get(key, []))
            if not items:
                continue
            block = TeamBlock(key=key, label=_team_display_label(key), entries=items)
            blocks.append(block)
            initial_snapshot.append(
                {
                    "key": block.key,
                    "label": block.label,
                    "names": block.to_names(),
                }
            )

        self.team_blocks = blocks
        self.initial_layout_snapshot = initial_snapshot
        self._preferences_cache = order_manager.load_preferences()

        self.show_view("order")

    def move_team(self, index: int, offset: int) -> None:
        if index is None:
            return
        target = index + offset
        if target < 0 or target >= len(self.team_blocks):
            return
        self.team_blocks[index], self.team_blocks[target] = (
            self.team_blocks[target],
            self.team_blocks[index],
        )
        self._refresh_order_view()

    def move_entry(self, team_index: Optional[int], entry_index: Optional[int], offset: int) -> None:
        if team_index is None or entry_index is None:
            return
        block = self.team_blocks[team_index]
        target = entry_index + offset
        if target < 0 or target >= len(block.entries):
            return
        block.entries[entry_index], block.entries[target] = (
            block.entries[target],
            block.entries[entry_index],
        )
        self._refresh_order_view()

    def remove_entry(self, team_index: Optional[int], entry_index: Optional[int]) -> None:
        if team_index is None or entry_index is None:
            return
        block = self.team_blocks[team_index]
        del block.entries[entry_index]
        if not block.entries:
            del self.team_blocks[team_index]
        self._refresh_order_view()

    def _refresh_order_view(self) -> None:
        view = self.views.get("order")
        if isinstance(view, OrderView):
            view.refresh_blocks()

    def reset_order(self) -> None:
        rebuilt_blocks: List[TeamBlock] = []
        for snapshot in self.initial_layout_snapshot:
            names = [
                self.entry_map[name]
                for name in snapshot["names"]
                if name in self.entry_map
            ]
            if not names:
                continue
            rebuilt_blocks.append(
                TeamBlock(
                    key=str(snapshot["key"]),
                    label=str(snapshot["label"]),
                    entries=names,
                )
            )
        self.team_blocks = rebuilt_blocks
        self._refresh_order_view()

    def apply_default_order(self) -> None:
        preferences = self._preferences_cache
        team_order_map = {key: idx for idx, key in enumerate(preferences.team_sequence)}
        self.team_blocks.sort(
            key=lambda block: team_order_map.get(block.key, len(team_order_map))
        )
        for block in self.team_blocks:
            member_sequence = preferences.member_sequences.get(block.key, [])
            block.entries = order_manager._sort_team_entries(  # type: ignore[attr-defined]
                list(block.entries),
                list(member_sequence),
            )
        self._refresh_order_view()

    def open_default_order_editor(self) -> None:
        DefaultOrderEditor(self)

    def reload_preferences(self) -> None:
        self._preferences_cache = order_manager.load_preferences()

    def collect_preference_teams(self, preferences) -> List[Dict[str, object]]:
        teams: List[Dict[str, object]] = []
        seen: set[str] = set()

        def append_team(key: str) -> None:
            normalized = _normalize_team_key(key)
            if normalized in seen:
                return
            teams.append(
                {
                    "key": normalized,
                    "label": _team_display_label(normalized),
                    "members": list(preferences.member_sequences.get(normalized, [])),
                }
            )
            seen.add(normalized)

        for team_key in preferences.team_sequence:
            append_team(team_key)
        for team_key in preferences.member_sequences:
            append_team(team_key)
        if not teams:
            append_team(UNGROUPED_TEAM_KEY)
        return teams

    def start_processing(self) -> None:
        if not self.team_blocks:
            messagebox.showwarning("警告", "処理するファイルがありません。")
            return
        if not self.selected_zip:
            messagebox.showerror("エラー", "ZIPファイルが読み込まれていません。")
            return

        order = [name for block in self.team_blocks for name in block.to_names()]
        if not order:
            messagebox.showwarning("警告", "処理対象が空です。")
            return

        order_manager.save(order, self.entry_map)

        job_id = uuid.uuid4().hex
        temp_zip = UPLOAD_DIR / f"{job_id}.zip"
        shutil.copy2(self.selected_zip, temp_zip)

        status_view = self.views.get("status")
        if isinstance(status_view, StatusView):
            status_view.update_status(
                {
                    "status": "queued",
                    "message": "処理を待機しています…",
                    "progress_current": 0,
                    "progress_total": len(order),
                }
            )

        self.show_view("status")

        self.status_queue = queue.Queue()
        original_name = self.zip_original_name or self.selected_zip.name
        self._processing_thread = threading.Thread(
            target=self._run_processing,
            args=(job_id, temp_zip, order, original_name),
            daemon=True,
        )
        self._processing_thread.start()
        self.root.after(200, self._poll_status_queue)

    def _run_processing(
        self,
        job_id: str,
        zip_path: Path,
        order: List[str],
        zip_original_name: str,
    ) -> None:
        work_root = WORK_DIR / job_id
        extract_dir = work_root / "extracted"
        pdf_dir = work_root / "pdf"
        try:
            extract_dir.mkdir(parents=True, exist_ok=True)
            pdf_dir.mkdir(parents=True, exist_ok=True)

            self.status_queue.put(
                {
                    "type": "status",
                    "status": "running",
                    "message": "ZIPファイルを展開しています…",
                    "progress_current": 0,
                    "progress_total": len(order),
                }
            )

            shutil.unpack_archive(str(zip_path), extract_dir)

            entry_subset = {
                name: self.entry_map[name]
                for name in order
                if name in self.entry_map
            }

            _apply_team_prefixes(extract_dir, entry_subset)

            ordered_entries = [entry_subset[name] for name in order if name in entry_subset]
            if not ordered_entries:
                raise RuntimeError("処理対象のドキュメントが見つかりませんでした。")

            report_number = _determine_report_number(zip_original_name, ordered_entries)

            pdf_paths: List[Path] = []
            for index, entry in enumerate(ordered_entries, start=1):
                self.status_queue.put(
                    {
                        "type": "status",
                        "status": "running",
                        "message": f"{entry.display_name} をPDFに変換しています…",
                        "progress_current": index - 1,
                        "progress_total": len(order),
                    }
                )
                source_path = extract_dir / entry.archive_name
                if not source_path.exists():
                    raise FileNotFoundError(
                        f"展開後のファイルが見つかりません: {entry.archive_name}"
                    )
                pdf_path = convert_word_to_pdf(source_path, pdf_dir)
                pdf_paths.append(pdf_path)
                self.status_queue.put(
                    {
                        "type": "status",
                        "status": "running",
                        "message": f"{entry.display_name} のPDF変換が完了しました。",
                        "progress_current": index,
                        "progress_total": len(order),
                    }
                )

            merged_path = work_root / f"第{report_number}回報告書.pdf"
            self.status_queue.put(
                {
                    "type": "status",
                    "status": "running",
                    "message": "PDFファイルを結合しています…",
                    "progress_current": len(order),
                    "progress_total": len(order),
                }
            )
            merge_pdfs(pdf_paths, merged_path)

            recipient_email = ""
            if EMAIL_CONFIG.is_configured:
                recipient_email = DEFAULT_RECIPIENT_EMAIL.strip()
            if recipient_email:
                try:
                    send_email_with_attachment(
                        config=EMAIL_CONFIG,
                        recipient=recipient_email,
                        subject=f"第{report_number}回報告書",
                        body="",
                        attachment_path=merged_path,
                    )
                except Exception:
                    # Email delivery is best-effort for the desktop client.
                    pass

            self.status_queue.put(
                {
                    "type": "completed",
                    "status": "completed",
                    "message": "PDFの結合が完了しました。",
                    "progress_current": len(order),
                    "progress_total": len(order),
                    "progress_percent": 100,
                    "pdf_path": str(merged_path),
                    "report_number": report_number,
                }
            )
        except ConversionError as exc:
            self.status_queue.put(
                {
                    "type": "failed",
                    "message": f"PDF変換に失敗しました: {exc}",
                }
            )
        except Exception as exc:  # noqa: BLE001
            self.status_queue.put(
                {
                    "type": "failed",
                    "message": f"エラーが発生しました: {exc}",
                }
            )
        finally:
            try:
                shutil.rmtree(work_root, ignore_errors=True)
                zip_path.unlink(missing_ok=True)
            finally:
                _cleanup_data_directories()

    def _poll_status_queue(self) -> None:
        try:
            while True:
                payload = self.status_queue.get_nowait()
                view = self.views.get("status")
                if not isinstance(view, StatusView):
                    continue
                if payload["type"] == "status":
                    view.update_status(payload)
                elif payload["type"] == "completed":
                    view.update_status(payload)
                    view.show_completion(
                        Path(str(payload.get("pdf_path"))),
                        str(payload.get("report_number", "")),
                    )
                elif payload["type"] == "failed":
                    view.show_failure(str(payload.get("message", "")))
        except queue.Empty:
            pass
        finally:
            if self._processing_thread and self._processing_thread.is_alive():
                self.root.after(200, self._poll_status_queue)

    def open_output_file(self, path: Optional[Path]) -> None:
        if not path or not path.exists():
            messagebox.showwarning("警告", "PDFファイルが見つかりません。")
            return
        if os.name == "nt":  # pragma: no cover - platform specific
            os.startfile(path)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":  # pragma: no cover
            subprocess.run(["open", str(path)], check=False)
        else:  # pragma: no cover
            subprocess.run(["xdg-open", str(path)], check=False)

    def reveal_output_folder(self, path: Optional[Path]) -> None:
        if not path:
            return
        folder = path.parent
        if os.name == "nt":  # pragma: no cover
            subprocess.run(["explorer", str(folder)], check=False)
        elif sys.platform == "darwin":  # pragma: no cover
            subprocess.run(["open", str(folder)], check=False)
        else:  # pragma: no cover
            subprocess.run(["xdg-open", str(folder)], check=False)

    def run(self) -> None:
        self.root.mainloop()


def main() -> None:
    app = DesktopApp()
    app.run()


if __name__ == "__main__":  # pragma: no cover - manual execution entrypoint
    main()