"""
PPT 黒背景 → 白背景 変換ツール GUI
"""

import io
import os
import platform
import subprocess
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import traceback

# PyInstaller の windowed モードでは stdout/stderr が None になるため、
# AttributeError を防ぐために StringIO に差し替える
if sys.stdout is None:
    sys.stdout = io.StringIO()
if sys.stderr is None:
    sys.stderr = io.StringIO()

# エラーログの保存先（.appの場合はデスクトップに出力）
_LOG_PATH = os.path.join(os.path.expanduser("~"), "Desktop", "PPT変換ツール_エラーログ.txt")


def _write_log(text: str) -> None:
    """エラーをファイルに書き出す（GUIに表示されない場合の保険）"""
    try:
        with open(_LOG_PATH, "a", encoding="utf-8") as f:
            f.write(text + "\n")
    except Exception:
        pass


try:
    from main import convert_pptx
except Exception as _import_err:
    _write_log(f"[import error]\n{traceback.format_exc()}")
    raise


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PPT 黒背景→白背景 変換ツール")
        self.resizable(False, False)
        self._input_path = ""
        self._output_path = ""
        self._build_ui()
        self._center()

    def _build_ui(self):
        # ---- タイトルエリア ----
        top = tk.Frame(self, bg="#2B2B2B", pady=16)
        top.pack(fill="x")
        tk.Label(
            top,
            text="PPT 黒背景 → 白背景 変換ツール",
            font=("", 14, "bold"),
            fg="white",
            bg="#2B2B2B",
        ).pack()
        tk.Label(
            top,
            text="ダークテーマの PowerPoint を白背景用に自動変換します",
            font=("", 9),
            fg="#AAAAAA",
            bg="#2B2B2B",
        ).pack(pady=(2, 0))

        # ---- ファイル選択エリア ----
        body = tk.Frame(self, padx=24, pady=16)
        body.pack(fill="both")

        tk.Label(body, text="① 変換したいファイルを選択", font=("", 10, "bold"), anchor="w").pack(
            fill="x"
        )

        file_row = tk.Frame(body)
        file_row.pack(fill="x", pady=(6, 0))

        self._file_label = tk.Label(
            file_row,
            text="（ファイルが選択されていません）",
            fg="#999999",
            anchor="w",
            wraplength=320,
            justify="left",
        )
        self._file_label.pack(side="left", fill="x", expand=True)

        tk.Button(
            file_row,
            text="選択...",
            command=self._choose_file,
            width=8,
        ).pack(side="right")

        # ---- 変換ボタン ----
        tk.Label(body, text="② 変換を実行", font=("", 10, "bold"), anchor="w").pack(
            fill="x", pady=(16, 6)
        )

        self._convert_btn = tk.Button(
            body,
            text="変換する",
            font=("", 12, "bold"),
            width=20,
            height=2,
            state="disabled",
            command=self._start_convert,
        )
        self._convert_btn.pack()

        # ---- ステータス表示 ----
        self._status_var = tk.StringVar(value="")
        self._status_lbl = tk.Label(
            body,
            textvariable=self._status_var,
            wraplength=380,
            font=("", 10),
        )
        self._status_lbl.pack(pady=(12, 0))

        # 保存先を開くボタン（変換完了後に表示）
        self._open_btn = tk.Button(
            body,
            text="保存先フォルダを開く",
            command=self._open_folder,
        )

        # 下余白
        tk.Frame(body, height=8).pack()

    def _center(self):
        """ウィンドウを画面中央に配置する"""
        self.update_idletasks()
        w = self.winfo_reqwidth()
        h = self.winfo_reqheight()
        x = (self.winfo_screenwidth() - w) // 2
        y = (self.winfo_screenheight() - h) // 3
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _choose_file(self):
        """ファイル選択ダイアログを開く"""
        path = filedialog.askopenfilename(
            title="変換する PowerPoint ファイルを選択",
            filetypes=[
                ("PowerPoint ファイル", "*.pptx"),
                ("すべてのファイル", "*.*"),
            ],
        )
        if not path:
            return
        self._input_path = path
        self._file_label.config(text=os.path.basename(path), fg="#111111")
        self._convert_btn.config(state="normal")
        self._status_var.set("")
        self._open_btn.pack_forget()

    def _start_convert(self):
        """変換を実行する（シンプルな同期処理）"""
        self._convert_btn.config(state="disabled")
        self._open_btn.pack_forget()
        self._status_lbl.config(text="変換中...", fg="#E07800")
        # 「変換中...」を画面に即座に反映する
        self.update_idletasks()

        try:
            out = convert_pptx(self._input_path)
            self._on_success(out)
        except Exception as exc:
            detail = traceback.format_exc()
            _write_log(f"[変換エラー]\n{detail}")
            self._on_error(str(exc), detail)

    def _on_success(self, output_path: str):
        """変換成功時の処理"""
        self._output_path = output_path
        name = os.path.basename(output_path)
        self._status_lbl.config(text=f"✓ 変換完了！\n→ {name}", fg="#2A8C2A")
        self._convert_btn.config(state="normal")
        self._open_btn.pack(pady=(8, 0))
        # ダイアログで確実に完了を知らせる
        self.lift()
        self.focus_force()
        messagebox.showinfo("変換完了", f"変換が完了しました！\n\n保存先:\n{output_path}")

    def _on_error(self, message: str, detail: str = ""):
        """変換失敗時の処理"""
        # エラー内容を画面のラベルに直接表示する（ダイアログが隠れる問題の対策）
        short_msg = message[:80] + "..." if len(message) > 80 else message
        self._status_var.set(f"✗ エラー:\n{short_msg}")
        self._status_lbl.config(text=f"✗ エラー:\n{short_msg}", fg="#CC2222")
        self._convert_btn.config(state="normal")
        self.update()
        # ダイアログも表示する（最前面に持ってくる）
        self.lift()
        self.focus_force()
        messagebox.showerror(
            "変換エラー",
            f"変換中にエラーが発生しました:\n\n{message}\n\n"
            f"詳細はデスクトップの「PPT変換ツール_エラーログ.txt」を確認してください。",
        )

    def _open_folder(self):
        """保存先フォルダをOSのファイルマネージャーで開く"""
        folder = os.path.dirname(os.path.abspath(self._output_path))
        system = platform.system()
        if system == "Windows":
            subprocess.run(["explorer", folder])
        elif system == "Darwin":
            subprocess.run(["open", folder])
        else:
            subprocess.run(["xdg-open", folder])


if __name__ == "__main__":
    app = App()
    app.mainloop()
