#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Split PDF
Code with ChatGPT5
必須インストール
必要パッケージ:  pip install pypdf tkinterdnd2
Python: 3.11+ 推奨（Windows 11 動作想定）
"""

import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pypdf import PdfReader, PdfWriter
from copy import copy
import platform

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
except ImportError:
    # messagebox は Tk() 未初期化だと動かないため、一時ウィンドウを作って表示する
    _root = tk.Tk()
    _root.withdraw()
    messagebox.showerror("エラー", "tkinterdnd2 が必要です。\n\nインストール: pip install tkinterdnd2")
    raise


def parse_page_range(range_str, max_pages):
    """'1-3,5,7-9' のような文字列をページ番号リスト（0始まり）に変換"""
    pages = set()
    errors = []
    if not range_str.strip():
        return list(range(max_pages)), errors  # 空なら全ページ
    for part in range_str.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            if part.count("-") != 1:
                errors.append(part)
                continue
            start_str, end_str = part.split("-")
            if not start_str or not end_str:
                errors.append(part)
                continue
            try:
                start = int(start_str)
                end = int(end_str)
            except ValueError:
                errors.append(part)
                continue
            if start < 1 or end < 1 or start > end:
                errors.append(part)
                continue
            if start > max_pages:
                continue
            pages.update(range(start - 1, min(end, max_pages)))
        else:
            try:
                p = int(part)
            except ValueError:
                errors.append(part)
                continue
            if 1 <= p <= max_pages:
                pages.add(p - 1)
            else:
                errors.append(part)
    return sorted(pages), errors


def crop_page(original_page, lower_left, upper_right):
    """ページをコピーして指定座標でトリミング"""
    new_page = copy(original_page)
    new_page.cropbox.lower_left = lower_left
    new_page.cropbox.upper_right = upper_right
    return new_page


def parse_split_ratio(ratio_str):
    """分割ライン位置(%)を 0.0-1.0 に変換"""
    if not ratio_str.strip():
        return 0.5, None
    try:
        ratio = float(ratio_str)
    except ValueError:
        return None, "分割ライン位置は数値で入力してください。"
    if ratio <= 0 or ratio >= 100:
        return None, "分割ライン位置は 1〜99 の範囲で指定してください。"
    return ratio / 100.0, None


def split_half_auto(input_path, output_path, page_range="", split_ratio=0.5):
    reader = PdfReader(input_path)
    writer = PdfWriter()

    target_pages, errors = parse_page_range(page_range, len(reader.pages))
    if errors:
        raise ValueError(f"ページ範囲が不正です: {', '.join(errors)}")
    target_set = set(target_pages)

    for idx, page in enumerate(reader.pages):
        if page_range.strip() and idx not in target_set:
            writer.add_page(page)
            continue
        width = float(page.mediabox.width)
        height = float(page.mediabox.height)

        delta = abs(split_ratio - 0.5)
        if width >= height:
            # 横向き → 左右に分割
            split_x = width * split_ratio
            adjust = width * delta
            if split_ratio >= 0.5:
                # 左が大きい → 左の外側をカット、右の外側に余白
                left = crop_page(page, (adjust, 0), (split_x, height))
                right = crop_page(page, (split_x, 0), (width + adjust, height))
            else:
                # 右が大きい → 右の外側をカット、左の外側に余白
                left = crop_page(page, (-adjust, 0), (split_x, height))
                right = crop_page(page, (split_x, 0), (width - adjust, height))
            writer.add_page(left)
            writer.add_page(right)
        else:
            # 縦向き → 上下に分割
            split_y = height * split_ratio
            adjust = height * delta
            if split_ratio >= 0.5:
                # 下が大きい → 下の外側をカット、上の外側に余白
                top = crop_page(page, (0, split_y), (width, height + adjust))
                bottom = crop_page(page, (0, adjust), (width, split_y))
            else:
                # 上が大きい → 上の外側をカット、下の外側に余白
                top = crop_page(page, (0, split_y), (width, height - adjust))
                bottom = crop_page(page, (0, -adjust), (width, split_y))
            writer.add_page(top)
            writer.add_page(bottom)

    with open(output_path, "wb") as f:
        writer.write(f)


def process_pdf(input_path):
    if not os.path.exists(input_path):
        messagebox.showerror("エラー", f"ファイルが見つかりません:\n{input_path}")
        return

    page_range = entry_range.get()
    auto_open = var_auto_open.get()
    split_ratio, ratio_err = parse_split_ratio(entry_split_ratio.get())
    if ratio_err:
        messagebox.showerror("エラー", ratio_err)
        return

    base, ext = os.path.splitext(input_path)
    output_path = base + "_split.pdf"
    try:
        split_half_auto(input_path, output_path, page_range, split_ratio)
        messagebox.showinfo("完了", f"変換完了:\n{output_path}")
        if auto_open:
            if platform.system() == "Windows":
                os.startfile(output_path)
            elif platform.system() == "Darwin":
                os.system(f"open '{output_path}'")
            else:
                os.system(f"xdg-open '{output_path}'")
    except Exception as e:
        messagebox.showerror("エラー", str(e))


def select_file():
    filepath = filedialog.askopenfilename(
        filetypes=[("PDF files", "*.pdf")],
        title="PDFを選択"
    )
    if filepath:
        process_pdf(filepath)


def drop(event):
    file_path = event.data.strip()
    if file_path.startswith("{") and file_path.endswith("}"):
        file_path = file_path[1:-1]
    if file_path.lower().endswith(".pdf"):
        process_pdf(file_path)
    else:
        messagebox.showerror("エラー", "PDFファイルをドロップしてください。")


# GUI作成
root = TkinterDnD.Tk()
root.title("PDF分割ツール")
root.geometry("420x300")
root.minsize(380, 280)
root.configure(bg="#f0f0f0")

style = ttk.Style()
style.theme_use("vista" if platform.system() == "Windows" else "clam")
default_font = ("Meiryo UI", 10)
style.configure(".", font=default_font)
style.configure("TLabel", font=default_font)
style.configure("TButton", font=default_font)
style.configure("TCheckbutton", font=default_font)

# ドラッグ＆ドロップエリア
drop_area = tk.Label(
    root,
    text="ここにPDFをドラッグ＆ドロップ\nまたは下のボタンから選択",
    font=("Meiryo UI", 11),
    fg="#555555",
    bg="#e8e8e8",
    relief="groove",
    borderwidth=2,
)
drop_area.pack(expand=True, fill="both", padx=12, pady=(12, 8))

# 設定エリア
frame_settings = ttk.Frame(root)
frame_settings.pack(fill="x", padx=12, pady=4)

# ページ範囲入力
ttk.Label(frame_settings, text="ページ範囲:").grid(row=0, column=0, sticky="e", padx=(0, 6))
entry_range = ttk.Entry(frame_settings, width=18)
entry_range.grid(row=0, column=1, sticky="w")
ttk.Label(frame_settings, text="例: 1-3,5,7-9", foreground="#000000").grid(row=0, column=2, padx=(8, 0))

# 分割ライン位置入力
ttk.Label(frame_settings, text="分割位置(%):").grid(row=1, column=0, sticky="e", padx=(0, 6), pady=4)
entry_split_ratio = ttk.Entry(frame_settings, width=8)
entry_split_ratio.insert(0, "50")
entry_split_ratio.grid(row=1, column=1, sticky="w", pady=4)
ttk.Label(frame_settings, text="例: 50", foreground="#000000").grid(row=1, column=2, padx=(8, 0))

# 下部エリア
frame_bottom = ttk.Frame(root)
frame_bottom.pack(fill="x", padx=12, pady=(4, 12))

var_auto_open = tk.BooleanVar(value=True)
chk_auto_open = ttk.Checkbutton(frame_bottom, text="分割後にファイルを開く", variable=var_auto_open)
chk_auto_open.pack(side="left")

btn = ttk.Button(frame_bottom, text="PDFを選択", command=select_file)
btn.pack(side="right")

# ドラッグ＆ドロップ対応
drop_area.drop_target_register(DND_FILES)
drop_area.dnd_bind("<<Drop>>", drop)

root.mainloop()
