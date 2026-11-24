import sys

import tkinter as tk
from tkinter import filedialog

ARCHIVE_TYPES = [("压缩包", "*.rar *.zip *.7z *.tar *.gz"), ("所有文件", "*.*")]


def main():
    kind = sys.argv[1] if len(sys.argv) > 1 else "dir"
    result = {"path": ""}

    if kind == "docx":
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        result["path"] = filedialog.askopenfilename(
            title="选择资料交接单模板",
            filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")]) or ""
        root.destroy()
        sys.stdout.reconfigure(encoding="utf-8")
        print(result["path"])
        return

    root = tk.Tk()
    root.title("选择资料文件夹")
    root.geometry("430x210")
    root.attributes("-topmost", True)
    root.resizable(False, False)

    tk.Label(root, text="选文件夹：会递归读取里面所有子文件夹和文件，图片也会 OCR",
             pady=14, wraplength=400).pack()

    box = tk.Frame(root)
    box.pack(pady=4)

    def pick_dir():
        result["path"] = filedialog.askdirectory(
            title="选择文件夹（会读取所有子文件夹、Word/Excel/PDF/图片）", parent=root) or ""
        root.destroy()

    def pick_zip():
        result["path"] = filedialog.askopenfilename(
            title="选择压缩包", parent=root, filetypes=ARCHIVE_TYPES) or ""
        root.destroy()

    tk.Button(box, text="选择文件夹（推荐）", width=18, command=pick_dir).pack(side="left", padx=8)
    tk.Button(box, text="选压缩包", width=10, command=pick_zip).pack(side="left", padx=8)
    tk.Button(root, text="取消", width=10, command=root.destroy).pack(pady=14)

    root.mainloop()
    sys.stdout.reconfigure(encoding="utf-8")
    print(result["path"])


if __name__ == "__main__":
    main()
