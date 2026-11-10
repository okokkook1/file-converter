"""
app.py - 檔案轉換與合併工具 (GUI 圖形化操作介面)
"""

import os
import subprocess
import threading
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import file_converter as fc

SUPPORTED_IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".webp", ".tiff", ".tif"}

BACKGROUND = "#f1f5f9"
SURFACE = "#ffffff"
INK = "#1e293b"
MUTED = "#475569"
BORDER = "#cbd5e1"
ACCENT = "#1d4ed8"
SELECTED = "#dbeafe"
FONT = "Microsoft JhengHei UI"


class FileConverterApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("多功能檔案轉換與合併工具")
        self.root.geometry("880x700")
        self.root.minsize(720, 580)

        # 檔案清單數據 [(display_name, full_path), ...]
        self.file_items = []

        self._setup_style()
        self._build_ui()
        self._check_dependencies()

    def _setup_style(self):
        self.root.configure(background=BACKGROUND)
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure("TFrame", background=BACKGROUND)
        style.configure("Surface.TFrame", background=SURFACE)
        style.configure("TLabel", background=BACKGROUND, foreground=INK, font=(FONT, 10))
        style.configure("Title.TLabel", background=SURFACE, font=(FONT, 19, "bold"))
        style.configure("Section.TLabel", font=(FONT, 12, "bold"))
        style.configure("Muted.TLabel", foreground=MUTED)
        style.configure("Status.TLabel", background=SURFACE, foreground=MUTED)
        style.configure("TButton", background=SURFACE, foreground=INK, bordercolor=BORDER,
                        font=(FONT, 10), padding=(12, 8))
        style.map("TButton", background=[("disabled", BACKGROUND), ("pressed", SELECTED), ("active", "#eff6ff")],
                  foreground=[("disabled", "#64748b")])
        style.configure("Accent.TButton", background=ACCENT, foreground=SURFACE, bordercolor=ACCENT,
                        font=(FONT, 10, "bold"), padding=(14, 9))
        style.map("Accent.TButton", background=[("pressed", "#1e3a8a"), ("active", "#1e40af")])
        style.configure("TNotebook", background=BACKGROUND, bordercolor=BORDER)
        style.configure("TNotebook.Tab", background=BACKGROUND, foreground=INK,
                        font=(FONT, 10, "bold"), padding=(16, 8))
        style.map("TNotebook.Tab", background=[("selected", SELECTED), ("active", "#eff6ff")])

    def _check_dependencies(self):
        soffice = fc.find_libreoffice()
        if not soffice:
            self.set_status("警告: 未偵測到 LibreOffice，DOCX/ODT 互轉與 ODT 合併功能將受限。")
        else:
            self.set_status(f"就緒 (已連結 LibreOffice: {Path(soffice).name})")

    def _build_ui(self):
        top_frame = ttk.Frame(self.root, style="Surface.TFrame", padding=(20, 14))
        top_frame.pack(fill=tk.X)
        ttk.Label(top_frame, text="檔案轉換與排序合併", style="Title.TLabel").pack(anchor=tk.W)
        ttk.Label(top_frame, text="加入檔案、確認順序，再選擇轉換或合併方式。",
                  style="Status.TLabel").pack(anchor=tk.W, pady=(3, 0))

        bottom_frame = ttk.Frame(self.root, style="Surface.TFrame", padding=(20, 8))
        bottom_frame.pack(side=tk.BOTTOM, fill=tk.X)
        self.open_dir_btn = ttk.Button(bottom_frame, text="開啟最近輸出資料夾", state=tk.DISABLED,
                                       command=self.open_last_output_dir)
        self.open_dir_btn.pack(side=tk.RIGHT)
        self.last_output_dir = None
        self.status_var = tk.StringVar(value="就緒")
        self.status_lbl = ttk.Label(bottom_frame, textvariable=self.status_var, style="Status.TLabel")
        self.status_lbl.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 12))
        self.status_lbl.bind("<Configure>", lambda event: self.status_lbl.configure(wraplength=event.width))

        content_frame = ttk.Frame(self.root, padding=(20, 12, 20, 0))
        list_heading = ttk.Frame(content_frame)
        list_heading.pack(fill=tk.X, pady=(0, 8))
        ttk.Label(list_heading, text="1  檔案清單", style="Section.TLabel").pack(side=tk.LEFT)
        self.file_count = ttk.Label(list_heading, style="Muted.TLabel")
        self.file_count.pack(side=tk.LEFT, padx=10)

        controls = ttk.Frame(content_frame)
        controls.pack(fill=tk.X, pady=(0, 8))
        ttk.Button(controls, text="新增檔案…", style="Accent.TButton",
                   command=self.add_files).pack(side=tk.LEFT, padx=(0, 8))
        for attr, label, command in (
            ("move_up_btn", "上移", self.move_up),
            ("move_down_btn", "下移", self.move_down),
            ("remove_btn", "移除選取", self.remove_selected),
            ("clear_btn", "清空清單", self.clear_all),
        ):
            button = ttk.Button(controls, text=label, width=8, command=command)
            button.pack(side=tk.LEFT, padx=(0, 8))
            setattr(self, attr, button)

        list_frame = ttk.Frame(content_frame)
        list_frame.pack(fill=tk.BOTH, expand=True)
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        self.listbox = tk.Listbox(
            list_frame, selectmode=tk.SINGLE, exportselection=False,
            font=(FONT, 11), height=5, activestyle="none", relief=tk.FLAT,
            background=SURFACE, foreground=INK, selectbackground=SELECTED, selectforeground=INK,
            highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT,
        )
        vertical = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.listbox.yview)
        horizontal = ttk.Scrollbar(list_frame, orient=tk.HORIZONTAL, command=self.listbox.xview)
        self.listbox.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        self.listbox.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        self.empty_hint = ttk.Label(list_frame, text="清單目前是空的。選擇「新增檔案」即可開始。",
                                    style="Status.TLabel")
        self.listbox.bind("<<ListboxSelect>>", self._update_file_controls)

        action_frame = ttk.Frame(self.root, padding=(20, 12))
        action_frame.pack(side=tk.BOTTOM, fill=tk.X)
        ttk.Label(action_frame, text="2  選擇操作", style="Section.TLabel").pack(anchor=tk.W, pady=(0, 8))
        notebook = ttk.Notebook(action_frame)
        notebook.pack(fill=tk.X)
        self.action_buttons = []
        for name, actions in (
            ("格式轉換", (
                ("DOCX → ODT", self.convert_docx_to_odt_action),
                ("ODT → DOCX", self.convert_odt_to_docx_action),
                ("Word → PDF", self.convert_doc_to_pdf_action),
                ("PDF → Word", self.convert_pdf_to_docx_action),
                ("圖片 → PDF", self.convert_images_to_pdf_action),
                ("PDF → 逐頁圖片", self.convert_pdf_to_images_action),
            )),
            ("依序合併", (
                ("自動辨識合併", self.auto_merge_action),
                ("合併 PDF", self.merge_pdf_action),
                ("圖片合併成 PDF", self.merge_images_action),
                ("合併 DOCX", self.merge_docx_action),
                ("合併 ODT", self.merge_odt_action),
            )),
        ):
            tab = ttk.Frame(notebook, style="Surface.TFrame", padding=8)
            notebook.add(tab, text=name)
            for column in range(3):
                tab.columnconfigure(column, weight=1, uniform="actions")
            for index, (label, action) in enumerate(actions):
                row, column = divmod(index, 3)
                button = ttk.Button(tab, text=label, command=lambda action=action: self.run_async(action))
                button.grid(row=row, column=column, padx=4, pady=4, sticky="ew")
                self.action_buttons.append(button)
        content_frame.pack(fill=tk.BOTH, expand=True)
        self._update_file_controls()

    def _update_file_controls(self, event=None):
        count = len(self.file_items)
        selected = self.listbox.curselection()
        self.file_count.configure(text=f"{count} 個檔案")
        if count:
            self.empty_hint.place_forget()
        else:
            self.empty_hint.place(in_=self.listbox, relx=0.5, rely=0.5, anchor=tk.CENTER)
        for button, enabled in (
            (self.move_up_btn, selected and selected[0] > 0),
            (self.move_down_btn, selected and selected[0] + 1 < count),
            (self.remove_btn, selected),
            (self.clear_btn, count),
        ):
            button.state(["!disabled" if enabled else "disabled"])
        for button in self.action_buttons:
            button.state(["!disabled" if count else "disabled"])

    # ---------- 清單管理與排序 ----------

    def add_files(self):
        filetypes = [
            ("所有支援檔案", "*.docx;*.doc;*.odt;*.pdf;*.png;*.jpg;*.jpeg;*.bmp;*.webp;*.tiff"),
            ("Word 文件 (*.docx;*.doc)", "*.docx;*.doc"),
            ("ODT 文件 (*.odt)", "*.odt"),
            ("PDF 檔案 (*.pdf)", "*.pdf"),
            ("圖片檔案 (*.png;*.jpg;*.jpeg;*.bmp;*.webp)", "*.png;*.jpg;*.jpeg;*.bmp;*.webp"),
            ("所有檔案 (*.*)", "*.*")
        ]
        chosen = filedialog.askopenfilenames(title="選取檔案", filetypes=filetypes)
        if not chosen:
            return

        for f in chosen:
            p = str(Path(f).resolve())
            if p not in self.file_items:
                self.file_items.append(p)

        self._refresh_listbox()
        self.set_status(f"已加入 {len(chosen)} 個檔案，清單目前共 {len(self.file_items)} 個。")

    def remove_selected(self):
        sel = self.listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        del self.file_items[idx]
        self._refresh_listbox(min(idx, len(self.file_items) - 1) if self.file_items else None)
        self.set_status(f"已移除項目，目前剩餘 {len(self.file_items)} 個檔案。")

    def clear_all(self):
        if not self.file_items:
            return
        self.file_items.clear()
        self._refresh_listbox()
        self.set_status("已清空清單。")

    def move_up(self):
        sel = self.listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        if idx <= 0:
            return
        # 交換
        self.file_items[idx - 1], self.file_items[idx] = self.file_items[idx], self.file_items[idx - 1]
        self._refresh_listbox(idx - 1)

    def move_down(self):
        sel = self.listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        if idx >= len(self.file_items) - 1:
            return
        # 交換
        self.file_items[idx + 1], self.file_items[idx] = self.file_items[idx], self.file_items[idx + 1]
        self._refresh_listbox(idx + 1)

    def _refresh_listbox(self, selected=None):
        self.listbox.delete(0, tk.END)
        for i, path in enumerate(self.file_items, 1):
            name = Path(path).name
            self.listbox.insert(tk.END, f"  {i:02d}   {name}   ·   {Path(path).parent}")
        if selected is not None:
            self.listbox.selection_set(selected)
            self.listbox.activate(selected)
            self.listbox.see(selected)
        self._update_file_controls()

    def set_status(self, text: str):
        self.status_var.set(text)

    def set_last_output(self, out_path: str):
        folder = Path(out_path).parent if Path(out_path).is_file() else Path(out_path)
        self.last_output_dir = str(folder)
        self.open_dir_btn.config(state=tk.NORMAL)

    def open_last_output_dir(self):
        if self.last_output_dir and os.path.exists(self.last_output_dir):
            if os.name == "nt":
                os.startfile(self.last_output_dir)
            else:
                subprocess.Popen(["xdg-open", self.last_output_dir])

    def run_async(self, func):
        """在背景執行緒執行避免 GUI 凍結"""
        def wrapper():
            try:
                func()
            except Exception as e:
                self.set_status(f"錯誤: {e}")
                messagebox.showerror("執行錯誤", str(e))

        t = threading.Thread(target=wrapper, daemon=True)
        t.start()

    # ---------- 轉換動作 ----------

    def convert_docx_to_odt_action(self):
        docx_files = [p for p in self.file_items if Path(p).suffix.lower() == ".docx"]
        if not docx_files:
            messagebox.showwarning("提示", "清單中沒有選取任何 .docx 檔案。")
            return

        self.set_status("正在進行 DOCX -> ODT 轉換...")
        count = 0
        last_out = None
        for docx in docx_files:
            out = fc.convert_docx_to_odt(docx)
            last_out = out
            count += 1
        self.set_last_output(last_out)
        self.set_status(f"轉換完成！已成功將 {count} 個 DOCX 轉為 ODT。")
        messagebox.showinfo("完成", f"已成功將 {count} 個 DOCX 檔案轉換為 ODT！")

    def convert_odt_to_docx_action(self):
        odt_files = [p for p in self.file_items if Path(p).suffix.lower() == ".odt"]
        if not odt_files:
            messagebox.showwarning("提示", "清單中沒有選取任何 .odt 檔案。")
            return

        self.set_status("正在進行 ODT -> DOCX 轉換...")
        count = 0
        last_out = None
        for odt in odt_files:
            out = fc.convert_odt_to_docx(odt)
            last_out = out
            count += 1
        self.set_last_output(last_out)
        self.set_status(f"轉換完成！已成功將 {count} 個 ODT 轉為 DOCX。")
        messagebox.showinfo("完成", f"已成功將 {count} 個 ODT 檔案轉換為 DOCX！")

    def convert_doc_to_pdf_action(self):
        doc_files = [p for p in self.file_items if Path(p).suffix.lower() in (".doc", ".docx")]
        if not doc_files:
            messagebox.showwarning("提示", "清單中沒有選取任何 Word 檔案 (.doc 或 .docx)。")
            return

        self.set_status("正在進行 Word -> PDF 轉換...")
        count = 0
        last_out = None
        for doc in doc_files:
            out = fc.convert_doc_to_pdf(doc)
            last_out = out
            count += 1
        self.set_last_output(last_out)
        self.set_status(f"轉換完成！已成功將 {count} 個 Word 文件轉為 PDF。")
        messagebox.showinfo("完成", f"已成功將 {count} 個 Word 文件轉換為 PDF！")

    def convert_pdf_to_docx_action(self):
        pdf_files = [p for p in self.file_items if Path(p).suffix.lower() == ".pdf"]
        if not pdf_files:
            messagebox.showwarning("提示", "清單中沒有選取任何 .pdf 檔案。")
            return

        self.set_status("正在進行 PDF -> Word (.docx) 轉換...")
        count = 0
        last_out = None
        for pdf in pdf_files:
            out = fc.convert_pdf_to_docx(pdf)
            last_out = out
            count += 1
        self.set_last_output(last_out)
        self.set_status(f"轉換完成！已成功將 {count} 個 PDF 轉為 Word DOCX 文件。")
        messagebox.showinfo("完成", f"已成功將 {count} 個 PDF 檔案轉換為 Word DOCX 文件！")

    def convert_images_to_pdf_action(self):
        imgs = [p for p in self.file_items if Path(p).suffix.lower() in SUPPORTED_IMAGE_EXTS]
        if not imgs:
            messagebox.showwarning("提示", "清單中沒有選取任何支援的圖片檔案。")
            return

        # 若多張圖片，詢問是否合併成一個 PDF 或個別轉換
        if len(imgs) > 1:
            ans = messagebox.askyesnocancel("圖片轉 PDF", "選擇「是」：依清單順序合併為單一多頁 PDF\n選擇「否」：每張圖片各自轉為獨立 PDF")
            if ans is None:
                return
            if ans:
                self.merge_images_action()
                return

        self.set_status("正在將圖片轉換為 PDF...")
        last_out = None
        for img in imgs:
            target = str(Path(img).with_suffix(".pdf"))
            fc.convert_images_to_pdf([img], target)
            last_out = target

        self.set_last_output(last_out)
        self.set_status(f"轉換完成！共轉換 {len(imgs)} 張圖片為 PDF。")
        messagebox.showinfo("完成", f"已完成 {len(imgs)} 張圖片轉 PDF！")

    def convert_pdf_to_images_action(self):
        pdfs = [p for p in self.file_items if Path(p).suffix.lower() == ".pdf"]
        if not pdfs:
            messagebox.showwarning("提示", "清單中沒有選取任何 .pdf 檔案。")
            return

        self.set_status("正在將 PDF 每頁渲染為圖片...")
        total_pages = 0
        last_dir = None
        for pdf in pdfs:
            imgs = fc.convert_pdf_to_images(pdf)
            total_pages += len(imgs)
            if imgs:
                last_dir = str(Path(imgs[0]).parent)

        if last_dir:
            self.set_last_output(last_dir)
        self.set_status(f"PDF 轉圖片完成！共產生 {total_pages} 張頁面圖片。")
        messagebox.showinfo("完成", f"已成功將 {len(pdfs)} 個 PDF 檔案轉出 {total_pages} 張圖片！")

    # ---------- 依序合併動作 ----------

    def auto_merge_action(self):
        if not self.file_items:
            messagebox.showwarning("提示", "清單中沒有任何檔案。")
            return

        exts = [Path(p).suffix.lower() for p in self.file_items]
        first_ext = exts[0]

        if all(e == ".pdf" for e in exts):
            self.merge_pdf_action()
        elif all(e in SUPPORTED_IMAGE_EXTS for e in exts):
            self.merge_images_action()
        elif all(e in (".docx", ".doc") for e in exts):
            self.merge_docx_action()
        elif all(e == ".odt" for e in exts):
            self.merge_odt_action()
        else:
            messagebox.showerror(
                "格式不一致",
                "清單中的檔案類型不一致，無法自動合併。\n請確保清單內的檔案均為同一類型（同為 PDF、同為圖片、同為 Word 或同為 ODT）。"
            )

    def merge_pdf_action(self):
        pdfs = [p for p in self.file_items if Path(p).suffix.lower() == ".pdf"]
        if len(pdfs) < 2:
            messagebox.showwarning("提示", "合併 PDF 至少需要 2 個 PDF 檔案。")
            return

        default_name = "merged_output.pdf"
        save_path = filedialog.asksaveasfilename(
            title="儲存合併後的 PDF 檔案",
            initialfile=default_name,
            defaultextension=".pdf",
            filetypes=[("PDF 檔案 (*.pdf)", "*.pdf")]
        )
        if not save_path:
            return

        self.set_status("正在依序合併 PDF...")
        fc.merge_pdfs(pdfs, save_path)
        self.set_last_output(save_path)
        self.set_status(f"合併成功！已儲存至: {Path(save_path).name}")
        messagebox.showinfo("合併完成", f"已依指定順序成功合併 {len(pdfs)} 個 PDF 檔案！\n儲存路徑：\n{save_path}")

    def merge_images_action(self):
        imgs = [p for p in self.file_items if Path(p).suffix.lower() in SUPPORTED_IMAGE_EXTS]
        if not imgs:
            messagebox.showwarning("提示", "清單中沒有選取任何支援的圖片檔案。")
            return

        default_name = "images_combined.pdf"
        save_path = filedialog.asksaveasfilename(
            title="儲存圖片合併後的 PDF 檔案",
            initialfile=default_name,
            defaultextension=".pdf",
            filetypes=[("PDF 檔案 (*.pdf)", "*.pdf")]
        )
        if not save_path:
            return

        self.set_status("正在將圖片依序合併為 PDF...")
        fc.merge_images(imgs, save_path)
        self.set_last_output(save_path)
        self.set_status(f"合併成功！已儲存至: {Path(save_path).name}")
        messagebox.showinfo("合併完成", f"已依指定順序將 {len(imgs)} 張圖片合併為單一多頁 PDF！\n儲存路徑：\n{save_path}")

    def merge_docx_action(self):
        docs = [p for p in self.file_items if Path(p).suffix.lower() in (".docx", ".doc")]
        if len(docs) < 2:
            messagebox.showwarning("提示", "合併 Word 文件至少需要 2 個 Word 檔案 (.docx 或 .doc)。")
            return

        default_name = "merged_output.docx"
        save_path = filedialog.asksaveasfilename(
            title="儲存合併後的 Word 文件",
            initialfile=default_name,
            defaultextension=".docx",
            filetypes=[("Word 文件 (*.docx)", "*.docx"), ("舊版 Word 文件 (*.doc)", "*.doc")]
        )
        if not save_path:
            return

        self.set_status("正在依序合併 Word 文件...")
        fc.merge_docx(docs, save_path)
        self.set_last_output(save_path)
        self.set_status(f"合併成功！已儲存至: {Path(save_path).name}")
        messagebox.showinfo("合併完成", f"已依指定順序成功合併 {len(docs)} 個 Word 文件！\n儲存路徑：\n{save_path}")

    def merge_odt_action(self):
        odts = [p for p in self.file_items if Path(p).suffix.lower() == ".odt"]
        if len(odts) < 2:
            messagebox.showwarning("提示", "合併 ODT 至少需要 2 個 ODT 檔案。")
            return

        default_name = "merged_output.odt"
        save_path = filedialog.asksaveasfilename(
            title="儲存合併後的 ODT 檔案",
            initialfile=default_name,
            defaultextension=".odt",
            filetypes=[("ODT 文件 (*.odt)", "*.odt")]
        )
        if not save_path:
            return

        self.set_status("正在依序合併 ODT...")
        fc.merge_odt(odts, save_path)
        self.set_last_output(save_path)
        self.set_status(f"合併成功！已儲存至: {Path(save_path).name}")
        messagebox.showinfo("合併完成", f"已依指定順序成功合併 {len(odts)} 個 ODT 文件！\n儲存路徑：\n{save_path}")


def main():
    root = tk.Tk()
    app = FileConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
