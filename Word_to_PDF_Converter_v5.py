import os
import sys
import re
import win32com.client
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
from collections import Counter
import pymupdf
from datetime import datetime


# --- Sorting Key Functions ---
def standard_sort_key(s):
    return os.path.basename(s).lower()


def natural_sort_key(s):
    s_base = os.path.basename(s)
    s_stem, s_ext = os.path.splitext(s_base)

    stem_key = [
        int(text) if text.isdigit() else text.lower()
        for text in re.split(r"([0-9]+)", s_stem)
    ]

    return stem_key, s_ext.lower()


def revision_aware_sort_key(s):
    s_base = os.path.basename(s)
    s_stem, s_ext = os.path.splitext(s_base)

    match = re.match(r"^(.*?)((?:[-_][a-zA-Z0-9]+)*)$", s_stem)

    if match:
        base_part, rev_part = match.groups()
    else:
        base_part, rev_part = s_stem, ""

    base_key = [
        int(text) if text.isdigit() else text.lower()
        for text in re.split(r"([0-9]+)", base_part)
    ]

    return base_key, rev_part.lower()


# --- Scrollable Frame Helper ---
class ScrollableFrame(tk.Frame):
    def __init__(self, container, *args, **kwargs):
        super().__init__(container, *args, **kwargs)

        self.canvas = tk.Canvas(
            self,
            borderwidth=0,
            highlightthickness=0
        )

        scrollbar = ttk.Scrollbar(
            self,
            orient="vertical",
            command=self.canvas.yview
        )

        self.scrollable_frame = tk.Frame(self.canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all")
            )
        )

        self.canvas.create_window(
            (0, 0),
            window=self.scrollable_frame,
            anchor="nw"
        )

        self.canvas.configure(yscrollcommand=scrollbar.set)

        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.canvas.bind("<Enter>", self._bind_mousewheel)
        self.canvas.bind("<Leave>", self._unbind_mousewheel)

    def _bind_mousewheel(self, event):
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _unbind_mousewheel(self, event):
        self.canvas.unbind_all("<MouseWheel>")

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(
            int(-1 * (event.delta / 120)),
            "units"
        )


# --- Excel Tab Selection Dialog ---
class ExcelTabSelectionDialog(tk.Toplevel):
    def __init__(self, parent, excel_file_info):
        super().__init__(parent)

        self.title("Select Excel Tabs for Conversion")
        self.geometry("450x450")
        self.resizable(False, False)

        self.excel_file_info = excel_file_info
        self.result = {}
        self.vars = {}

        self.transient(parent)
        self.grab_set()

        lbl = tk.Label(
            self,
            text="Select the tabs you want to convert to PDF:",
            font=("Arial", 10, "bold")
        )
        lbl.pack(pady=10, padx=10, anchor="w")

        scroll_frame = ScrollableFrame(self)
        scroll_frame.pack(fill="both", expand=True, padx=10, pady=5)

        inner_frame = scroll_frame.scrollable_frame

        for file, sheets in excel_file_info.items():
            is_single_tab = len(sheets) == 1
            file_var = tk.BooleanVar(value=is_single_tab)

            self.vars[file] = {
                "var": file_var,
                "sheets": {}
            }

            cb = tk.Checkbutton(
                inner_frame,
                text=file,
                variable=file_var,
                font=("Arial", 9, "bold"),
                command=lambda f=file: self.toggle_file(f)
            )
            cb.pack(anchor="w", padx=5, pady=(5, 0))

            for sheet in sheets:
                sheet_var = tk.BooleanVar(value=is_single_tab)
                self.vars[file]["sheets"][sheet] = sheet_var

                scb = tk.Checkbutton(
                    inner_frame,
                    text=sheet,
                    variable=sheet_var,
                    command=lambda f=file: self.check_file_state(f)
                )
                scb.pack(anchor="w", padx=25)

        scroll_lbl = tk.Label(
            self,
            text="↓ Scroll down for more files/tabs if needed ↓",
            font=("Arial", 8, "italic"),
            fg="gray"
        )
        scroll_lbl.pack(pady=(5, 0))

        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=10)

        tk.Button(
            btn_frame,
            text="Confirm Selection",
            width=20,
            command=self.on_confirm
        ).pack()

    def toggle_file(self, file):
        state = self.vars[file]["var"].get()

        for sheet_var in self.vars[file]["sheets"].values():
            sheet_var.set(state)

    def check_file_state(self, file):
        all_checked = all(
            var.get()
            for var in self.vars[file]["sheets"].values()
        )

        self.vars[file]["var"].set(all_checked)

    def on_confirm(self):
        for file, data in self.vars.items():
            selected_sheets = [
                sheet
                for sheet, var in data["sheets"].items()
                if var.get()
            ]

            if selected_sheets:
                self.result[file] = selected_sheets

        self.destroy()


# --- Contract Checks Selection Dialog ---
class ContractChecksOptionsDialog(tk.Toplevel):
    def __init__(self, parent, callback):
        super().__init__(parent)

        self.title("Select Checks to Run")
        self.geometry("480x620")
        self.resizable(False, False)

        self.callback = callback

        self.transient(parent)
        self.grab_set()

        tk.Label(
            self,
            text="Select optional contract checks:",
            font=("Arial", 10, "bold")
        ).pack(pady=10, padx=10, anchor="w")

        self.hf_var = tk.BooleanVar(value=True)

        tk.Checkbutton(
            self,
            text="Header/Footer Contract Number Check",
            variable=self.hf_var
        ).pack(anchor="w", padx=20)

        self.font_var = tk.BooleanVar(value=True)

        tk.Checkbutton(
            self,
            text="Font Uniformity Check (Word files)",
            variable=self.font_var
        ).pack(anchor="w", padx=20)

        self.keyword_var = tk.BooleanVar(value=False)

        tk.Checkbutton(
            self,
            text="Keyword / Guide Text Check",
            variable=self.keyword_var,
            command=self.toggle_keyword_options
        ).pack(anchor="w", padx=20, pady=(12, 0))

        self.keyword_options_frame = tk.Frame(self)
        self.keyword_options_frame.pack(
            anchor="w",
            padx=45,
            pady=(5, 10),
            fill=tk.X
        )

        tk.Label(
            self.keyword_options_frame,
            text="Contract Type:"
        ).pack(anchor="w")

        self.keyword_contract_type_var = tk.StringVar(
            value="Design-Bid-Build Contract"
        )

        self.keyword_contract_type_combo = ttk.Combobox(
            self.keyword_options_frame,
            textvariable=self.keyword_contract_type_var,
            values=[
                "Design-Bid-Build Contract",
                "EPC Contract"
            ],
            state="disabled",
            width=40
        )
        self.keyword_contract_type_combo.pack(anchor="w", pady=(0, 8))

        self.keyword_contract_type_combo.bind(
            "<<ComboboxSelected>>",
            self.update_keyword_role_dropdown
        )

        tk.Label(
            self.keyword_options_frame,
            text="Project Role:"
        ).pack(anchor="w")

        self.keyword_role_var = tk.StringVar(value="Owner")

        self.keyword_role_combo = ttk.Combobox(
            self.keyword_options_frame,
            textvariable=self.keyword_role_var,
            values=[
                "Owner",
                "Contractor",
                "Engineer"
            ],
            state="disabled",
            width=40
        )
        self.keyword_role_combo.pack(anchor="w", pady=(0, 10))

        tk.Label(
            self.keyword_options_frame,
            text="Additional keywords to flag, one per line, optional:"
        ).pack(anchor="w")

        self.keyword_text = tk.Text(
            self.keyword_options_frame,
            height=8,
            width=46,
            state="disabled"
        )
        self.keyword_text.pack(anchor="w")

        tk.Label(
            self,
            text=(
                "Automatically flags blue-box-style headings, "
                "bracketed text, parenthetical text, highlighted text, "
                "shaded text, and blue font text."
            ),
            font=("Arial", 8, "italic"),
            fg="gray",
            wraplength=410,
            justify=tk.LEFT
        ).pack(anchor="w", padx=45, pady=(0, 10))

        tk.Label(
            self,
            text="Review mode after scan:",
            font=("Arial", 9, "bold")
        ).pack(anchor="w", padx=45)

        self.review_mode_var = tk.StringVar(value="Report only")
        self.review_mode_combo = ttk.Combobox(
            self,
            textvariable=self.review_mode_var,
            values=[
                "Report only",
                "Export report to .txt",
                "Export report and apply review comments to Word files"
            ],
            state="readonly",
            width=48
        )
        self.review_mode_combo.pack(anchor="w", padx=45, pady=(0, 10))

        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=(10, 20))

        tk.Button(
            btn_frame,
            text="Run Selected Checks",
            width=30,
            command=self.on_run
        ).pack(pady=(0, 5))

        tk.Button(
            btn_frame,
            text="Cancel and return to main window",
            width=30,
            command=self.destroy
        ).pack()

    def toggle_keyword_options(self):
        if self.keyword_var.get():
            self.keyword_contract_type_combo.config(state="readonly")
            self.keyword_role_combo.config(state="readonly")
            self.keyword_text.config(state="normal")
            self.update_keyword_role_dropdown()
        else:
            self.keyword_contract_type_combo.config(state="disabled")
            self.keyword_role_combo.config(state="disabled")
            self.keyword_text.config(state="disabled")

    def update_keyword_role_dropdown(self, event=None):
        contract_type = self.keyword_contract_type_var.get()

        if contract_type == "Design-Bid-Build Contract":
            roles = [
                "Owner",
                "Contractor",
                "Engineer"
            ]
        elif contract_type == "EPC Contract":
            roles = [
                "Contractor",
                "SubContractor"
            ]
        else:
            roles = []

        self.keyword_role_combo["values"] = roles

        if roles:
            self.keyword_role_var.set(roles[0])
        else:
            self.keyword_role_var.set("")

    def on_run(self):
        keywords = [
            line.strip()
            for line in self.keyword_text.get("1.0", tk.END).splitlines()
            if line.strip()
        ]

        options = {
            "hf": self.hf_var.get(),
            "font": self.font_var.get(),
            "keyword": self.keyword_var.get(),
            "keyword_contract_type": self.keyword_contract_type_var.get(),
            "keyword_role": self.keyword_role_var.get(),
            "keywords_to_check": keywords,
            "review_mode": self.review_mode_var.get()
        }

        self.destroy()
        self.callback(options)


# --- Contract Checks Results Dialog ---
class ChecksResultsDialog(tk.Toplevel):
    def __init__(self, parent, results_dict):
        super().__init__(parent)

        self.title("Contract Checks Summary")
        self.geometry("720x620")

        self.results_dict = results_dict

        self.transient(parent)
        self.grab_set()

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(
            fill=tk.BOTH,
            expand=True,
            padx=10,
            pady=(10, 5)
        )

        if "hf" in results_dict:
            self.build_tab("Header/Footer", results_dict["hf"])

        if "font" in results_dict:
            header_text = (
                f"Determined Standard Font: "
                f"{results_dict.get('font_meta', 'Unknown')}"
            )
            self.build_tab(
                "Font Uniformity",
                results_dict["font"],
                header_text
            )

        if "keyword" in results_dict:
            keyword_header = results_dict.get("keyword_meta", None)
            self.build_tab(
                "Keyword / Guide Flags",
                results_dict["keyword"],
                keyword_header
            )

        warning_text = (
            "These checks are rudimentary. Any flagged issues should be "
            "double checked by the user. If edits are needed, make them in the "
            "local native Word or Excel files. Close Word and Excel before "
            "returning to this converter and starting PDF conversion."
        )

        tk.Label(
            self,
            text=warning_text,
            font=("Arial", 8, "italic"),
            fg="firebrick",
            wraplength=680,
            justify=tk.LEFT
        ).pack(padx=15, pady=(5, 5))

        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=(5, 10))

        tk.Button(
            btn_frame,
            text="Export Summaries to .txt",
            width=25,
            command=self.export_to_txt
        ).pack(side=tk.LEFT, padx=10)

        tk.Button(
            btn_frame,
            text="Close",
            width=15,
            command=self.destroy
        ).pack(side=tk.LEFT, padx=10)

    def build_tab(self, title, items, header_text=None):
        frame = tk.Frame(self.notebook)
        self.notebook.add(frame, text=title)

        if header_text:
            tk.Label(
                frame,
                text=header_text,
                font=("Arial", 10, "bold"),
                fg="navy",
                wraplength=660,
                justify=tk.LEFT
            ).pack(pady=(10, 0), padx=10, anchor="w")

        scroll_frame = ScrollableFrame(frame)
        scroll_frame.pack(fill="both", expand=True, padx=5, pady=5)

        inner_frame = scroll_frame.scrollable_frame

        for item in items:
            row_frame = tk.Frame(inner_frame)
            row_frame.pack(fill=tk.X, pady=5)

            if item["status"] == "ok":
                icon_lbl = tk.Label(
                    row_frame,
                    text="✓",
                    fg="green",
                    font=("Arial", 12, "bold"),
                    width=3
                )

                desc = tk.Label(
                    row_frame,
                    text=item["file"],
                    font=("Arial", 9),
                    wraplength=640,
                    justify=tk.LEFT
                )

                icon_lbl.pack(side=tk.LEFT, anchor="n")
                desc.pack(side=tk.LEFT, anchor="n", pady=(2, 0))

            else:
                icon_lbl = tk.Label(
                    row_frame,
                    text="⚠",
                    fg="goldenrod",
                    font=("Arial", 12, "bold"),
                    width=3
                )
                icon_lbl.pack(side=tk.LEFT, anchor="n")

                text_frame = tk.Frame(row_frame)
                text_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)

                tk.Label(
                    text_frame,
                    text=item["file"],
                    font=("Arial", 9, "bold"),
                    wraplength=640,
                    justify=tk.LEFT
                ).pack(anchor="w")

                tk.Label(
                    text_frame,
                    text="Issues: " + " | ".join(item["issues"]),
                    font=("Arial", 8),
                    fg="firebrick",
                    wraplength=640,
                    justify=tk.LEFT
                ).pack(anchor="w")

    def export_to_txt(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text Files", "*.txt")],
            title="Save Summary Report"
        )

        if not file_path:
            return

        try:
            with open(file_path, "w", encoding="utf-8") as f:
                f.write("CONTRACT CHECKS SUMMARY REPORT\n")
                f.write("=" * 40 + "\n\n")

                checks = [
                    ("hf", "HEADER / FOOTER ISSUES"),
                    ("font", "FONT UNIFORMITY ISSUES"),
                    ("keyword", "KEYWORD / GUIDE FLAGS")
                ]

                for key, title in checks:
                    if key in self.results_dict:
                        f.write(f"--- {title} ---\n")

                        if key == "font" and "font_meta" in self.results_dict:
                            f.write(
                                f"Determined Standard Font: "
                                f"{self.results_dict['font_meta']}\n\n"
                            )

                        if key == "keyword" and "keyword_meta" in self.results_dict:
                            f.write(f"{self.results_dict['keyword_meta']}\n\n")

                        issues_found = False

                        for item in self.results_dict[key]:
                            if item["status"] == "warning":
                                issues_found = True
                                f.write(f"[!] {item['file']}\n")

                                for issue in item["issues"]:
                                    f.write(f"    - {issue}\n")

                        if not issues_found:
                            f.write("No issues found in this category.\n")

                        f.write("\n")

            messagebox.showinfo(
                "Export Successful",
                f"Summary saved to:\n{file_path}"
            )

        except Exception as e:
            messagebox.showerror(
                "Export Error",
                f"Failed to save text file: {e}"
            )


# --- Reorder Dialog Class ---
class ReorderDialog(simpledialog.Dialog):
    def __init__(self, parent, file_list):
        self.file_list = file_list
        self.result_list = None
        self._move_job = None

        super().__init__(parent, title="Reorder Files for Combined PDF")

    def body(self, master):
        controls_frame = tk.Frame(master)
        controls_frame.pack(pady=5, padx=5, fill=tk.X)

        sort_label = tk.Label(controls_frame, text="Sort Method:")
        sort_label.grid(row=0, column=0, padx=(0, 5), sticky="w")

        self.sort_method_var = tk.StringVar()

        sort_options = [
            "Revision Sort",
            "Natural Sort (Default)",
            "Standard Sort (A-Z)"
        ]

        self.sort_combo = ttk.Combobox(
            controls_frame,
            textvariable=self.sort_method_var,
            values=sort_options,
            state="readonly",
            width=25
        )
        self.sort_combo.grid(row=0, column=1, sticky="ew")
        self.sort_combo.current(1)
        self.sort_combo.bind("<<ComboboxSelected>>", self.apply_view_rules)

        self.move_covers_var = tk.BooleanVar()

        cover_check = tk.Checkbutton(
            controls_frame,
            text="Auto-move Covers to Ends",
            variable=self.move_covers_var,
            command=self.apply_view_rules
        )
        cover_check.grid(
            row=1,
            column=0,
            columnspan=3,
            pady=(5, 0),
            sticky="w"
        )

        controls_frame.columnconfigure(1, weight=1)

        list_frame = tk.Frame(master)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.listbox = tk.Listbox(
            list_frame,
            yscrollcommand=scrollbar.set,
            selectmode=tk.SINGLE,
            width=80,
            height=25
        )
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar.config(command=self.listbox.yview)

        self.apply_view_rules()

        btn_frame = tk.Frame(master)
        btn_frame.pack(pady=10)

        up_button = tk.Button(btn_frame, text="▲ Move Up", width=15)
        up_button.pack(side=tk.LEFT, padx=5)
        up_button.bind("<ButtonPress-1>", lambda e: self.start_move("up"))
        up_button.bind("<ButtonRelease-1>", lambda e: self.stop_move())

        down_button = tk.Button(btn_frame, text="▼ Move Down", width=15)
        down_button.pack(side=tk.LEFT, padx=5)
        down_button.bind("<ButtonPress-1>", lambda e: self.start_move("down"))
        down_button.bind("<ButtonRelease-1>", lambda e: self.stop_move())

        return self.listbox

    def apply_view_rules(self, event=None):
        method = self.sort_method_var.get()

        if "Revision" in method:
            sort_key = revision_aware_sort_key
        elif "Natural" in method:
            sort_key = natural_sort_key
        else:
            sort_key = standard_sort_key

        self.file_list.sort(key=sort_key)

        if self.move_covers_var.get():
            front_covers = []
            back_covers = []
            middle_files = []

            for f in self.file_list:
                fname_lower = os.path.basename(f).lower()

                if "front cover" in fname_lower:
                    front_covers.append(f)
                elif "back cover" in fname_lower:
                    back_covers.append(f)
                else:
                    middle_files.append(f)

            self.file_list = front_covers + middle_files + back_covers

        self._update_listbox_ui()

    def _update_listbox_ui(self):
        sel_idx = self.listbox.curselection()

        self.listbox.delete(0, tk.END)

        for path in self.file_list:
            self.listbox.insert(tk.END, os.path.basename(path))

        if sel_idx and len(sel_idx) > 0 and sel_idx[0] < self.listbox.size():
            self.listbox.selection_set(sel_idx[0])

    def start_move(self, direction):
        if self._move_job:
            self.after_cancel(self._move_job)

        if direction == "up":
            self.move_up()
        else:
            self.move_down()

        self._move_job = self.after(
            100,
            lambda: self.start_move(direction)
        )

    def stop_move(self):
        if self._move_job:
            self.after_cancel(self._move_job)

        self._move_job = None

    def move_up(self):
        sel_idx = self.listbox.curselection()

        if not sel_idx or sel_idx[0] == 0:
            return

        idx = sel_idx[0]

        self.file_list.insert(idx - 1, self.file_list.pop(idx))

        text = self.listbox.get(idx)
        self.listbox.delete(idx)
        self.listbox.insert(idx - 1, text)

        self.listbox.selection_set(idx - 1)
        self.listbox.see(idx - 1)

    def move_down(self):
        sel_idx = self.listbox.curselection()

        if not sel_idx or sel_idx[0] == self.listbox.size() - 1:
            return

        idx = sel_idx[0]

        self.file_list.insert(idx + 1, self.file_list.pop(idx))

        text = self.listbox.get(idx)
        self.listbox.delete(idx)
        self.listbox.insert(idx + 1, text)

        self.listbox.selection_set(idx + 1)
        self.listbox.see(idx + 1)

    def apply(self):
        self.result_list = self.file_list


# --- Main Application GUI ---
class BatchConverterApp:
    def __init__(self, root):
        self.root = root

        self.root.title("BMcD Batch Converter")
        self.root.geometry("600x580")
        self.root.resizable(False, False)

        self.source_folder = ""
        self.all_files = []
        self.excel_tabs_to_convert = {}

        top_frame = tk.Frame(self.root)
        top_frame.pack(pady=10, padx=10, fill=tk.X)

        self.select_btn = tk.Button(
            top_frame,
            text="Select Folder with Word/Excel/PDF files",
            command=self.select_folder
        )
        self.select_btn.pack(
            side=tk.LEFT,
            expand=True,
            fill=tk.X,
            padx=(0, 5)
        )

        self.help_btn = tk.Button(
            top_frame,
            text="Help",
            width=10,
            command=self.show_help
        )
        self.help_btn.pack(side=tk.LEFT)

        list_frame = tk.Frame(self.root)
        list_frame.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        list_label = tk.Label(
            list_frame,
            text="Files to be Processed:",
            anchor="w"
        )
        list_label.pack(fill=tk.X)

        scrollbar = tk.Scrollbar(list_frame, orient=tk.VERTICAL)

        self.file_listbox = tk.Listbox(
            list_frame,
            yscrollcommand=scrollbar.set
        )

        scrollbar.config(command=self.file_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        bottom_frame = tk.Frame(self.root)
        bottom_frame.pack(pady=10, padx=10, fill=tk.X)

        self.status_label = tk.Label(
            bottom_frame,
            text="Please select a folder to begin.",
            wraplength=580
        )
        self.status_label.pack(fill=tk.X)

        self.progress_bar = ttk.Progressbar(
            bottom_frame,
            length=300,
            mode="determinate"
        )
        self.progress_bar.pack(pady=5, fill=tk.X)

        self.precheck_btn = tk.Button(
            bottom_frame,
            text="Optional checks for contracts",
            state=tk.DISABLED,
            command=self.open_checks_dialog
        )
        self.precheck_btn.pack(pady=(5, 15), fill=tk.X)

        self.start_btn = tk.Button(
            bottom_frame,
            text="Start Conversion & Combine",
            state=tk.DISABLED,
            command=self.run_conversion_and_combine,
            height=2,
            font=("Arial", 9, "bold")
        )
        self.start_btn.pack(pady=(0, 5), fill=tk.X)

    def show_help(self):
        help_text = (
            "Converts Word and Excel files to PDF and combines them.\n\n"
            "On Word files, comments and track changes will be turned OFF "
            "automatically to ensure a clean export.\n\n"
            "Optional Checks:\n"
            "Audit Word/Excel files for header/footer references, font uniformity, "
            "user-entered keywords, bracketed text, parentheses text, and "
            "guide-style headings before conversion.\n\n"
            "---\n"
            "Version 5.0, April 22, 2026\n"
            "Author: jgutierrez2@burnsmcd.com"
        )

        messagebox.showinfo("Help", help_text)

    def select_folder(self):
        folder = filedialog.askdirectory(title="Select Folder")

        if not folder:
            return

        self.source_folder = folder

        self.status_label.config(text=f"Folder Selected: {self.source_folder}")
        self.file_listbox.delete(0, tk.END)

        self.start_btn.config(state=tk.DISABLED)
        self.precheck_btn.config(state=tk.DISABLED)

        valid_extensions = (
            ".docx",
            ".doc",
            ".docm",
            ".xlsx",
            ".xls",
            ".xlsm",
            ".pdf"
        )

        self.all_files = sorted(
            [
                f
                for f in os.listdir(self.source_folder)
                if f.lower().endswith(valid_extensions)
                and not f.startswith("~$")
            ],
            key=natural_sort_key
        )

        if not self.all_files:
            messagebox.showwarning(
                "No Files",
                "No compatible Word, Excel, or PDF files found."
            )
            return

        excel_files = [
            f
            for f in self.all_files
            if f.lower().endswith((".xlsx", ".xls", ".xlsm"))
        ]

        if excel_files:
            self.analyze_and_prompt_excel_tabs(excel_files)

        for filename in self.all_files:
            self.file_listbox.insert(tk.END, filename)

            if filename in self.excel_tabs_to_convert:
                for tab in self.excel_tabs_to_convert[filename]:
                    self.file_listbox.insert(tk.END, f"      ↳ {tab}")

        self.start_btn.config(state=tk.NORMAL)
        self.precheck_btn.config(state=tk.NORMAL)

    def analyze_and_prompt_excel_tabs(self, excel_files):
        self.status_label.config(text="Analyzing Excel tabs...")
        self.root.update_idletasks()

        excel_app = None
        excel_file_info = {}

        try:
            excel_app = win32com.client.Dispatch("Excel.Application")
            excel_app.Visible = False
            excel_app.DisplayAlerts = False

            for file in excel_files:
                path = os.path.abspath(
                    os.path.join(self.source_folder, file)
                )

                try:
                    wb = excel_app.Workbooks.Open(path, ReadOnly=True)

                    excel_file_info[file] = [
                        sheet.Name
                        for sheet in wb.Sheets
                    ]

                    wb.Close(False)

                except Exception as e:
                    print(f"Error reading {file}: {e}")

        except Exception:
            pass

        finally:
            if excel_app:
                excel_app.Quit()

        if excel_file_info:
            dialog = ExcelTabSelectionDialog(self.root, excel_file_info)
            self.root.wait_window(dialog)

            self.excel_tabs_to_convert = dialog.result

            new_all_files = []

            for f in self.all_files:
                if f in excel_files:
                    if f in self.excel_tabs_to_convert:
                        new_all_files.append(f)
                else:
                    new_all_files.append(f)

            self.all_files = new_all_files

        self.status_label.config(text=f"Folder Selected: {self.source_folder}")

    def open_checks_dialog(self):
        ContractChecksOptionsDialog(self.root, self.execute_contract_checks)

    # --- Text / Guide Content Helpers ---
    def _normalize_text(self, text):
        if not text:
            return ""

        text = str(text)
        text = text.replace("\r", " ")
        text = text.replace("\n", " ")
        text = text.replace("\t", " ")
        text = re.sub(r"\s+", " ", text)

        return text.strip()

    def _find_brackets_and_parentheses(self, text):
        issues = []

        clean_text = self._normalize_text(text)

        bracket_matches = re.findall(r"\[[^\]]+\]", clean_text)
        parenthesis_matches = re.findall(r"\([^)]*\)", clean_text)

        for match in bracket_matches:
            issues.append(f"Bracketed text found: {match}")

        for match in parenthesis_matches:
            issues.append(f"Parenthetical text found: {match}")

        return issues

    def _find_blue_box_heading_like_content(self, doc):
        issues = []

        try:
            for para in doc.Paragraphs:
                try:
                    style_name = str(para.Range.Style.NameLocal)
                    para_text = self._normalize_text(para.Range.Text)

                    if not para_text:
                        continue

                    style_name_lower = style_name.lower()
                    para_text_lower = para_text.lower()

                    style_matches = any(
                        token in style_name_lower
                        for token in [
                            "blue",
                            "box",
                            "guide",
                            "heading",
                            "instruction",
                            "specifier",
                            "note",
                            "comment",
                            "editor"
                        ]
                    )

                    text_matches = any(
                        token in para_text_lower
                        for token in [
                            "blue box",
                            "guide note",
                            "guide heading",
                            "instruction heading",
                            "spec writer",
                            "specwriter",
                            "note:",
                            "comment:",
                            "instruction:"
                        ]
                    )

                    shading_or_blue_font = False

                    try:
                        shading_color = para.Range.Shading.BackgroundPatternColor

                        if shading_color not in [None, -16777216, 9999999, 0]:
                            shading_or_blue_font = True
                    except Exception:
                        pass

                    try:
                        if self._is_blue_rgb(para.Range.Font.Color):
                            shading_or_blue_font = True
                    except Exception:
                        pass

                    if style_matches or text_matches or shading_or_blue_font:
                        issues.append(
                            f"Blue-box/guide heading-like content found: "
                            f"{para_text[:180]}"
                        )

                except Exception:
                    pass

        except Exception:
            pass

        return list(dict.fromkeys(issues))

    def _is_blue_rgb(self, rgb_value):
        try:
            if rgb_value is None:
                return False

            rgb_value = int(rgb_value)

            if rgb_value < 0:
                return False

            red = rgb_value & 255
            green = (rgb_value >> 8) & 255
            blue = (rgb_value >> 16) & 255

            return blue > 100 and blue > red + 40 and blue > green + 20

        except Exception:
            return False

    def _find_guide_heading_styles(self, doc):
        """
        Flags paragraphs whose Word style name appears to be a guide note,
        blue box, instruction heading, or spec-writer instruction style.

        If you know the exact style name later, add it to exact_style_names.
        """
        issues = []

        exact_style_names = [
            # Add exact style names here if known, for example:
            # "Blue Box",
            # "Guide Note",
            # "Spec Writer Instruction",
        ]

        style_keywords = [
            "blue",
            "guide",
            "instruction",
            "specifier",
            "spec writer",
            "specwriter",
            "editor",
            "note",
            "comment",
            "shaded"
        ]

        try:
            for para in doc.Paragraphs:
                try:
                    style_name = str(para.Range.Style.NameLocal)
                    para_text = self._normalize_text(para.Range.Text)

                    if not para_text:
                        continue

                    style_name_lower = style_name.lower()

                    exact_match = style_name in exact_style_names

                    keyword_match = any(
                        keyword in style_name_lower
                        for keyword in style_keywords
                    )

                    if exact_match or keyword_match:
                        issues.append(
                            f"Guide/blue-box style found "
                            f"('{style_name}'): {para_text[:180]}"
                        )

                except Exception:
                    pass

        except Exception:
            pass

        return list(dict.fromkeys(issues))

    def _find_highlight_shading_blue_font(self, doc):
        issues = []

        try:
            for para in doc.Paragraphs:
                rng = para.Range
                para_text = self._normalize_text(rng.Text)

                if not para_text:
                    continue

                try:
                    if rng.HighlightColorIndex != 0:
                        issues.append(
                            f"Highlighted text found: {para_text[:180]}"
                        )
                except Exception:
                    pass

                try:
                    shading_color = rng.Shading.BackgroundPatternColor

                    if shading_color not in [None, -16777216, 9999999, 0]:
                        issues.append(
                            f"Shaded text found: {para_text[:180]}"
                        )
                except Exception:
                    pass

                try:
                    if self._is_blue_rgb(rng.Font.Color):
                        issues.append(
                            f"Blue font text found: {para_text[:180]}"
                        )
                except Exception:
                    pass

        except Exception:
            pass

        return list(dict.fromkeys(issues))

    def _prepare_word_document_for_scan(self, doc):
        try:
            if hasattr(doc, "TrackRevisions"):
                doc.TrackRevisions = False
        except Exception:
            pass

        try:
            if hasattr(doc, "ActiveWindow") and hasattr(doc.ActiveWindow, "View"):
                view = doc.ActiveWindow.View

                for attr in [
                    "ShowRevisionsAndComments",
                    "ShowComments",
                    "ShowFormatChanges",
                    "ShowInkAnnotations"
                ]:
                    if hasattr(view, attr):
                        setattr(view, attr, False)
        except Exception:
            pass

    def _export_summary_report(self, results_dict):
        if not self.source_folder:
            return ""

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_path = os.path.join(
            self.source_folder,
            f"contract_check_report_{timestamp}.txt"
        )

        lines = ["Contract Check Summary", "=" * 40, ""]

        if "hf" in results_dict:
            lines.append("Header/Footer Results")
            lines.append("-" * 20)
            for entry in results_dict["hf"]:
                lines.append(f"File: {entry['file']}")
                lines.append(f"Status: {entry['status']}")
                if entry.get("issues"):
                    for issue in entry["issues"]:
                        lines.append(f"  - {issue}")
                lines.append("")

        if "keyword" in results_dict:
            lines.append("Keyword/Guide Results")
            lines.append("-" * 20)
            if results_dict.get("keyword_meta"):
                lines.append(results_dict["keyword_meta"])
                lines.append("")
            for entry in results_dict["keyword"]:
                lines.append(f"File: {entry['file']}")
                lines.append(f"Status: {entry['status']}")
                if entry.get("issues"):
                    for issue in entry["issues"]:
                        lines.append(f"  - {issue}")
                lines.append("")

        if "font" in results_dict:
            lines.append("Font Results")
            lines.append("-" * 20)
            lines.append(f"Standard font: {results_dict.get('font_meta', 'Unknown')}")
            lines.append("")
            for entry in results_dict["font"]:
                lines.append(f"File: {entry['file']}")
                lines.append(f"Status: {entry['status']}")
                if entry.get("issues"):
                    for issue in entry["issues"]:
                        lines.append(f"  - {issue}")
                lines.append("")

        try:
            with open(report_path, "w", encoding="utf-8") as fh:
                fh.write("\n".join(lines))
            return report_path
        except Exception:
            return ""

    def _apply_review_notes_to_word_files(self, results_dict):
        review_targets = []

        for entry in results_dict.get("hf", []):
            if entry.get("status") == "warning":
                review_targets.append((entry["file"], entry.get("issues", [])))

        for entry in results_dict.get("keyword", []):
            if entry.get("status") == "warning":
                review_targets.append((entry["file"], entry.get("issues", [])))

        for entry in results_dict.get("font", []):
            if entry.get("status") == "warning":
                review_targets.append((entry["file"], entry.get("issues", [])))

        if not review_targets:
            return

        word = None

        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0

            for filename, issues in review_targets:
                if not filename.lower().endswith((".docx", ".doc", ".docm")):
                    continue

                input_path = os.path.abspath(
                    os.path.join(self.source_folder, filename)
                )

                if not os.path.exists(input_path):
                    continue

                if not messagebox.askyesno(
                    "Apply review comments",
                    f"Add review comments to {filename}?\n\n"
                    f"{len(issues)} issue(s) were found."
                ):
                    continue

                doc = None

                try:
                    doc = word.Documents.Open(input_path, ReadOnly=False)
                    start_range = doc.Range(0, 0)
                    start_range.Collapse(0)
                    start_range.InsertAfter("\n")

                    note_text = "Contract check review note:\n"
                    note_text += "\n".join(issues[:12])

                    if len(issues) > 12:
                        note_text += "\n..."

                    doc.Comments.Add(start_range, note_text)
                    doc.Save()
                except Exception as exc:
                    messagebox.showwarning(
                        "Review Note Failed",
                        f"Could not add review comments to {filename}: {exc}"
                    )
                finally:
                    if doc is not None:
                        doc.Close(False)

        except Exception as exc:
            messagebox.showerror("Review Notes Failed", str(exc))
        finally:
            if word:
                word.Quit()

    # --- Advanced Contract Checks Engine ---
    def execute_contract_checks(self, options):
        if not any([options["hf"], options["font"], options["keyword"]]):
            return

        native_files = [
            f
            for f in self.all_files
            if f.lower().endswith(
                (".docx", ".doc", ".docm", ".xlsx", ".xls", ".xlsm")
            )
        ]

        if not native_files:
            messagebox.showinfo(
                "No Files",
                "No Word or Excel files found to check."
            )
            return

        self.start_btn.config(state=tk.DISABLED)
        self.precheck_btn.config(state=tk.DISABLED)

        self.status_label.config(text="Running selected contract checks...")
        self.root.update_idletasks()

        keywords_to_check = options.get("keywords_to_check", [])
        keyword_selection_label = ""

        if options["keyword"]:
            keyword_contract_type = options.get("keyword_contract_type", "")
            keyword_role = options.get("keyword_role", "")

            keyword_selection_label = (
                f"{keyword_contract_type} - {keyword_role}"
            )

        final_results = {}

        if options["hf"]:
            final_results["hf"] = []

        if options["font"]:
            final_results["font"] = []

        if options["keyword"]:
            final_results["keyword"] = []

            if keywords_to_check:
                keyword_text = ", ".join(keywords_to_check)
            else:
                keyword_text = "No manual keywords entered"

            final_results["keyword_meta"] = (
                f"Keyword / Guide Text Check Type: "
                f"{keyword_selection_label}\n"
                f"Manual Keywords: {keyword_text}\n"
                f"Automatic Flags: blue-box-style headings, "
                f"bracketed text, parenthetical text, highlighted text, "
                f"shaded text, and blue font text"
            )

        review_mode = options.get("review_mode", "Report only")

        font_tracker = {}
        word = None
        excel = None

        self.progress_bar["maximum"] = len(native_files)

        try:
            for i, filename in enumerate(native_files):
                self.progress_bar["value"] = i + 1
                self.root.update_idletasks()

                ext = os.path.splitext(filename)[1].lower()
                base_name = os.path.splitext(filename)[0]

                input_path = os.path.abspath(
                    os.path.join(self.source_folder, filename)
                )

                hf_issues = []
                keyword_issues = []
                doc_font = None

                try:
                    if ext in [".docx", ".doc", ".docm"]:
                        if not word:
                            word = win32com.client.Dispatch("Word.Application")
                            word.Visible = False
                            word.DisplayAlerts = 0

                        doc = word.Documents.Open(input_path, ReadOnly=True)

                        if options["font"]:
                            try:
                                doc_font = doc.Styles(-1).Font.Name
                            except Exception:
                                doc_font = "Unknown"

                            font_tracker[filename] = doc_font

                        if options["hf"]:
                            h_text = ""
                            f_text = ""

                            for sec in doc.Sections:
                                for hf_type in [1, 2, 3]:
                                    try:
                                        h_text += (
                                            sec.Headers(hf_type).Range.Text
                                            + " "
                                        )
                                    except Exception:
                                        pass

                                    try:
                                        f_text += (
                                            sec.Footers(hf_type).Range.Text
                                            + " "
                                        )
                                    except Exception:
                                        pass

                            hf_issues = self._analyze_hf_text(
                                base_name,
                                h_text,
                                f_text
                            )

                        if options["keyword"]:
                            raw_content_text = doc.Content.Text
                            content_text = self._normalize_text(
                                raw_content_text
                            ).lower()

                            for kw in keywords_to_check:
                                search_kw = self._normalize_text(kw).lower()

                                if search_kw and search_kw in content_text:
                                    keyword_issues.append(
                                        f"Manual keyword found: {kw}"
                                    )

                            keyword_issues.extend(
                                self._find_brackets_and_parentheses(
                                    raw_content_text
                                )
                            )

                            keyword_issues.extend(
                                self._find_guide_heading_styles(doc)
                            )

                            keyword_issues.extend(
                                self._find_blue_box_heading_like_content(doc)
                            )

                            keyword_issues.extend(
                                self._find_highlight_shading_blue_font(doc)
                            )

                        doc.Close(False)

                    elif ext in [".xlsx", ".xls", ".xlsm"]:
                        if not excel:
                            excel = win32com.client.Dispatch("Excel.Application")
                            excel.Visible = False
                            excel.DisplayAlerts = False

                        wb = excel.Workbooks.Open(input_path, ReadOnly=True)

                        sheets_to_check = self.excel_tabs_to_convert.get(
                            filename,
                            [s.Name for s in wb.Sheets]
                        )

                        all_cell_text = ""

                        for sheet_name in sheets_to_check:
                            ws = wb.Sheets(sheet_name)

                            if options["hf"]:
                                h_text = ""
                                f_text = ""
                                ps = ws.PageSetup

                                for attr in [
                                    "LeftHeader",
                                    "CenterHeader",
                                    "RightHeader"
                                ]:
                                    val = getattr(ps, attr, "")

                                    if val:
                                        h_text += val + " "

                                        if "&F" in val.upper():
                                            h_text += base_name + " "

                                for attr in [
                                    "LeftFooter",
                                    "CenterFooter",
                                    "RightFooter"
                                ]:
                                    val = getattr(ps, attr, "")

                                    if val:
                                        f_text += val + " "

                                        if "&F" in val.upper():
                                            f_text += base_name + " "

                                hf_issues.extend(
                                    self._analyze_hf_text(
                                        base_name,
                                        h_text,
                                        f_text
                                    )
                                )

                            if options["keyword"]:
                                try:
                                    all_cell_text += (
                                        str(ws.UsedRange.Value) + " "
                                    )
                                except Exception:
                                    pass

                        if options["keyword"]:
                            clean_cell_text = self._normalize_text(
                                all_cell_text
                            ).lower()

                            for kw in keywords_to_check:
                                search_kw = self._normalize_text(kw).lower()

                                if search_kw and search_kw in clean_cell_text:
                                    keyword_issues.append(
                                        f"Manual keyword found: {kw}"
                                    )

                            keyword_issues.extend(
                                self._find_brackets_and_parentheses(
                                    all_cell_text
                                )
                            )

                        wb.Close(False)

                except Exception as e:
                    hf_issues.append(f"Failed to read metadata ({e})")

                if options["hf"]:
                    hf_issues = list(dict.fromkeys(hf_issues))

                    final_results["hf"].append(
                        {
                            "file": filename,
                            "status": "warning" if hf_issues else "ok",
                            "issues": hf_issues
                        }
                    )

                if options["keyword"]:
                    keyword_issues = list(dict.fromkeys(keyword_issues))

                    final_results["keyword"].append(
                        {
                            "file": filename,
                            "status": "warning" if keyword_issues else "ok",
                            "issues": keyword_issues
                        }
                    )

            if options["font"] and font_tracker:
                counts = Counter(font_tracker.values())
                standard_font = counts.most_common(1)[0][0]

                final_results["font_meta"] = standard_font

                for f_name, f_type in font_tracker.items():
                    if f_type != standard_font:
                        final_results["font"].append(
                            {
                                "file": f_name,
                                "status": "warning",
                                "issues": [
                                    f"Uses {f_type} "
                                    f"(Standard is {standard_font})"
                                ]
                            }
                        )
                    else:
                        final_results["font"].append(
                            {
                                "file": f_name,
                                "status": "ok",
                                "issues": []
                            }
                        )

            if review_mode in [
                "Export report to .txt",
                "Export report and apply review comments to Word files"
            ]:
                report_path = self._export_summary_report(final_results)

                if report_path:
                    messagebox.showinfo(
                        "Export Successful",
                        f"Summary saved to:\n{report_path}"
                    )

            if review_mode == "Export report and apply review comments to Word files":
                self._apply_review_notes_to_word_files(final_results)

            ChecksResultsDialog(self.root, final_results)

        except Exception as e:
            messagebox.showerror(
                "Check Failed",
                f"An error occurred: {e}"
            )

        finally:
            if word:
                word.Quit()

            if excel:
                excel.Quit()

            self.start_btn.config(state=tk.NORMAL)
            self.precheck_btn.config(state=tk.NORMAL)

            self.status_label.config(
                text=f"Folder Selected: {self.source_folder}"
            )

            self.progress_bar["value"] = 0

    def _analyze_hf_text(self, base_name, raw_h_text, raw_f_text):
        issues = []

        clean_h = re.sub(r"\s+", " ", raw_h_text).strip()
        clean_f = re.sub(r"\s+", " ", raw_f_text).strip()

        match = re.match(
            r"^([\d]+\.[\d]+\.[\d]+)(?:[\.\s_-]+(.*))?$",
            base_name
        )

        if match and match.group(1):
            proj_contract = match.group(1).strip()
            spec_num = match.group(2).strip() if match.group(2) else ""

            if proj_contract not in clean_f:
                issues.append(
                    f"Footer missing project/contract number "
                    f"'{proj_contract}'"
                )

            if spec_num and spec_num not in clean_h:
                issues.append(
                    f"Header missing spec number/name '{spec_num}'"
                )

        else:
            if base_name not in clean_h and base_name not in clean_f:
                issues.append(
                    f"Filename '{base_name}' not found anywhere in "
                    f"header or footer"
                )

        return issues

    def run_conversion_and_combine(self):
        self.start_btn.config(state=tk.DISABLED)
        self.precheck_btn.config(state=tk.DISABLED)
        self.select_btn.config(state=tk.DISABLED)

        self.status_label.config(
            text="Processing... The application may be unresponsive during conversion."
        )

        self.root.update_idletasks()

        files_to_convert = [
            f
            for f in self.all_files
            if not f.lower().endswith(".pdf")
        ]

        output_folder = os.path.join(self.source_folder, "Converted_PDFs")

        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        self.progress_bar["maximum"] = len(files_to_convert)

        new_pdf_paths = []
        failed_files = []
        word = None
        excel = None

        try:
            for i, filename in enumerate(files_to_convert):
                self.status_label.config(text=f"Converting: {filename}")
                self.progress_bar["value"] = i + 1
                self.root.update_idletasks()

                ext = os.path.splitext(filename)[1].lower()

                input_path = os.path.abspath(
                    os.path.join(self.source_folder, filename)
                )

                try:
                    if ext in [".docx", ".doc", ".docm"]:
                        output_path = os.path.abspath(
                            os.path.join(
                                output_folder,
                                os.path.splitext(filename)[0] + ".pdf"
                            )
                        )

                        if not word:
                            word = win32com.client.Dispatch("Word.Application")
                            word.Visible = False
                            word.DisplayAlerts = 0

                        doc = word.Documents.Open(input_path, ReadOnly=True)

                        word.ActiveWindow.View.ShowRevisionsAndComments = False
                        word.ActiveWindow.View.RevisionsView = 0

                        doc.ExportAsFixedFormat(output_path, 17)
                        doc.Close(False)

                        new_pdf_paths.append(output_path)

                    elif ext in [".xlsx", ".xls", ".xlsm"]:
                        if not excel:
                            excel = win32com.client.Dispatch("Excel.Application")
                            excel.Visible = False
                            excel.DisplayAlerts = False

                        wb = excel.Workbooks.Open(input_path, ReadOnly=True)

                        sheets_to_convert = self.excel_tabs_to_convert.get(
                            filename,
                            []
                        )

                        if len(sheets_to_convert) > 0:
                            if len(sheets_to_convert) == 1 and len(wb.Sheets) == 1:
                                output_path = os.path.abspath(
                                    os.path.join(
                                        output_folder,
                                        os.path.splitext(filename)[0] + ".pdf"
                                    )
                                )

                                wb.Sheets(sheets_to_convert[0]).Select()
                                wb.ActiveSheet.ExportAsFixedFormat(
                                    0,
                                    output_path
                                )

                                new_pdf_paths.append(output_path)

                            else:
                                base_name = os.path.splitext(filename)[0]

                                for sheet_name in sheets_to_convert:
                                    tab_output_path = os.path.abspath(
                                        os.path.join(
                                            output_folder,
                                            f"{base_name} - {sheet_name}.pdf"
                                        )
                                    )

                                    wb.Sheets(sheet_name).Select()
                                    wb.ActiveSheet.ExportAsFixedFormat(
                                        0,
                                        tab_output_path
                                    )

                                    new_pdf_paths.append(tab_output_path)

                        wb.Close(False)

                except Exception as e:
                    failed_files.append(f"{filename}: {e}")

            if word:
                word.Quit()

            if excel:
                excel.Quit()

            self.root.withdraw()

            original_pdfs = [
                os.path.join(self.source_folder, f)
                for f in self.all_files
                if f.lower().endswith(".pdf")
            ]

            all_pdf_paths = new_pdf_paths + original_pdfs
            all_pdf_paths.sort(key=natural_sort_key)

            proceed_to_combine = False
            dialog_result_ok = False

            if len(all_pdf_paths) > 1:
                if messagebox.askyesno(
                    "Combine Files?",
                    f"Conversion complete.\n"
                    f"Found {len(all_pdf_paths)} total PDFs.\n\n"
                    f"Combine them into a single PDF?"
                ):
                    proceed_to_combine = True

            elif all_pdf_paths:
                messagebox.showinfo(
                    "Process Complete",
                    "Conversion finished. Only one PDF exists, "
                    "so no combination was performed."
                )

            else:
                messagebox.showinfo(
                    "Process Complete",
                    "Conversion finished. No PDFs were found to combine."
                )

            if proceed_to_combine:
                dialog = ReorderDialog(self.root, all_pdf_paths)

                if dialog.result_list:
                    dialog_result_ok = True
                    self.process_merging(dialog.result_list, output_folder)

            if failed_files:
                messagebox.showwarning(
                    "Conversion Issues",
                    "The following files failed to convert:\n\n"
                    + "\n".join(failed_files)
                )

            if not (proceed_to_combine and dialog_result_ok):
                messagebox.showinfo(
                    "Process Complete",
                    "File conversion is finished."
                )

        except Exception as e:
            if word:
                word.Quit()

            if excel:
                excel.Quit()

            messagebox.showerror("Critical Error", f"Engine failure: {e}")

        finally:
            sys.exit(0)

    def process_merging(self, ordered_paths, output_folder):
        try:
            merged_doc = pymupdf.open()

            toc = []
            page_labels = []
            current_page_count = 0

            for pdf_path in ordered_paths:
                filename = os.path.basename(pdf_path)
                doc_title = os.path.splitext(filename)[0]

                with pymupdf.open(pdf_path) as src_doc:
                    num_pages = src_doc.page_count

                    toc.append(
                        [
                            1,
                            doc_title,
                            current_page_count + 1
                        ]
                    )

                    page_labels.append(
                        {
                            "startpage": current_page_count,
                            "prefix": doc_title + "-",
                            "style": "D",
                            "firstpagenum": 1
                        }
                    )

                    merged_doc.insert_pdf(src_doc)
                    current_page_count += num_pages

            merged_doc.set_toc(toc)

            if hasattr(merged_doc, "set_page_labels"):
                merged_doc.set_page_labels(page_labels)

            combined_path = os.path.join(output_folder, "_combined.pdf")

            merged_doc.save(
                combined_path,
                garbage=4,
                deflate=True
            )

            merged_doc.close()

            if messagebox.askyesno(
                "Success!",
                "Files combined successfully!\n"
                "Saved in 'Converted_PDFs' folder.\n\n"
                "Would you like to open the output folder?"
            ):
                os.startfile(output_folder)

        except Exception as e:
            messagebox.showerror(
                "Merge Error",
                f"Could not combine files: {e}"
            )


def show_initial_warning():
    dialog = tk.Toplevel()
    dialog.title("Warning")
    dialog.resizable(False, False)

    dialog.update_idletasks()

    width = 450
    height = 120

    x = (dialog.winfo_screenwidth() // 2) - (width // 2)
    y = (dialog.winfo_screenheight() // 2) - (height // 2)

    dialog.geometry(f"{width}x{height}+{x}+{y}")

    label = tk.Label(
        dialog,
        text=(
            "THIS TOOL REQUIRES WORD AND EXCEL TO BE CLOSED, OTHERWISE IT WILL "
            "KILL ANY OPEN PROCESSES.\nPLEASE SAVE YOUR WORK IN THOSE PROGRAMS "
            "BEFORE PROCEEDING."
        ),
        wraplength=430,
        justify=tk.CENTER
    )

    label.pack(pady=10, padx=10)

    user_choice = tk.BooleanVar(value=False)

    def on_continue():
        user_choice.set(True)
        dialog.destroy()

    def on_exit():
        user_choice.set(False)
        dialog.destroy()

    button_frame = tk.Frame(dialog)
    button_frame.pack(pady=10)

    tk.Button(
        button_frame,
        text="Continue",
        width=10,
        command=on_continue
    ).pack(side=tk.LEFT, padx=10)

    tk.Button(
        button_frame,
        text="Exit",
        width=10,
        command=on_exit
    ).pack(side=tk.LEFT, padx=10)

    dialog.protocol("WM_DELETE_WINDOW", on_exit)

    dialog.grab_set()
    dialog.wait_window()

    return user_choice.get()


if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()

    if show_initial_warning():
        main_app_window = tk.Toplevel(root)
        app = BatchConverterApp(main_app_window)

        main_app_window.protocol(
            "WM_DELETE_WINDOW",
            lambda: sys.exit(0)
        )

        root.mainloop()

    else:
        sys.exit(0)