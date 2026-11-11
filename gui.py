"""
gui.py
Minimal, professional Tkinter user interface for Excel Financial Summary Automation Tool.
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from excel_processor import process_excel_summary, ExcelAutomationError


class ExcelAutomationApp(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Excel Financial Summary Automation Tool")
        self.geometry("780x640")
        self.minsize(700, 560)

        # Set default font & colors
        self.configure(bg="#F3F4F6")

        # Variables
        self.source_path_var = tk.StringVar()
        self.template_path_var = tk.StringVar()
        self.output_path_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Ready. Select files and click Process Excel.")

        # Default paths if present in current directory
        base_dir = os.path.dirname(os.path.abspath(__file__))
        default_tpl = os.path.join(base_dir, "template.xlsx")
        if os.path.exists(default_tpl):
            self.template_path_var.set(default_tpl)

        default_src = os.path.join(base_dir, "sample_source_23-10-2026.xlsx")
        if os.path.exists(default_src):
            self.source_path_var.set(default_src)
            default_out = os.path.join(base_dir, "processed_output.xlsx")
            self.output_path_var.set(default_out)

        self._create_widgets()

    def _create_widgets(self):
        # Header Banner
        header_frame = tk.Frame(self, bg="#1E293B", padx=20, pady=16)
        header_frame.pack(fill=tk.X)

        title_lbl = tk.Label(
            header_frame,
            text="Excel Financial Summary Automation",
            font=("Segoe UI", 15, "bold"),
            fg="#FFFFFF",
            bg="#1E293B"
        )
        title_lbl.pack(anchor="w")

        sub_lbl = tk.Label(
            header_frame,
            text="Extracts KSM, SDM & KSM HUF financial summaries into the target template with strict format preservation.",
            font=("Segoe UI", 9),
            fg="#94A3B8",
            bg="#1E293B"
        )
        sub_lbl.pack(anchor="w", pady=(3, 0))

        # Main Body Frame
        body = tk.Frame(self, bg="#F3F4F6", padx=24, pady=16)
        body.pack(fill=tk.BOTH, expand=True)

        # File Selection Card
        card = tk.LabelFrame(
            body,
            text=" File Configuration ",
            font=("Segoe UI", 10, "bold"),
            bg="#FFFFFF",
            fg="#1E293B",
            padx=16,
            pady=16,
            relief=tk.GROOVE,
            bd=1
        )
        card.pack(fill=tk.X, pady=(0, 14))

        # 1. Source File
        self._build_file_picker_row(
            card,
            row=0,
            label="Source Excel File:",
            var=self.source_path_var,
            browse_command=self._browse_source,
            hint="Workbook with 4 sheets (e.g. 23-10-2026.xlsx)"
        )

        # 2. Template File
        self._build_file_picker_row(
            card,
            row=1,
            label="Template Excel File:",
            var=self.template_path_var,
            browse_command=self._browse_template,
            hint="Template with {PUTDATE}, {B3}, {B5}, etc. placeholders"
        )

        # 3. Output File
        self._build_file_picker_row(
            card,
            row=2,
            label="Output File Location:",
            var=self.output_path_var,
            browse_command=self._browse_output,
            hint="Destination file (e.g. processed_output.xlsx)"
        )

        # Action Button Frame
        action_frame = tk.Frame(body, bg="#F3F4F6")
        action_frame.pack(fill=tk.X, pady=(0, 14))

        self.btn_process = tk.Button(
            action_frame,
            text="  Process Excel  ",
            font=("Segoe UI", 11, "bold"),
            bg="#0284C7",
            fg="#FFFFFF",
            activebackground="#0369A1",
            activeforeground="#FFFFFF",
            padx=20,
            pady=8,
            relief=tk.FLAT,
            cursor="hand2",
            command=self._start_processing
        )
        self.btn_process.pack(side=tk.LEFT)

        # Status Bar / Message area
        self.status_label = tk.Label(
            action_frame,
            textvariable=self.status_var,
            font=("Segoe UI", 9, "italic"),
            fg="#4B5563",
            bg="#F3F4F6",
            anchor="w"
        )
        self.status_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(16, 0))

        # Results & Summary Frame
        results_card = tk.LabelFrame(
            body,
            text=" Extraction Summary & Output Details ",
            font=("Segoe UI", 10, "bold"),
            bg="#FFFFFF",
            fg="#1E293B",
            padx=16,
            pady=12,
            relief=tk.GROOVE,
            bd=1
        )
        results_card.pack(fill=tk.BOTH, expand=True)

        # Treeview table to display extracted values
        columns = ("Section", "Field", "Source Cell", "Extracted Value")
        self.tree = ttk.Treeview(results_card, columns=columns, show="headings", height=8)

        # Configure columns
        self.tree.heading("Section", text="Account Section")
        self.tree.heading("Field", text="Financial Field")
        self.tree.heading("Source Cell", text="Source Coordinate")
        self.tree.heading("Extracted Value", text="Extracted Financial Value")

        self.tree.column("Section", width=110, anchor="center")
        self.tree.column("Field", width=140, anchor="w")
        self.tree.column("Source Cell", width=100, anchor="center")
        self.tree.column("Extracted Value", width=260, anchor="w")

        # Scrollbar for treeview
        scrollbar = ttk.Scrollbar(results_card, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Bottom Button Bar for Output
        bottom_bar = tk.Frame(self, bg="#E5E7EB", padx=20, pady=10)
        bottom_bar.pack(fill=tk.X, side=tk.BOTTOM)

        self.btn_open_file = tk.Button(
            bottom_bar,
            text="Open Generated Excel",
            font=("Segoe UI", 9),
            bg="#FFFFFF",
            fg="#1F2937",
            padx=12,
            pady=4,
            relief=tk.SOLID,
            bd=1,
            cursor="hand2",
            state=tk.DISABLED,
            command=self._open_output_file
        )
        self.btn_open_file.pack(side=tk.RIGHT, padx=(8, 0))

        self.btn_open_folder = tk.Button(
            bottom_bar,
            text="Open Output Folder",
            font=("Segoe UI", 9),
            bg="#FFFFFF",
            fg="#1F2937",
            padx=12,
            pady=4,
            relief=tk.SOLID,
            bd=1,
            cursor="hand2",
            state=tk.DISABLED,
            command=self._open_output_folder
        )
        self.btn_open_folder.pack(side=tk.RIGHT)

        self.lbl_path_info = tk.Label(
            bottom_bar,
            text="No output generated yet.",
            font=("Segoe UI", 8),
            fg="#6B7280",
            bg="#E5E7EB"
        )
        self.lbl_path_info.pack(side=tk.LEFT)

    def _build_file_picker_row(self, parent, row, label, var, browse_command, hint):
        row_frame = tk.Frame(parent, bg="#FFFFFF")
        row_frame.grid(row=row, column=0, sticky="ew", pady=5)
        parent.columnconfigure(0, weight=1)

        lbl = tk.Label(row_frame, text=label, font=("Segoe UI", 9, "bold"), fg="#374151", bg="#FFFFFF", width=18, anchor="w")
        lbl.pack(side=tk.LEFT)

        entry = tk.Entry(row_frame, textvariable=var, font=("Segoe UI", 9), fg="#111827", bg="#F9FAFB", relief=tk.SOLID, bd=1)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(4, 8), ipady=3)

        btn = tk.Button(
            row_frame,
            text="Browse...",
            font=("Segoe UI", 8),
            bg="#E5E7EB",
            fg="#1F2937",
            relief=tk.SOLID,
            bd=1,
            padx=8,
            pady=2,
            cursor="hand2",
            command=browse_command
        )
        btn.pack(side=tk.RIGHT)

    def _browse_source(self):
        f = filedialog.askopenfilename(
            title="Select Source Excel Workbook",
            filetypes=[("Excel Files", "*.xlsx *.xlsm"), ("All Files", "*.*")]
        )
        if f:
            self.source_path_var.set(f)
            # Auto-suggest output path in same directory if not set
            if not self.output_path_var.get() or self.output_path_var.get().endswith("processed_output.xlsx"):
                src_dir = os.path.dirname(f)
                self.output_path_var.set(os.path.join(src_dir, "processed_output.xlsx"))

    def _browse_template(self):
        f = filedialog.askopenfilename(
            title="Select Template Excel Workbook",
            filetypes=[("Excel Files", "*.xlsx *.xlsm"), ("All Files", "*.*")]
        )
        if f:
            self.template_path_var.set(f)

    def _browse_output(self):
        f = filedialog.asksaveasfilename(
            title="Choose Output Excel File Location",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
        )
        if f:
            self.output_path_var.set(f)

    def _start_processing(self):
        source = self.source_path_var.get().strip()
        template = self.template_path_var.get().strip()
        output = self.output_path_var.get().strip()

        # Update status
        self.status_var.set("Processing workbook... please wait.")
        self.status_label.config(fg="#0284C7")
        self.btn_process.config(state=tk.DISABLED)

        # Run processing in background thread to keep UI completely responsive
        threading.Thread(target=self._run_process_worker, args=(source, template, output), daemon=True).start()

    def _run_process_worker(self, source, template, output):
        try:
            result = process_excel_summary(source, template, output)
            self.after(0, self._on_success, result)
        except ExcelAutomationError as e:
            self.after(0, self._on_error, str(e))
        except Exception as e:
            self.after(0, self._on_error, f"Unexpected system error:\n{e}")

    def _on_success(self, result):
        self.btn_process.config(state=tk.NORMAL)
        self.status_var.set("Success! Output workbook generated.")
        self.status_label.config(fg="#16A34A")

        # Enable output buttons
        self.last_output_file = result["output_file"]
        self.btn_open_file.config(state=tk.NORMAL)
        self.btn_open_folder.config(state=tk.NORMAL)
        self.lbl_path_info.config(text=f"Output: {result['output_file']}", fg="#15803D")

        # Clear and populate table
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Add Date row
        self.tree.insert("", tk.END, values=("Header", "Date", "A1", result["date"]))

        # Add 12 financial rows
        mapping = [
            ("KSM", "Stock Value", "B3"),
            ("KSM", "MTF Debit", "B5"),
            ("KSM", "Booked PNL", "B11"),
            ("KSM", "MTM", "B9"),
            ("SDM", "Stock Value", "E3"),
            ("SDM", "MTF Debit", "E5"),
            ("SDM", "Booked PNL", "E11"),
            ("SDM", "MTM", "E9"),
            ("KSM HUF", "Stock Value", "H3"),
            ("KSM HUF", "MTF Debit", "H5"),
            ("KSM HUF", "Booked PNL", "H11"),
            ("KSM HUF", "MTM", "H9"),
        ]

        cells = result["extracted_cells"]
        for section, field, coord in mapping:
            val = cells.get(coord, "N/A")
            self.tree.insert("", tk.END, values=(section, field, coord, val))

        messagebox.showinfo(
            "Excel Automation Complete",
            f"Successfully processed financial summary!\n\n"
            f"• Source Sheet: {result['source_sheet']}\n"
            f"• Date: {result['date']}\n"
            f"• Replaced Placeholders: {result['replacement_count']}\n"
            f"• Saved To:\n{result['output_file']}"
        )

    def _on_error(self, err_msg):
        self.btn_process.config(state=tk.NORMAL)
        self.status_var.set("Error during processing. See details.")
        self.status_label.config(fg="#DC2626")
        messagebox.showerror("Processing Error", err_msg)

    def _open_output_file(self):
        if hasattr(self, "last_output_file") and os.path.exists(self.last_output_file):
            try:
                os.startfile(self.last_output_file)
            except Exception as e:
                messagebox.showerror("Cannot Open File", f"Error opening file:\n{e}")

    def _open_output_folder(self):
        if hasattr(self, "last_output_file") and os.path.exists(self.last_output_file):
            folder = os.path.dirname(os.path.abspath(self.last_output_file))
            try:
                os.startfile(folder)
            except Exception as e:
                messagebox.showerror("Cannot Open Folder", f"Error opening folder:\n{e}")


def launch_gui():
    app = ExcelAutomationApp()
    app.mainloop()


if __name__ == "__main__":
    launch_gui()
