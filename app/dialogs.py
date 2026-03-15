import os
import customtkinter as ctk
import pandas as pd
from tkinter import filedialog, messagebox, ttk


class ResolutionDialog(ctk.CTkToplevel):
    def __init__(self, parent, company_name, candidates, callback):
        super().__init__(parent)
        self.title("Resolve Ambiguity")
        self.geometry("700x700")
        self.callback = callback
        self.result = None
        self.attributes("-topmost", True)

        self.label = ctk.CTkLabel(
            self,
            text=f"Multiple results found for:\n'{company_name}'",
            font=("Arial", 16, "bold"),
        )
        self.label.pack(pady=10)

        self.sublabel = ctk.CTkLabel(
            self, text="Please select the correct company from the list below:"
        )
        self.sublabel.pack(pady=5)

        self.scroll_frame = ctk.CTkScrollableFrame(self, width=600, height=400)
        self.scroll_frame.pack(pady=10, padx=20)

        self.radio_var = ctk.IntVar(value=-1)
        self.candidates = candidates

        for idx, cand in enumerate(candidates):
            text = f"{cand['name']}\n({cand['reg_no']}) - {cand['type']}"
            rb = ctk.CTkRadioButton(
                self.scroll_frame, text=text, variable=self.radio_var, value=idx
            )
            rb.pack(anchor="w", pady=5, padx=10)

        self.manual_label = ctk.CTkLabel(
            self, text="Or search manually with a new keyword:"
        )
        self.manual_label.pack(pady=(10, 0))

        self.manual_entry = ctk.CTkEntry(
            self, width=300, placeholder_text="New keyword..."
        )
        self.manual_entry.pack(pady=5)

        self.btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.btn_frame.pack(pady=20)

        ctk.CTkButton(
            self.btn_frame, text="Confirm Selection", command=self.on_confirm
        ).pack(side="left", padx=10)
        ctk.CTkButton(
            self.btn_frame, text="Search Manual Keyword", command=self.on_manual_search
        ).pack(side="left", padx=10)
        ctk.CTkButton(
            self.btn_frame, text="Skip This Company", command=self.on_skip, fg_color="red"
        ).pack(side="left", padx=10)

        self.protocol("WM_DELETE_WINDOW", self.on_skip)

    def on_confirm(self):
        idx = self.radio_var.get()
        if idx >= 0:
            self.callback("SELECT", self.candidates[idx])
            self.destroy()
        elif idx == -1:
            messagebox.showwarning("Warning", "Please select a company.")

    def on_manual_search(self):
        term = self.manual_entry.get().strip()
        if term:
            self.callback("MANUAL_SEARCH", term)
            self.destroy()

    def on_skip(self):
        self.callback("SKIP", None)
        self.destroy()


class PasteDialog(ctk.CTkToplevel):
    def __init__(self, parent, callback):
        super().__init__(parent)
        self.title("Paste Company List")
        self.geometry("500x600")
        self.callback = callback
        self.attributes("-topmost", True)

        ctk.CTkLabel(self, text="Paste company names below (one per line):").pack(
            pady=10
        )

        self.textbox = ctk.CTkTextbox(self, width=450, height=450)
        self.textbox.pack(pady=10, padx=10)

        ctk.CTkButton(self, text="Add Companies", command=self.on_add).pack(pady=10)

    def on_add(self):
        text = self.textbox.get("1.0", "end").strip()
        if text:
            companies = [line.strip() for line in text.splitlines() if line.strip()]
            self.callback(companies)
        self.destroy()


class EditDialog(ctk.CTkToplevel):
    def __init__(self, parent, current_data, callback):
        super().__init__(parent)
        self.title("Edit Company Data")
        self.geometry("400x500")
        self.callback = callback
        self.attributes("-topmost", True)

        self.fields = {}
        labels = [
            "Company Name", "Status", "BRN",
            "Old Reg No", "New Reg No", "Found Name", "Type",
        ]

        for idx, lbl in enumerate(labels):
            ctk.CTkLabel(self, text=lbl).pack(pady=(10, 0))
            entry = ctk.CTkEntry(self, width=300)
            entry.insert(0, str(current_data[idx]))
            entry.pack(pady=5)
            self.fields[lbl] = entry

        ctk.CTkButton(self, text="Save Changes", command=self.on_save).pack(pady=20)

    def on_save(self):
        new_values = [
            self.fields[lbl].get()
            for lbl in [
                "Company Name", "Status", "BRN",
                "Old Reg No", "New Reg No", "Found Name", "Type",
            ]
        ]
        self.callback(new_values)
        self.destroy()


class HistoryDialog(ctk.CTkToplevel):
    def __init__(self, parent, history_manager, resume_callback):
        super().__init__(parent)
        self.title("History / Resume")
        self.geometry("600x400")
        self.history = history_manager
        self.resume_callback = resume_callback
        self.attributes("-topmost", True)

        columns = ("id", "name", "date", "items", "status")
        self.tree = ttk.Treeview(
            self, columns=columns, show="headings", selectmode="browse"
        )
        self.tree.heading("id", text="ID")
        self.tree.heading("name", text="Session Name")
        self.tree.heading("date", text="Date")
        self.tree.heading("items", text="Total Items")
        self.tree.heading("status", text="Status")

        self.tree.column("id", width=50)
        self.tree.column("name", width=200)
        self.tree.column("date", width=150)
        self.tree.column("items", width=80)
        self.tree.column("status", width=100)

        self.tree.pack(fill="both", expand=True, padx=10, pady=10)

        btn_frame = ctk.CTkFrame(self)
        btn_frame.pack(fill="x", padx=10, pady=10)

        ctk.CTkButton(btn_frame, text="Resume Selected", command=self.on_resume).pack(
            side="right"
        )

        self.load_sessions()

    def load_sessions(self):
        sessions = self.history.get_all_sessions()
        for s in sessions:
            self.tree.insert("", "end", values=s)

    def on_resume(self):
        selected = self.tree.selection()
        if not selected:
            return
        vals = self.tree.item(selected, "values")
        session_id = vals[0]
        self.resume_callback(session_id)
        self.destroy()


class ExportDialog(ctk.CTkToplevel):
    def __init__(self, parent, data, available_columns):
        super().__init__(parent)
        self.title("Custom Export")
        self.geometry("400x500")
        self.data = data
        self.available_columns = available_columns
        self.attributes("-topmost", True)

        ctk.CTkLabel(
            self, text="Select Columns to Export:", font=("Arial", 14, "bold")
        ).pack(pady=10)

        self.check_vars = {}
        self.scroll = ctk.CTkScrollableFrame(self, height=250)
        self.scroll.pack(fill="x", padx=20, pady=5)

        for col in available_columns:
            var = ctk.BooleanVar(value=True)
            self.check_vars[col] = var
            ctk.CTkCheckBox(self.scroll, text=col, variable=var).pack(
                anchor="w", pady=2
            )

        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(pady=20, fill="x", padx=20)

        ctk.CTkButton(btn_frame, text="Copy to Clipboard", command=self.do_copy).pack(
            fill="x", pady=5
        )
        ctk.CTkButton(
            btn_frame, text="Export to Excel", command=self.do_excel, fg_color="green"
        ).pack(fill="x", pady=5)
        ctk.CTkButton(
            btn_frame, text="Export to CSV", command=self.do_csv, fg_color="#555"
        ).pack(fill="x", pady=5)

    def get_selected_data(self):
        selected_cols = [col for col, var in self.check_vars.items() if var.get()]
        if not selected_cols:
            messagebox.showwarning("Warning", "No columns selected!")
            return None

        filtered_data = []
        for row in self.data:
            new_row = {k: row.get(k, "") for k in selected_cols}
            filtered_data.append(new_row)
        return pd.DataFrame(filtered_data)

    def do_copy(self):
        df = self.get_selected_data()
        if df is not None:
            try:
                df.to_clipboard(index=False, sep="\t")
                messagebox.showinfo("Success", "Copied to clipboard!")
                self.destroy()
            except Exception as e:
                messagebox.showerror("Error", f"Clipboard error: {e}")

    def do_excel(self):
        df = self.get_selected_data()
        if df is not None:
            path = filedialog.asksaveasfilename(
                defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")]
            )
            if path:
                try:
                    df.to_excel(path, index=False)
                    messagebox.showinfo("Success", f"Saved to {os.path.basename(path)}")
                    self.destroy()
                except Exception as e:
                    messagebox.showerror("Error", f"Excel error: {e}")

    def do_csv(self):
        df = self.get_selected_data()
        if df is not None:
            path = filedialog.asksaveasfilename(
                defaultextension=".csv", filetypes=[("CSV Files", "*.csv")]
            )
            if path:
                try:
                    df.to_csv(path, index=False, encoding="utf-8-sig")
                    messagebox.showinfo("Success", f"Saved to {os.path.basename(path)}")
                    self.destroy()
                except Exception as e:
                    messagebox.showerror("Error", f"CSV error: {e}")
