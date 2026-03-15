import customtkinter as ctk
import pandas as pd
import threading
import queue
import time
import os
import re
import random
import logging
from concurrent.futures import ThreadPoolExecutor
from tkinter import filedialog, messagebox, ttk
from app.scraper import CTOSScraper
from app.history import TaskHistoryManager
from app.dialogs import ResolutionDialog, PasteDialog, EditDialog, HistoryDialog, ExportDialog
import tkinter as tk

logger = logging.getLogger(__name__)

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

PENDING_STATUSES = {"Pending", "Searching...", "Processing...", "IN_PROGRESS"}


class CTOSApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("CTOS Company Scraper")
        self.geometry("1100x700")

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # State
        self.scraper = None
        self.stop_event = threading.Event()
        self.queue = queue.Queue()
        self.input_df = None
        self.output_data = []
        self.output_data_lock = threading.Lock()
        self.user_resolution_event = threading.Event()
        self.user_resolution_data = None
        self.processing_single = False
        self.paused = False
        self.pause_cond = threading.Condition()
        self.resolution_lock = threading.Lock()
        self.prog_lock = threading.Lock()
        self.prog_counter = 0

        self.history = TaskHistoryManager()
        self.current_session_id = None

        self._setup_sidebar()
        self._setup_main_area()

    def _setup_sidebar(self):
        self.sidebar = ctk.CTkFrame(self, width=250, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_rowconfigure(8, weight=1)

        ctk.CTkLabel(
            self.sidebar, text="CTOS Scraper", font=ctk.CTkFont(size=20, weight="bold")
        ).grid(row=0, column=0, padx=20, pady=20)

        # Single Search Section
        ctk.CTkLabel(self.sidebar, text="Single Search:", anchor="w").grid(
            row=1, column=0, padx=20, pady=(10, 0), sticky="w"
        )
        self.entry_single = ctk.CTkEntry(
            self.sidebar, placeholder_text="Enter Company Name"
        )
        self.entry_single.grid(row=2, column=0, padx=20, pady=5)

        self.btn_search_single = ctk.CTkButton(
            self.sidebar, text="Add One", command=self.add_single_item
        )
        self.btn_search_single.grid(row=3, column=0, padx=20, pady=5)

        ctk.CTkLabel(self.sidebar, text="or", text_color="gray").grid(row=4, column=0)

        # Batch Section
        self.btn_load = ctk.CTkButton(
            self.sidebar, text="Load CSV File", command=self.load_csv
        )
        self.btn_load.grid(row=5, column=0, padx=20, pady=5)

        self.btn_paste = ctk.CTkButton(
            self.sidebar, text="Paste Text List", command=self.open_paste_dialog
        )
        self.btn_paste.grid(row=6, column=0, padx=20, pady=5)

        # History Button
        self.btn_history = ctk.CTkButton(
            self.sidebar,
            text="History / Resume",
            command=self.open_history_dialog,
            fg_color="#555",
        )
        self.btn_history.grid(row=8, column=0, padx=20, pady=(20, 5))

        # Unattended Mode Checkbox
        self.unattended_var = ctk.BooleanVar(value=False)
        self.chk_unattended = ctk.CTkCheckBox(
            self.sidebar, text="Unattended Mode", variable=self.unattended_var
        )
        self.chk_unattended.grid(row=9, column=0, padx=20, pady=5)

        # Thread & Delay Settings
        self.settings_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        self.settings_frame.grid(row=10, column=0, padx=20, pady=5, sticky="ew")

        ctk.CTkLabel(self.settings_frame, text="Threads (1-5):").pack(anchor="w")
        self.threads_var = ctk.IntVar(value=1)
        self.slider_threads = ctk.CTkSlider(
            self.settings_frame,
            from_=1,
            to=5,
            number_of_steps=4,
            variable=self.threads_var,
        )
        self.slider_threads.pack(fill="x", pady=2)
        self.lbl_threads_val = ctk.CTkLabel(
            self.settings_frame, textvariable=self.threads_var
        )
        self.lbl_threads_val.pack()

        ctk.CTkLabel(self.settings_frame, text="Delay (0-5s):").pack(anchor="w")
        self.delay_var = ctk.IntVar(value=0)
        self.slider_delay = ctk.CTkSlider(
            self.settings_frame,
            from_=0,
            to=5,
            number_of_steps=5,
            variable=self.delay_var,
        )
        self.slider_delay.pack(fill="x", pady=2)
        self.lbl_delay_val = ctk.CTkLabel(
            self.settings_frame, textvariable=self.delay_var
        )
        self.lbl_delay_val.pack()

        self.btn_start_batch = ctk.CTkButton(
            self.sidebar,
            text="Start",
            command=self.start_batch_search,
            fg_color="green",
            state="disabled",
        )
        self.btn_start_batch.grid(row=11, column=0, padx=20, pady=5)

        self.btn_stop = ctk.CTkButton(
            self.sidebar,
            text="Stop",
            command=self.stop_process,
            fg_color="red",
            state="disabled",
        )
        self.btn_stop.grid(row=12, column=0, padx=20, pady=5)

        self.btn_pause = ctk.CTkButton(
            self.sidebar,
            text="Pause",
            command=self.toggle_pause,
            fg_color="orange",
            state="disabled",
        )
        self.btn_pause.grid(row=13, column=0, padx=20, pady=5)

        self.btn_export = ctk.CTkButton(
            self.sidebar, text="Export Results", command=self.export_csv, state="disabled"
        )
        self.btn_export.grid(row=14, column=0, padx=20, pady=20, sticky="s")

    def _setup_main_area(self):
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)
        self.main_frame.grid_rowconfigure(1, weight=1)
        self.main_frame.grid_columnconfigure(0, weight=1)

        # Status Bar
        self.status_bar = ctk.CTkFrame(self.main_frame, height=30)
        self.status_bar.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        self.lbl_status = ctk.CTkLabel(self.status_bar, text="Ready")
        self.lbl_status.pack(side="left", padx=10)

        self.progress_bar = ctk.CTkProgressBar(self.status_bar)
        self.progress_bar.set(0)
        self.progress_bar.pack(side="right", padx=10, pady=5)

        # Treeview Table
        self.style = ttk.Style()
        self.style.theme_use("default")
        self.style.configure(
            "Treeview",
            background="#2b2b2b",
            foreground="white",
            fieldbackground="#2b2b2b",
            rowheight=35,
            font=("Arial", 13),
        )
        self.style.configure(
            "Treeview.Heading", font=("Arial", 14, "bold"), rowheight=40
        )
        self.style.map("Treeview", background=[("selected", "#1f538d")])

        columns = (
            "input", "status", "brn", "old_reg",
            "new_reg", "found_name", "type", "confidence",
        )
        self.tree = ttk.Treeview(
            self.main_frame, columns=columns, show="headings", selectmode="browse"
        )

        self.tree.heading("input", text="Input Company")
        self.tree.heading("status", text="Status")
        self.tree.heading("brn", text="BRN")
        self.tree.heading("old_reg", text="Old Reg No")
        self.tree.heading("new_reg", text="New Reg No")
        self.tree.heading("found_name", text="Found Name")
        self.tree.heading("type", text="Type")
        self.tree.heading("confidence", text="Confidence")

        self.tree.column("input", width=250)
        self.tree.column("status", width=100)
        self.tree.column("brn", width=150)
        self.tree.column("old_reg", width=100)
        self.tree.column("new_reg", width=120)
        self.tree.column("found_name", width=250)
        self.tree.column("type", width=80)
        self.tree.column("confidence", width=120)

        self.all_columns = list(columns)
        self.visible_columns = [
            "input", "status", "brn", "found_name", "type", "confidence",
        ]
        self.tree["displaycolumns"] = self.visible_columns

        self.tree.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)

        # Header Menu
        self.header_menu = tk.Menu(self, tearoff=0)

        # Scrollbar
        scrollbar = ttk.Scrollbar(
            self.main_frame, orient=tk.VERTICAL, command=self.tree.yview
        )
        self.tree.configure(yscroll=scrollbar.set)
        scrollbar.grid(row=1, column=1, sticky="ns")

        # Context Menu
        self.menu = tk.Menu(self, tearoff=0)
        self.menu.add_command(label="Edit Data", command=self.edit_selected)
        self.menu.add_command(
            label="Re-scrape This Company", command=self.rescrape_selected
        )
        self.menu.add_command(
            label="Export Selected (Custom)", command=self.export_selected_custom
        )
        self.menu.add_separator()
        self.menu.add_command(label="Delete Row", command=self.delete_selected)
        self.tree.bind("<Button-3>", self.on_tree_right_click)

    def on_tree_right_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region == "heading":
            self.show_header_menu(event)
        elif region == "cell":
            self.show_context_menu(event)

    def show_header_menu(self, event):
        self.header_menu.delete(0, "end")
        for col in self.all_columns:
            curr_visible = self.tree["displaycolumns"]
            if curr_visible == ("#all",):
                is_vis = True
            else:
                is_vis = col in curr_visible

            label = f"{col} {'(Visible)' if is_vis else ''}"
            self.header_menu.add_command(
                label=label, command=lambda c=col: self.toggle_column(c)
            )
        self.header_menu.post(event.x_root, event.y_root)

    def toggle_column(self, col):
        current = list(self.tree["displaycolumns"])
        if current == ["#all"]:
            current = list(self.all_columns)

        if col in current:
            current.remove(col)
        else:
            current.append(col)
            current.sort(key=lambda x: self.all_columns.index(x))

        self.tree["displaycolumns"] = current

    def show_context_menu(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.menu.post(event.x_root, event.y_root)

    def edit_selected(self):
        selected_item = self.tree.selection()
        if not selected_item:
            return
        values = self.tree.item(selected_item, "values")
        EditDialog(
            self, values, lambda new_vals: self.update_row_data(selected_item, new_vals)
        )

    def update_row_data(self, item, new_vals):
        self.tree.item(item, values=new_vals)

    def export_selected_custom(self):
        selected = self.tree.selection()
        if not selected:
            return
        vals = self.tree.item(selected, "values")
        company_name = vals[0]

        row_data = None
        with self.output_data_lock:
            for item in self.output_data:
                if item.get("Company Name") == company_name:
                    row_data = item
                    break

        if row_data:
            cols = [
                "Company Name", "Status", "BRN", "Old Reg No",
                "New Reg No", "Found Name", "Type", "Confidence", "Details",
            ]
            ExportDialog(self, [row_data], cols)
        else:
            messagebox.showwarning(
                "Warning", "Could not find detailed data for export."
            )

    def rescrape_selected(self):
        selected_item = self.tree.selection()
        if not selected_item:
            return
        values = self.tree.item(selected_item, "values")
        company_name = values[0]

        if (
            self.btn_start_batch.cget("state") == "disabled"
            and not self.processing_single
        ):
            messagebox.showwarning(
                "Busy",
                "Cannot re-scrape while batch is running. Please Pause or Stop first.",
            )
            return

        self.entry_single.delete(0, "end")
        self.entry_single.insert(0, company_name)

        self.tree.item(
            selected_item, values=(company_name, "Searching...", "", "", "", "", "", "")
        )

        self.run_process_single(company_name, selected_item[0])

    def delete_selected(self):
        selected_item = self.tree.selection()
        if not selected_item:
            return
        vals = self.tree.item(selected_item, "values")
        company_name = vals[0]
        self.tree.delete(selected_item)

        with self.output_data_lock:
            self.output_data = [
                item
                for item in self.output_data
                if item.get("Company Name") != company_name
            ]

    def open_history_dialog(self):
        HistoryDialog(self, self.history, self.resume_session)

    def resume_session(self, session_id):
        results = self.history.get_session_results(session_id)
        if not results:
            return

        self.current_session_id = session_id
        self.lbl_status.configure(text=f"Resuming Session {session_id}...")

        self.tree.delete(*self.tree.get_children())
        with self.output_data_lock:
            self.output_data = []

        companies_to_process = []
        item_ids_to_process = []

        for res in results:
            comp_input = res["company_input"]
            status = res["status"]
            reg = res["reg_no"]
            data = res["data"]

            old_reg, new_reg = "", ""
            if "/" in reg:
                parts = reg.split("/")
                old_reg = parts[0].strip()
                new_reg = parts[1].strip() if len(parts) > 1 else ""
            else:
                old_reg = reg

            fname = data.get("name", "")
            ctype = data.get("type", "")
            conf = data.get("confidence", "History")

            iid = self.tree.insert(
                "",
                "end",
                values=(comp_input, status, reg, old_reg, new_reg, fname, ctype, conf),
            )

            with self.output_data_lock:
                self.output_data.append(
                    {
                        "Company Name": comp_input,
                        "Status": status,
                        "Confidence": conf,
                        "BRN": reg,
                        "Old Reg No": old_reg,
                        "New Reg No": new_reg,
                        "Found Name": fname,
                        "Type": ctype,
                        "Details": str(data),
                    }
                )

            if status in PENDING_STATUSES:
                companies_to_process.append(comp_input)
                item_ids_to_process.append(iid)

        if companies_to_process:
            self.processing_single = False
            self.stop_event.clear()
            self.paused = False
            self.btn_stop.configure(state="normal")
            self.btn_pause.configure(state="normal", text="Pause")
            self.btn_start_batch.configure(state="disabled")
            self.btn_search_single.configure(state="disabled")
            self.btn_load.configure(state="disabled")

            t = threading.Thread(
                target=self.run_process,
                args=(companies_to_process, item_ids_to_process),
                daemon=True,
            )
            t.start()
            self.after(100, self.check_queue)
        else:
            messagebox.showinfo(
                "Resume", "All items in this session are already completed."
            )

        if self.output_data:
            self.btn_export.configure(state="normal")

    def open_paste_dialog(self):
        PasteDialog(self, self.add_companies_from_text)

    def add_companies_from_text(self, company_list):
        if not company_list:
            return

        new_df = pd.DataFrame({"Company Name": company_list})

        if self.input_df is None:
            self.input_df = new_df
        else:
            self.input_df = pd.concat([self.input_df, new_df], ignore_index=True)

        self.lbl_status.configure(
            text=f"Added {len(company_list)} companies. Total: {len(self.input_df)}"
        )
        self.btn_start_batch.configure(state="normal")

        for comp in company_list:
            self.tree.insert("", "end", values=(comp, "Pending", "", "", "", "", ""))

    def toggle_pause(self):
        with self.pause_cond:
            self.paused = not self.paused
            if self.paused:
                self.btn_pause.configure(text="Resume")
                self.lbl_status.configure(text="Paused")
            else:
                self.btn_pause.configure(text="Pause")
                self.lbl_status.configure(text="Processing...")
                self.pause_cond.notify()

    def load_csv(self):
        path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if path:
            try:
                self.input_df = pd.read_csv(path, header=None)
                first_row = self.input_df.iloc[0].astype(str).str.lower().tolist()
                if "company name" in first_row:
                    self.input_df = pd.read_csv(path)
                    if "Company Name" not in self.input_df.columns:
                        self.input_df.columns = [c.title() for c in self.input_df.columns]
                else:
                    self.input_df = pd.read_csv(path, header=None)
                    self.input_df.rename(columns={0: "Company Name"}, inplace=True)

                self.lbl_status.configure(
                    text=f"Loaded {len(self.input_df)} rows from {os.path.basename(path)}"
                )
                self.btn_start_batch.configure(state="normal")

                for item in self.tree.get_children():
                    self.tree.delete(item)

                for _, row in self.input_df.iterrows():
                    self.tree.insert(
                        "",
                        "end",
                        values=(row.get("Company Name", ""), "Pending", "", "", "", "", ""),
                    )

            except Exception as e:
                messagebox.showerror("Error", f"Failed to read file: {e}")

    def add_single_item(self):
        term = self.entry_single.get().strip()
        if not term:
            return

        self.tree.insert("", 0, values=(term, "Pending", "", "", "", "", "", ""))
        self.entry_single.delete(0, "end")

        self.btn_start_batch.configure(state="normal")

    def run_process_single(self, company, item_id):
        self.processing_single = True
        self.btn_search_single.configure(state="disabled")
        self.btn_start_batch.configure(state="disabled")
        self.btn_load.configure(state="disabled")

        t = threading.Thread(
            target=self.run_process, args=([company], [item_id]), daemon=True
        )
        t.start()
        self.after(100, self.check_queue)

    def start_batch_search(self):
        companies_to_req = []
        item_ids_to_req = []

        for item in self.tree.get_children():
            vals = self.tree.item(item, "values")
            status = vals[1]
            if status == "Pending":
                companies_to_req.append(vals[0])
                item_ids_to_req.append(item)

        if not companies_to_req:
            messagebox.showinfo("Info", "No pending items to search.")
            return

        self.current_session_id = self.history.create_session(
            "Batch " + str(int(time.time())), len(companies_to_req), companies_to_req
        )
        self.lbl_status.configure(text=f"Session {self.current_session_id} Started.")

        self.processing_single = False
        self.stop_event.clear()
        self.paused = False
        self.btn_stop.configure(state="normal")
        self.btn_pause.configure(state="normal", text="Pause")
        self.btn_start_batch.configure(state="disabled")
        self.btn_search_single.configure(state="disabled")
        self.btn_load.configure(state="disabled")

        t = threading.Thread(
            target=self.run_process,
            args=(companies_to_req, item_ids_to_req),
            daemon=True,
        )
        t.start()
        self.after(100, self.check_queue)

    def stop_process(self):
        self.stop_event.set()
        self.btn_stop.configure(state="disabled", text="Stopping...")

    def run_process(self, companies, item_ids):
        max_workers = self.threads_var.get()
        delay_sec = self.delay_var.get()
        total = len(companies)

        task_queue = queue.Queue()
        for i, (comp, iid) in enumerate(zip(companies, item_ids)):
            task_queue.put((i, comp, iid))

        self.prog_counter = 0

        def worker_func():
            try:
                scraper = CTOSScraper(headless=False)
            except Exception as e:
                logger.error("Failed to init scraper: %s", e)
                return

            try:
                self._worker_loop(scraper, task_queue, total, delay_sec)
            finally:
                scraper.close_driver()

        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = [executor.submit(worker_func) for _ in range(max_workers)]
            for f in futures:
                f.result()

        self.queue.put(("DONE", None))

    def _worker_loop(self, scraper, task_queue, total, delay_sec):
        """Main loop for a single worker thread."""
        while not task_queue.empty() and not self.stop_event.is_set():
            try:
                idx, company, item_id = task_queue.get(timeout=1)
            except queue.Empty:
                break

            self.queue.put(
                ("UPDATE_ROW", (item_id, company, "Processing...", "", "", "", "", "", ""))
            )

            if self.paused:
                self.queue.put(
                    ("UPDATE_ROW", (item_id, company, "Paused", "", "", "", "", "", ""))
                )
                with self.pause_cond:
                    self.pause_cond.wait_for(
                        lambda: not self.paused or self.stop_event.is_set()
                    )

            if self.stop_event.is_set():
                return

            if delay_sec > 0:
                time.sleep(random.uniform(0.5, delay_sec))

            final_result, status_text, confidence_level = self._search_single_company(
                scraper, company
            )

            reg_raw = ""
            if final_result:
                reg_raw = final_result.get("reg_no", "") or final_result.get("req_no", "")
            old_reg, new_reg = "", ""
            if "/" in reg_raw:
                parts = reg_raw.split("/")
                old_reg = parts[0].strip()
                new_reg = parts[1].strip() if len(parts) > 1 else ""
            else:
                old_reg = reg_raw

            fname = final_result.get("name", "") if final_result else ""
            ctype = final_result.get("type", "") if final_result else ""

            self.queue.put(
                (
                    "UPDATE_ROW",
                    (
                        item_id, company, status_text, reg_raw,
                        old_reg, new_reg, fname, ctype, confidence_level,
                    ),
                )
            )

            if self.current_session_id:
                self.history.update_result(
                    self.current_session_id,
                    company,
                    status_text,
                    reg_raw,
                    final_result if final_result else {},
                )

            with self.output_data_lock:
                self.output_data.append(
                    {
                        "Company Name": company,
                        "Status": status_text,
                        "Confidence": confidence_level,
                        "BRN": reg_raw,
                        "Old Reg No": old_reg,
                        "New Reg No": new_reg,
                        "Found Name": fname,
                        "Type": ctype,
                        "Details": str(final_result),
                    }
                )

            with self.prog_lock:
                self.prog_counter += 1
                self.queue.put(("PROG", self.prog_counter / total))

    def _search_single_company(self, scraper, company):
        """Searches for one company, handling ambiguity. Returns (result, status, confidence)."""
        search_term = company
        unattended = self.unattended_var.get()

        while True:
            search_res = scraper.search_company(search_term)

            if search_res["status"] in ["FOUND", "FOUND_PARTIAL"]:
                return search_res["data"], "Found", "High (Direct)"

            if search_res["status"] == "AMBIGUOUS":
                return self._handle_ambiguous(scraper, company, search_res, unattended)

            if search_res["status"] == "NOT_FOUND":
                if unattended:
                    return None, "Not Found", "None"

                action, data = self._request_resolution(company, [])
                if action == "MANUAL_SEARCH":
                    search_term = data
                    continue
                return None, "Not Found", "None"

            return None, "Error", "None"

    def _handle_ambiguous(self, scraper, company, search_res, unattended):
        """Handle ambiguous results. Returns (final_result, status, confidence)."""
        candidates = search_res["data"]
        comps = [c for c in candidates if "COMPANY" in c.get("type", "").upper()]
        selected_candidate = None
        auto_method = ""

        if len(comps) == 1:
            selected_candidate = comps[0]
            auto_method = "High (Unique Co)"
        elif len(comps) > 1:
            selected_candidate, auto_method = self._try_auto_match(company, comps)

        if selected_candidate:
            final_result = self._resolve_candidate(scraper, selected_candidate)
            final_result["confidence"] = auto_method
            return final_result, "Found (Auto)", auto_method

        if unattended:
            return None, "Ambiguous (Skipped)", "Low"

        action, data = self._request_resolution(company, search_res["data"])

        if action == "SELECT":
            final_result = self._resolve_candidate(scraper, data)
            return final_result, "Found (Manual)", "Manual"
        elif action == "MANUAL_SEARCH":
            new_res = scraper.search_company(data)
            if new_res["status"] in ["FOUND", "FOUND_PARTIAL"]:
                return new_res["data"], "Found", "High (Direct)"
            if new_res["status"] == "AMBIGUOUS":
                return self._handle_ambiguous(scraper, company, new_res, unattended)
            return None, "Not Found", "None"
        else:
            return None, "Skipped", "None"

    def _try_auto_match(self, company, comps):
        """Try strict then loose name matching. Returns (candidate, method) or (None, '')."""

        def normalize_base(txt):
            txt = txt.upper()
            txt = re.sub(r"SDN\.?\s*BHD\.?", "", txt, flags=re.IGNORECASE).strip()
            txt = re.sub(r"PLT", "", txt, flags=re.IGNORECASE).strip()
            txt = re.sub(r"ENTERPRISE", "", txt, flags=re.IGNORECASE).strip()
            return txt

        def clean_symbols(txt):
            return re.sub(r"\W+", "", txt)

        def remove_parens(txt):
            return re.sub(r"\(.*?\)", "", txt)

        norm_input_strict = clean_symbols(normalize_base(company))
        strict_matches = [
            c
            for c in comps
            if clean_symbols(normalize_base(c["name"])) == norm_input_strict
        ]

        if len(strict_matches) == 1:
            return strict_matches[0], "High (Strict Match)"

        norm_input_loose = clean_symbols(remove_parens(normalize_base(company)))
        loose_matches = [
            c
            for c in comps
            if clean_symbols(remove_parens(normalize_base(c["name"]))) == norm_input_loose
        ]

        if len(loose_matches) == 1:
            return loose_matches[0], "High (Loose Match)"

        return None, ""

    def _resolve_candidate(self, scraper, candidate):
        """Click a candidate for details or return summary data."""
        if scraper.fast_mode:
            return CTOSScraper._without_element(candidate)

        click_res = scraper._click_and_scrape_details(candidate)
        if click_res["status"] in ["FOUND", "FOUND_PARTIAL"]:
            return click_res["data"]
        return CTOSScraper._without_element(candidate)

    def _request_resolution(self, company, candidates):
        """Thread-safe interactive resolution via the main thread GUI."""
        with self.resolution_lock:
            self.queue.put(("RESOLVE", (company, candidates)))
            self.user_resolution_event.wait()
            self.user_resolution_event.clear()
            return self.user_resolution_data

    def check_queue(self):
        try:
            while True:
                msg_type, data = self.queue.get_nowait()

                if msg_type == "UPDATE_ROW":
                    item_id, col1, col2, col3, col4, col5, col6, col7, col8 = data
                    self.tree.item(
                        item_id, values=(col1, col2, col3, col4, col5, col6, col7, col8)
                    )
                    self.tree.see(item_id)
                elif msg_type == "PROG":
                    self.progress_bar.set(data)
                elif msg_type == "RESOLVE":
                    ResolutionDialog(self, *data, self.handle_resolution)
                elif msg_type == "DONE":
                    self.reset_ui()
                    messagebox.showinfo("Completed", "Process Finished")
                    return
        except queue.Empty:
            pass
        self.after(100, self.check_queue)

    def handle_resolution(self, action, data):
        self.user_resolution_data = (action, data)
        self.user_resolution_event.set()

    def reset_ui(self):
        self.btn_start_batch.configure(state="normal")
        self.btn_search_single.configure(state="normal")
        self.btn_load.configure(state="normal")
        self.btn_stop.configure(state="disabled", text="Stop")
        self.btn_pause.configure(state="disabled", text="Pause")
        self.btn_export.configure(state="normal")
        self.progress_bar.set(0)

    def export_csv(self):
        with self.output_data_lock:
            if not self.output_data:
                messagebox.showwarning("Warning", "No data to export.")
                return
            data_copy = list(self.output_data)

        cols = [
            "Company Name", "Status", "BRN", "Old Reg No",
            "New Reg No", "Found Name", "Type", "Confidence", "Details",
        ]
        ExportDialog(self, data_copy, cols)


