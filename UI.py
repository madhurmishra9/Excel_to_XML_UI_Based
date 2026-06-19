"""Tkinter front-end for the excel_xml backend (../Excel_to_xml).

Surfaces every operation the backend exposes:

  * Excel -> XML   (excel_to_xml)
  * XML -> Excel   (xml_to_excel)
  * Compare        (check_xml_data)
  * Fetch remote   (remote_to_excel / remote_to_xml)
  * File search    (file_search)

The backend is expected in a sibling folder named ``Excel_to_xml``.  Set the
``EXCEL_XML_BACKEND`` environment variable to point elsewhere if needed.

The backend's heavy dependencies (openpyxl, xlrd, beautifulsoup4, lxml,
requests, feedparser) are imported lazily inside the backend, so importing the
package here is cheap -- but an operation will fail if its dependency is not
installed.  Run this UI with the backend's virtual environment (or otherwise
install ``../Excel_to_xml/requirements.txt``) so those imports succeed.
"""

import contextlib
import io
import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText


# --------------------------------------------------------------------------- #
# Locate and import the excel_xml backend.
# --------------------------------------------------------------------------- #
def _add_backend_to_path():
    """Put the folder that contains the ``excel_xml`` package on ``sys.path``."""
    candidates = []
    override = os.environ.get("EXCEL_XML_BACKEND")
    if override:
        candidates.append(override)
    here = os.path.dirname(os.path.abspath(__file__))
    candidates.append(os.path.normpath(os.path.join(here, "..", "Excel_to_xml")))
    for path in candidates:
        if os.path.isdir(os.path.join(path, "excel_xml")):
            if path not in sys.path:
                sys.path.insert(0, path)
            return path
    return None


_BACKEND_DIR = _add_backend_to_path()

try:
    from excel_xml import (
        check_xml_data,
        excel_to_xml,
        file_search,
        remote_to_excel,
        remote_to_xml,
        xml_to_excel,
    )
except Exception as exc:  # backend missing or broken -- fail with a clear message
    _root = tk.Tk()
    _root.withdraw()
    messagebox.showerror(
        "excel_xml backend not found",
        "Could not import the 'excel_xml' backend.\n\n"
        "Expected it in a sibling folder named 'Excel_to_xml', or set the "
        "EXCEL_XML_BACKEND environment variable to the folder that contains "
        "the 'excel_xml' package.\n\n"
        "Import error: %s" % exc,
    )
    raise SystemExit(1)


EXCEL_OPEN_TYPES = [("Excel files", "*.xlsx *.xlsm *.xls"), ("All files", "*.*")]
EXCEL_SAVE_TYPES = [("Excel workbook", "*.xlsx"), ("All files", "*.*")]
FETCH_KINDS = ["auto", "rss", "xml", "html", "data"]


class ConverterApp:
    """The whole UI: a notebook of operations sharing one output log."""

    def __init__(self, root):
        self.root = root
        self._busy = False
        self._action_buttons = []

        root.title("Excel ↔ XML Converter")
        root.geometry("780x580")
        root.minsize(640, 480)

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(side="top", fill="x", padx=8, pady=(8, 4))

        self._build_excel_to_xml_tab()
        self._build_xml_to_excel_tab()
        self._build_compare_tab()
        self._build_fetch_tab()
        self._build_search_tab()

        self._build_output(root)

    # ------------------------------------------------------------------ #
    # Shared widgets / helpers
    # ------------------------------------------------------------------ #
    def _build_output(self, root):
        self.status = tk.StringVar(value="Ready")
        ttk.Label(root, textvariable=self.status, relief="sunken", anchor="w").pack(
            side="bottom", fill="x"
        )

        frame = ttk.LabelFrame(root, text="Output")
        frame.pack(side="bottom", fill="both", expand=True, padx=8, pady=4)

        self.output = ScrolledText(frame, height=12, state="disabled", wrap="word")
        self.output.pack(side="left", fill="both", expand=True, padx=(4, 0), pady=4)

        ttk.Button(frame, text="Clear", command=self.clear_log).pack(
            side="right", anchor="n", padx=4, pady=4
        )

    def _path_row(self, parent, label, row, browse_cmd):
        """Add a 'label / entry / Browse' row and return its StringVar."""
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=4, pady=6)
        var = tk.StringVar()
        ttk.Entry(parent, textvariable=var).grid(
            row=row, column=1, sticky="we", padx=4, pady=6
        )
        ttk.Button(parent, text="Browse", command=lambda: browse_cmd(var)).grid(
            row=row, column=2, padx=4, pady=6
        )
        return var

    def _add_action(self, parent, text, row, command):
        btn = ttk.Button(parent, text=text, command=command)
        btn.grid(row=row, column=1, sticky="e", padx=4, pady=10)
        self._action_buttons.append(btn)
        return btn

    # ---- file dialogs ------------------------------------------------- #
    def _browse_excel_open(self, var):
        path = filedialog.askopenfilename(
            title="Select Excel file", filetypes=EXCEL_OPEN_TYPES
        )
        if path:
            var.set(path)

    def _browse_excel_save(self, var):
        path = filedialog.asksaveasfilename(
            title="Save Excel workbook as",
            defaultextension=".xlsx",
            filetypes=EXCEL_SAVE_TYPES,
        )
        if path:
            var.set(path)

    def _browse_folder(self, var):
        path = filedialog.askdirectory(title="Select folder")
        if path:
            var.set(path)

    # ---- logging / status --------------------------------------------- #
    def log(self, text):
        self.output.configure(state="normal")
        self.output.insert("end", text + "\n")
        self.output.see("end")
        self.output.configure(state="disabled")

    def clear_log(self):
        self.output.configure(state="normal")
        self.output.delete("1.0", "end")
        self.output.configure(state="disabled")

    def set_status(self, text):
        self.status.set(text)

    # ---- threaded execution ------------------------------------------- #
    def run_async(self, label, work, on_success):
        """Run ``work()`` off the UI thread.

        Anything ``work`` prints to stdout is captured and written to the log
        (the backend's compare/search report results that way).  On success
        ``on_success(result)`` runs on the main thread; on failure the
        exception is shown in a dialog and logged.
        """
        if self._busy:
            return
        self._set_busy(True)
        self.set_status("%s ..." % label)

        def worker():
            buf = io.StringIO()
            try:
                with contextlib.redirect_stdout(buf):
                    result = work()
            except Exception as exc:  # surfaced on the main thread below
                self.root.after(0, self._finish, label, on_success, None, buf.getvalue(), exc)
                return
            self.root.after(0, self._finish, label, on_success, result, buf.getvalue(), None)

        threading.Thread(target=worker, daemon=True).start()

    def _finish(self, label, on_success, result, captured, exc):
        self._set_busy(False)
        if captured.strip():
            self.log(captured.rstrip())
        if exc is not None:
            self.set_status("%s failed" % label)
            self.log("ERROR: %s" % exc)
            messagebox.showerror(label, str(exc))
            return
        on_success(result)
        self.set_status("%s done" % label)

    def _set_busy(self, busy):
        self._busy = busy
        state = "disabled" if busy else "normal"
        for btn in self._action_buttons:
            btn.configure(state=state)

    # ------------------------------------------------------------------ #
    # Tabs
    # ------------------------------------------------------------------ #
    def _build_excel_to_xml_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Excel → XML")
        tab.columnconfigure(1, weight=1)

        excel_var = self._path_row(tab, "Excel file:", 0, self._browse_excel_open)
        out_var = self._path_row(tab, "Output folder:", 1, self._browse_folder)
        ttk.Label(
            tab, text="One .xml file per sheet. Blank folder -> <name>_xml in the working directory."
        ).grid(row=2, column=1, sticky="w", padx=4)

        def run():
            excel = excel_var.get().strip()
            if not excel:
                messagebox.showwarning("Missing input", "Select an Excel file to convert.")
                return
            out = out_var.get().strip()
            self.run_async(
                "Excel -> XML",
                lambda: excel_to_xml(excel, out),
                lambda written: self.log(
                    "Wrote %d XML file(s):\n  %s" % (len(written), "\n  ".join(written))
                ),
            )

        self._add_action(tab, "Convert", 3, run)

    def _build_xml_to_excel_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="XML → Excel")
        tab.columnconfigure(1, weight=1)

        xml_var = self._path_row(tab, "XML folder:", 0, self._browse_folder)
        out_var = self._path_row(tab, "Output Excel:", 1, self._browse_excel_save)
        ttk.Label(
            tab, text="Combines every .xml in the folder into one workbook (one sheet per file)."
        ).grid(row=2, column=1, sticky="w", padx=4)

        def run():
            xml_dir = xml_var.get().strip()
            if not xml_dir:
                messagebox.showwarning(
                    "Missing input", "Select the folder containing the XML files."
                )
                return
            out = out_var.get().strip()
            self.run_async(
                "XML -> Excel",
                lambda: xml_to_excel(xml_dir, out),
                lambda path: self.log("Wrote %s" % path),
            )

        self._add_action(tab, "Convert", 3, run)

    def _build_compare_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Compare")
        tab.columnconfigure(1, weight=1)

        excel_var = self._path_row(tab, "Excel file:", 0, self._browse_excel_open)
        xml_var = self._path_row(tab, "XML folder:", 1, self._browse_folder)
        ttk.Label(
            tab, text="Reports any cell that differs between the Excel file and the XML folder."
        ).grid(row=2, column=1, sticky="w", padx=4)

        def run():
            excel = excel_var.get().strip()
            xml_dir = xml_var.get().strip()
            if not excel or not xml_dir:
                messagebox.showwarning(
                    "Missing input", "Select both an Excel file and an XML folder."
                )
                return
            self.clear_log()
            self.log("Comparing\n  %s\nagainst\n  %s\n" % (excel, xml_dir))
            self.run_async(
                "Compare",
                lambda: check_xml_data(excel, xml_dir),  # prints diffs to stdout
                self._compare_done,
            )

        self._add_action(tab, "Compare", 3, run)

    def _compare_done(self, result):
        if result == 0:
            self.log("100% Match")
            messagebox.showinfo("Compare", "100% Match")
        else:
            self.log("\nDifferences found (listed above).")

    def _build_fetch_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Fetch")
        tab.columnconfigure(1, weight=1)

        ttk.Label(tab, text="Source URL / URI:").grid(
            row=0, column=0, sticky="w", padx=4, pady=6
        )
        source_var = tk.StringVar()
        ttk.Entry(tab, textvariable=source_var).grid(
            row=0, column=1, columnspan=2, sticky="we", padx=4, pady=6
        )

        ttk.Label(tab, text="Kind:").grid(row=1, column=0, sticky="w", padx=4, pady=6)
        kind_var = tk.StringVar(value="auto")
        ttk.Combobox(
            tab, textvariable=kind_var, values=FETCH_KINDS, state="readonly", width=12
        ).grid(row=1, column=1, sticky="w", padx=4, pady=6)

        ttk.Label(tab, text="Target:").grid(row=2, column=0, sticky="w", padx=4, pady=6)
        self._fetch_target_var = tk.StringVar(value="excel")
        target_frame = ttk.Frame(tab)
        target_frame.grid(row=2, column=1, sticky="w", padx=4, pady=6)
        ttk.Radiobutton(
            target_frame, text="Excel workbook", variable=self._fetch_target_var,
            value="excel", command=self._update_fetch_out_label,
        ).pack(side="left")
        ttk.Radiobutton(
            target_frame, text="XML folder", variable=self._fetch_target_var,
            value="xml", command=self._update_fetch_out_label,
        ).pack(side="left", padx=(10, 0))

        self._fetch_out_label = ttk.Label(tab, text="Output Excel:")
        self._fetch_out_label.grid(row=3, column=0, sticky="w", padx=4, pady=6)
        out_var = tk.StringVar()
        ttk.Entry(tab, textvariable=out_var).grid(
            row=3, column=1, sticky="we", padx=4, pady=6
        )

        def browse_out():
            if self._fetch_target_var.get() == "excel":
                self._browse_excel_save(out_var)
            else:
                self._browse_folder(out_var)

        ttk.Button(tab, text="Browse", command=browse_out).grid(
            row=3, column=2, padx=4, pady=6
        )
        ttk.Label(
            tab,
            text="RSS/Atom, arbitrary XML, HTML tables, or a cloud URI "
            "(s3://, gs://, az://, gdrive://). Blank output uses a default name.",
        ).grid(row=4, column=1, sticky="w", padx=4)

        def run():
            source = source_var.get().strip()
            if not source:
                messagebox.showwarning("Missing input", "Enter a URL or cloud URI to fetch.")
                return
            kind = kind_var.get()
            out = out_var.get().strip()
            if self._fetch_target_var.get() == "excel":
                self.run_async(
                    "Fetch -> Excel",
                    lambda: remote_to_excel(source, out, kind),
                    lambda path: self.log("Wrote %s" % path),
                )
            else:
                self.run_async(
                    "Fetch -> XML",
                    lambda: remote_to_xml(source, out, kind),
                    lambda written: self.log(
                        "Wrote %d XML file(s):\n  %s" % (len(written), "\n  ".join(written))
                    ),
                )

        self._add_action(tab, "Fetch", 5, run)

    def _update_fetch_out_label(self):
        if self._fetch_target_var.get() == "excel":
            self._fetch_out_label.configure(text="Output Excel:")
        else:
            self._fetch_out_label.configure(text="Output folder:")

    def _build_search_tab(self):
        tab = ttk.Frame(self.notebook)
        self.notebook.add(tab, text="Search")
        tab.columnconfigure(1, weight=1)

        ttk.Label(tab, text="File name:").grid(row=0, column=0, sticky="w", padx=4, pady=6)
        name_var = tk.StringVar()
        ttk.Entry(tab, textvariable=name_var).grid(
            row=0, column=1, sticky="we", padx=4, pady=6
        )
        ttk.Label(
            tab, text="Exact file name; searches every local drive (Windows only)."
        ).grid(row=1, column=1, sticky="w", padx=4)

        def run():
            name = name_var.get().strip()
            if not name:
                messagebox.showwarning("Missing input", "Enter a file name to search for.")
                return
            self.clear_log()
            self.log("Searching all drives for %r ..." % name)
            self.run_async(
                "Search",
                lambda: file_search(name),
                lambda results: self.log("\n%d match(es)." % len(results)),
            )

        self._add_action(tab, "Search", 2, run)


def main():
    root = tk.Tk()
    ConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
