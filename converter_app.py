"""Simple Tkinter application to batch convert Microsoft Word documents to PDF using Word COM automation."""

from __future__ import annotations

import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText
from typing import Iterable
import traceback

import pythoncom
import win32com.client


class WordToPDFConverterApp:
    """Tkinter GUI for selecting Word files and converting them to PDF."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Word to PDF Converter")
        self.root.geometry("600x480")
        self.root.resizable(False, False)

        self.selected_files: list[str] = []
        self.output_dir_var = tk.StringVar()
        self.progress_var = tk.DoubleVar(value=0.0)
        self.status_var = tk.StringVar(value="Select Word files to begin.")
        self._conversion_thread: threading.Thread | None = None

        self._build_widgets()

    def _build_widgets(self) -> None:
        padding = {"padx": 10, "pady": 5}

        button_frame = tk.Frame(self.root)
        button_frame.pack(fill=tk.X, **padding)

        add_button = tk.Button(button_frame, text="Add Word Files", command=self.add_files)
        add_button.pack(side=tk.LEFT)

        convert_button = tk.Button(button_frame, text="Start Conversion", command=self.start_conversion)
        convert_button.pack(side=tk.LEFT, padx=(10, 0))

        clear_button = tk.Button(button_frame, text="Clear List", command=self.clear_file_list)
        clear_button.pack(side=tk.LEFT, padx=(10, 0))

        file_frame = tk.LabelFrame(self.root, text="Selected Files")
        file_frame.pack(fill=tk.BOTH, expand=True, **padding)

        self.file_listbox = tk.Listbox(file_frame, height=10)
        self.file_listbox.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        output_frame = tk.Frame(self.root)
        output_frame.pack(fill=tk.X, **padding)

        output_label = tk.Label(output_frame, text="Output Folder:")
        output_label.pack(side=tk.LEFT)

        output_entry = tk.Entry(output_frame, textvariable=self.output_dir_var, width=45)
        output_entry.pack(side=tk.LEFT, padx=(5, 5))

        browse_button = tk.Button(output_frame, text="Browse", command=self.choose_output_directory)
        browse_button.pack(side=tk.LEFT)

        progress_frame = tk.Frame(self.root)
        progress_frame.pack(fill=tk.X, **padding)

        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X)

        self.status_label = tk.Label(self.root, textvariable=self.status_var, anchor="w")
        self.status_label.pack(fill=tk.X, padx=10)

        log_frame = tk.LabelFrame(self.root, text="Log")
        log_frame.pack(fill=tk.BOTH, expand=True, **padding)

        self.log_widget = ScrolledText(log_frame, height=10, state=tk.DISABLED)
        self.log_widget.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.convert_button = convert_button
        self.add_button = add_button
        self.clear_button = clear_button
        self.browse_button = browse_button

    def add_files(self) -> None:
        filetypes = [("Word documents", "*.doc;*.docx"), ("All files", "*.*")]
        new_files = filedialog.askopenfilenames(title="Select Word files", filetypes=filetypes)
        added = 0
        for path in new_files:
            normalized = os.path.normpath(path)
            if normalized not in self.selected_files:
                self.selected_files.append(normalized)
                self.file_listbox.insert(tk.END, normalized)
                added += 1

        if self.selected_files and not self.output_dir_var.get():
            self.output_dir_var.set(os.path.dirname(self.selected_files[0]))

        if added:
            self._set_status(f"Loaded {added} new file(s). Total: {len(self.selected_files)}")
        elif new_files:
            self._set_status("No new files added (duplicates skipped).")

    def clear_file_list(self) -> None:
        self.selected_files.clear()
        self.file_listbox.delete(0, tk.END)
        self.progress_var.set(0.0)
        self._set_status("Cleared file list.")

    def choose_output_directory(self) -> None:
        directory = filedialog.askdirectory(title="Select output folder")
        if directory:
            self.output_dir_var.set(os.path.normpath(directory))
            self._set_status(f"Output folder set to: {self.output_dir_var.get()}")

    def start_conversion(self) -> None:
        if self._conversion_thread and self._conversion_thread.is_alive():
            messagebox.showinfo("Conversion in progress", "Please wait for the current batch to finish.")
            return

        if not self.selected_files:
            messagebox.showwarning("No files selected", "Add at least one Word document to convert.")
            return

        output_dir = self.output_dir_var.get()
        if not output_dir:
            messagebox.showwarning("No output folder", "Choose an output folder before converting.")
            return

        if not os.path.isdir(output_dir):
            messagebox.showerror("Invalid folder", "The selected output folder does not exist.")
            return

        self._prepare_for_conversion()

        files = tuple(self.selected_files)
        self._conversion_thread = threading.Thread(target=self._run_conversion, args=(files, output_dir), daemon=True)
        self._conversion_thread.start()

    def _prepare_for_conversion(self) -> None:
        self.progress_var.set(0.0)
        self._clear_log()
        self._set_status("Starting conversion...")
        self._toggle_controls(state=tk.DISABLED)

    def _run_conversion(self, files: Iterable[str], output_dir: str) -> None:
        com_initialized = False
        word_app = None
        try:
            pythoncom.CoInitialize()
            com_initialized = True
            self._log_from_thread('Launching Microsoft Word...')
            word_app = win32com.client.DispatchEx('Word.Application')
            word_app.Visible = False

            files_list = list(files)
            total = len(files_list)
            if total == 0:
                self._log_from_thread('No files to convert.')
                self.root.after(0, self._set_status, 'No files to convert.')
                return

            self._log_from_thread(f'Converting {total} file(s)...')
            success_count = 0
            failures: list[tuple[str, str]] = []

            for index, file_path in enumerate(files_list, start=1):
                self._log_from_thread(f'Converting: {file_path}')
                try:
                    pdf_path = self._build_output_path(file_path, output_dir)
                    saved_path = self._convert_file(word_app, file_path, pdf_path)
                    success_count += 1
                    self._log_from_thread(f'[OK] Saved: {saved_path}')
                except Exception as exc:  # noqa: BLE001 - surface conversion issues
                    error_message = str(exc)
                    failures.append((file_path, error_message))
                    self._log_from_thread(f'[ERROR] Failed: {file_path} -> {error_message}')
                finally:
                    progress = (index / total) * 100.0
                    self.root.after(0, self.progress_var.set, progress)

            summary_message = f'Converted {success_count} of {total} file(s).'
            if failures:
                summary_message += ' Check log for details.'
            self._log_from_thread(summary_message)
            self.root.after(0, self._set_status, summary_message)

            if failures:
                details = '\n'.join(f'- {os.path.basename(path)}: {reason}' for path, reason in failures)
                message = 'Some files could not be converted.\n\n' + details
                self.root.after(0, messagebox.showwarning, 'Conversion completed with errors', message)
            else:
                self.root.after(0, messagebox.showinfo, 'Conversion complete', summary_message)
        except Exception as exc:  # noqa: BLE001 - unexpected failure
            self._log_thread_error('Unexpected error during conversion', exc)
        finally:
            if word_app is not None:
                word_app.Quit()
            if com_initialized:
                pythoncom.CoUninitialize()
            self.root.after(0, self._toggle_controls, tk.NORMAL)

    def _log_thread_error(self, prefix: str, exc: Exception) -> None:
        message = f"{prefix}: {exc}"
        self._log_from_thread(message)
        trace = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__)).strip()
        if trace:
            for line in trace.splitlines():
                self._log_from_thread(line)
        self.root.after(0, self._set_status, message)
        self.root.after(0, messagebox.showerror, "Conversion error", message)

    def _convert_file(self, word_app, doc_path: str, pdf_path: str) -> str:
        wd_format_pdf = 17  # Word WdExportFormat constant
        doc_path_abs = os.path.abspath(doc_path)
        pdf_path_abs = os.path.abspath(pdf_path)

        if os.path.exists(pdf_path_abs):
            os.remove(pdf_path_abs)

        doc = None
        try:
            doc = word_app.Documents.Open(doc_path_abs, ReadOnly=True)
            doc.ExportAsFixedFormat(pdf_path_abs, wd_format_pdf)
        finally:
            if doc is not None:
                doc.Close(False)

        return pdf_path_abs

    def _build_output_path(self, file_path: str, output_dir: str) -> str:
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        return os.path.join(output_dir, f"{base_name}.pdf")

    def _toggle_controls(self, state: str) -> None:
        for widget in (self.convert_button, self.add_button, self.clear_button, self.browse_button):
            widget.config(state=state)

    def _clear_log(self) -> None:
        self.log_widget.config(state=tk.NORMAL)
        self.log_widget.delete("1.0", tk.END)
        self.log_widget.config(state=tk.DISABLED)

    def _log_from_thread(self, message: str) -> None:
        self.root.after(0, self._append_log, message)

    def _append_log(self, message: str) -> None:
        self.log_widget.config(state=tk.NORMAL)
        self.log_widget.insert(tk.END, message + "\n")
        self.log_widget.see(tk.END)
        self.log_widget.config(state=tk.DISABLED)

    def _set_status(self, message: str) -> None:
        self.status_var.set(message)


def main() -> None:
    root = tk.Tk()
    app = WordToPDFConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()




