#!/usr/bin/env python3
"""Tkinter GUI launcher for docfill.

This launcher is a thin GUI wrapper around docfill.py. It helps choose
JSON/files and run the CLI utility without typing the command manually.
"""
from __future__ import annotations

import os
import platform
import shlex
import subprocess
import sys
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

DOC_EXTENSIONS = {".doc", ".docx", ".odt"}
JSON_EXTENSION = ".json"


class DocfillLauncher(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("docfill launcher")
        self.minsize(840, 580)

        self.python_path_var = tk.StringVar(value=self._guess_python_executable())
        self.docfill_path_var = tk.StringVar(value=self._guess_docfill_script())
        self.suffix_var = tk.StringVar(value=".rendered")
        self.libre_office_exec_var = tk.StringVar()
        self.extra_args_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Готово.")
        self.verbose_var = tk.BooleanVar(value=False)
        self.check_var = tk.BooleanVar(value=False)
        self.ignore_case_var = tk.BooleanVar(value=False)
        self.in_place_var = tk.BooleanVar(value=False)
        self.running = False
        self.process: subprocess.Popen[str] | None = None

        self._build_ui()
        self._configure_window_geometry()
        self._update_command_preview()


    def _configure_window_geometry(self) -> None:
        self.update_idletasks()

        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()

        target_w = min(max(self.winfo_reqwidth() + 20, 900), max(640, int(screen_w * 0.92)))
        target_h = min(max(self.winfo_reqheight() + 20, 620), max(480, int(screen_h * 0.85)))

        x = max((screen_w - target_w) // 2, 0)
        y = max((screen_h - target_h) // 3, 0)

        self.maxsize(screen_w, screen_h)
        self.geometry(f"{target_w}x{target_h}+{x}+{y}")

    def _guess_python_executable(self) -> str:
        return sys.executable or "python"

    def _guess_docfill_script(self) -> str:
        candidates = []

        try:
            here = Path(__file__).resolve().parent
            candidates.append(here / "docfill.py")
        except Exception:
            pass

        cwd = Path.cwd()
        candidates.append(cwd / "docfill.py")

        for candidate in candidates:
            if candidate.is_file():
                return str(candidate)

        return "docfill.py"

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(4, weight=1)
        self.rowconfigure(6, weight=0)

        top = ttk.Frame(self, padding=10)
        top.grid(row=0, column=0, sticky="nsew")
        top.columnconfigure(1, weight=1)

        ttk.Label(top, text="Python:").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        python_entry = ttk.Entry(top, textvariable=self.python_path_var)
        python_entry.grid(row=0, column=1, sticky="ew", pady=4)
        ttk.Button(top, text="Выбрать...", command=self._choose_python).grid(row=0, column=2, sticky="ew", padx=(8, 0), pady=4)

        ttk.Label(top, text="docfill.py:").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        docfill_entry = ttk.Entry(top, textvariable=self.docfill_path_var)
        docfill_entry.grid(row=1, column=1, sticky="ew", pady=4)
        ttk.Button(top, text="Выбрать...", command=self._choose_docfill).grid(row=1, column=2, sticky="ew", padx=(8, 0), pady=4)

        self.python_path_var.trace_add("write", lambda *_: self._update_command_preview())
        self.docfill_path_var.trace_add("write", lambda *_: self._update_command_preview())
        self.suffix_var.trace_add("write", lambda *_: self._update_command_preview())
        self.libre_office_exec_var.trace_add("write", lambda *_: self._update_command_preview())
        self.extra_args_var.trace_add("write", lambda *_: self._update_command_preview())
        self.verbose_var.trace_add("write", lambda *_: self._update_command_preview())
        self.check_var.trace_add("write", lambda *_: self._update_command_preview())
        self.ignore_case_var.trace_add("write", lambda *_: self._update_command_preview())
        self.in_place_var.trace_add("write", lambda *_: self._update_command_preview())

        files_frame = ttk.Frame(self, padding=(10, 0, 10, 0))
        files_frame.grid(row=1, column=0, sticky="nsew")
        files_frame.columnconfigure(0, weight=1)
        files_frame.columnconfigure(1, weight=1)
        files_frame.rowconfigure(0, weight=1)

        self._build_file_list(
            parent=files_frame,
            column=0,
            title="JSON-файлы с плейсхолдерами",
            add_command=self._add_json_files,
            remove_command=lambda: self._remove_selected(self.json_listbox),
            clear_command=lambda: self._clear_list(self.json_listbox),
            attr_name="json_listbox",
        )
        self._build_file_list(
            parent=files_frame,
            column=1,
            title="Документы",
            add_command=self._add_document_files,
            remove_command=lambda: self._remove_selected(self.doc_listbox),
            clear_command=lambda: self._clear_list(self.doc_listbox),
            attr_name="doc_listbox",
        )

        options = ttk.LabelFrame(self, text="Параметры запуска", padding=10)
        options.grid(row=2, column=0, sticky="ew", padx=10, pady=(10, 0))
        for idx in range(4):
            options.columnconfigure(idx, weight=1)

        ttk.Checkbutton(options, text="--ignore-case", variable=self.ignore_case_var).grid(row=0, column=0, sticky="w", pady=2)
        ttk.Checkbutton(options, text="--check", variable=self.check_var).grid(row=0, column=1, sticky="w", pady=2)
        ttk.Checkbutton(options, text="-v / --verbose", variable=self.verbose_var).grid(row=0, column=2, sticky="w", pady=2)
        ttk.Checkbutton(options, text="--in-place", variable=self.in_place_var).grid(row=0, column=3, sticky="w", pady=2)

        ttk.Label(options, text="--suffix:").grid(row=1, column=0, sticky="w", pady=(8, 2))
        ttk.Entry(options, textvariable=self.suffix_var).grid(row=1, column=1, sticky="ew", pady=(8, 2), padx=(0, 8))

        ttk.Label(options, text="--libreoffice-exec:").grid(row=1, column=2, sticky="w", pady=(8, 2))
        ttk.Entry(options, textvariable=self.libre_office_exec_var).grid(row=1, column=3, sticky="ew", pady=(8, 2))

        ttk.Label(options, text="Дополнительные аргументы:").grid(row=2, column=0, sticky="w", pady=(8, 2))
        ttk.Entry(options, textvariable=self.extra_args_var).grid(row=2, column=1, columnspan=3, sticky="ew", pady=(8, 2))

        preview_frame = ttk.LabelFrame(self, text="Команда", padding=10)
        preview_frame.grid(row=3, column=0, sticky="ew", padx=10, pady=(10, 0))
        preview_frame.columnconfigure(0, weight=1)
        self.command_preview = tk.Text(preview_frame, height=3, wrap="word")
        self.command_preview.grid(row=0, column=0, sticky="ew")
        self.command_preview.configure(state="disabled")

        output_frame = ttk.LabelFrame(self, text="Вывод", padding=10)
        output_frame.grid(row=4, column=0, sticky="nsew", padx=10, pady=(10, 0))
        output_frame.columnconfigure(0, weight=1)
        output_frame.rowconfigure(0, weight=1)
        self.output_text = tk.Text(output_frame, wrap="word", height=12)
        self.output_text.grid(row=0, column=0, sticky="nsew")
        output_scroll = ttk.Scrollbar(output_frame, orient="vertical", command=self.output_text.yview)
        output_scroll.grid(row=0, column=1, sticky="ns")
        self.output_text.configure(yscrollcommand=output_scroll.set)

        buttons = ttk.Frame(self, padding=10)
        buttons.grid(row=5, column=0, sticky="ew")
        buttons.columnconfigure(0, weight=1)

        left_buttons = ttk.Frame(buttons)
        left_buttons.grid(row=0, column=0, sticky="w")
        ttk.Button(left_buttons, text="Запустить", command=self._start_run).pack(side="left")
        ttk.Button(left_buttons, text="Остановить", command=self._stop_run).pack(side="left", padx=(8, 0))
        ttk.Button(left_buttons, text="Очистить вывод", command=self._clear_output).pack(side="left", padx=(8, 0))

        ttk.Label(buttons, textvariable=self.status_var).grid(row=0, column=1, sticky="e")

        hints = ttk.LabelFrame(self, text="Подсказки", padding=10)
        hints.grid(row=6, column=0, sticky="nsew", padx=10, pady=(0, 10))
        hints.columnconfigure(0, weight=1)
        hints.rowconfigure(0, weight=1)
        hint_text = (
            "• JSON-файлы должны иметь расширение .json.\n"
            "• Документы распознаются по расширениям .doc, .docx и .odt.\n"
            "• Для .doc может понадобиться LibreOffice.\n"
            "• Этот launcher не меняет docfill.py, а просто запускает его с выбранными аргументами."
        )
        ttk.Label(hints, text=hint_text, justify="left").grid(row=0, column=0, sticky="nw")

    def _build_file_list(self, parent, column, title, add_command, remove_command, clear_command, attr_name):
        frame = ttk.LabelFrame(parent, text=title, padding=10)
        frame.grid(row=0, column=column, sticky="nsew", padx=(0 if column == 0 else 6, 6 if column == 0 else 0))
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        listbox = tk.Listbox(frame, selectmode=tk.EXTENDED)
        listbox.grid(row=0, column=0, sticky="nsew")
        setattr(self, attr_name, listbox)
        listbox.bind("<<ListboxSelect>>", lambda _e: self._update_command_preview())

        scroll = ttk.Scrollbar(frame, orient="vertical", command=listbox.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        listbox.configure(yscrollcommand=scroll.set)

        buttons = ttk.Frame(frame)
        buttons.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        ttk.Button(buttons, text="Добавить...", command=add_command).pack(side="left")
        ttk.Button(buttons, text="Удалить выбранное", command=remove_command).pack(side="left", padx=(8, 0))
        ttk.Button(buttons, text="Очистить", command=clear_command).pack(side="left", padx=(8, 0))

    def _choose_python(self) -> None:
        path = filedialog.askopenfilename(
            title="Выберите Python",
            filetypes=[("Python executable", "python.exe python3.exe py.exe *.exe"), ("All files", "*.*")],
        )
        if path:
            self.python_path_var.set(path)

    def _choose_docfill(self) -> None:
        path = filedialog.askopenfilename(
            title="Выберите docfill.py",
            filetypes=[("Python files", "*.py"), ("All files", "*.*")],
        )
        if path:
            self.docfill_path_var.set(path)

    def _add_json_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Выберите JSON-файлы",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        )
        self._add_paths(self.json_listbox, paths, required_suffix=JSON_EXTENSION)

    def _add_document_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Выберите документы",
            filetypes=[("Supported documents", "*.doc *.docx *.odt"), ("All files", "*.*")],
        )
        self._add_paths(self.doc_listbox, paths, required_suffixes=DOC_EXTENSIONS)

    def _add_paths(self, listbox: tk.Listbox, paths, required_suffix: str | None = None, required_suffixes: set[str] | None = None) -> None:
        existing = set(listbox.get(0, tk.END))
        added_any = False
        rejected: list[str] = []
        for path in paths:
            suffix = Path(path).suffix.lower()
            ok = True
            if required_suffix is not None:
                ok = suffix == required_suffix
            if required_suffixes is not None:
                ok = suffix in required_suffixes
            if not ok:
                rejected.append(path)
                continue
            if path not in existing:
                listbox.insert(tk.END, path)
                existing.add(path)
                added_any = True
        if rejected:
            messagebox.showwarning(
                "Неподдерживаемые файлы",
                "Некоторые файлы были пропущены из-за неподдерживаемого расширения:\n\n" + "\n".join(rejected),
            )
        if added_any:
            self._update_command_preview()

    def _remove_selected(self, listbox: tk.Listbox) -> None:
        indices = list(listbox.curselection())
        if not indices:
            return
        for idx in reversed(indices):
            listbox.delete(idx)
        self._update_command_preview()

    def _clear_list(self, listbox: tk.Listbox) -> None:
        listbox.delete(0, tk.END)
        self._update_command_preview()

    def _quote_for_preview(self, value: str) -> str:
        if platform.system().lower().startswith("win"):
            if not value or any(ch in value for ch in ' \t"'):
                escaped = value.replace('"', '\\"')
                return f'"{escaped}"'
            return value
        return shlex.quote(value)

    def _parse_extra_args(self) -> list[str]:
        text = self.extra_args_var.get().strip()
        if not text:
            return []
        try:
            return shlex.split(text, posix=(os.name != "nt"))
        except ValueError as exc:
            raise ValueError(f"Не удалось разобрать дополнительные аргументы: {exc}") from exc

    def _build_command(self) -> list[str]:
        python_exec = self.python_path_var.get().strip() or self._guess_python_executable()
        docfill_script = self.docfill_path_var.get().strip() or self._guess_docfill_script()
        json_files = list(self.json_listbox.get(0, tk.END))
        doc_files = list(self.doc_listbox.get(0, tk.END))

        cmd = [python_exec, docfill_script]
        if self.ignore_case_var.get():
            cmd.append("--ignore-case")
        if self.check_var.get():
            cmd.append("--check")
        if self.verbose_var.get():
            cmd.append("--verbose")
        if self.in_place_var.get():
            cmd.append("--in-place")

        suffix = self.suffix_var.get().strip()
        if suffix and suffix != ".rendered":
            cmd.extend(["--suffix", suffix])

        libre_exec = self.libre_office_exec_var.get().strip()
        if libre_exec:
            cmd.extend(["--libreoffice-exec", libre_exec])

        cmd.extend(self._parse_extra_args())
        cmd.extend(json_files)
        cmd.extend(doc_files)
        return cmd

    def _validate_before_run(self) -> bool:
        python_exec = self.python_path_var.get().strip()
        docfill_script = self.docfill_path_var.get().strip()
        json_files = list(self.json_listbox.get(0, tk.END))
        doc_files = list(self.doc_listbox.get(0, tk.END))

        if not python_exec:
            messagebox.showerror("Ошибка", "Не указан путь к Python.")
            return False
        if not docfill_script:
            messagebox.showerror("Ошибка", "Не указан путь к docfill.py.")
            return False
        if not Path(docfill_script).is_file():
            messagebox.showerror("Ошибка", f"Файл docfill.py не найден:\n{docfill_script}")
            return False
        if not json_files:
            messagebox.showerror("Ошибка", "Добавьте хотя бы один JSON-файл.")
            return False
        if not doc_files:
            messagebox.showerror("Ошибка", "Добавьте хотя бы один документ.")
            return False
        for path in json_files:
            if Path(path).suffix.lower() != JSON_EXTENSION:
                messagebox.showerror("Ошибка", f"JSON-файл должен оканчиваться на .json:\n{path}")
                return False
        for path in doc_files:
            if Path(path).suffix.lower() not in DOC_EXTENSIONS:
                messagebox.showerror("Ошибка", f"Неподдерживаемый формат документа:\n{path}")
                return False
        try:
            self._parse_extra_args()
        except ValueError as exc:
            messagebox.showerror("Ошибка", str(exc))
            return False
        return True

    def _update_command_preview(self) -> None:
        try:
            cmd = self._build_command()
            preview = " ".join(self._quote_for_preview(part) for part in cmd)
        except ValueError as exc:
            preview = f"Ошибка в дополнительных аргументах: {exc}"
        self.command_preview.configure(state="normal")
        self.command_preview.delete("1.0", tk.END)
        self.command_preview.insert("1.0", preview)
        self.command_preview.configure(state="disabled")

    def _append_output(self, text: str) -> None:
        self.output_text.insert(tk.END, text)
        self.output_text.see(tk.END)

    def _clear_output(self) -> None:
        self.output_text.delete("1.0", tk.END)
        self.status_var.set("Вывод очищен.")

    def _set_running_state(self, running: bool) -> None:
        self.running = running
        self.status_var.set("Выполняется..." if running else self.status_var.get())

    def _start_run(self) -> None:
        if self.running:
            messagebox.showinfo("docfill launcher", "Команда уже выполняется.")
            return
        if not self._validate_before_run():
            return

        cmd = self._build_command()
        self._append_output("\n> " + " ".join(self._quote_for_preview(p) for p in cmd) + "\n\n")
        self.status_var.set("Выполняется...")
        self._set_running_state(True)

        thread = threading.Thread(target=self._run_subprocess, args=(cmd,), daemon=True)
        thread.start()

    def _run_subprocess(self, cmd: list[str]) -> None:
        try:
            self.process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
            )
            assert self.process.stdout is not None
            for line in self.process.stdout:
                self.after(0, self._append_output, line)
            returncode = self.process.wait()
            self.after(0, self._finish_run, returncode)
        except FileNotFoundError as exc:
            self.after(0, self._run_failed, f"Не удалось запустить процесс: {exc}")
        except Exception as exc:
            self.after(0, self._run_failed, f"Ошибка запуска: {exc}")
        finally:
            self.process = None

    def _finish_run(self, returncode: int) -> None:
        self._set_running_state(False)
        if returncode == 0:
            self.status_var.set("Готово. Команда завершилась успешно.")
            self._append_output("\n[launcher] Команда завершилась успешно.\n")
        else:
            self.status_var.set(f"Готово. Код возврата: {returncode}.")
            self._append_output(f"\n[launcher] Команда завершилась с кодом {returncode}.\n")

    def _run_failed(self, message: str) -> None:
        self._set_running_state(False)
        self.status_var.set("Ошибка запуска.")
        self._append_output(f"\n[launcher] {message}\n")
        messagebox.showerror("Ошибка", message)

    def _stop_run(self) -> None:
        if not self.running or self.process is None:
            self.status_var.set("Нет активного процесса.")
            return
        try:
            self.process.terminate()
            self.status_var.set("Отправлен сигнал остановки...")
            self._append_output("\n[launcher] Отправлен сигнал остановки.\n")
        except Exception as exc:
            messagebox.showerror("Ошибка", f"Не удалось остановить процесс: {exc}")


def main() -> int:
    app = DocfillLauncher()
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
