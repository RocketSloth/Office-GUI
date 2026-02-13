import json
import math
import queue
import re
import subprocess
import sys
import threading
from dataclasses import dataclass
from pathlib import Path
from tkinter import PhotoImage, Tk, messagebox
from tkinter import scrolledtext
from tkinter import ttk


APP_TITLE = "TJ Task Center"
ROOT_DIR = Path(__file__).resolve().parent
CONFIG_PATH = ROOT_DIR / "tasks.json"
ASSETS_DIR = ROOT_DIR / "assets"
LOGO_CANDIDATES = ("tj_logo.png", "tj_logo.gif", "tj_logo.ppm", "tj_logo.pgm")
ICON_CANDIDATES = ("tj_icon.ico",)


BG_APP = "#0B1017"
BG_PANEL = "#131C29"
BG_CARD = "#1A2535"
FG_PRIMARY = "#F5F8FF"
FG_MUTED = "#93A4BC"
ACCENT = "#13C6D7"
ACCENT_HOVER = "#1DB5FF"
WARN = "#F59E0B"
DANGER = "#E05454"


@dataclass
class Task:
    name: str
    description: str
    command: list[str] | str
    cwd: Path


def pretty_name(stem: str) -> str:
    with_spaces = re.sub(r"(?<!^)(?=[A-Z])", " ", stem.replace("_", " ").strip())
    return " ".join(word.capitalize() for word in with_spaces.split())


def expand_placeholders(value: str) -> str:
    return (
        value.replace("{python}", sys.executable)
        .replace("{project_dir}", str(ROOT_DIR))
        .replace("{assets_dir}", str(ASSETS_DIR))
    )


def parse_task(raw: dict) -> Task | None:
    name = str(raw.get("name", "")).strip()
    if not name:
        return None

    description = str(raw.get("description", "")).strip()
    command = raw.get("command")
    if not command:
        return None

    if isinstance(command, str):
        parsed_command: list[str] | str = expand_placeholders(command)
    elif isinstance(command, list):
        parsed_command = [expand_placeholders(str(part)) for part in command if str(part).strip()]
        if not parsed_command:
            return None
    else:
        return None

    raw_cwd = str(raw.get("cwd", "{project_dir}"))
    cwd = Path(expand_placeholders(raw_cwd))
    if not cwd.is_absolute():
        cwd = (ROOT_DIR / cwd).resolve()

    if not cwd.exists():
        cwd = ROOT_DIR

    return Task(name=name, description=description, command=parsed_command, cwd=cwd)


def discover_tasks() -> list[Task]:
    tasks: list[Task] = []

    if CONFIG_PATH.exists():
        try:
            data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            raw_tasks = data.get("tasks", [])
            if isinstance(raw_tasks, list):
                for raw_task in raw_tasks:
                    if isinstance(raw_task, dict):
                        task = parse_task(raw_task)
                        if task:
                            tasks.append(task)
        except Exception:
            tasks = []

    configured_scripts: set[str] = set()
    for task in tasks:
        script_name = find_script_name(task.command)
        if script_name:
            configured_scripts.add(script_name.lower())

    for script in sorted(ROOT_DIR.glob("*.py")):
        if script.name == Path(__file__).name:
            continue
        if script.name.lower() in configured_scripts:
            continue
        tasks.append(
            Task(
                name=pretty_name(script.stem),
                description=f"Run {script.name}",
                command=[sys.executable, script.name],
                cwd=ROOT_DIR,
            )
        )
    return tasks


def find_script_name(command: list[str] | str) -> str | None:
    if isinstance(command, list):
        for part in reversed(command):
            part_clean = part.strip().strip("\"'")
            if part_clean.lower().endswith(".py"):
                return Path(part_clean).name
        return None

    matches = re.findall(r"([A-Za-z0-9_./\\ -]+\.py)", command)
    if matches:
        return Path(matches[-1].strip().strip("\"'")).name
    return None


class TaskCenter(Tk):
    def __init__(self, tasks: list[Task]) -> None:
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("980x620")
        self.minsize(860, 520)
        self.configure(bg=BG_APP)

        self.tasks = tasks
        self.process: subprocess.Popen | None = None
        self.output_queue: queue.Queue[str | tuple[str, int, str]] = queue.Queue()
        self.logo_image: PhotoImage | None = None
        self.task_buttons: list[ttk.Button] = []
        self.style = ttk.Style(self)

        self._configure_theme()
        self._set_icon()
        self._build_ui()
        self.after(120, self._drain_output_queue)

    def _configure_theme(self) -> None:
        try:
            self.style.theme_use("clam")
        except Exception:
            pass

        self.style.configure(".", background=BG_APP, foreground=FG_PRIMARY)
        self.style.configure("Root.TFrame", background=BG_APP)
        self.style.configure("Panel.TFrame", background=BG_PANEL)
        self.style.configure("Card.TFrame", background=BG_CARD)
        self.style.configure("Accent.TFrame", background=ACCENT)

        self.style.configure(
            "Title.TLabel",
            background=BG_PANEL,
            foreground=FG_PRIMARY,
            font=("Segoe UI Semibold", 20),
        )
        self.style.configure(
            "SubTitle.TLabel",
            background=BG_PANEL,
            foreground=FG_MUTED,
            font=("Segoe UI", 10),
        )
        self.style.configure(
            "PanelHeading.TLabel",
            background=BG_PANEL,
            foreground=FG_PRIMARY,
            font=("Segoe UI Semibold", 11),
        )
        self.style.configure(
            "Panel.TLabel",
            background=BG_PANEL,
            foreground=FG_PRIMARY,
        )
        self.style.configure(
            "SectionHint.TLabel",
            background=BG_PANEL,
            foreground=FG_MUTED,
            font=("Segoe UI", 9),
        )
        self.style.configure(
            "TaskName.TLabel",
            background=BG_CARD,
            foreground=FG_PRIMARY,
            font=("Segoe UI Semibold", 10),
        )
        self.style.configure(
            "TaskHint.TLabel",
            background=BG_CARD,
            foreground=FG_MUTED,
            font=("Segoe UI", 9),
        )
        self.style.configure(
            "StatusLabel.TLabel",
            background=BG_PANEL,
            foreground=FG_MUTED,
            font=("Segoe UI", 9),
        )
        self.style.configure(
            "StatusReady.TLabel",
            background=BG_PANEL,
            foreground=ACCENT,
            font=("Segoe UI Semibold", 9),
        )
        self.style.configure(
            "StatusRunning.TLabel",
            background=BG_PANEL,
            foreground=WARN,
            font=("Segoe UI Semibold", 9),
        )
        self.style.configure(
            "FallbackLogo.TLabel",
            background=BG_PANEL,
            foreground=ACCENT,
            font=("Segoe UI Black", 30),
        )

        self.style.configure(
            "Task.TButton",
            background=BG_CARD,
            foreground=FG_PRIMARY,
            borderwidth=0,
            padding=(10, 8),
            font=("Segoe UI Semibold", 10),
        )
        self.style.map(
            "Task.TButton",
            background=[
                ("disabled", "#2A3343"),
                ("pressed", ACCENT),
                ("active", ACCENT_HOVER),
            ],
            foreground=[("disabled", "#5A6880"), ("!disabled", FG_PRIMARY)],
        )

        self.style.configure(
            "Danger.TButton",
            background="#3B2326",
            foreground="#FFD9D9",
            borderwidth=0,
            padding=(10, 7),
            font=("Segoe UI Semibold", 9),
        )
        self.style.map(
            "Danger.TButton",
            background=[("disabled", "#2B2E35"), ("pressed", DANGER), ("active", "#C64040")],
            foreground=[("disabled", "#626B78"), ("!disabled", "#FFE6E6")],
        )

    def _set_icon(self) -> None:
        for icon_name in ICON_CANDIDATES:
            icon_path = ASSETS_DIR / icon_name
            if icon_path.exists():
                try:
                    self.iconbitmap(default=str(icon_path))
                    return
                except Exception:
                    return

    def _load_logo(self) -> PhotoImage | None:
        for logo_name in LOGO_CANDIDATES:
            logo_path = ASSETS_DIR / logo_name
            if logo_path.exists():
                try:
                    image = PhotoImage(file=str(logo_path))
                    if image.width() > 360:
                        scale = max(1, math.ceil(image.width() / 360))
                        image = image.subsample(scale, scale)
                    return image
                except Exception:
                    continue
        return None

    def _build_ui(self) -> None:
        main = ttk.Frame(self, padding=14, style="Root.TFrame")
        main.pack(fill="both", expand=True)

        header_shell = ttk.Frame(main, style="Panel.TFrame")
        header_shell.pack(fill="x")

        accent = ttk.Frame(header_shell, style="Accent.TFrame", height=4)
        accent.pack(fill="x")

        header = ttk.Frame(header_shell, style="Panel.TFrame", padding=(14, 12))
        header.pack(fill="x")

        self.logo_image = self._load_logo()
        if self.logo_image:
            logo_label = ttk.Label(header, image=self.logo_image, style="Panel.TLabel")
            logo_label.pack(side="left")
        else:
            logo_label = ttk.Label(header, text="TJ", style="FallbackLogo.TLabel")
            logo_label.pack(side="left", padx=(0, 10))

        title_frame = ttk.Frame(header, style="Panel.TFrame")
        title_frame.pack(side="left", fill="x", expand=True, padx=(12, 0))
        ttk.Label(title_frame, text=APP_TITLE, style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            title_frame,
            text="Click a button to run your daily work scripts.",
            style="SubTitle.TLabel",
        ).pack(anchor="w")

        tasks_count = len(self.tasks)
        tasks_label = f"{tasks_count} task{'s' if tasks_count != 1 else ''} loaded"
        ttk.Label(header, text=tasks_label, style="SubTitle.TLabel").pack(side="right", padx=(10, 0))

        body = ttk.Frame(main, style="Root.TFrame")
        body.pack(fill="both", expand=True, pady=(12, 0))

        left = ttk.Frame(body, style="Panel.TFrame", padding=12, width=320)
        left.pack(side="left", fill="y")
        left.pack_propagate(False)

        right = ttk.Frame(body, style="Panel.TFrame", padding=12)
        right.pack(side="left", fill="both", expand=True, padx=(12, 0))

        ttk.Label(left, text="Tasks", style="PanelHeading.TLabel").pack(anchor="w")
        ttk.Label(left, text="Run any script with one click.", style="SectionHint.TLabel").pack(anchor="w", pady=(0, 10))

        task_holder = ttk.Frame(left, style="Panel.TFrame")
        task_holder.pack(fill="both", expand=True)

        for task in self.tasks:
            card = ttk.Frame(task_holder, style="Card.TFrame", padding=10)
            card.pack(fill="x", pady=4)
            ttk.Label(card, text=task.name, style="TaskName.TLabel").pack(anchor="w")
            button = ttk.Button(card, text="Run Task", style="Task.TButton", command=lambda t=task: self._run_task(t))
            button.pack(fill="x")
            self.task_buttons.append(button)
            if task.description:
                ttk.Label(
                    card,
                    text=task.description,
                    style="TaskHint.TLabel",
                    wraplength=280,
                    justify="left",
                ).pack(anchor="w", pady=(6, 0))

        if not self.tasks:
            ttk.Label(task_holder, text="No tasks found.", style="SectionHint.TLabel").pack(anchor="w")

        controls = ttk.Frame(right, style="Panel.TFrame")
        controls.pack(fill="x", pady=(0, 8))

        ttk.Label(controls, text="Status:", style="StatusLabel.TLabel").pack(side="left")
        self.status_var = ttk.Label(controls, text="Ready", style="StatusReady.TLabel")
        self.status_var.pack(side="left")

        self.stop_button = ttk.Button(
            controls,
            text="Stop Current Task",
            command=self._stop_task,
            state="disabled",
            style="Danger.TButton",
        )
        self.stop_button.pack(side="right")

        self.output = scrolledtext.ScrolledText(
            right,
            wrap="word",
            height=24,
            font=("Consolas", 10),
            state="disabled",
        )
        self.output.pack(fill="both", expand=True)
        self.output.configure(
            background="#0F1724",
            foreground="#E3ECFF",
            insertbackground=ACCENT,
            selectbackground="#2A4E74",
            highlightthickness=1,
            highlightbackground="#2B3B52",
            highlightcolor=ACCENT,
            borderwidth=0,
            padx=10,
            pady=10,
        )
        self.output.tag_configure("error", foreground="#FF8F8F")
        self.output.tag_configure("success", foreground="#66E3A6")
        self.output.tag_configure("info", foreground="#8EC9FF")
        self.output.tag_configure("progress", foreground="#F2CC71")

        self._append_output("Task center started.\n")
        if not self.tasks:
            self._append_output("No tasks were found. Add scripts or update tasks.json.\n")

    def _append_output(self, text: str) -> None:
        tag = None
        stripped = text.strip()
        if "[ERROR]" in stripped:
            tag = "error"
        elif "[DONE]" in stripped:
            tag = "success"
        elif "[PROGRESS]" in stripped:
            tag = "progress"
        elif "[INFO]" in stripped:
            tag = "info"

        self.output.configure(state="normal")
        if tag:
            self.output.insert("end", text, tag)
        else:
            self.output.insert("end", text)
        self.output.see("end")
        self.output.configure(state="disabled")

    def _set_running_state(self, running: bool) -> None:
        for button in self.task_buttons:
            button.configure(state="disabled" if running else "normal")
        self.stop_button.configure(state="normal" if running else "disabled")
        if running:
            self.status_var.configure(text="Running", style="StatusRunning.TLabel")
        else:
            self.status_var.configure(text="Ready", style="StatusReady.TLabel")

    def _run_task(self, task: Task) -> None:
        if self.process and self.process.poll() is None:
            messagebox.showwarning(APP_TITLE, "A task is already running. Stop it first.")
            return

        self._append_output(f"\n=== {task.name} ===\n")
        self._append_output(f"Working directory: {task.cwd}\n")
        self._append_output(f"Command: {task.command}\n\n")
        self._set_running_state(True)

        worker = threading.Thread(target=self._execute_task, args=(task,), daemon=True)
        worker.start()

    def _execute_task(self, task: Task) -> None:
        command = task.command
        shell_mode = isinstance(command, str)

        try:
            self.process = subprocess.Popen(
                command,
                cwd=str(task.cwd),
                shell=shell_mode,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
            )
        except Exception as exc:
            self.output_queue.put(f"[ERROR] Failed to start task: {exc}\n")
            self.output_queue.put(("__COMPLETE__", 1, task.name))
            return

        assert self.process.stdout is not None
        for line in self.process.stdout:
            self.output_queue.put(line)

        return_code = self.process.wait()
        self.output_queue.put(("__COMPLETE__", return_code, task.name))

    def _stop_task(self) -> None:
        if not self.process or self.process.poll() is not None:
            return
        self._append_output("\n[INFO] Stop requested. Terminating task...\n")
        try:
            self.process.terminate()
        except Exception as exc:
            self._append_output(f"[ERROR] Could not stop task: {exc}\n")

    def _drain_output_queue(self) -> None:
        while not self.output_queue.empty():
            item = self.output_queue.get()
            if isinstance(item, tuple) and item and item[0] == "__COMPLETE__":
                _, return_code, task_name = item
                if return_code == 0:
                    self._append_output(f"\n[DONE] {task_name} completed successfully.\n")
                else:
                    self._append_output(f"\n[DONE] {task_name} exited with code {return_code}.\n")
                self.process = None
                self._set_running_state(False)
            else:
                self._append_output(str(item))

        self.after(120, self._drain_output_queue)


def main() -> int:
    tasks = discover_tasks()
    app = TaskCenter(tasks)
    app.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
