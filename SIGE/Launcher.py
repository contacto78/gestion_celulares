import os
import shutil
import subprocess
import sys
import tkinter as tk
from pathlib import Path

APP_TITLE = "Portal de Gestión"


def base_dir() -> Path:
    """Obtiene la carpeta donde realmente está el launcher, no el directorio actual."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def find_python() -> str:
    """Localiza Python de forma independiente de la carpeta donde esté instalado."""
    if not getattr(sys, "frozen", False):
        return sys.executable

    root = base_dir()
    candidates = [
        root / "python.exe",
        root / "pythonw.exe",
        root / "runtime" / "python.exe",
        root / "runtime" / "pythonw.exe",
    ]
    for candidate in candidates:
        if candidate.exists():
            return str(candidate)

    for name in ("pythonw.exe", "python.exe", "py.exe"):
        found = shutil.which(name)
        if found:
            return found

    raise FileNotFoundError(
        "No se encontró Python para ejecutar Gestion_Celular\\app_celulares.py."
    )


def launch_cellular():
    root = base_dir()
    target = root / "Gestion_Celular" / "app_celulares.py"
    if not target.is_file():
        show_error(
            "No se encontró la aplicación Gestión Celular.\n\n"
            f"Ruta esperada:\n{target}"
        )
        return

    try:
        python = find_python()
        kwargs = {"cwd": str(target.parent)}
        if os.name == "nt":
            kwargs["creationflags"] = subprocess.CREATE_NEW_PROCESS_GROUP
        subprocess.Popen([python, str(target)], **kwargs)
    except Exception as exc:
        show_error(f"No se pudo iniciar Gestión Celular.\n\n{exc}")


def show_error(message):
    from tkinter import messagebox
    messagebox.showerror(APP_TITLE, message, parent=root_window)


def draw_phone(canvas):
    canvas.create_rectangle(38, 12, 82, 88, outline="#1f5f8b", width=4)
    canvas.create_rectangle(44, 22, 76, 74, outline="#1f5f8b", width=2)
    canvas.create_oval(56, 79, 64, 87, fill="#1f5f8b", outline="")


def draw_building(canvas):
    canvas.create_rectangle(28, 22, 92, 88, outline="#1f5f8b", width=4)
    for x in (39, 55, 71):
        for y in (34, 51, 68):
            canvas.create_rectangle(x, y, x + 8, y + 9, outline="#1f5f8b", width=2)
    canvas.create_polygon(24, 22, 60, 8, 96, 22, outline="#1f5f8b", fill="", width=4)


def draw_taxi(canvas):
    canvas.create_polygon(24, 57, 35, 36, 48, 29, 73, 29, 87, 42, 96, 57, 96, 72, 24, 72,
                          outline="#1f5f8b", fill="", width=4)
    canvas.create_rectangle(47, 34, 71, 47, outline="#1f5f8b", width=3)
    canvas.create_oval(31, 63, 45, 77, outline="#1f5f8b", width=4)
    canvas.create_oval(75, 63, 89, 77, outline="#1f5f8b", width=4)
    canvas.create_text(60, 22, text="TAXI", fill="#1f5f8b", font=("Segoe UI", 8, "bold"))


def draw_parking(canvas):
    canvas.create_rectangle(26, 10, 94, 88, outline="#1f5f8b", width=4)
    canvas.create_text(60, 50, text="P", fill="#1f5f8b", font=("Segoe UI", 46, "bold"))


ICON_DRAWERS = [draw_phone, draw_building, draw_taxi, draw_parking]


def create_card(parent, title, drawer, command):
    card = tk.Frame(parent, bg="white", bd=1, relief="solid")
    card.pack(side="left", fill="both", expand=True, padx=10, pady=10)

    icon = tk.Canvas(card, width=120, height=105, bg="white", highlightthickness=0)
    icon.pack(pady=(18, 4))
    drawer(icon)

    button = tk.Button(
        card,
        text=title,
        command=command,
        font=("Segoe UI", 12, "bold"),
        bg="#1f5f8b",
        fg="white",
        activebackground="#17496b",
        activeforeground="white",
        relief="flat",
        cursor="hand2",
        padx=12,
        pady=12,
    )
    button.pack(fill="x", padx=18, pady=(4, 18))

    icon.bind("<Button-1>", lambda _e: command())
    return card


root_window = tk.Tk()
root_window.title(APP_TITLE)
root_window.geometry("1120x500")
root_window.minsize(900, 430)
root_window.configure(bg="#f2f5f8")

header = tk.Frame(root_window, bg="#1f5f8b", height=105)
header.pack(fill="x")
header.pack_propagate(False)

tk.Label(
    header,
    text="Portal de Gestión",
    font=("Segoe UI", 24, "bold"),
    bg="#1f5f8b",
    fg="white",
).pack(pady=(20, 2))

tk.Label(
    header,
    text="Seleccione el sistema que desea utilizar",
    font=("Segoe UI", 11),
    bg="#1f5f8b",
    fg="white",
).pack()

content = tk.Frame(root_window, bg="#f2f5f8")
content.pack(fill="both", expand=True, padx=35, pady=28)

create_card(content, "Gestion Celular", draw_phone, launch_cellular)
create_card(content, "Gestión Inmobiliario", draw_building, lambda: None)
create_card(content, "Gestión de Taxi", draw_taxi, lambda: None)
create_card(content, "Gestión de Estacionamiento", draw_parking, lambda: None)

footer = tk.Label(
    root_window,
    text=f"Ubicación del portal: {base_dir()}",
    font=("Segoe UI", 8),
    bg="#f2f5f8",
    fg="#667085",
)
footer.pack(pady=(0, 10))

root_window.mainloop()
