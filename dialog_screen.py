import tkinter as tk
from tkinter import messagebox, ttk


def seleccionar_tipos_documento(parent, tipos, preseleccionados=None):
    preseleccionados = set(preseleccionados or []) & set(tipos) or set(tipos)
    resultado = {"valor": None}

    ventana = tk.Toplevel(parent)
    ventana.title("Tipos de documento")
    ventana.resizable(False, False)
    ventana.transient(parent)

    marco = ttk.Frame(ventana, padding=14)
    marco.pack(fill="both", expand=True)

    ttk.Label(
        marco, text="Seleccione los tipos de documento a incluir en el informe de IVA:"
    ).pack(anchor="w", pady=(0, 8))

    variables = {t: tk.BooleanVar(value=t in preseleccionados) for t in tipos}
    var_todos = tk.BooleanVar(value=all(v.get() for v in variables.values()))

    def alternar_todos():
        for v in variables.values():
            v.set(var_todos.get())

    def actualizar_todos():
        var_todos.set(all(v.get() for v in variables.values()))

    ttk.Checkbutton(
        marco, text="Seleccionar todos", variable=var_todos, command=alternar_todos
    ).pack(anchor="w")
    ttk.Separator(marco).pack(fill="x", pady=6)

    for tipo, var in variables.items():
        ttk.Checkbutton(
            marco, text=tipo, variable=var, command=actualizar_todos
        ).pack(anchor="w", padx=(12, 0))

    def aceptar():
        seleccion = [t for t, v in variables.items() if v.get()]
        if not seleccion:
            messagebox.showwarning(
                "Tipos de documento",
                "Debe seleccionar al menos un tipo de documento.",
                parent=ventana,
            )
            return
        resultado["valor"] = seleccion
        ventana.destroy()

    barra = ttk.Frame(marco)
    barra.pack(fill="x", pady=(12, 0))
    ttk.Button(barra, text="Cancelar", command=ventana.destroy).pack(side="right")
    ttk.Button(barra, text="Aceptar", command=aceptar).pack(side="right", padx=(0, 6))

    ventana.protocol("WM_DELETE_WINDOW", ventana.destroy)
    ventana.bind("<Return>", lambda _e: aceptar())
    ventana.bind("<Escape>", lambda _e: ventana.destroy())

    ventana.update_idletasks()
    x = parent.winfo_rootx() + (parent.winfo_width() - ventana.winfo_width()) // 2
    y = parent.winfo_rooty() + (parent.winfo_height() - ventana.winfo_height()) // 2
    ventana.geometry(f"+{max(x, 0)}+{max(y, 0)}")
    ventana.grab_set()
    ventana.focus_set()
    parent.wait_window(ventana)

    return resultado["valor"]