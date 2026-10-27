from tkinter import ttk

from core.config import ICON_FILE
from core.errors import logger

PALETTE = {
    "bg": "#F5F6F8",
    "surface": "#FFFFFF",
    "surface_alt": "#F8FAFC",
    "sidebar_bg": "#FFFFFF",
    "sidebar_fg": "#475569",
    "sidebar_hover": "#F1F5F9",
    "sidebar_active": "#EEF3FA",
    "sidebar_accent": "#2F5D9A",
    "sidebar_section": "#94A3B8",
    "border": "#E5E7EB",
    "border_strong": "#D1D5DB",
    "text": "#1F2937",
    "text_muted": "#6B7280",
    "primary": "#2F5D9A",
    "primary_hover": "#284F84",
    "primary_active": "#21426E",
    "primary_soft": "#EEF3FA",
    "primary_fg": "#FFFFFF",
    "disabled": "#B6BDC8",
    "success": "#3E9B6E",
    "warning": "#C98A1B",
    "danger": "#C94A4A",
    "console_bg": "#1F2933",
    "console_fg": "#E5E7EB",
}

_FUENTE = ["Helvetica"]
_FUENTE_TITULO = [None]


def fuente():
    return _FUENTE[0]


def fuente_titulo(tamano):
    return (_FUENTE_TITULO[0], tamano) if _FUENTE_TITULO[0] else (fuente(), tamano, "bold")


def aplicar_icono(ventana):
    if not ICON_FILE.exists():
        logger.info(f"Icono no encontrado en {ICON_FILE}")
        return
    try:
        ventana.iconbitmap(ICON_FILE)
    except Exception as e:
        logger.error(f"Error al establecer el icono de la aplicación: {e}")


def configurar_estilos(root):
    windows = "win32" in root.tk.call("tk", "windowingsystem")
    _FUENTE[0] = "Segoe UI" if windows else "Helvetica"
    _FUENTE_TITULO[0] = "Segoe UI Semibold" if windows else None
    f, p = fuente(), PALETTE
    style = ttk.Style(root)
    style.theme_use("clam")

    style.configure(".", font=(f, 10), background=p["bg"], foreground=p["text"])
    style.configure("TFrame", background=p["bg"])
    style.configure("Card.TFrame", background=p["surface"])
    style.configure("CardAlt.TFrame", background=p["surface_alt"])

    etiquetas = {
        "Header.TLabel": (fuente_titulo(17), p["text"], p["bg"]),
        "PageDesc.TLabel": ((f, 10), p["text_muted"], p["bg"]),
        "CardHeader.TLabel": (fuente_titulo(12), p["text"], p["surface"]),
        "Subheader.TLabel": ((f, 10), p["text_muted"], p["bg"]),
        "SubheaderCard.TLabel": ((f, 10), p["text_muted"], p["surface"]),
        "FieldLabel.TLabel": ((f, 10), p["text"], p["surface"]),
        "FieldLabelOptional.TLabel": ((f, 10), p["text_muted"], p["surface"]),
    }
    for nombre, (font, fg, bg) in etiquetas.items():
        style.configure(nombre, font=font, foreground=fg, background=bg)

    for nombre, tam, pad in (("TButton", 10, (14, 7)), ("Secondary.TButton", 10, (14, 7)), ("Toggle.TButton", 9, (10, 5))):
        style.configure(nombre, font=(f, tam), padding=pad, background=p["surface"], foreground=p["text"], relief="flat",
                        bordercolor=p["border_strong"], lightcolor=p["surface"], darkcolor=p["surface"])
        style.map(nombre, background=[("disabled", p["surface"]), ("active", p["surface_alt"])],
                  foreground=[("disabled", p["disabled"])], bordercolor=[("active", p["text_muted"])])

    campo = dict(fieldbackground=p["surface"], bordercolor=p["border_strong"], lightcolor=p["surface"],
                 darkcolor=p["surface"], foreground=p["text"])
    style.configure("TCombobox", padding=(8, 5), background=p["surface"], arrowsize=14, arrowcolor=p["text_muted"], **campo)
    style.map("TCombobox", fieldbackground=[("readonly", p["surface"])], background=[("readonly", p["surface"])],
              bordercolor=[("focus", p["primary"])], lightcolor=[("focus", p["primary"])],
              selectbackground=[("readonly", p["surface"])], selectforeground=[("readonly", p["text"])])
    style.configure("TEntry", padding=(10, 6), **campo)
    style.map("TEntry", fieldbackground=[("readonly", p["surface_alt"])],
              bordercolor=[("focus", p["primary"])], lightcolor=[("focus", p["primary"])])
    root.option_add("*TCombobox*Listbox.font", (f, 10))
    root.option_add("*TCombobox*Listbox.selectBackground", p["primary_soft"])
    root.option_add("*TCombobox*Listbox.selectForeground", p["text"])

    for nombre, bg in (("TCheckbutton", p["bg"]), ("Card.TCheckbutton", p["surface"])):
        style.configure(nombre, background=bg, foreground=p["text"], font=(f, 10), indicatorbackground=p["surface"],
                        indicatorforeground=p["primary_fg"], upperbordercolor=p["border_strong"],
                        lowerbordercolor=p["border_strong"])
        style.map(nombre, background=[("active", bg)],
                  indicatorbackground=[("selected", p["primary"]), ("pressed", p["primary_soft"])])

    style.configure("Treeview", background=p["surface"], fieldbackground=p["surface"], foreground=p["text"],
                    rowheight=28, borderwidth=0, bordercolor=p["border"], font=(f, 10))
    style.map("Treeview", background=[("selected", p["primary_soft"])], foreground=[("selected", p["text"])])
    style.configure("Treeview.Heading", font=fuente_titulo(9), background=p["surface_alt"], foreground=p["text_muted"],
                    relief="flat", bordercolor=p["border"], lightcolor=p["surface_alt"], darkcolor=p["border"],
                    padding=(8, 6))
    style.map("Treeview.Heading", background=[("active", p["border"])])

    style.configure("TScrollbar", background=p["border"], troughcolor=p["surface_alt"], bordercolor=p["surface_alt"],
                    lightcolor=p["border"], darkcolor=p["border"], arrowcolor=p["text_muted"], arrowsize=12,
                    gripcount=0, relief="flat")
    style.map("TScrollbar", background=[("active", p["border_strong"])])

    style.configure("Status.TLabel", font=(f, 9), padding=(8, 8), background=p["surface"], foreground=p["text_muted"])
    style.configure("StatusDot.TLabel", font=(f, 10), background=p["surface"], foreground=p["text_muted"])
    style.configure("TSeparator", background=p["border"])
    style.configure("Card.TLabelframe", background=p["surface"], borderwidth=1, relief="solid", bordercolor=p["border"])
    style.configure("Card.TLabelframe.Label", font=fuente_titulo(10), foreground=p["text"], background=p["surface"])
    style.configure("TProgressbar", thickness=4, background=p["primary"], troughcolor=p["border"], borderwidth=0,
                    lightcolor=p["primary"], darkcolor=p["primary"], bordercolor=p["border"])
    return style