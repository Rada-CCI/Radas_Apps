import re
import subprocess
import threading
import ctypes
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import winreg
import os
import shutil
from pathlib import Path
import app_version as _av
import sys

# Optional Pillow for better image resizing; fallback to Tk PhotoImage
try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except Exception:
    PIL_AVAILABLE = False


TYPE_MAP = {
    "Console": "C",
    "Laptop": "L",
    "Rack PC": "R",
    "Desktop": "D",
}


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin() != 0
    except Exception:
        return False


def _resource_path(rel_path: str) -> str:
    """Return an absolute path to a resource, working when running in a PyInstaller one-file bundle.

    If running frozen, PyInstaller extracts files to sys._MEIPASS; otherwise use the script directory.
    """
    base = getattr(sys, '_MEIPASS', None)
    if base:
        return os.path.join(base, rel_path)
    return os.path.join(os.path.dirname(__file__), rel_path)


def sanitize_custom(name: str) -> str:
    # Convert spaces to -, allow only letters, numbers, - and _
    name = name.replace(" ", "-")
    # Remove disallowed characters
    name = re.sub(r"[^A-Za-z0-9-_]", "", name)
    return name[:15]


class ScrollableFrame(ttk.Frame):
    """A simple vertically-scrollable frame for ttk widgets."""
    def __init__(self, container, *args, **kwargs):
        super().__init__(container, *args, **kwargs)
        canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0)
        vsb = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        self.inner = ttk.Frame(canvas)

        self.inner_id = canvas.create_window((0, 0), window=self.inner, anchor="nw")

        canvas.configure(yscrollcommand=vsb.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        def _on_frame_config(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _on_canvas_config(event):
            # keep inner width in sync with canvas width
            canvas.itemconfig(self.inner_id, width=event.width)

        self.inner.bind("<Configure>", _on_frame_config)
        canvas.bind('<Configure>', _on_canvas_config)

        # Windows mousewheel
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all('<MouseWheel>', _on_mousewheel)


class ToolTip:
    """Very small tooltip helper for tkinter widgets."""
    def __init__(self, widget, text, delay=500):
        self.widget = widget
        self.text = text
        self.delay = delay
        self.tipwindow = None
        self.id = None
        widget.bind("<Enter>", self.schedule)
        widget.bind("<Leave>", self.hide)

    def schedule(self, _=None):
        self.cancel()
        self.id = self.widget.after(self.delay, self.show)

    def cancel(self):
        if self.id:
            try:
                self.widget.after_cancel(self.id)
            except Exception:
                pass
            self.id = None

    def show(self):
        if self.tipwindow or not self.text:
            return
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 10
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        label = ttk.Label(tw, text=self.text, justify='left', background='#FFFFE0', relief='solid', borderwidth=1)
        label.pack(ipadx=6, ipady=3)

    def hide(self, _=None):
        self.cancel()
        if self.tipwindow:
            try:
                self.tipwindow.destroy()
            except Exception:
                pass
            self.tipwindow = None


class ComputerNamerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        # App name and version
        self.APP_NAME = "CCI_New_PC_Setup"
        self.VERSION = _av.read_version()
        self.title(f"{self.APP_NAME} v{self.VERSION}")
        self.resizable(True, True)
        self.geometry("820x620")
        self.minsize(720, 480)
        padding = {"padx": 8, "pady": 6}

        # Variables
        self.customer_var = tk.StringVar()
        self.type_var = tk.StringVar(value="Console")
        self.serial_var = tk.StringVar()
        self.custom_var = tk.StringVar()
        self.generated_var = tk.StringVar()
        self.tz_var = tk.StringVar()
        self.bg_path_var = tk.StringVar(value=r"H:\Shared drives\Marketing and Advertising\Branding_ Bios_Logos\Desktop Wallpaper\1920x1080CCI Desktop Wallpaper, 169 - Support Info.png")
        self.apply_power_var = tk.BooleanVar(value=True)
        self.apply_taskbar_var = tk.BooleanVar(value=True)
        self.apply_datetime_var = tk.BooleanVar(value=True)
        self.apply_notifications_var = tk.BooleanVar(value=True)
        self.apply_windowsupdate_var = tk.BooleanVar(value=False)
        self.apply_personalization_var = tk.BooleanVar(value=True)
        # New: fine-grained power button / lid actions
        self.power_button_do_nothing_var = tk.BooleanVar(value=False)
        self.sleep_button_do_nothing_var = tk.BooleanVar(value=False)
        self.lid_close_do_nothing_var = tk.BooleanVar(value=False)

        # assets directory (used for icons and default background)
        self.assets_dir = Path(os.path.dirname(__file__)) / "assets"
        self.assets_dir.mkdir(parents=True, exist_ok=True)

        # Notebook (tabs)
        notebook = ttk.Notebook(self)
        notebook.grid(column=0, row=0, sticky="nsew", **padding)
        # make the root expand so notebook and tabs can grow
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # Naming tab (separate, scrollable)
        naming_frame = ttk.Frame(notebook)
        notebook.add(naming_frame, text="Computer Name")

        # Use a ScrollableFrame so the inputs remain reachable on small windows
        name_scroll = ScrollableFrame(naming_frame)
        name_scroll.grid(row=0, column=0, sticky="nsew", padx=(8,0), pady=8)
        naming_frame.grid_rowconfigure(0, weight=1)
        naming_frame.grid_columnconfigure(0, weight=1)

        frm = name_scroll.inner
        frm.columnconfigure(0, weight=0)
        frm.columnconfigure(1, weight=1)

        ttk.Label(frm, text="Customer Name:").grid(column=0, row=0, sticky="w")
        ttk.Entry(frm, textvariable=self.customer_var, width=40).grid(column=1, row=0, sticky="we")
        ttk.Label(frm, text="Type of Computer:").grid(column=0, row=1, sticky="w")
        type_menu = ttk.OptionMenu(frm, self.type_var, self.type_var.get(), *TYPE_MAP.keys())
        type_menu.grid(column=1, row=1, sticky="w")

        ttk.Label(frm, text="Serial Number:").grid(column=0, row=2, sticky="w")
        vcmd = (self.register(self._validate_serial), "%P")
        ttk.Entry(frm, textvariable=self.serial_var, validate="key", validatecommand=vcmd, width=40).grid(column=1, row=2, sticky="we")

        ttk.Label(frm, text="Custom Name (optional):").grid(column=0, row=3, sticky="w")
        c_entry = ttk.Entry(frm, textvariable=self.custom_var, width=40)
        c_entry.grid(column=1, row=3, sticky="we")
        self.custom_var.trace_add("write", self._on_custom_change)

        ttk.Separator(frm, orient="horizontal").grid(column=0, row=4, columnspan=2, sticky="ew", pady=8)
        ttk.Label(frm, text="New Computer Name:").grid(column=0, row=5, sticky="w")
        ttk.Label(frm, textvariable=self.generated_var, foreground="blue").grid(column=1, row=5, sticky="w")
        self.accept_btn = ttk.Button(frm, text="Accept new computer name", command=self._on_accept, state="disabled")
        self.accept_btn.grid(column=0, row=6, columnspan=2, pady=(12, 0))
        for var in (self.customer_var, self.type_var, self.serial_var):
            var.trace_add("write", self._update_generated)
        self._update_generated()

        # System settings tab (left = scrollable inputs, right = preview)
        sys_frame = ttk.Frame(notebook)
        notebook.add(sys_frame, text="System Settings")
        sys_frame.grid_rowconfigure(0, weight=1)
        sys_frame.grid_columnconfigure(0, weight=1)
        sys_frame.grid_columnconfigure(1, weight=0)

        # Left: inputs (scrollable)
        left = ScrollableFrame(sys_frame)
        left.grid(row=0, column=0, sticky="nsew", padx=(8,0), pady=8)
        inputs = left.inner
        inputs.columnconfigure(0, weight=1)

        # Right: checklist & per-function toggles
        right = ttk.Frame(sys_frame)
        right.grid(row=0, column=1, sticky="nsew", padx=(8,8), pady=8)
        right.columnconfigure(0, weight=1)

        ttk.Label(inputs, text="Apply the following system settings (best-effort):").grid(column=0, row=0, sticky="w", pady=(0,6))

        # Use a Treeview so settings can be grouped and expanded/collapsed
        self.settings_tree = ttk.Treeview(inputs, show='tree')
        self.settings_tree.grid(column=0, row=1, sticky="nsew")
        inputs.rowconfigure(1, weight=1)

        # Build a mapping from tree items to BooleanVars so toggles update state
        self._tree_item_map = {}

        # Helper to insert items and remember the var mapping
        def _ins(parent, text, var):
            iid = self.settings_tree.insert(parent, 'end', text=text)
            self._tree_item_map[iid] = var
            # initialize display text with checkbox marker
            self._update_tree_item_text(iid)
            return iid

        # Top-level groups
        power_grp = self.settings_tree.insert('', 'end', text='Power')
        taskbar_grp = self.settings_tree.insert('', 'end', text='Taskbar')
        datetime_grp = self.settings_tree.insert('', 'end', text='Date & Time')
        notifications_grp = self.settings_tree.insert('', 'end', text='Notifications')
        windowsupdate_grp = self.settings_tree.insert('', 'end', text='Windows Update')
        personalization_grp = self.settings_tree.insert('', 'end', text='Personalization')

        # Power group children
        _ins(power_grp, 'Power (never sleep/display/disk)', self.apply_power_var)
        _ins(power_grp, 'Power button -> Do nothing', self.power_button_do_nothing_var)
        _ins(power_grp, 'Sleep button -> Do nothing', self.sleep_button_do_nothing_var)
        _ins(power_grp, 'Lid close -> Do nothing (laptops)', self.lid_close_do_nothing_var)

        # Taskbar
        _ins(taskbar_grp, 'Taskbar settings', self.apply_taskbar_var)

        # Date & Time
        _ins(datetime_grp, 'Date & Time (sync and timezone if provided)', self.apply_datetime_var)

        # Notifications
        _ins(notifications_grp, 'Notifications / Do Not Disturb', self.apply_notifications_var)

        # Windows Update
        _ins(windowsupdate_grp, 'Windows Update (run then stop service)', self.apply_windowsupdate_var)

        # Personalization
        _ins(personalization_grp, 'Personalization (dark + accent + background)', self.apply_personalization_var)

        # Allow toggling items with single-click; clicking a top-level/group toggles all children
        self.settings_tree.bind('<Button-1>', self._on_tree_click)

        # Time Zone entry remains below the tree
        ttk.Label(inputs, text="Time Zone ID (optional):").grid(column=0, row=2, sticky="w", pady=(8,0))
        ttk.Entry(inputs, textvariable=self.tz_var, width=40).grid(column=0, row=3, sticky="w")

        # action buttons
        ttk.Button(inputs, text="Preview actions", command=self._preview_system_actions).grid(column=0, row=12, pady=(12,0), sticky="w")
        ttk.Button(inputs, text="Apply selected settings", command=self._apply_system_settings).grid(column=0, row=12, pady=(12,0), sticky="e")

        # Right-side: function toggles and planned changes checklist
    # 'Select functions to run' section removed - settings are controlled by the left Treeview

        checklist_frame = ttk.LabelFrame(right, text="Planned changes")
        checklist_frame.grid(column=0, row=1, sticky="nsew", pady=(8,0))
        checklist_frame.columnconfigure(0, weight=0)
        checklist_frame.columnconfigure(1, weight=1)

        # Allow the right column to expand for the live preview area
        right.grid_rowconfigure(2, weight=1)

        # Live preview area (embedded in System Settings tab)
        preview_out_frame = ttk.LabelFrame(right, text="Preview output")
        preview_out_frame.grid(column=0, row=2, sticky="nsew", pady=(8,0))
        preview_out_frame.columnconfigure(0, weight=1)
        preview_out_frame.rowconfigure(0, weight=1)

        # toolbar for preview
        pf_toolbar = ttk.Frame(preview_out_frame)
        pf_toolbar.grid(column=0, row=1, sticky="ew", pady=(6,4))
        pf_toolbar.columnconfigure(0, weight=1)

        ttk.Button(pf_toolbar, text="Open last preview", command=lambda: self._open_last_preview()).pack(side='left')
        ttk.Button(pf_toolbar, text="Clear", command=lambda: self._clear_preview()).pack(side='left', padx=(6,4))
        ttk.Button(pf_toolbar, text="Save...", command=lambda: self._save_preview()).pack(side='left')

        # text widget for live output
        self.preview_text = tk.Text(preview_out_frame, wrap='none')
        self.preview_text.grid(column=0, row=0, sticky='nsew')
        self.preview_text.config(state='disabled')

        vsb2 = ttk.Scrollbar(preview_out_frame, orient='vertical', command=self.preview_text.yview)
        vsb2.grid(column=1, row=0, sticky='ns')
        self.preview_text.configure(yscrollcommand=vsb2.set)
        hsb2 = ttk.Scrollbar(preview_out_frame, orient='horizontal', command=self.preview_text.xview)
        hsb2.grid(column=0, row=2, sticky='ew')
        self.preview_text.configure(xscrollcommand=hsb2.set)

        # track last preview file
        self._last_preview_path = None

        # Define the granular planned actions and how they map to the top-level checkboxes
        self._detailed_plan = [
            ("Power button -> Do nothing", lambda: self.apply_power_var.get() and self.power_button_do_nothing_var.get()),
            ("Sleep button -> Do nothing", lambda: self.apply_power_var.get() and self.sleep_button_do_nothing_var.get()),
            ("Lid close -> Do nothing", lambda: self.apply_power_var.get() and self.lid_close_do_nothing_var.get()),
            ("Disk timeout (AC)", lambda: self.apply_power_var.get()),
            ("Disk timeout (DC)", lambda: self.apply_power_var.get()),
            ("Monitor timeout (AC)", lambda: self.apply_power_var.get()),
            ("Monitor timeout (DC)", lambda: self.apply_power_var.get()),
            ("Standby timeout (AC)", lambda: self.apply_power_var.get()),
            ("Standby timeout (DC)", lambda: self.apply_power_var.get()),
            ("Taskbar alignment (center)", lambda: self.apply_taskbar_var.get()),
            ("Taskbar size (default)", lambda: self.apply_taskbar_var.get()),
            ("Open Taskbar settings UI", lambda: self.apply_taskbar_var.get()),
            ("Set Time Zone (if provided)", lambda: self.apply_datetime_var.get() and bool(self.tz_var.get().strip())),
            ("Sync time now (w32tm /resync)", lambda: self.apply_datetime_var.get()),
            ("Disable toasts (PushNotifications::ToastEnabled=0)", lambda: self.apply_notifications_var.get()),
            ("Set Focus Assist policy (QuietHours)", lambda: self.apply_notifications_var.get()),
            ("Open Notifications settings UI", lambda: self.apply_notifications_var.get()),
            ("Install/Use PSWindowsUpdate module", lambda: self.apply_windowsupdate_var.get()),
            ("Run Windows Update via PSWindowsUpdate", lambda: self.apply_windowsupdate_var.get()),
            ("Stop and disable wuauserv", lambda: self.apply_windowsupdate_var.get()),
            ("Set dark theme (Apps/System)", lambda: self.apply_personalization_var.get()),
            ("Set accent color", lambda: self.apply_personalization_var.get()),
            ("Set desktop background (provided or sample)", lambda: self.apply_personalization_var.get()),
        ]

        # create small image icons (try to load PNGs from assets; otherwise create simple colored squares)
        try:
            ok_path = self.assets_dir / "icon_ok.png"
            off_path = self.assets_dir / "icon_off.png"
            if PIL_AVAILABLE and ok_path.exists() and off_path.exists():
                self.icon_ok = ImageTk.PhotoImage(Image.open(ok_path).resize((16,16), Image.LANCZOS))
                self.icon_off = ImageTk.PhotoImage(Image.open(off_path).resize((16,16), Image.LANCZOS))
            elif ok_path.exists() and off_path.exists():
                self.icon_ok = tk.PhotoImage(file=str(ok_path))
                self.icon_off = tk.PhotoImage(file=str(off_path))
            else:
                # fallback: create small colored squares
                self.icon_ok = tk.PhotoImage(width=16, height=16)
                self.icon_ok.put((("#00a000",),), to=(0,0,15,15))
                self.icon_off = tk.PhotoImage(width=16, height=16)
                self.icon_off.put((("#888888",),), to=(0,0,15,15))
        except Exception:
            # last-resort fallback to simple text markers
            self.icon_ok = None
            self.icon_off = None

        # create labels for each detailed item (icon + description)
        self.detailed_labels = {}
        r = 0
        for desc, checker in self._detailed_plan:
            if self.icon_off:
                icon = ttk.Label(checklist_frame, image=self.icon_off)
                icon.image = self.icon_off
            else:
                icon = ttk.Label(checklist_frame, text=" ", width=2)
            icon.grid(column=0, row=r, sticky="w", padx=(6,4))
            lbl = ttk.Label(checklist_frame, text=desc)
            lbl.grid(column=1, row=r, sticky="w", pady=2)
            self.detailed_labels[desc] = (icon, lbl, checker)
            r += 1

        # Attach short explanatory tooltips to each checklist item (icon + label)
        tooltip_map = {
            "Disk timeout (AC)": "Set disk idle timeout on AC power (powercfg).",
            "Disk timeout (DC)": "Set disk idle timeout on battery (powercfg).",
            "Monitor timeout (AC)": "Set display timeout on AC power (powercfg).",
            "Monitor timeout (DC)": "Set display timeout on battery (powercfg).",
            "Standby timeout (AC)": "Set system standby timeout on AC power (powercfg).",
            "Standby timeout (DC)": "Set system standby timeout on battery (powercfg).",
            "Taskbar alignment (center)": "Adjust taskbar alignment to center (registry/Explorer settings).",
            "Taskbar size (default)": "Reset taskbar size to default (registry/Explorer settings).",
            "Open Taskbar settings UI": "Opens the Taskbar settings UI for manual confirmation.",
            "Set Time Zone (if provided)": "Set the system time zone (tzutil).",
            "Sync time now (w32tm /resync)": "Force a time sync now using w32tm.",
            "Disable toasts (PushNotifications::ToastEnabled=0)": "Disable toast notifications via registry.",
            "Set Focus Assist policy (QuietHours)": "Enable Quiet Hours / Focus Assist via registry or settings.",
            "Open Notifications settings UI": "Opens the Notifications settings UI for manual confirmation.",
            "Install/Use PSWindowsUpdate module": "Acquire PSWindowsUpdate module to run Windows Update from PowerShell.",
            "Run Windows Update via PSWindowsUpdate": "Run Windows Update checks/installs via PSWindowsUpdate.",
            "Stop and disable wuauserv": "Stops and disables the Windows Update service (wuauserv).",
            "Set dark theme (Apps/System)": "Switch system and apps to Dark theme via registry/settings.",
            "Set accent color": "Set a Windows accent color to match corporate theme.",
            "Set desktop background (provided or sample)": "Set the desktop wallpaper to the provided image or sample.",
        }
        for desc, (icon, lbl, _) in self.detailed_labels.items():
            txt = tooltip_map.get(desc, "")
            if txt:
                try:
                    ToolTip(icon, txt)
                    ToolTip(lbl, txt)
                except Exception:
                    pass

        # attach var traces to update the checklist when selections change
        for v in (self.apply_power_var, self.apply_taskbar_var, self.apply_datetime_var, self.apply_notifications_var, self.apply_windowsupdate_var, self.apply_personalization_var, self.tz_var, self.bg_path_var, self.power_button_do_nothing_var, self.sleep_button_do_nothing_var, self.lid_close_do_nothing_var):
            try:
                def _trace_cb(*_args, var=v):
                    # update main checklist labels
                    try:
                        self._update_checklist()
                    except Exception:
                        pass
                    # if we have a settings tree, update mapped items
                    try:
                        if hasattr(self, '_tree_item_map'):
                            for iid, mapped_var in self._tree_item_map.items():
                                if mapped_var is var:
                                    self._update_tree_item_text(iid)
                    except Exception:
                        pass

                v.trace_add("write", _trace_cb)
            except Exception:
                pass

        # ensure checklist reflects initial checkbox states
        self._update_checklist()

        # Move background preview to its own tab for a larger preview area
        preview_tab = ttk.Frame(notebook)
        notebook.add(preview_tab, text="Background Preview")
        preview_tab.grid_rowconfigure(0, weight=1)
        preview_tab.grid_columnconfigure(0, weight=1)

        preview_frame = ttk.LabelFrame(preview_tab, text="Background preview")
        preview_frame.grid(column=0, row=0, padx=8, pady=8, sticky="nsew")
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)
        self.preview_label = ttk.Label(preview_frame)
        self.preview_label.grid(column=0, row=0, sticky="nsew", padx=8, pady=8)

        # Background selection controls (moved from System Settings)
        ctrl_frame = ttk.Frame(preview_frame)
        ctrl_frame.grid(column=0, row=1, sticky="ew", padx=8, pady=(6,8))
        ctrl_frame.columnconfigure(0, weight=1)
        ttk.Label(ctrl_frame, text="Background image (optional):").grid(column=0, row=0, sticky="w")
        self.bg_entry = ttk.Entry(ctrl_frame, textvariable=self.bg_path_var, width=60)
        self.bg_entry.grid(column=0, row=1, sticky="ew", pady=(2,0))
        ttk.Button(ctrl_frame, text="Browse…", command=self._browse_bg).grid(column=1, row=1, sticky="w", padx=(6,0))
        # Apply background now button
        ctrl_frame.columnconfigure(1, weight=0)
        ttk.Button(ctrl_frame, text="Apply background", command=self._apply_background).grid(column=2, row=1, sticky="w", padx=(6,0))
        ttk.Label(ctrl_frame, text="If empty, a sample corporate wallpaper will be used.").grid(column=0, row=2, columnspan=3, sticky="w", pady=(6,0))

        # When bg path changes, reload preview and update checklist
        try:
            self.bg_path_var.trace_add("write", lambda *_, __=None: (self._load_preview_image(self.bg_path_var.get()), self._update_checklist()))
        except Exception:
            try:
                self.bg_path_var.trace("w", lambda *args: (self._load_preview_image(self.bg_path_var.get()), self._update_checklist()))
            except Exception:
                pass

        # Ensure default background exists in assets and load initial preview
        self.default_bg_path = self.assets_dir / "default_background.png"
        self._ensure_embedded_background()
        initial = str(self.default_bg_path) if self.default_bg_path.exists() else self.bg_path_var.get()
        self._load_preview_image(initial)

    def _validate_serial(self, new_value: str) -> bool:
        if new_value == "":
            return True
        return new_value.isdigit()

    def _on_custom_change(self, *_):
        raw = self.custom_var.get()
        sanitized = sanitize_custom(raw)
        if sanitized != raw:
            try:
                traces = self.custom_var.trace_info()
                for t in traces:
                    self.custom_var.trace_remove(t[0], t[1])
            except Exception:
                pass
            self.custom_var.set(sanitized)
            self.custom_var.trace_add("write", self._on_custom_change)
        self._update_generated()

    def _update_generated(self, *_):
        custom = self.custom_var.get().strip()
        if custom:
            new_name = sanitize_custom(custom)
        else:
            cust = re.sub(r"\s+", "-", self.customer_var.get().strip())
            type_char = TYPE_MAP.get(self.type_var.get(), "D")
            serial = self.serial_var.get().strip()
            parts = [p for p in [cust, type_char, serial] if p]
            new_name = "-".join(parts)
            new_name = re.sub(r"\s+", "-", new_name)
        self.generated_var.set(new_name)
        if new_name:
            self.accept_btn.config(state="normal")
        else:
            self.accept_btn.config(state="disabled")

    def _update_tree_item_text(self, item_iid):
        """Update a tree item's displayed text to include a checkbox marker based on its mapped BooleanVar."""
        var = self._tree_item_map.get(item_iid)
        if var is None:
            return
        base_text = self.settings_tree.item(item_iid, 'text')
        # strip any existing marker
        if base_text.startswith('[x] ' ) or base_text.startswith('[ ] '):
            base_text = base_text[4:]
        marker = '[x]' if var.get() else '[ ]'
        self.settings_tree.item(item_iid, text=f"{marker} {base_text}")

    def _on_tree_doubleclick(self, event):
        iid = self.settings_tree.identify_row(event.y)
        if not iid:
            return
        var = self._tree_item_map.get(iid)
        if var is None:
            # might be a group; ignore
            return
        # toggle the var
        var.set(not var.get())
        # update text
        self._update_tree_item_text(iid)


    def _on_tree_click(self, event):
        """Single-click handler: toggles leaf items, or selects/deselects all descendants for groups."""
        iid = self.settings_tree.identify_row(event.y)
        if not iid:
            return

        # If this item is mapped to a BooleanVar, toggle it
        if iid in self._tree_item_map:
            var = self._tree_item_map[iid]
            var.set(not var.get())
            self._update_tree_item_text(iid)
            try:
                self._update_checklist()
            except Exception:
                pass
            return

        # Otherwise treat it as a group: gather descendant mapped items and toggle them all
        def _gather_descendants(item):
            items = []
            for child in self.settings_tree.get_children(item):
                if child in self._tree_item_map:
                    items.append(child)
                items.extend(_gather_descendants(child))
            return items

        descendants = _gather_descendants(iid)
        if not descendants:
            return

        # If any descendant is unchecked, set all to True; otherwise set all to False
        any_unchecked = any(not self._tree_item_map[d].get() for d in descendants)
        for d in descendants:
            self._tree_item_map[d].set(any_unchecked)
            self._update_tree_item_text(d)
        try:
            self._update_checklist()
        except Exception:
            pass


    def _on_accept(self):
        name = self.generated_var.get()
        if not name:
            return
        msg = (
            f"Change the computer name to '{name}'?\n\nThis operation requires Administrator privileges. "
            "The rename will not fully take effect until the system restarts."
        )
        if not messagebox.askokcancel("Confirm rename", msg):
            return
        if not is_admin():
            messagebox.showerror("Administrator required", "This operation must be run as Administrator. Please run the script elevated and try again.")
            return
        try:
            subprocess.run([
                "powershell", "-NoProfile", "-NonInteractive", "-Command",
                f"Rename-Computer -NewName '{name}' -Force"
            ], check=True)
            if messagebox.askyesno("Restart now?", "Rename queued successfully. Do you want to restart now for the change to take effect?"):
                subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", "Restart-Computer -Force"], check=True)
        except subprocess.CalledProcessError as e:
            messagebox.showerror("Error", f"Failed to rename computer:\n{e}")

    def _browse_bg(self):
        path = filedialog.askopenfilename(title="Select background image", filetypes=[("Images", "*.jpg;*.jpeg;*.png;*.bmp;*.gif"), ("All files", "*")])
        if path:
            self.bg_path_var.set(path)
            self._load_preview_image(path)

    def _ensure_embedded_background(self):
        try:
            if self.default_bg_path.exists():
                return
            src = Path(self.bg_path_var.get())
            if src.exists() and src.is_file():
                try:
                    shutil.copy2(str(src), str(self.default_bg_path))
                except Exception:
                    pass
        except Exception:
            pass

    def _load_preview_image(self, path_or_obj):
        try:
            path = str(path_or_obj)
            if not path or not os.path.exists(path):
                self.preview_label.config(image="")
                self.preview_label.image = None
                return
            if PIL_AVAILABLE:
                img = Image.open(path)
                max_w, max_h = 360, 360
                img.thumbnail((max_w, max_h), Image.LANCZOS)
                tk_img = ImageTk.PhotoImage(img)
                self.preview_label.config(image=tk_img)
                self.preview_label.image = tk_img
            else:
                try:
                    tk_img = tk.PhotoImage(file=path)
                    self.preview_label.config(image=tk_img)
                    self.preview_label.image = tk_img
                except Exception:
                    self.preview_label.config(text=os.path.basename(path))
                    self.preview_label.image = None
        except Exception:
            try:
                self.preview_label.config(image="")
                self.preview_label.image = None
            except Exception:
                pass

    def _update_checklist(self):
        """Refresh the planned changes checklist labels to reflect current selections."""
        # Update detailed labels (icon + description) using their checker functions
        for desc, trio in self.detailed_labels.items():
            icon, lbl, checker = trio
            try:
                applies = bool(checker())
            except Exception:
                applies = False
            # If we have image icons, switch images; otherwise fallback to text markers
            if applies:
                if getattr(self, 'icon_ok', None) is not None and getattr(icon, 'image', None) is not None:
                    icon.config(image=self.icon_ok)
                    icon.image = self.icon_ok
                else:
                    icon.config(text="✓", foreground="green")
                lbl.config(foreground="black")
            else:
                if getattr(self, 'icon_off', None) is not None and getattr(icon, 'image', None) is not None:
                    icon.config(image=self.icon_off)
                    icon.image = self.icon_off
                else:
                    icon.config(text="—", foreground="gray")
                lbl.config(foreground="gray")

    def _preview_system_actions(self):
        script = _resource_path('configure-windows.ps1')
        if not os.path.exists(script):
            messagebox.showerror("Script not found", f"Could not find {script}")
            return
        try:
            pscmd = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", script, "-DryRun"]
            if self.tz_var.get().strip():
                pscmd.extend(["-TimeZoneId", self.tz_var.get().strip()])
            if self.bg_path_var.get().strip():
                pscmd.extend(["-BackgroundPath", self.bg_path_var.get().strip()])
            elif self.default_bg_path.exists():
                pscmd.extend(["-BackgroundPath", str(self.default_bg_path)])
            # Pass fine-grained power action switches
            if self.power_button_do_nothing_var.get():
                pscmd.append("-PowerButtonDoNothing")
            if self.sleep_button_do_nothing_var.get():
                pscmd.append("-SleepButtonDoNothing")
            if self.lid_close_do_nothing_var.get():
                pscmd.append("-LidCloseDoNothing")
            # Pass Do* function selection switches only when user toggles them
            if self.do_powersettings_var.get():
                pscmd.append("-DoPowerSettings")
            if self.do_taskbar_var.get():
                pscmd.append("-DoTaskbar")
            if self.do_datetime_var.get():
                pscmd.append("-DoDateTime")
            if self.do_notifications_var.get():
                pscmd.append("-DoNotifications")
            if self.do_windowsupdate_var.get():
                pscmd.append("-DoWindowsUpdate")
            if self.do_personalization_var.get():
                pscmd.append("-DoPersonalization")
            if self.do_powerbuttonactions_var.get():
                pscmd.append("-DoPowerButtonActions")
            # refresh checklist to reflect user's current choices before running
            self._update_checklist()
            # Run powershell and capture output to a temporary file, then show it in a pop-up
            import tempfile
            tf = tempfile.NamedTemporaryFile(delete=False, suffix='.txt')
            tf.close()
            with open(tf.name, 'w', encoding='utf-8', errors='replace') as fh:
                subprocess.run(pscmd, check=False, stdout=fh, stderr=subprocess.STDOUT)
            # Load and display the output in the GUI
            try:
                with open(tf.name, 'r', encoding='utf-8', errors='replace') as fh:
                    out = fh.read()
            except Exception:
                out = "(Could not read output file)"
            self._show_ps_output(out, title="Preview output (DryRun)")
            messagebox.showinfo("Preview actions", "PowerShell configurator finished DryRun — output shown in the preview window.")
        except subprocess.CalledProcessError as e:
            messagebox.showerror("Preview failed", f"Preview run failed: {e}")

    def _show_ps_output(self, text, title="Output"):
        """Show PowerShell output text in a modal read-only window with scrollbar."""
        win = tk.Toplevel(self)
        win.title(title)
        win.geometry("720x480")
        win.transient(self)
        win.grab_set()

        frm = ttk.Frame(win)
        frm.pack(fill='both', expand=True)

        txt = tk.Text(frm, wrap='none')
        txt.insert('1.0', text)
        txt.config(state='disabled')
        txt.pack(side='left', fill='both', expand=True)

        vsb = ttk.Scrollbar(frm, orient='vertical', command=txt.yview)
        vsb.pack(side='right', fill='y')
        txt.configure(yscrollcommand=vsb.set)

        # Horizontal scrollbar
        hsb = ttk.Scrollbar(win, orient='horizontal', command=txt.xview)
        hsb.pack(side='bottom', fill='x')
        txt.configure(xscrollcommand=hsb.set)

        btn_frame = ttk.Frame(win)
        btn_frame.pack(fill='x')
        def _save():
            p = filedialog.asksaveasfilename(defaultextension='.txt', filetypes=[('Text', '*.txt'), ('All', '*.*')])
            if p:
                try:
                    with open(p, 'w', encoding='utf-8', errors='replace') as fh:
                        fh.write(text)
                except Exception as e:
                    messagebox.showerror('Save failed', str(e))

        ttk.Button(btn_frame, text='Save...', command=_save).pack(side='right', padx=6, pady=6)
        ttk.Button(btn_frame, text='Close', command=win.destroy).pack(side='right', pady=6)

    def _apply_system_settings(self):
        script = _resource_path('configure-windows.ps1')
        if not os.path.exists(script):
            messagebox.showerror("Script not found", f"Could not find {script}")
            return
        if not is_admin():
            messagebox.showerror("Administrator required", "Applying system settings requires Administrator privileges. Run this script elevated and try again.")
            return
        if not messagebox.askyesno("Apply settings", "Apply system settings now? This will run the PowerShell configurator script which may change registry and stop services."):
            return
        try:
            cmd = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", script]
            if self.tz_var.get().strip():
                cmd.extend(["-TimeZoneId", self.tz_var.get().strip()])
            if self.bg_path_var.get().strip():
                cmd.extend(["-BackgroundPath", self.bg_path_var.get().strip()])
            elif self.default_bg_path.exists():
                cmd.extend(["-BackgroundPath", str(self.default_bg_path)])
            # Pass fine-grained power action switches
            if self.power_button_do_nothing_var.get():
                cmd.append("-PowerButtonDoNothing")
            if self.sleep_button_do_nothing_var.get():
                cmd.append("-SleepButtonDoNothing")
            if self.lid_close_do_nothing_var.get():
                cmd.append("-LidCloseDoNothing")
            # Pass Do* function selection switches only when user toggles them
            if self.do_powersettings_var.get():
                cmd.append("-DoPowerSettings")
            if self.do_taskbar_var.get():
                cmd.append("-DoTaskbar")
            if self.do_datetime_var.get():
                cmd.append("-DoDateTime")
            if self.do_notifications_var.get():
                cmd.append("-DoNotifications")
            if self.do_windowsupdate_var.get():
                cmd.append("-DoWindowsUpdate")
            if self.do_personalization_var.get():
                cmd.append("-DoPersonalization")
            if self.do_powerbuttonactions_var.get():
                cmd.append("-DoPowerButtonActions")
            # refresh checklist to reflect user's current choices before running
            self._update_checklist()
            subprocess.run(cmd, check=True)
            messagebox.showinfo("Apply settings", "PowerShell configurator executed. Review the PowerShell output for details.")
        except subprocess.CalledProcessError as e:
            messagebox.showerror("Apply failed", f"PowerShell configurator failed: {e}")

    def _apply_background(self):
        """Apply the selected background immediately by running the PowerShell configurator with -BackgroundPath.

        This operation requires Administrator privileges because setting the system wallpaper may change system settings.
        """
        script = _resource_path('configure-windows.ps1')
        if not os.path.exists(script):
            messagebox.showerror("Script not found", f"Could not find {script}")
            return
        if not is_admin():
            messagebox.showerror("Administrator required", "Applying the background requires Administrator privileges. Run this program elevated and try again.")
            return
        bg = self.bg_path_var.get().strip() or (str(self.default_bg_path) if self.default_bg_path.exists() else "")
        if not bg or not os.path.exists(bg):
            messagebox.showerror("Background not found", "Selected background file does not exist.")
            return
        if not messagebox.askyesno("Apply background", f"Apply the selected background now?\n{bg}"):
            return
        try:
            cmd = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", script, "-BackgroundPath", bg]
            subprocess.run(cmd, check=True)
            messagebox.showinfo("Background applied", "Background applied successfully.")
        except subprocess.CalledProcessError as e:
            messagebox.showerror("Apply failed", f"Failed to apply background: {e}")

if __name__ == "__main__":
    app = ComputerNamerApp()
    app.mainloop()
