import re
import subprocess
import ctypes
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import winreg
import os
import shutil
from pathlib import Path
import app_version as _av

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

        left = ScrollableFrame(sys_frame)
        left.grid(row=0, column=0, sticky="nsew", padx=(8,0), pady=8)

        inputs = left.inner
        inputs.columnconfigure(0, weight=1)

        ttk.Label(inputs, text="Apply the following system settings (best-effort):").grid(column=0, row=0, sticky="w", pady=(0,6))
        ttk.Checkbutton(inputs, text="Power (never sleep/display/disk)", variable=self.apply_power_var).grid(column=0, row=1, sticky="w")
        ttk.Checkbutton(inputs, text="Taskbar settings", variable=self.apply_taskbar_var).grid(column=0, row=2, sticky="w")
        ttk.Checkbutton(inputs, text="Date & Time (sync and timezone if provided)", variable=self.apply_datetime_var).grid(column=0, row=3, sticky="w")
        ttk.Checkbutton(inputs, text="Notifications / Do Not Disturb", variable=self.apply_notifications_var).grid(column=0, row=4, sticky="w")
        ttk.Checkbutton(inputs, text="Windows Update (run then stop service)", variable=self.apply_windowsupdate_var).grid(column=0, row=5, sticky="w")
        ttk.Checkbutton(inputs, text="Personalization (dark + accent + background)", variable=self.apply_personalization_var).grid(column=0, row=6, sticky="w")

        ttk.Label(inputs, text="Time Zone ID (optional):").grid(column=0, row=7, sticky="w", pady=(8,0))
        ttk.Entry(inputs, textvariable=self.tz_var, width=40).grid(column=0, row=8, sticky="w")

        ttk.Label(inputs, text="Background image (optional):").grid(column=0, row=9, sticky="w", pady=(8,0))
        ttk.Entry(inputs, textvariable=self.bg_path_var, width=40).grid(column=0, row=10, sticky="w")
        ttk.Button(inputs, text="Browse…", command=self._browse_bg).grid(column=0, row=11, sticky="w", pady=(4,0))

        ttk.Button(inputs, text="Preview actions", command=self._preview_system_actions).grid(column=0, row=12, pady=(12,0), sticky="w")
        ttk.Button(inputs, text="Apply selected settings", command=self._apply_system_settings).grid(column=0, row=12, pady=(12,0), sticky="e")

        # Detailed planned changes checklist (icons + per-setting list)
        checklist_frame = ttk.LabelFrame(inputs, text="Planned changes")
        checklist_frame.grid(column=0, row=13, columnspan=2, sticky="ew", pady=(12,8))
        checklist_frame.columnconfigure(0, weight=0)
        checklist_frame.columnconfigure(1, weight=1)

        # Define the granular planned actions and how they map to the top-level checkboxes
        self._detailed_plan = [
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

        # create labels for each detailed item (icon + description)
        self.detailed_labels = {}
        r = 0
        for desc, checker in self._detailed_plan:
            icon = ttk.Label(checklist_frame, text=" ", width=2)
            icon.grid(column=0, row=r, sticky="w", padx=(6,4))
            lbl = ttk.Label(checklist_frame, text=desc)
            lbl.grid(column=1, row=r, sticky="w", pady=2)
            self.detailed_labels[desc] = (icon, lbl, checker)
            r += 1

        # attach var traces to update the checklist when selections change
        for v in (self.apply_power_var, self.apply_taskbar_var, self.apply_datetime_var, self.apply_notifications_var, self.apply_windowsupdate_var, self.apply_personalization_var, self.tz_var, self.bg_path_var):
            try:
                v.trace_add("write", lambda *_, __=None: self._update_checklist())
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

        # Ensure assets dir exists and copy default background into assets if available
        self.assets_dir = Path(os.path.dirname(__file__)) / "assets"
        self.assets_dir.mkdir(parents=True, exist_ok=True)
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
            if applies:
                icon.config(text="✓", foreground="green")
                lbl.config(foreground="black")
            else:
                icon.config(text="—", foreground="gray")
                lbl.config(foreground="gray")

    def _preview_system_actions(self):
        script = os.path.join(os.path.dirname(__file__), 'configure-windows.ps1')
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
            # refresh checklist to reflect user's current choices before running
            self._update_checklist()
            subprocess.run(pscmd, check=True)
            messagebox.showinfo("Preview actions", "PowerShell configurator ran in DryRun mode. Check output in the PowerShell window that appeared.")
        except subprocess.CalledProcessError as e:
            messagebox.showerror("Preview failed", f"Preview run failed: {e}")

    def _apply_system_settings(self):
        script = os.path.join(os.path.dirname(__file__), 'configure-windows.ps1')
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
            # refresh checklist to reflect user's current choices before running
            self._update_checklist()
            subprocess.run(cmd, check=True)
            messagebox.showinfo("Apply settings", "PowerShell configurator executed. Review the PowerShell output for details.")
        except subprocess.CalledProcessError as e:
            messagebox.showerror("Apply failed", f"PowerShell configurator failed: {e}")


if __name__ == "__main__":
    app = ComputerNamerApp()
    app.mainloop()
