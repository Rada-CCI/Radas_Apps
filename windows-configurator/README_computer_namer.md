Computer Namer GUI

What this does
- Provides a small Tkinter GUI to generate a new computer name from three fields or a custom name.
- Enforces custom-name rules: max 15 chars, only letters/numbers/-/_ allowed, spaces converted to '-'.
- If a custom name is provided it overrides the generated name from the other fields.
- When you accept, the script uses PowerShell's Rename-Computer; Administrator rights are required and a restart is needed to apply the name change.

Inputs
1. Customer Name — free text.
2. Type of Computer — dropdown: Console, Laptop, Rack PC, Desktop.
   - Mapped to single character: Console=C, Laptop=L, Rack PC=R, Desktop=D.
3. Serial Number — digits only.
4. Custom Name (optional) — enforced rules above.

How generated name is formed
- If Custom Name is present and non-empty: sanitized custom name is used.
- Otherwise: [Customer Name]-[TypeChar]-[SerialNumber] with spaces replaced by '-'.

Run
- Requires Python 3 and Tkinter installed (usually included on Windows Python).
- To run interactively (recommended):

```powershell
python "..\windows-configurator\computer_namer_gui.py"
```

Notes
- The script must be run elevated to perform the rename. If not elevated, the GUI will prompt and refuse to perform the change.
- The rename command is executed via PowerShell; the system must be restarted for the new name to appear everywhere.
- This tool only changes the computer name. It does not attempt to change domain membership, Active Directory records, or other inventory systems.
