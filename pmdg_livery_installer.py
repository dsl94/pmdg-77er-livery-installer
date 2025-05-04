import os
import shutil
import zipfile
import json
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
import subprocess

CONFIG_PATH = Path.home() / ".pmdg_livery_config.json"

def load_config():
    if CONFIG_PATH.exists():
        with open(CONFIG_PATH, 'r') as f:
            return json.load(f)
    return {}

def save_config(data):
    with open(CONFIG_PATH, 'w') as f:
        json.dump(data, f)

def update_folder_label():
    config = load_config()
    folder = config.get("community_folder", "Not set")
    folder_label_var.set(f"Selected Community folder:\n{folder}")

def select_community_folder():
    folder = filedialog.askdirectory(title="Select your PMDG Community Folder (contains pmdg-aircraft-77er-liveries)")
    if folder:
        config = load_config()
        config["community_folder"] = folder
        save_config(config)
        update_folder_label()
        messagebox.showinfo("Saved", f"Community folder saved:\n{folder}")
        return folder
    return None

def extract_livery(zip_path, community_folder):
    try:
        simobjects_dir = Path(community_folder) / "pmdg-aircraft-77er-liveries" / "SimObjects" / "Airplanes"
        community_root = Path(community_folder) / "pmdg-aircraft-77er-liveries"
        simobjects_dir.mkdir(parents=True, exist_ok=True)

        folder_name = os.path.splitext(os.path.basename(zip_path))[0]
        output_folder = simobjects_dir / folder_name
        output_folder.mkdir(parents=True, exist_ok=True)

        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(output_folder)

        cfg_path = output_folder / "aircraft.cfg"
        livery_json_path = output_folder / "livery.json"
        options_ini_path = output_folder / "options.ini"

        if not livery_json_path.exists():
            raise Exception("livery.json not found")

        # Load atcId
        with open(livery_json_path, 'r') as f:
            livery_data = json.load(f)

        atc_id = livery_data.get("atcId", "")
        if not atc_id:
            raise Exception("atcId not found in livery.json")

        renamed_ini = output_folder / f"{atc_id}.ini"
        if options_ini_path.exists():
            options_ini_path.rename(renamed_ini)

        # Determine ini destination
        ini_target_paths = [
            Path.home() / "AppData" / "Local" / "Packages" / "Microsoft.FlightSimulator_8wekyb3d8bbwe" / "LocalState" / "Packages" / "pmdg-aircraft-77er" / "work" / "Aircraft",
            Path.home() / "AppData" / "Roaming" / "Microsoft Flight Simulator" / "Packages" / "pmdg-aircraft-77er" / "work" / "Aircraft"
        ]
        ini_target = next((p for p in ini_target_paths if p.exists()), None)
        if not ini_target:
            raise Exception("Could not find PMDG Aircraft folder (work/Aircraft)")

        shutil.copy(renamed_ini, ini_target)

        # Regenerate layout
        layout_json = community_root / "layout.json"
        layout_gen_exe = community_root / "MSFSLayoutGenerator.exe"
        gen_layout_bat = community_root / "GEN_LAYOUT.bat"

        if layout_gen_exe.exists():
            subprocess.run([str(layout_gen_exe), str(layout_json)])
        elif gen_layout_bat.exists():
            subprocess.run([str(gen_layout_bat)], shell=True)

        return f"Livery installed: '{folder_name}' with atcId: '{atc_id}'"
    except Exception as e:
        return f"Error: {str(e)}"

def browse_zip():
    zip_path = filedialog.askopenfilename(filetypes=[("ZIP files", "*.zip")])
    if not zip_path:
        return

    config = load_config()
    community = config.get("community_folder")

    if not community or not os.path.exists(community):
        community = select_community_folder()
        if not community:
            messagebox.showwarning("Canceled", "Community folder not selected.")
            return

    result = extract_livery(zip_path, community)
    messagebox.showinfo("Result", result)

def change_community_folder():
    select_community_folder()

# GUI
root = tk.Tk()
root.title("PMDG 77ER Livery Installer")

frame = tk.Frame(root, padx=20, pady=20)
frame.pack()

tk.Label(frame, text="Select a PMDG 77ER livery .zip file:").pack()

tk.Button(frame, text="Install Livery (.zip)", command=browse_zip).pack(pady=10)
tk.Button(frame, text="Change Community Folder", command=change_community_folder).pack(pady=5)

# Show current community folder
folder_label_var = tk.StringVar()
update_folder_label()
tk.Label(frame, textvariable=folder_label_var, wraplength=400, justify="center", fg="blue").pack(pady=10)

root.mainloop()
