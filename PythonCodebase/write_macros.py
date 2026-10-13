import sys
import shutil
from pathlib import Path
from datetime import datetime

# Ensure UTF-8 output encoding for Windows terminal compatibility
if sys.platform == 'win32' and hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass


def ensure_access_vbom():
    """
    Ensure 'Trust access to the VBA project object model' is enabled in Windows Registry
    so Excel allows programmatic modification of macros via COM.
    """
    if sys.platform != 'win32':
        return
    try:
        import winreg
        for ver in ["16.0", "15.0", "14.0"]:
            key_path = rf"Software\Microsoft\Office\{ver}\Excel\Security"
            try:
                with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                    try:
                        val, _ = winreg.QueryValueEx(key, "AccessVBOM")
                        if val != 1:
                            winreg.SetValueEx(key, "AccessVBOM", 0, winreg.REG_DWORD, 1)
                    except FileNotFoundError:
                        winreg.SetValueEx(key, "AccessVBOM", 0, winreg.REG_DWORD, 1)
            except Exception:
                pass
    except Exception as e:
        print(f"[NOTE] Registry AccessVBOM check: {e}")


def strip_vba_attributes(vba_code):
    """
    Strips leading file-level 'Attribute VB_...' headers from extracted VBA files.
    The Excel CodeModule object requires the code body without these file-level headers.
    """
    lines = vba_code.splitlines()
    body_lines = []
    in_header = True
    for line in lines:
        stripped = line.strip()
        if in_header:
            if stripped.startswith("Attribute VB_") or stripped == "":
                continue
            else:
                in_header = False
                body_lines.append(line)
        else:
            body_lines.append(line)
    return "\r\n".join(body_lines)


def create_backup(excel_path):
    """
    Creates a timestamped backup of the Excel workbook before modifying it.
    """
    backup_dir = excel_path.parent / "Backups"
    backup_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_file = backup_dir / f"{excel_path.stem}_{timestamp}{excel_path.suffix}"
    shutil.copy2(excel_path, backup_file)
    return backup_file


def write_vba_to_workbook(excel_path, vba_source_dir, backup=True):
    """
    Writes extracted/edited VBA files from vba_source_dir back into excel_path.
    
    Structure expected in vba_source_dir:
      ├── microsoft excel objects/  (*.cls)
      └── modules/                  (*.bas)
      
    Args:
        excel_path (Path or str): Path to the target Excel file.
        vba_source_dir (Path or str): Path to the source folder containing VBA code.
        backup (bool): Whether to create a backup file before writing.
        
    Returns:
        tuple (int, int): Count of updated objects and modules.
    """
    try:
        import win32com.client
    except ImportError:
        print("[ERROR] Missing required library 'pywin32'.")
        print("Please install it using: pip install pywin32")
        return (0, 0)

    excel_file = Path(excel_path).resolve()
    vba_dir = Path(vba_source_dir).resolve()

    if not excel_file.exists():
        print(f"[ERROR] Target Excel file not found: {excel_file}")
        return (0, 0)

    if not vba_dir.exists() or not vba_dir.is_dir():
        print(f"[ERROR] Source VBA folder not found: {vba_dir}")
        return (0, 0)

    print('=' * 80)
    print('VBA WRITE / INJECTION PROCESS')
    print('=' * 80)
    print(f"Target Excel File: {excel_file}")
    print(f"Source VBA Folder: {vba_dir}")
    print('=' * 80)

    if backup:
        try:
            bak = create_backup(excel_file)
            print(f"[BACKUP] Created backup copy: {bak.name}\n")
        except Exception as e:
            print(f"[WARNING] Could not create backup: {e}\n")

    ensure_access_vbom()

    excel_app = win32com.client.Dispatch("Excel.Application")
    excel_app.Visible = False
    excel_app.DisplayAlerts = False
    excel_app.EnableEvents = False        # Prevent Workbook_Open events from firing
    excel_app.AskToUpdateLinks = False

    workbook = None
    obj_updated = 0
    mod_updated = 0

    try:
        workbook = excel_app.Workbooks.Open(str(excel_file))
        vb_project = workbook.VBProject

        # Map existing components by case-insensitive name
        components = {comp.Name.lower(): comp for comp in vb_project.VBComponents}

        # 1. Process Microsoft Excel Objects (Worksheets & Workbook)
        excel_objs_dir = vba_dir / "microsoft excel objects"
        if excel_objs_dir.exists():
            cls_files = sorted(list(excel_objs_dir.glob("*.cls")) + list(excel_objs_dir.glob("*.bas")))
            for f in cls_files:
                comp_name = f.stem.lower()
                with open(f, 'r', encoding='utf-8', errors='replace') as fp:
                    raw_code = fp.read()
                clean_code = strip_vba_attributes(raw_code)

                if comp_name in components:
                    comp = components[comp_name]
                    cm = comp.CodeModule
                    if cm.CountOfLines > 0:
                        cm.DeleteLines(1, cm.CountOfLines)
                    if clean_code.strip():
                        cm.AddFromString(clean_code)
                    print(f"  [Excel Object] Updated {comp.Name} from {f.name}")
                    obj_updated += 1
                else:
                    print(f"  [WARNING] Component {f.stem} not found in workbook {excel_file.name}. Skipping.")

        # 2. Process Standard Modules
        modules_dir = vba_dir / "modules"
        if modules_dir.exists():
            bas_files = sorted(list(modules_dir.glob("*.bas")) + list(modules_dir.glob("*.cls")))
            for f in bas_files:
                comp_name = f.stem.lower()
                with open(f, 'r', encoding='utf-8', errors='replace') as fp:
                    raw_code = fp.read()
                clean_code = strip_vba_attributes(raw_code)

                if comp_name in components:
                    comp = components[comp_name]
                    cm = comp.CodeModule
                    if cm.CountOfLines > 0:
                        cm.DeleteLines(1, cm.CountOfLines)
                    if clean_code.strip():
                        cm.AddFromString(clean_code)
                    print(f"  [Module]       Updated {comp.Name} from {f.name}")
                    mod_updated += 1
                else:
                    # New module created in VS Code: add it to the VBProject
                    vbext_ct_std_module = 1
                    new_comp = vb_project.VBComponents.Add(vbext_ct_std_module)
                    new_comp.Name = f.stem
                    components[new_comp.Name.lower()] = new_comp
                    cm = new_comp.CodeModule
                    if cm.CountOfLines > 0:
                        cm.DeleteLines(1, cm.CountOfLines)
                    if clean_code.strip():
                        cm.AddFromString(clean_code)
                    print(f"  [Module]       Created new module {new_comp.Name} from {f.name}")
                    mod_updated += 1

        print(f"\nSaving changes to {excel_file.name}...")
        workbook.Save()
        print("[SUCCESS] Workbook saved successfully!")
        print(f"Summary: {obj_updated} Excel object(s) updated, {mod_updated} module(s) updated.\n")

    except Exception as e:
        print(f"\n[ERROR] Failed to write VBA back to {excel_file.name}: {e}")
        import traceback
        traceback.print_exc()
    finally:
        if workbook is not None:
            try:
                workbook.Close(False)
            except Exception:
                pass
        try:
            excel_app.Quit()
        except Exception:
            pass

    return (obj_updated, mod_updated)


def find_default_directories():
    """
    Locates the ExcelFiles and VBAFiles directories relative to the script or project workspace.
    """
    script_dir = Path(__file__).resolve().parent
    workspace_dir = script_dir.parent

    # Candidates for ExcelFiles
    candidates_excel = [
        workspace_dir / "ExcelFiles",
        Path.cwd() / "ExcelFiles",
        script_dir / "ExcelFiles",
    ]
    excel_dir = None
    for c in candidates_excel:
        if c.exists() and c.is_dir():
            excel_dir = c
            break
    if not excel_dir:
        excel_dir = workspace_dir / "ExcelFiles"

    # Candidates for VBAFiles
    candidates_vba = [
        workspace_dir / "VBAFiles",
        Path.cwd() / "VBAFiles",
        script_dir / "VBAFiles",
    ]
    vba_parent_dir = None
    for c in candidates_vba:
        if c.exists() and c.is_dir():
            vba_parent_dir = c
            break
    if not vba_parent_dir:
        vba_parent_dir = workspace_dir / "VBAFiles"

    return excel_dir, vba_parent_dir


def write_all_from_folder(vba_parent_dir, excel_dir):
    """
    Iterates through each source-file folder in vba_parent_dir,
    locates the corresponding Excel workbook in excel_dir,
    and updates its VBA macros.
    """
    vba_path = Path(vba_parent_dir).resolve()
    excel_path = Path(excel_dir).resolve()

    if not vba_path.exists() or not vba_path.is_dir():
        print(f"[ERROR] VBAFiles directory not found: {vba_path}")
        return {}

    if not excel_path.exists() or not excel_path.is_dir():
        print(f"[ERROR] ExcelFiles directory not found: {excel_path}")
        return {}

    # List subdirectories in VBAFiles (each is a source file folder)
    source_folders = sorted([d for d in vba_path.iterdir() if d.is_dir()])

    if not source_folders:
        print(f"No source folders found in {vba_path}")
        return {}

    print('=' * 80)
    print('BATCH VBA WRITE-BACK PROCESS')
    print('=' * 80)
    print(f"VBA Source Directory:  {vba_path}")
    print(f"Excel Target Directory: {excel_path}")
    print(f"Found {len(source_folders)} workbook folder(s) to process:")
    for d in source_folders:
        print(f"  - {d.name}")
    print('=' * 80 + '\n')

    results = {}
    for folder in source_folders:
        # Match Excel file in ExcelFiles with same stem
        target_excel = None
        for ext in ('.xlsm', '.xls', '.xlsb', '.xlam'):
            candidate = excel_path / f"{folder.name}{ext}"
            if candidate.exists():
                target_excel = candidate
                break

        if not target_excel:
            print(f"[WARNING] No matching Excel workbook found in {excel_path.name} for folder '{folder.name}'")
            continue

        obj_cnt, mod_cnt = write_vba_to_workbook(target_excel, folder)
        results[folder.name] = (obj_cnt, mod_cnt)

    print('=' * 80)
    print('ALL WORKBOOKS UPDATED SUCCESSFULLY!')
    print('=' * 80)
    for name, (o_cnt, m_cnt) in results.items():
        print(f"  {name}: {o_cnt} object(s), {m_cnt} module(s) pushed.")
    print('=' * 80)

    return results


if __name__ == "__main__":
    default_excel_dir, default_vba_dir = find_default_directories()

    # Allow command line overrides
    if len(sys.argv) > 1:
        arg = Path(sys.argv[1]).resolve()
        if arg.is_file():
            # If user provided a specific Excel file: find its VBA folder
            folder = default_vba_dir / arg.stem
            write_vba_to_workbook(arg, folder)
            sys.exit(0)
        elif arg.is_dir():
            # If user provided a folder inside VBAFiles
            if (arg / "microsoft excel objects").exists() or (arg / "modules").exists():
                # Locate corresponding Excel file
                for ext in ('.xlsm', '.xls', '.xlsb', '.xlam'):
                    cand = default_excel_dir / f"{arg.name}{ext}"
                    if cand.exists():
                        write_vba_to_workbook(cand, arg)
                        sys.exit(0)
                print(f"[ERROR] Could not find Excel workbook matching '{arg.name}' in {default_excel_dir}")
                sys.exit(1)
            else:
                default_vba_dir = arg

    write_all_from_folder(default_vba_dir, default_excel_dir)

