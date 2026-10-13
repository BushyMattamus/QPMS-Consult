import sys
import hashlib
import json
import shutil
from pathlib import Path
from datetime import datetime

PROJECT_ROOT = Path(__file__).resolve().parent.parent
PRODUCTION_DIR = PROJECT_ROOT / "ExcelFiles"
STAGING_DIR = PROJECT_ROOT / "StagedExcelFiles"
TEST_RUNS_DIR = PROJECT_ROOT / ".test_runs"

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


def sha256_file(file_path):
    """Return the SHA-256 digest of a file without loading it all into memory."""
    digest = hashlib.sha256()
    with open(file_path, "rb") as file:
        for chunk in iter(lambda: file.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def vba_source_files(source_dir):
    """Return source files that are consumed by the VBA writer."""
    source_dir = Path(source_dir)
    files = []
    for subdirectory in ("microsoft excel objects", "modules"):
        source_subdirectory = source_dir / subdirectory
        if source_subdirectory.is_dir():
            files.extend(
                source_file
                for source_file in source_subdirectory.iterdir()
                if source_file.is_file() and source_file.suffix.lower() in {".bas", ".cls"}
            )
    return sorted(files, key=lambda path: str(path.relative_to(source_dir)).casefold())


def sha256_vba_source(source_dir):
    """Hash all VBA files consumed by the writer, including their relative paths."""
    source_dir = Path(source_dir)
    digest = hashlib.sha256()
    for source_file in vba_source_files(source_dir):
        relative_path = source_file.relative_to(source_dir).as_posix().casefold()
        digest.update(relative_path.encode("utf-8"))
        digest.update(b"\0")
        digest.update(sha256_file(source_file).encode("ascii"))
        digest.update(b"\0")
    return digest.hexdigest()


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
        raise RuntimeError(
            "Missing required library 'pywin32'. Install it using: pip install pywin32"
        )

    excel_file = Path(excel_path).resolve()
    vba_dir = Path(vba_source_dir).resolve()

    if excel_file.parent == PRODUCTION_DIR.resolve():
        raise ValueError(
            "Direct writes to ExcelFiles are disabled. Build the workbook under "
            "StagedExcelFiles and promote it only after successful tests."
        )

    if not excel_file.exists():
        raise FileNotFoundError(f"Target Excel file not found: {excel_file}")

    if not vba_dir.exists() or not vba_dir.is_dir():
        raise FileNotFoundError(f"Source VBA folder not found: {vba_dir}")

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
        raise
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
    locates the corresponding production workbook in excel_dir,
    and writes a fresh VBA-updated copy to staging.
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
    print('BATCH STAGED WORKBOOK BUILD')
    print('=' * 80)
    print(f"VBA Source Directory:  {vba_path}")
    print(f"Production Source Directory: {excel_path}")
    print(f"Staging Target Directory: {STAGING_DIR}")
    print(f"Found {len(source_folders)} workbook folder(s) to process:")
    for d in source_folders:
        print(f"  - {d.name}")
    print('=' * 80 + '\n')

    results = {}
    for folder in source_folders:
        # Match Excel file in ExcelFiles with same stem
        source_excel = None
        for ext in ('.xlsm', '.xls', '.xlsb', '.xlam'):
            candidate = excel_path / f"{folder.name}{ext}"
            if candidate.exists():
                source_excel = candidate
                break

        if not source_excel:
            print(f"[WARNING] No matching Excel workbook found in {excel_path.name} for folder '{folder.name}'")
            continue

        target_excel = stage_single_workbook(source_excel, folder)
        print(f"[STAGED] {target_excel}")
        with open(target_excel.with_suffix(".release.json"), "r", encoding="utf-8") as file:
            metadata = json.load(file)
        obj_cnt = metadata["objects_updated"]
        mod_cnt = metadata["modules_updated"]
        results[folder.name] = (obj_cnt, mod_cnt)

    print('=' * 80)
    print('ALL MATCHING WORKBOOKS STAGED SUCCESSFULLY!')
    print('=' * 80)
    for name, (o_cnt, m_cnt) in results.items():
        print(f"  {name}: {o_cnt} object(s), {m_cnt} module(s) injected.")
    print('=' * 80)

    return results


def find_changed_vba_targets(vba_parent_dir, excel_dir, staging_dir=STAGING_DIR):
    """Find workbooks whose VBA source differs from the last staged source snapshot."""
    vba_path = Path(vba_parent_dir).resolve()
    excel_path = Path(excel_dir).resolve()
    stage_path = Path(staging_dir).resolve()
    changed_targets = []
    for source_dir in sorted(vba_path.iterdir(), key=lambda path: path.name.casefold()):
        if not source_dir.is_dir() or not vba_source_files(source_dir):
            continue

        workbook_stem = source_dir.name
        production_file = None
        for extension in (".xlsm", ".xls", ".xlsb", ".xlam"):
            candidate = excel_path / f"{workbook_stem}{extension}"
            if candidate.is_file():
                production_file = candidate
                break
        if production_file is None:
            continue

        staged_path = stage_path / production_file.name
        metadata_path = staged_path.with_suffix(".release.json")
        if not metadata_path.is_file():
            changed_targets.append((production_file, source_dir))
            continue

        with open(metadata_path, "r", encoding="utf-8") as file:
            metadata = json.load(file)
        if metadata.get("vba_source_sha256") != sha256_vba_source(source_dir):
            changed_targets.append((production_file, source_dir))

    return changed_targets


def stage_changed_workbooks(vba_parent_dir, excel_dir):
    """Stage every workbook whose VBA source changed since its previous stage."""
    targets = find_changed_vba_targets(vba_parent_dir, excel_dir)
    if not targets:
        print("[UNCHANGED] No VBA source changes found since the last staging build.")
        return []

    print(f"[DETECTED] VBA source changes for {len(targets)} workbook(s):")
    staged_paths = []
    for source_excel, source_dir in targets:
        print(f"  - {source_excel.name}")
        staged_path = stage_single_workbook(source_excel, source_dir)
        staged_paths.append(staged_path)
        print(f"[STAGED] Built unverified workbook: {staged_path}")

    return staged_paths


def stage_single_workbook(source_excel, vba_source_dir, staging_dir=STAGING_DIR):
    """Create a fresh staged copy of a production workbook and inject VBA into it."""
    source_path = Path(source_excel).resolve()
    source_dir = Path(vba_source_dir).resolve()
    stage_dir = Path(staging_dir).resolve()
    if source_path.parent != PRODUCTION_DIR.resolve():
        raise ValueError(f"Source workbook must be directly inside {PRODUCTION_DIR}")
    if not source_path.is_file():
        raise FileNotFoundError(f"Production workbook not found: {source_path}")
    if not source_dir.is_dir():
        raise FileNotFoundError(f"Source VBA folder not found: {source_dir}")

    stage_dir.mkdir(parents=True, exist_ok=True)
    staged_path = stage_dir / source_path.name
    evidence_path = staged_path.with_suffix(".release.json")
    test_results = TEST_RUNS_DIR / f"{staged_path.name}_test_results.json"
    source_hash = sha256_file(source_path)
    vba_source_hash = sha256_vba_source(source_dir)
    for stale_evidence in (evidence_path, test_results):
        if stale_evidence.exists():
            stale_evidence.unlink()

    shutil.copy2(source_path, staged_path)
    objects_updated, modules_updated = write_vba_to_workbook(staged_path, source_dir, backup=False)
    if sha256_file(source_path) != source_hash:
        raise RuntimeError("Production workbook changed while it was being staged; rebuild the candidate.")
    if sha256_vba_source(source_dir) != vba_source_hash:
        raise RuntimeError("VBA source changed while the workbook was being staged; rebuild the candidate.")

    metadata = {
        "production_path": str(source_path),
        "production_sha256": source_hash,
        "staged_path": str(staged_path.resolve()),
        "staged_sha256": sha256_file(staged_path),
        "vba_source_path": str(source_dir),
        "vba_source_sha256": vba_source_hash,
        "built_at": datetime.now().isoformat(),
        "objects_updated": objects_updated,
        "modules_updated": modules_updated,
    }
    with open(evidence_path, "w", encoding="utf-8") as file:
        json.dump(metadata, file, indent=2)
    return staged_path


def resolve_production_workbook(argument, production_dir=PRODUCTION_DIR):
    """Resolve a workbook name or path to a workbook in the production directory."""
    requested = Path(argument)
    candidate = requested if requested.is_absolute() else Path.cwd() / requested
    production_path = Path(production_dir).resolve()
    if candidate.is_file():
        candidate = candidate.resolve()
    else:
        candidate = production_path / requested.name
        if not candidate.is_file() and not requested.suffix:
            candidate = next(
                (
                    production_path / f"{requested.name}{extension}"
                    for extension in (".xlsm", ".xls", ".xlsb", ".xlam")
                    if (production_path / f"{requested.name}{extension}").is_file()
                ),
                candidate,
            )
        candidate = candidate.resolve()
    if candidate.parent != production_path:
        raise ValueError(f"Workbook must be directly inside {Path(production_dir).resolve()}")
    if not candidate.is_file():
        raise FileNotFoundError(f"Production workbook not found: {candidate}")
    return candidate


def main():
    import argparse

    parser = argparse.ArgumentParser(
        description="Build staged Excel workbooks from production files and VBA source."
    )
    parser.add_argument(
        "target",
        nargs="?",
        help="Production workbook path/name or VBA source folder; omit to stage all workbooks with VBA changes since the last stage.",
    )
    parser.add_argument(
        "--all",
        action="store_true",
        help="Stage every production workbook with a matching VBA source folder.",
    )
    args = parser.parse_args()

    try:
        production_dir, vba_dir = find_default_directories()
        if args.all:
            if args.target:
                parser.error("TARGET cannot be combined with --all")
            results = write_all_from_folder(vba_dir, production_dir)
            return 0 if results else 1

        if args.target:
            target = Path(args.target).resolve()
            if target.is_file():
                source_excel = resolve_production_workbook(target, production_dir)
                source_dir = vba_dir / source_excel.stem
            elif target.is_dir():
                if (target / "microsoft excel objects").exists() or (target / "modules").exists():
                    source_dir = target
                    source_excel = resolve_production_workbook(target.name, production_dir)
                else:
                    parser.error("Target must be a production workbook or a workbook's VBA source folder")
            else:
                source_excel = resolve_production_workbook(args.target, production_dir)
                source_dir = vba_dir / source_excel.stem
            staged_path = stage_single_workbook(source_excel, source_dir)
            print(f"[STAGED] Built unverified workbook: {staged_path}")
        else:
            staged_paths = stage_changed_workbooks(vba_dir, production_dir)
            if not staged_paths:
                return 0
        return 0
    except (FileNotFoundError, RuntimeError, ValueError) as error:
        print(f"[ERROR] {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
