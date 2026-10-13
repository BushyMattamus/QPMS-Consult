import sys
from pathlib import Path

# Ensure UTF-8 output encoding for Windows terminal compatibility
if sys.platform == 'win32' and hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass


def is_excel_object(vba_code, stream_path, vba_filename):
    """
    Determine if a VBA component is a Microsoft Excel Object (Worksheet / Workbook).
    
    Checks for:
      - 'VB_Customizable = True' attribute (standard MS-OVBA marker for document modules)
      - Worksheet / Workbook CLSIDs (00020820 / 00020819)
      - Standard naming patterns (ThisWorkbook, Sheet*)
    """
    if 'VB_Customizable = True' in vba_code:
        return True
    if '00020820-0000-0000-C000-000000000046' in vba_code or '00020819-0000-0000-C000-000000000046' in vba_code:
        return True
    sp = stream_path.lower()
    fn = vba_filename.lower()
    if 'thisworkbook' in sp or 'thisworkbook' in fn:
        return True
    if 'sheet' in sp or fn.startswith('sheet'):
        return True
    return False


def get_vba_parser():
    try:
        from oletools.olevba import VBA_Parser
        return VBA_Parser
    except ImportError as exc:
        raise RuntimeError(
            "Missing required library 'oletools'. Install project dependencies with: "
            "python -m pip install -r requirements.txt"
        ) from exc


def extract_macros_oletools(file_path, output_dir=None, verbose=False):
    """
    Extract macro code from VBA projects using oletools.VBA_Parser.
    
    Args:
        file_path (str or Path): Path to the Excel file (.xlsm, .xls, .xlsb, etc.).
        output_dir (str or Path, optional): Directory to save extracted files.
                                            If provided, separates macros into:
                                              - 'microsoft excel objects/'
                                              - 'modules/'
        verbose (bool): If True, print full VBA source code to console.
        
    Returns:
        list[dict]: Extracted macro items with metadata and code.
    """
    VBA_Parser = get_vba_parser()

    target_path = Path(file_path).resolve()
    
    print('=' * 80)
    print('VBA MACRO EXTRACTION (OLETOOLS)')
    print('=' * 80)
    print(f'Source File: {target_path}')
    print('=' * 80)
    
    if not target_path.exists():
        raise FileNotFoundError(f"Target file does not exist: {target_path}")
    
    results = []
    
    try:
        vbaparser = VBA_Parser(str(target_path))
        
        if vbaparser.detect_vba_macros():
            # extract_macros() yields (filename, stream_path, vba_filename, vba_code)
            for macro in vbaparser.extract_macros():
                filename, stream_path, vba_filename, vba_code = macro[:4]
                
                # Clean up null characters and normalize line endings to prevent double blank lines (\r\r\n)
                if vba_code:
                    clean_code = '\r\n'.join(vba_code.replace('\x00', '').splitlines())
                else:
                    clean_code = ""
                
                is_obj = is_excel_object(clean_code, stream_path, vba_filename)
                category = "microsoft excel objects" if is_obj else "modules"
                
                results.append({
                    "filename": filename,
                    "stream_path": stream_path,
                    "vba_filename": vba_filename,
                    "vba_code": clean_code,
                    "category": category,
                    "is_excel_object": is_obj,
                })
            
            print(f'Found {len(results)} VBA component(s) in {target_path.name}:\n')
            for r in results:
                cat_label = "Excel Object" if r["is_excel_object"] else "Module"
                print(f"  [{cat_label:<12}] {r['vba_filename']} ({r['stream_path']})")
                if verbose:
                    print('=' * 80)
                    print(r["vba_code"])
                    print('=' * 80)
                    
            # Save extracted macros to organized directories
            if output_dir:
                out_path = Path(output_dir).resolve()
                excel_objects_dir = out_path / "microsoft excel objects"
                modules_dir = out_path / "modules"
                
                excel_objects_dir.mkdir(parents=True, exist_ok=True)
                modules_dir.mkdir(parents=True, exist_ok=True)
                
                print('\n' + '-' * 80)
                print(f"SAVING TO: {out_path}")
                print(f"  -> Microsoft Excel Objects: {excel_objects_dir}")
                print(f"  -> Modules:                 {modules_dir}")
                print('-' * 80)
                
                obj_count = 0
                mod_count = 0
                
                for r in results:
                    mod_name = r['vba_filename'] or r['stream_path'].split('/')[-1]
                    if r["is_excel_object"]:
                        if not mod_name.endswith(('.cls', '.bas', '.frm', '.vba', '.txt')):
                            mod_name += '.cls'
                        file_dest = excel_objects_dir / mod_name
                        obj_count += 1
                    else:
                        if not mod_name.endswith(('.bas', '.cls', '.frm', '.vba', '.txt')):
                            mod_name += '.bas'
                        file_dest = modules_dir / mod_name
                        mod_count += 1
                    
                    with open(file_dest, 'w', encoding='utf-8', errors='replace', newline='') as f:
                        f.write(r['vba_code'])
                        
                    print(f"  Saved [{r['category']}]: {mod_name}")
                    
                print(f"\nSummary for {target_path.name}: {obj_count} excel object(s), {mod_count} module(s) saved.")
                print()
        else:
            print("No VBA macros found in the file.")
            
        vbaparser.close()
        
    except Exception as e:
        print(f'\n[ERROR] Failed to extract macros: {e}')
        import traceback
        traceback.print_exc()
        raise RuntimeError(f"Failed to extract macros from {target_path}") from e
        
    return results


def extract_all_from_folder(excel_dir, vba_parent_dir):
    """
    Processes all Excel files in excel_dir, extracting VBA macros into
    source-file-named folders inside vba_parent_dir.
    
    Args:
        excel_dir (Path or str): Directory containing source Excel files.
        vba_parent_dir (Path or str): Parent directory to store extracted VBA folders.
    """
    excel_path = Path(excel_dir).resolve()
    vba_path = Path(vba_parent_dir).resolve()
    
    print('=' * 80)
    print('BATCH VBA EXTRACTION PROCESS')
    print('=' * 80)
    print(f"Source Folder: {excel_path}")
    print(f"Output Folder: {vba_path}")
    print('=' * 80)
    
    if not excel_path.exists() or not excel_path.is_dir():
        print(f"[ERROR] Excel files folder does not exist: {excel_path}")
        raise FileNotFoundError(f"Excel files folder does not exist: {excel_path}")

    get_vba_parser()
    
    # Supported Excel macro extensions
    extensions = ('*.xlsm', '*.xls', '*.xlsb', '*.xlam')
    excel_files = []
    for ext in extensions:
        excel_files.extend(excel_path.glob(ext))
        
    # Exclude temporary Excel lock files (e.g. ~$Filename.xlsm)
    excel_files = sorted([f for f in excel_files if not f.name.startswith('~$')])
    
    if not excel_files:
        print(f"No Excel files found in {excel_path}")
        return {}
    
    print(f"Found {len(excel_files)} Excel file(s) to process:\n")
    for f in excel_files:
        print(f"  - {f.name}")
    print()
    
    summary = {}
    failures = []
    for excel_file in excel_files:
        # Create folder reflecting the original source file's name
        source_folder = vba_path / excel_file.stem
        try:
            macros = extract_macros_oletools(excel_file, output_dir=source_folder)
            summary[excel_file.name] = len(macros)
        except Exception as exc:
            failures.append((excel_file.name, exc))
        
    print('=' * 80)
    print('EXTRACTION SUMMARY')
    print('=' * 80)
    for fname, count in summary.items():
        print(f"  {fname}: {count} macro item(s) extracted")
    for fname, error in failures:
        print(f"  [FAILED] {fname}: {error}")
    print('=' * 80)

    if failures:
        raise RuntimeError(f"Failed to extract {len(failures)} of {len(excel_files)} workbook(s).")

    return summary


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


if __name__ == "__main__":
    try:
        default_excel_dir, default_vba_dir = find_default_directories()

        # Allow command line overrides if specified
        if len(sys.argv) > 1:
            target = Path(sys.argv[1]).resolve()
            if target.is_file():
                # Process single file into VBAFiles / <Source File Name>
                dest_folder = default_vba_dir / target.stem
                extract_macros_oletools(target, output_dir=dest_folder)
                sys.exit(0)
            elif target.is_dir():
                default_excel_dir = target
            else:
                raise FileNotFoundError(f"Input path does not exist: {target}")

        extract_all_from_folder(default_excel_dir, default_vba_dir)
    except (FileNotFoundError, RuntimeError, ValueError) as exc:
        print(f"[ERROR] {exc}", file=sys.stderr)
        sys.exit(1)