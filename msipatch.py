import os
import shutil
import subprocess
import argparse
import uuid

class MSIPathResolver:
    def __init__(self, msi_arch="x86"):
        self.msi_arch = msi_arch.lower()
        self.FOLDER_MAP = {
            "windows": ("WindowsFolder", "TARGETDIR", "Windows"),
            "system32": (lambda arch:
                ("System64Folder", "WindowsFolder", "System64") if arch == "x64"
                else ("SystemFolder", "WindowsFolder", "System32")),
            "systemfolder": (lambda arch:
                ("System64Folder", "WindowsFolder", "System64") if arch == "x64"
                else ("SystemFolder", "WindowsFolder", "System32")),
            "syswow64": (lambda arch:
                ("SystemFolder", "WindowsFolder", "SysWOW64") if arch == "x64"
                else ("System64Folder", "WindowsFolder", "SysWOW64")),
            "fonts": ("FontsFolder", "WindowsFolder", "Fonts"),

            "program files": (lambda arch:
                ("ProgramFiles64Folder", "TARGETDIR", "ProgramFiles") if arch == "x64"
                else ("ProgramFilesFolder", "TARGETDIR", "ProgramFiles")),
            "program files (x86)": ("ProgramFilesFolder", "TARGETDIR", "ProgramFiles (x86)"),

            "common files": (lambda arch:
                ("CommonFiles64Folder", "ProgramFiles64Folder", "Common Files") if arch == "x64"
                else ("CommonFilesFolder", "ProgramFilesFolder", "Common Files")),

            "programdata": ("CommonAppDataFolder", "TARGETDIR", "CommonAppData"),

            "users": ("ProfilesFolder", "TARGETDIR", "Users"),
            
            # User personal folders base
            "personalfolder": ("PersonalFolder", "TARGETDIR", "Documents"),
            "appdata": ("AppDataFolder", "TARGETDIR", "AppData"),
            "localappdata": ("LocalAppDataFolder", "TARGETDIR", "LocalAppData"),
            "desktop": ("DesktopFolder", "TARGETDIR", "Desktop"),
            "documents": ("PersonalFolder", "TARGETDIR", "Documents"),
            "downloads": ("DownloadsFolder", "PersonalFolder", "Downloads"),
            "start menu": ("StartMenuFolder", "TARGETDIR", "Start Menu"),
            "programs": ("ProgramsFolder", "StartMenuFolder", "Programs"),
            "startup": ("StartupFolder", "TARGETDIR", "Startup"),
            "sendto": ("SendToFolder", "TARGETDIR", "SendTo"),
            "templates": ("TemplatesFolder", "TARGETDIR", "Templates"),
            "favorites": ("FavoritesFolder", "TARGETDIR", "Favorites"),
            "recent": ("RecentFolder", "TARGETDIR", "Recent"),
            "nethood": ("NetHoodFolder", "TARGETDIR", "NetHood"),
            "printhood": ("PrintHoodFolder", "TARGETDIR", "PrintHood"),
            "common desktop": ("CommonDesktopFolder", "TARGETDIR", "Desktop"),
            "common programs": ("CommonProgramsFolder", "TARGETDIR", "Programs"),
            "common start menu": ("CommonStartMenuFolder", "TARGETDIR", "Start Menu"),
            "common startup": ("CommonStartupFolder", "TARGETDIR", "Startup"),
            "admin tools": ("AdminToolsFolder", "TARGETDIR", "AdminTools"),
            "common admin tools": ("CommonAdminToolsFolder", "TARGETDIR", "AdminTools"),
            "internet": ("InternetFolder", "TARGETDIR", "Internet"),
            "pictures": ("MyPicturesFolder", "PersonalFolder", "My Pictures"),
            "music": ("MyMusicFolder", "PersonalFolder", "My Music"),
            "videos": ("MyVideoFolder", "PersonalFolder", "My Videos"),
            "temp": ("TempFolder", "TARGETDIR", "Temp"),

            # Folder IDs (optional aliases to themselves or arch-based)
            "system64folder": ("System64Folder", "WindowsFolder", "System64"),
            "programfilesfolder": ("ProgramFilesFolder", "TARGETDIR", "ProgramFiles"),
            "programfiles64folder": ("ProgramFiles64Folder", "TARGETDIR", "ProgramFiles"),
            "commonfilesfolder": ("CommonFilesFolder", "ProgramFilesFolder", "Common Files"),
            "commonfiles64folder": ("CommonFiles64Folder", "ProgramFiles64Folder", "Common Files"),
            "windowsfolder": ("WindowsFolder", "TARGETDIR", "Windows"),
        }
    

    def list_available_directories(self):
        # Return all folder keys sorted nicely for display
        keys = sorted(self.FOLDER_MAP.keys())
        print("Available destination directories:")
        for k in keys:
            print(f"  - {k}")


    def get_required_directory_entries(self, input_path):
        # Normalize path, split by backslash
        parts = input_path.replace("/", "\\").split("\\")
        if not parts:
            return []

        entries = []

        def add_folder(folder_key):
            folder_key_lower = folder_key.lower()
            entry = self.FOLDER_MAP.get(folder_key_lower)
            if entry is None:
                return None

            # If entry is a lambda, evaluate with arch
            if callable(entry):
                folder_id, parent, default_dir = entry(self.msi_arch)
            else:
                folder_id, parent, default_dir = entry

            # Add parent folder first if needed
            if parent and parent != "TARGETDIR":
                # Avoid infinite recursion on self-parenting
                if parent.lower() != folder_key_lower:
                    add_folder(parent)

            # Add folder if not already present
            if not any(e[0] == folder_id for e in entries):
                entries.append((folder_id, parent, default_dir))

            return folder_id

        parent_id = add_folder(parts[0])
        for part in parts[1:]:
            dir_id = self._make_directory_id(part, parent_id)
            entries.append((dir_id, parent_id, part))
            parent_id = dir_id

        return entries

    def _make_directory_id(self, name, parent_id):
        safe_name = ''.join(c for c in name.title() if c.isalnum())
        return f"{parent_id}_{safe_name}" if parent_id else safe_name


class InstallerDatabaseTables:

    def __init__(self, arch, temp_dir):
        self.temp_dir = temp_dir
        self.feature_idt = os.path.join(self.temp_dir, "Feature.idt")
        self.directory_idt = os.path.join(temp_dir, "Directory.idt")
        self.file_idt = os.path.join(temp_dir, "File.idt")
        self.comp_idt = os.path.join(temp_dir, "Component.idt")
        self.media_idt = os.path.join(temp_dir, "Media.idt")
        self.featcomp_idt = os.path.join(temp_dir, "FeatureComponents.idt")
        self.customact_idt = os.path.join(temp_dir, "CustomAction.idt")
        self.instexecseq_idt = os.path.join(temp_dir, "InstallExecuteSequence.idt")
        self.binary_idt = os.path.join(temp_dir, "Binary.idt")
        self.msi_arch = arch


    def get_last_sequence_from_media_idt(self):
        with open(self.media_idt, "r") as f:
            lines = f.readlines()

        data_line = lines[3].strip()
        columns = data_line.split("\t")
        last_sequence = columns[1]
        return int(last_sequence)
    

    def get_top_level_feature(self):
        with open(self.feature_idt, "r") as f:
            lines = f.readlines()
            for line in lines:
                if line.strip() == "" or line.startswith("s") or line.startswith("Feature"):
                    continue
                cols = line.strip().split("\t")
                if len(cols) >= 2 and cols[1].strip() == "":
                    print(f"[+] Top-level feature detected: {cols[0]}")
                    return cols[0]
        raise RuntimeError("No top-level feature found in Feature.idt.")
    

    def get_next_install_sequence_number(self):
        sequence_numbers = set()
        step = 1
        after_action = "InstallFiles"
        after_action_seq = None

        with open(self.instexecseq_idt, 'r', encoding='utf-8') as f:
            lines = f.readlines()

        data_lines = lines[3:] if lines[0].startswith("s72") else lines

        for line in data_lines:
            parts = line.strip().split('\t')
            if len(parts) < 3:
                continue
            action, condition, seq = parts
            try:
                seq_num = int(seq)
            except ValueError:
                continue
            sequence_numbers.add(seq_num)
            if action == after_action:
                after_action_seq = seq_num

        if after_action_seq is None:
            raise ValueError(f"Action '{after_action}' not found in {self.instexecseq_idt}")

        candidate = after_action_seq + step
        while candidate in sequence_numbers:
            candidate += step

        print(f"[+] Install sequence number detected: {candidate}")
        return candidate


    def modify_file_idt(self, file_cab_name, component_name, file_dest_name, filesize, file_sequence_number):
        with open(self.file_idt, "a") as f:
            f.write(f"{file_cab_name}\t{component_name}\t{file_dest_name}\t{filesize}\t\t\t0\t{file_sequence_number}\n")
        print("[i] Modified File.idt")


    def modify_component_idt(self, component_name, guid, file_dest_dir, file_cab_name):
        with open(self.comp_idt, "a") as f:
            f.write(f"{component_name}\t{{{guid}}}\t{file_dest_dir}\t256\t\t{file_cab_name}\n")
        print("[i] Modified Component.idt")


    def modify_feature_components_idt(self, top_feature, component_name):
        with open(self.featcomp_idt, "a") as f:
            f.write(f"{top_feature}\t{component_name}\n")
        print("[i] Modified FeatureComponents.idt")


    def modify_media_idt(self, file_sequence_number):
        with open(self.media_idt, "r") as f:
            lines = f.readlines()
            
        header = lines[:3]
        data = lines[3:]
        parts = data[0].rstrip("\n").split("\t")
        parts[1] = str(file_sequence_number)
        data[0] = "\t".join(parts) + "\n"

        with open(self.media_idt, "w") as f:
            f.writelines(header + data)
        print("[i] Modified Media.idt")


    def modify_directory_idt(self, destination_path):
        file_dest_dir_key = None
        with open(self.directory_idt, "r") as f:
            lines = f.readlines()

        existing_dirs = {line.split("\t")[0] for line in lines if "\t" in line}
        needed_entries = MSIPathResolver(self.msi_arch).get_required_directory_entries(destination_path)

        insert_index = len(lines)
        for i, line in enumerate(lines):
            if line.strip() == "Directory\tDirectory_Parent\tDefaultDir":
                insert_index = i + 3
                break

        for entry in needed_entries:
            if entry[0] not in existing_dirs:
                print(f"[i] Adding missing {entry[0]} entry to Directory.idt")
                lines.insert(insert_index, f"{entry[0]}\t{entry[1]}\t{entry[2]}\n")
                insert_index += 1
                file_dest_dir_key = entry[0]

        with open(self.directory_idt, "w") as f:
            f.writelines(lines)
        return file_dest_dir_key
    

    def modify_binary_idt(self, file_path):
        file_name = os.path.basename(file_path).split(".")[0]
        dest_path = os.path.join('Binary', f"Binary.{file_name}")

        shutil.copy(file_path, dest_path)
        with open(self.binary_idt, "a") as f:
            f.write(f"{file_name}\tBinary.{file_name}\n")
        print("[i] Modified Binary.idt")
        return file_name


    def modify_install_execute_sequence_idt(self, action_name):
        install_condition = ""
        install_sequence_number = self.get_next_install_sequence_number()
        with open(self.instexecseq_idt, "a") as f:
            f.write(f"{action_name}\t{install_condition}\t{install_sequence_number}\n")
        print("[i] Modified InstallExecuteSequence.idt")
    

    def modify_custom_action_idt(self, action_name, action_type, action_source, action_target):
        if not os.path.exists(self.customact_idt):
            with open(self.customact_idt, "w") as f:
                f.write("Action\tType\tSource\tTarget\n")
                f.write("s72\ti2\ts72\tL0\n")
                f.write("CustomAction\tAction\n")

        with open(self.customact_idt, "a") as f:
            f.write(f"{action_name}\t{action_type}\t{action_source}\t{action_target}\n")
        print("[i] Modified CustomAction.idt")


    def file_dropper(self, component_name, file_cab_name, file_dest_name, file_dest_dir, file_path):
        print("[+] Modifying IDT files...")
        filesize = os.path.getsize(file_path)
        file_sequence_number = self.get_last_sequence_from_media_idt() + 1
        guid = str(uuid.uuid4()).upper()
        top_feature = self.get_top_level_feature()

        file_dest_dir_key = self.modify_directory_idt(file_dest_dir)
        self.modify_file_idt(file_cab_name, component_name, file_dest_name, filesize, file_sequence_number)
        self.modify_component_idt(component_name, guid, file_dest_dir_key, file_cab_name)
        self.modify_feature_components_idt(top_feature, component_name)
        self.modify_media_idt(file_sequence_number)
        
        return ["Directory.idt", "Component.idt", "File.idt", "FeatureComponents.idt", "Media.idt"]


    def run_custom_exe_action(self, file_path, action_name, is_async, exe_args):
        print("[+] Modifying IDT files...")
        action_source = self.modify_binary_idt(file_path)
        self.modify_install_execute_sequence_idt(action_name)

        action_type = 2
        if is_async:
            action_type += 192
        
        action_target = exe_args
        
        self.modify_custom_action_idt(action_name, action_type, action_source, action_target)
        return ["Binary.idt", "CustomAction.idt", "InstallExecuteSequence.idt"]
 

    def run_custom_dll_action(self, file_path, action_name, is_async, exported_function):
        print("[+] Modifying IDT files...")
        action_source = self.modify_binary_idt(file_path)
        self.modify_install_execute_sequence_idt(action_name)

        action_type = 1
        if is_async:
            action_type += 64
        
        action_target = exported_function

        self.modify_custom_action_idt(action_name, action_type, action_source, action_target)
        return ["Binary.idt", "CustomAction.idt", "InstallExecuteSequence.idt"]


    def run_custom_preinstalled_exe_action(self, action_name, command):
        action_type = 226
        action_source = "INSTALLDIR"
        print(command)
        action_target = command
        
        self.modify_install_execute_sequence_idt(action_name)
        self.modify_custom_action_idt(action_name, action_type, action_source, action_target)
        return ["CustomAction.idt", "InstallExecuteSequence.idt"]


class CabinetFile:
    def __init__(self, cab_path):
        self.cab_path = cab_path
        self.extract_dir = os.path.splitext(cab_path)[0]


    @staticmethod
    def find_cab_file(streams_dir):
        for entry in os.listdir(streams_dir):
            if entry.lower().endswith(".cab"):
                return os.path.join(streams_dir, entry)
        return None


    def extract(self):
        os.makedirs(self.extract_dir, exist_ok=True)
        cmd = ["cabextract", "-d", self.extract_dir, self.cab_path]
        print(f"[+] Extracting CAB: {self.cab_path}")
        result = subprocess.run(cmd, text=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        if result.returncode != 0:
            raise RuntimeError(f"cabextract failed:\n{result.stderr}")
        print("[+] CAB extracted to:", self.extract_dir)


    def copy_file_to_extracted_dir(self, file_path, file_cab_name):
        if not self.extract_dir:
            raise ValueError("CAB must be extracted before copying files.")
        
        dest_path = os.path.join(self.extract_dir, file_cab_name)
        shutil.copy(file_path, dest_path)
        print(f"[+] File copied to: {dest_path}")


    def get_file_name_list_in_order(self):
        result = subprocess.run(
            ["cabextract", "-l", self.cab_path],
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )

        if result.returncode != 0:
            raise RuntimeError(f"cabextract failed:\n{result.stderr}")

        file_list = []
        for line in result.stdout.splitlines():
            line = line.strip()
            if not line or line.startswith("Extracting from") or "files" in line or "archive" in line:
                continue
            parts = line.split()
            if len(parts) >= 5:
                file_list.append(parts[-1])
        return file_list


    def rebuild_cab_from_dir(self, additional_file_name):
        if not self.extract_dir:
            raise ValueError("CAB must be extracted before rebuilding.")
        
        original_cwd = os.getcwd()
        os.chdir(self.extract_dir)
        try:
            file_list = self.get_file_name_list_in_order()
            file_list.append(additional_file_name)
            if not file_list:
                raise RuntimeError("No files to pack in CAB.")

            cmd = ["gcab", "-c", "-z", "-n", self.cab_path] + file_list
            print(f"[+] Rebuilding CAB")
            result = subprocess.run(cmd, text=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            if result.returncode != 0:
                raise RuntimeError(f"gcab failed:\n{result.stderr}")
            print(f"[+] CAB rebuilt: {self.cab_path}")
        finally:
            os.chdir(original_cwd)


class MSITool:
    def __init__(self, original_msi_file_path):
        self.temp_dir = self._create_temp_folder()
        self.binary_dir = "Binary"
        self.original_msi_file_path = original_msi_file_path
        self.patched_msi_file_path = os.path.join(os.getcwd(), "patched.msi")


    def __del__(self):
        try:
            if os.path.isdir(self.temp_dir):
                shutil.rmtree(self.temp_dir)
                print(f"[+] Temporary directory '{self.temp_dir}' deleted.")
                shutil.rmtree(self.binary_dir)
                print(f"[+] Binary directory '{self.binary_dir}' deleted.")
        except Exception as e:
            print(f"[!] Failed to delete temp directory: {e}")


    def _create_temp_folder(self):
        temp_dir = "msipatch_temp"
        os.path.join(os.getcwd(), temp_dir)
        os.makedirs(temp_dir, exist_ok=True)
        return temp_dir


    def _run_command(self, cmd, error_message):
        print("[i] Running:", " ".join(cmd))
        result = subprocess.run(cmd, text=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        if result.returncode != 0:
            print(f"[-] {error_message}")
            print(result.stderr)
            raise RuntimeError(error_message)
        return result


    def dump(self):
        cmd = ["msidump", "-t", "-s", "-d", self.temp_dir, self.original_msi_file_path]
        self._run_command(cmd, "msidump failed")
        print("[+] MSI dumped successfully.")


    def rebuild_msi_from_idts(self, idt_list_in_order):
        print("[+] Rebuilding MSI from IDTs...")
        shutil.copy(self.original_msi_file_path, self.patched_msi_file_path)
        for idt in idt_list_in_order:
            idt_path = os.path.join(self.temp_dir, idt)
            cmd = ["msibuild", self.patched_msi_file_path, "-i", idt_path]
            self._run_command(cmd, f"Failed to insert {idt}")



    def embed_cab_file(self, cab_path):
        cab_name = os.path.basename(cab_path)
        cmd = ["msibuild", self.patched_msi_file_path, "-a", cab_name, cab_path]
        self._run_command(cmd, f"Failed to embed CAB file {cab_name}")
        print(f"[+] Embedded CAB file '{cab_name}' into MSI.")


def inject_file_into_msi(arch, msi_file_path, inject_file_path, file_cab_name, file_dest_dir, file_dest_name, component_name):
    msitool = MSITool(msi_file_path)
    msitool.dump()

    cab_file_path = CabinetFile.find_cab_file(os.path.join(os.getcwd(), msitool.temp_dir, "_Streams"))
    cabinet_file = CabinetFile(cab_file_path)
    cabinet_file.extract()
    cabinet_file.copy_file_to_extracted_dir(inject_file_path, file_cab_name)
    cabinet_file.rebuild_cab_from_dir(file_cab_name)

    idt = InstallerDatabaseTables(arch, msitool.temp_dir)
    idt_list_in_order = idt.file_dropper(component_name, file_cab_name, file_dest_name, file_dest_dir, inject_file_path)

    msitool.rebuild_msi_from_idts(idt_list_in_order)
    msitool.embed_cab_file(cab_file_path)


def inject_custom_action_into_msi(args):
    msitool = MSITool(args.msi)
    msitool.dump()

    idt = InstallerDatabaseTables(args.msi, msitool.temp_dir)

    if args.action_type == 'exe':
        idt_list_in_order = idt.run_custom_exe_action(args.action_file, args.action_name, args.is_async, args.action_args)

    if args.action_type == 'dll':
        idt_list_in_order = idt.run_custom_dll_action(args.action_file, args.action_name, args.is_async, args.action_export)

    if args.action_type == 'cmd':
        idt_list_in_order = idt.run_custom_preinstalled_exe_action(args.action_name, args.action_args)

    msitool.rebuild_msi_from_idts(idt_list_in_order)


def check_required_tools_is_installed():
    # Map each tool to the package that provides it
    required_tools = {
        "msidump": "msitools",
        "msibuild": "msitools",
        "cabextract": "cabextract",
        "gcab": "gcab"
    }

    missing_packages = {}
    for tool, package in required_tools.items():
        if shutil.which(tool) is None:
            missing_packages[package] = True

    if missing_packages:
        print("The following required tools are missing:\n")
        for package in missing_packages:
            print(f" - Missing tools from package: {package}")
        print("\nYou can install them with the following command:\n")
        print(f"sudo apt update && sudo apt install {' '.join(missing_packages.keys())}")
        return False
    
    return True


def parse_args():
    parser = argparse.ArgumentParser(
        description="MSIPatch: Inject files or custom actions into an existing MSI file.",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter
    )

    parser.add_argument("-l", "--list", action="store_true", help="List available directory keys and exit")
    parser.add_argument("-m", "--msi", metavar="MSI", help="Path to the original MSI file")

    # File dropper
    parser.add_argument("-i", "--file", metavar="FILE", help="File to inject into the MSI")
    parser.add_argument("-d", "--dest", metavar="DIR", default="system32", help="Destination folder key (e.g., system32, desktop)")
    parser.add_argument("-n", "--name", metavar="NAME", help="Target filename when dropped")
    parser.add_argument("-c", "--cab", metavar="CAB", help="Filename inside the CAB archive")
    parser.add_argument("-C", "--comp", metavar="COMP", default="MyComponent", help="MSI Component name")

    # Custom action config
    parser.add_argument("--add-action", action="store_true", help="Add a custom action")
    parser.add_argument("--action-name", metavar="ACTION", help="Custom action name")
    parser.add_argument("--action-type", choices=["exe", "dll", "cmd"], help="Type of custom action")
    parser.add_argument("--action-file", metavar="FILE", help="EXE or DLL file to embed (not required for cmd)")
    parser.add_argument("--action-export", metavar="FUNC", help="Exported DLL function (only for DLL actions)")
    parser.add_argument("--action-args", metavar="ARGS", default="", help="Arguments for EXE (or cmd string for CMD)")
    parser.add_argument("--async", dest="is_async", action="store_true", help="Run the custom action asynchronously")

    parser.add_argument("-a", "--arch", choices=["x86", "x64"], default="x86", metavar="ARCH", help="Target architecture")

    args = parser.parse_args()

    # Basic validation
    if not args.list and not args.msi:
        parser.error("--msi (-m) is required unless --list (-l) is used.")

    # File injection validation
    if args.file:
        missing = []
        if not args.dest:
            missing.append("--dest (-d)")
        if not args.name:
            missing.append("--name (-n)")
        if not args.cab:
            missing.append("--cab (-c)")
        if not args.comp:
            missing.append("--comp (-C)")
        if missing:
            parser.error(f"The following arguments are required with --file: {', '.join(missing)}")

    # Custom action validation
    if args.add_action:
        if not args.action_name:
            parser.error("--action-name is required with --add-action")
        if not args.action_type:
            parser.error("--action-type is required with --add-action")

        if args.action_type in ["exe", "dll"] and not args.action_file:
            parser.error("--action-file is required for EXE or DLL actions")

        if args.action_type == "dll" and not args.action_export:
            parser.error("--action-export is required for DLL actions")

    return args

def main():
    args = parse_args()

    if not check_required_tools_is_installed():
        return

    if args.list:
        MSIPathResolver(args.arch).list_available_directories()
        return

    if args.file:
        inject_file_into_msi(
            arch=args.arch,
            msi_file_path=args.msi,
            inject_file_path=args.file,
            file_cab_name=args.cab,
            file_dest_dir=args.dest,
            file_dest_name=args.name,
            component_name=args.comp
        )
        return
    
    if args.add_action:
        inject_custom_action_into_msi(args)
        return


if __name__ == "__main__":
    main()
