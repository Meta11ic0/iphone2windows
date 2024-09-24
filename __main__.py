import argparse
import glob
import os
import pathlib
import pythoncom
from datetime import datetime
from win32comext.shell import shell, shellcon


# helper function
def remove_prefix(str, prefix):
    if not str.startswith(prefix):
        raise Exception(f"'{str}' should start with '{prefix}")
    return str[len(prefix):]


def get_desktop_shell():
    return shell.SHGetDesktopFolder()


def get_full_path(file_shell):
    return file_shell.GetDisplayName(shellcon.SIGDN_DESKTOPABSOLUTEEDITING)


# 
def get_child_folder_shell(parent_folder_shell, child_folder_name: str):
    for child_pidl in parent_folder_shell:
        folder_name = parent_folder_shell.GetDisplayNameOf(child_pidl, shellcon.SHGDN_NORMAL)
        if folder_name == child_folder_name:
            return parent_folder_shell.BindToObject(child_pidl, None, shell.IID_IShellFolder)
    raise Exception(f"Cannot find {child_folder_name}")


# safetly, only for folder
def get_folder_shell_from_str(folder_str):
    current_folder_shell = get_desktop_shell()
    folders = folder_str.split("\\")
    for folder in folders:
        try:
            current_folder_shell = get_child_folder_shell(current_folder_shell, folder)
        except BaseException as exception:
            raise Exception(f"Cannot get shell folder for {folder_str} (at '{folder}')") from exception
    return current_folder_shell


# quickly, without check
def get_shell_from_str(path_str):
    try:
        return shell.SHCreateItemFromParsingName(path_str, None, shell.IID_IShellItem)
    except BaseException as exception:
        raise Exception(f"Cannot get shell for {path_str}") from exception



# 
def get_files_dict_from_shell(folder_shell):
    result = {}
    for folder_pidl in folder_shell.EnumObjects(0, shellcon.SHCONTF_FOLDERS):
        child_folder_shell = folder_shell.BindToObject(folder_pidl, None, shell.IID_IShellFolder)
        child_folder_name = folder_shell.GetDisplayNameOf(folder_pidl, shellcon.SHGDN_FORADDRESSBAR)
        print(f"Listing folder '{child_folder_name}'")
        result |= get_files_dict_from_shell(child_folder_shell)
    for file_pidl in folder_shell.EnumObjects(0, shellcon.SHCONTF_NONFOLDERS):
        file_directory_pidl = shell.SHGetIDListFromObject(folder_shell)
        file_shell = shell.SHCreateShellItem(file_directory_pidl, None, file_pidl)
        file_full_path_str = get_full_path(file_shell)
        result[file_full_path_str] = file_shell
    return result


def copy_file(source_file_shell, destination_folder_shell, desination_file_name):
    print(f"Copying '{get_full_path(source_file_shell)}' to '{get_full_path(destination_folder_shell)}'")

    file_operation_object = pythoncom.CoCreateInstance(shell.CLSID_FileOperation,
                                     None,
                                     pythoncom.CLSCTX_ALL,
                                     shell.IID_IFileOperation)
    file_operation_object.CopyItem(source_file_shell, destination_folder_shell, desination_file_name)
    file_operation_object.PerformOperations()


# get imported files from record 
def read_records(record_folder):
    imported_files_set = set()
    if record_folder is not None:
        if not os.path.exists(record_folder):
            raise Exception(f"{record_folder} does not exist")
        if not os.path.isdir(record_folder):
            raise Exception(f"{record_folder} is not a folder")
        else:
            for file_name in glob.glob(os.path.join(record_folder, "*.txt")):
                print(f"Loading imported files list from '{file_name}'")
                with open(file_name, "r") as files:
                    for line in files:
                        file_name = line.strip()
                        imported_files_set.add(file_name)
    print(f"Loaded {len(imported_files_set)} imported files")
    return imported_files_set


# get all file from folder
# return a dict key is a file relative path, e.g. "202301_a\IMG_1694.jpg"
# value is windows shell of this file "IMG_1694.jpg" 
def get_files_dict(source_folder_str, imported_files_set):
    files_dict = {}
    source_folder_shell = get_folder_shell_from_str(source_folder_str)
    all_files_dict = get_files_dict_from_shell(source_folder_shell)
    for path in sorted(all_files_dict.keys()):
        file_shell = all_files_dict[path]
        full_path_str = get_full_path(file_shell)
        relative_path_str = remove_prefix(full_path_str, source_folder_str)
        relative_path_str = remove_prefix(relative_path_str, '\\')
        if relative_path_str not in imported_files_set:
            files_dict[relative_path_str] = file_shell
    return files_dict


# using windows shell to import
def import_files(files_dict, destination_folder_str):
    destination_directorys_dict = {}
    for file_relative_path_str in sorted(files_dict.keys()):
        desination_file_full_path_str = os.path.join(destination_folder_str, file_relative_path_str)
        desination_file_directory_str = os.path.dirname(desination_file_full_path_str)
        desination_file_name = os.path.basename(desination_file_full_path_str)
        # if desination file directory not exits, build directory and mark
        if desination_file_directory_str not in destination_directorys_dict:
            pathlib.Path(desination_file_directory_str).mkdir(parents=True, exist_ok=True)
            destination_folder_shell = get_shell_from_str(desination_file_directory_str)
            destination_directorys_dict[desination_file_directory_str] = destination_folder_shell
        file_shell = files_dict[file_relative_path_str]
        copy_file(file_shell, destination_folder_shell, desination_file_name)
 

# record imported files in this process  
def write_record(record_folder, files_dict):
    time_str = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    recored_file = os.path.join(record_folder, f"record_{time_str}.txt")
    print(f"Writing '{recored_file}'")
    with open(recored_file, "w") as files:
        for file_name in sorted(files_dict.keys()):
            files.write(f"{file_name}\n")


def main(args):
    print(f"Program args: {args.__dict__}")
    imported_files_set = read_records(args.record_folder)
    files_dict = get_files_dict(args.source, imported_files_set)
    print(f"Import {len(files_dict)} files")
    if args.skip_copy:
        print(f"skip-copy mode, skipping copying")
    elif len(files_dict) > 0:
        import_files(files_dict, args.destination)
        write_record(args.record_folder, files_dict)
    else:
        print(f"Nothing to copy")
        

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument('source')
    parser.add_argument('destination')
    parser.add_argument('--metadata-folder', required=False)
    parser.add_argument('--skip-copy', required=False, action='store_true')
    main(parser.parse_args())
