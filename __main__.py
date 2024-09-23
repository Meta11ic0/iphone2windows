import argparse
import glob
import os
import pathlib
from datetime import datetime

import win32utils
from win32utils import CopyParams

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
def get_files_dict(folder_str):
    source_folder_shell = win32utils.get_folder_shell_from_str(folder_str)
    files_dict = win32utils.walk_dcim(source_folder_shell)
    return files_dict


def remove_prefix(str, prefix):
    if not str.startswith(prefix):
        raise Exception(f"'{str}' should start with '{prefix}")
    return str[len(prefix):]


# based on imported files which get from record
# classify files into 'imported' and 'not imported' categories.
def classify_files(source_folder_str, files_dict, imported_files_set):
    not_imported_files_set = set()
    files_to_import_dict = {}
    for path in sorted(files_dict.keys()):
        file_shell = files_dict[path]
        file_full_path_str = win32utils.get_file_full_path(file_shell)
        file_relative_path_str = remove_prefix(file_full_path_str, source_folder_str)
        file_relative_path_str = remove_prefix(file_relative_path_str, '\\')
        if file_relative_path_str not in imported_files_set:
            files_to_import_dict[file_relative_path_str] = file_shell
        else:
            not_imported_files_set.add(file_relative_path_str)
    return  not_imported_files_set, files_to_import_dict


# using windows shell to import
def import_files(files_dict, destination_folder_str):
    target_folder_shell_item_by_path = {}
    copy_params_list = []
    for desination_file_relative_path_str in sorted(files_dict.keys()):
        desination_file_full_path_str = os.path.join(destination_folder_str, desination_file_relative_path_str)
        desination_folder_str = os.path.dirname(desination_file_full_path_str)
        desination_filesname = os.path.basename(desination_file_full_path_str)
        if desination_folder_str not in target_folder_shell_item_by_path:
            pathlib.Path(desination_folder_str).mkdir(parents=True, exist_ok=True)
            target_folder_shell_item = win32utils.get_shell_item_from_path(desination_folder_str)
            target_folder_shell_item_by_path[desination_folder_str] = target_folder_shell_item
        file_shell = files_dict[desination_file_relative_path_str]
        copy_params = CopyParams(file_shell, target_folder_shell_item_by_path[desination_folder_str],
                                 desination_filesname)
        copy_params_list.append(copy_params)
    win32utils.copy_multiple_files(copy_params_list)


def write_record(record_folder, files_path_set):
    time_str = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    imported_files_metadata_path = os.path.join(record_folder, f"imported_{time_str}.txt")
    print(f"Writing '{imported_files_metadata_path}'")
    with open(imported_files_metadata_path, "w") as files:
        for file_name in sorted(list(files_path_set)):
            files.write(f"{file_name}\n")


def main(args):
    print(f"Program args: {args.__dict__}")

    source_folder_str = args.source
    destination_folder_str = args.destination

    imported_files_set = read_records(args.record_folder)

    all_files_dict = get_files_dict(source_folder_str)

    not_import_files_set, files_to_import_dict = classify_files(
        source_folder_str, all_files_dict, imported_files_set)

    print(f"Import {len(files_to_import_dict)} files")

    if args.skip_copy:
        print(f"skip-copy mode, skipping copying")
    elif len(files_to_import_dict) > 0:
        import_files(files_to_import_dict, destination_folder_str)
        write_record(args.record_folder, files_to_import_dict)
    else:
        print(f"Nothing to copy")
        

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument('source')
    parser.add_argument('destination')
    parser.add_argument('--metadata-folder', required=False)
    parser.add_argument('--skip-copy', required=False, action='store_true')
    main(parser.parse_args())
