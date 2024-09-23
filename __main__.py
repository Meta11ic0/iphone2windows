import argparse
import glob
import os
import pathlib
import re
from datetime import datetime

import win32utils
from win32utils import CopyParams


# Loads paths of already imported files into a set
def load_already_imported_files_names(metadata_folder):
    already_imported_files_set = set()
    if metadata_folder is not None:
        if not os.path.exists(metadata_folder):
            raise Exception(f"{metadata_folder} does not exist")
        if not os.path.isdir(metadata_folder):
            raise Exception(f"{metadata_folder} is not a folder")
        else:
            for filesname in glob.glob(os.path.join(metadata_folder, "*.txt")):
                print(f"Loading imported files list from '{filesname}'")
                with open(filesname, "r") as files:
                    for line in files:
                        filesname = line.strip()
                        already_imported_files_set.add(filesname)
    print(f"Loaded {len(already_imported_files_set)} imported files")
    return already_imported_files_set


def classify_files(source_folder_str, files_dict, already_imported_files_set):
    imported_files_set = set()
    not_imported_files_set = set()
    shell_items_to_copy = {}
    for path in sorted(files_dict.keys()):
        source_files_shell_item = files_dict[path]
        source_files_absolute_path = win32utils.get_absolute_name(source_files_shell_item)
        files_relative_path = remove_prefix(source_files_absolute_path, source_folder_str)
        files_relative_path = remove_prefix(files_relative_path, '\\')
        if files_relative_path not in already_imported_files_set:
            shell_items_to_copy[files_relative_path] = source_files_shell_item
            imported_files_set.add(files_relative_path)
        else:
            not_imported_files_set.add(files_relative_path)
    return imported_files_set, not_imported_files_set, shell_items_to_copy

def remove_prefix(str, prefix):
    if not str.startswith(prefix):
        raise Exception(f"'{str}' should start with '{prefix}")
    return str[len(prefix):]


def copy_using_windows_shell(files_dict, destination_base_path_str):
    target_folder_shell_item_by_path = {}
    copy_params_list = []
    for destination_files_path in sorted(files_dict.keys()):
        desination_full_path = os.path.join(destination_base_path_str, destination_files_path)
        desination_folder = os.path.dirname(desination_full_path)
        desination_filesname = os.path.basename(desination_full_path)
        if desination_folder not in target_folder_shell_item_by_path:
            pathlib.Path(desination_folder).mkdir(parents=True, exist_ok=True)
            target_folder_shell_item = win32utils.get_shell_item_from_path(desination_folder)
            target_folder_shell_item_by_path[desination_folder] = target_folder_shell_item
        source_files_shell_item = files_dict[destination_files_path]
        copy_params = CopyParams(source_files_shell_item, target_folder_shell_item_by_path[desination_folder],
                                 desination_filesname)
        copy_params_list.append(copy_params)
    win32utils.copy_multiple_files(copy_params_list)


def write_imported_files_list_to_metadata_folder(metadata_folder, files_path_set):
    time_str = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    imported_files_metadata_path = os.path.join(metadata_folder, f"imported_{time_str}.txt")
    print(f"Writing '{imported_files_metadata_path}'")
    with open(imported_files_metadata_path, "w") as files:
        for filesname in sorted(list(files_path_set)):
            files.write(f"{filesname}\n")


def main(args):
    print(f"Program args: {args.__dict__}")

    source_folder_str = args.source
    destination_folder_str = args.destination

    already_imported_files_set = load_already_imported_files_names(args.metadata_folder)

    source_folder_shell = win32utils.get_shell_folder_from_absolute_display_name(
        source_folder_str)

    files_dict = win32utils.walk_dcim(source_folder_shell)

    imported_files_set, not_import_files_set, files_to_import_dict = classify_files(
        source_folder_str, files_dict, already_imported_files_set)

    print(f"Import {len(imported_files_set)} files")

    if args.skip_copy:
        print(f"skip-copy mode, skipping copying")
    elif len(files_to_import_dict) > 0:
        copy_using_windows_shell(files_to_import_dict, destination_folder_str)
    else:
        print(f"Nothing to copy")

    if len(imported_files_set) > 0:
        write_imported_files_list_to_metadata_folder(args.metadata_folder, imported_files_set)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument('source')
    parser.add_argument('destination')
    parser.add_argument('--metadata-folder', required=False)
    parser.add_argument('--skip-copy', required=False, action='store_true')
    main(parser.parse_args())
