import argparse
import pathlib
import pythoncom
import logging
import sys
from datetime import datetime
from win32comext.shell import shell, shellcon
from tqdm import tqdm

def set_logger(log_level, time_str):
    log_dir = pathlib.Path('./log')
    log_dir.mkdir(exist_ok=True)
    log_file = log_dir / f'{time_str}_runlog.log'
    logger = logging.getLogger('file_copy_logger')
    logger.setLevel(logging.DEBUG)
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setLevel(logging.INFO)
    console_handler = logging.StreamHandler()
    console_handler.setLevel(log_level)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    return logger

logger = None  # Will be set in main()

def log_and_raise(message, exception_type=Exception, log_level=logging.ERROR, exc_info=True):
    logger.log(log_level, message, exc_info=exc_info)
    raise exception_type(message)


def remove_prefix(str, prefix):
    if not str.startswith(prefix):
        raise Exception(f"'{str}' should start with '{prefix}")
    return str[len(prefix):]


def get_desktop_shell():
    return shell.SHGetDesktopFolder()


def get_file_full_path(file_shell):
    return file_shell.GetDisplayName(shellcon.SIGDN_DESKTOPABSOLUTEEDITING)


def get_folder_full_path(folder_shell):
    folder_pidl = shell.SHGetIDListFromObject(folder_shell)
    folder_item = shell.SHCreateShellItem(None, None, folder_pidl)
    folder_full_path_str = folder_item.GetDisplayName(shellcon.SIGDN_DESKTOPABSOLUTEEDITING)
    return folder_full_path_str


def get_child_folder_shell(parent_folder_shell, child_folder_name: str):
    for child_pidl in parent_folder_shell:
        folder_name = parent_folder_shell.GetDisplayNameOf(child_pidl, shellcon.SHGDN_NORMAL)
        if folder_name == child_folder_name:
            return parent_folder_shell.BindToObject(child_pidl, None, shell.IID_IShellFolder)
    raise Exception(f"Cannot find {child_folder_name}")


def get_folder_shell_from_str(folder_str):
    current_folder_shell = get_desktop_shell()
    folders = pathlib.Path(folder_str).parts
    for folder in folders:
        try:
            current_folder_shell = get_child_folder_shell(current_folder_shell, folder)
        except BaseException:
            log_and_raise(f"Cannot get shell folder for {folder_str} (at '{folder}')", exc_info=True)
    return current_folder_shell


def get_shell_from_str(path_str):
    try:
        return shell.SHCreateItemFromParsingName(path_str, None, shell.IID_IShellItem)
    except BaseException:
        log_and_raise(f"Cannot get shell for {path_str}", exc_info=True)


def get_files_dict_from_shell(folder_shell):
    result = {}
    folder_count = 0
    file_count = 0
    repeated_file_count = 0
    repeated_files = []
    flags = shellcon.SHCONTF_FOLDERS | shellcon.SHCONTF_NONFOLDERS
    
    for item_pidl in sorted(folder_shell.EnumObjects(0, flags)):
        is_folder = folder_shell.GetAttributesOf([item_pidl], shellcon.SFGAO_FOLDER)
        if is_folder:
            child_folder_shell = folder_shell.BindToObject(item_pidl, None, shell.IID_IShellFolder)
            child_folder_name = folder_shell.GetDisplayNameOf(item_pidl, shellcon.SHGDN_FORADDRESSBAR)
            logger.debug(f"Listing folder '{child_folder_name}'")
            folder_count += 1
            sub_result = get_files_dict_from_shell(child_folder_shell)
            result.update(sub_result)
        else:
            file_directory_pidl = shell.SHGetIDListFromObject(folder_shell)
            file_shell = shell.SHCreateShellItem(file_directory_pidl, None, item_pidl)
            file_full_path_str = get_file_full_path(file_shell)
            file_count += 1
            if file_full_path_str in result:
                repeated_file_count += 1
                repeated_files.append(f"{file_full_path_str}")
            else:
                result[file_full_path_str] = file_shell
    folder_name = get_folder_full_path(folder_shell)
    logger.info(f"{folder_name}: "
                f"folder count: {folder_count}, "
                f"file count: {file_count}, "
                f"repeated file count: {repeated_file_count}, "
                f"repeated files: {repeated_files}")
    return result


def copy_file(source_file_shell, destination_folder_shell, destination_file_name):
    logger.debug(f"Copying '{get_file_full_path(source_file_shell)}' to '{get_file_full_path(destination_folder_shell)}'")
    file_operation_object = pythoncom.CoCreateInstance(shell.CLSID_FileOperation,
                                     None,
                                     pythoncom.CLSCTX_ALL,
                                     shell.IID_IFileOperation)
    file_operation_object.CopyItem(source_file_shell, destination_folder_shell, destination_file_name)
    file_operation_object.PerformOperations()


def read_records(record_folder_str):
    imported_files_set = set()
    record_folder_path = pathlib.Path(record_folder_str)
    if not record_folder_path.exists():
        record_folder_path.mkdir(parents=True, exist_ok=True)
        logger.info(f"Created directory: {record_folder_path}")
    if not record_folder_path.is_dir():
        log_and_raise(f"{record_folder_str} is not a folder")
    else:
        for file_path in record_folder_path.glob("*.txt"):
            logger.debug(f"Loading imported files list from '{file_path}'")
            with open(file_path, "r") as files:
                for line in files:
                    file_name = line.strip()
                    imported_files_set.add(file_name)
    logger.info(f"Loaded {len(imported_files_set)} imported files")
    return imported_files_set


def get_files_dict(source_folder_str, imported_files_set):
    files_dict = {}
    source_folder_shell = get_folder_shell_from_str(source_folder_str)
    all_files_dict = get_files_dict_from_shell(source_folder_shell)
    for path in sorted(all_files_dict.keys()):
        file_shell = all_files_dict[path]
        full_path_str = get_file_full_path(file_shell)
        relative_path_str = remove_prefix(full_path_str, source_folder_str)
        relative_path_str = remove_prefix(relative_path_str, '\\')
        if relative_path_str not in imported_files_set:
            files_dict[relative_path_str] = file_shell
    return files_dict

def import_files(files_dict, destination_folder_str, record_folder_str, time_str):
    destination_directories_dict = {}
    success_import_count = 0
    record_file_path = pathlib.Path(record_folder_str) / f"record_{time_str}.txt"
    with open(record_file_path, "w", encoding='utf-8') as record_file:
        for file_relative_path_str in tqdm(sorted(files_dict.keys()), desc="Copying files"):
            destination_file_full_path = pathlib.Path(destination_folder_str) / file_relative_path_str
            destination_file_directory = destination_file_full_path.parent
            destination_file_name = destination_file_full_path.name
            if str(destination_file_directory) not in destination_directories_dict:
                destination_file_directory.mkdir(parents=True, exist_ok=True)
                destination_folder_shell = get_shell_from_str(str(destination_file_directory))
                destination_directories_dict[str(destination_file_directory)] = destination_folder_shell
            file_shell = files_dict[file_relative_path_str]
            try:
                copy_file(file_shell, destination_folder_shell, destination_file_name)
                success_import_count += 1
                record_file.write(f"{file_relative_path_str}\n")
                record_file.flush()
            except Exception as e:
                logger.error(f"Failed to copy {file_relative_path_str}: {str(e)}", exc_info=True)
    logger.info(f"Record file written: {record_file_path}")
    return success_import_count



def write_record(record_folder_str, files_dict):
    time_str = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    record_folder_path = pathlib.Path(record_folder_str)
    record_file_path = record_folder_path / f"record_{time_str}.txt"
    logger.info(f"Writing '{record_file_path}'")
    with open(record_file_path, "w") as files:
        for file_name in sorted(files_dict.keys()):
            files.write(f"{file_name}\n")

def main(args):
    global logger
    current_time_str = datetime.now().strftime('%Y%m%d_%H%M%S')
    logger = set_logger(logging.DEBUG, current_time_str)
    logger.info(f"Program args: {args.__dict__}")
    source_folder_str = args.source
    destination_folder_str = args.destination
    record_folder_str = args.record_folder
    try:
        imported_files_set = read_records(record_folder_str)
        files_dict = get_files_dict(source_folder_str, imported_files_set)
        total_files = len(files_dict)
        logger.info(f"Found {total_files} files to import")
        if total_files > 0:
            if args.skip_copy:
                logger.info("Skip-copy mode activated")
                success_import_count = 0
            else:
                success_import_count = import_files(files_dict, destination_folder_str, record_folder_str, current_time_str)
                logger.info(f"Summary: {success_import_count} out of {total_files} files successfully copied")
        else:
            logger.info("Nothing to copy")
    except Exception as e:
        logger.error(f"An unexpected error occurred: {str(e)}", exc_info=True)
        sys.exit(1)

if __name__ == "__main__":
    '''
    parser = argparse.ArgumentParser()
    parser.add_argument('source', help="Source directory")
    parser.add_argument('destination', help="Destination directory")
    parser.add_argument('--record_folder', default="./record", help="Folder to store record files")
    parser.add_argument('--skip-copy', action='store_true', help="Skip copying files (for testing)")
    args = parser.parse_args()
    '''
    class Args:
        source = r"此电脑\Apple iPhone\Internal Storage"
        destination = r"D:\iphone14"
        record_folder= r".\record"
        skip_copy = False
    args = Args()
    main(args)
