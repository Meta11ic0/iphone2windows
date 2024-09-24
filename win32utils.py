from dataclasses import dataclass

import pythoncom
from win32comext.shell import shell, shellcon


@dataclass
class CopyParams:
    source_file_shell: object
    destination_folder_shell: object
    desination_file_name: str


def get_desktop_shell_folder():
    return shell.SHGetDesktopFolder()


# returns the child shell folder of a shell folder with a given name
def get_child_shell_folder_with_display_name(parent_shell_folder, child_folder_name: str):
    for child_pidl in parent_shell_folder:
        child_display_name = parent_shell_folder.GetDisplayNameOf(child_pidl, shellcon.SHGDN_NORMAL)
        if child_display_name == child_folder_name:
            return parent_shell_folder.BindToObject(child_pidl, None, shell.IID_IShellFolder)
    raise Exception(f"Cannot find {child_folder_name}")


# returns a shell folder for a string path, e.g. "This PC\Apple iPhone\Internal Storage\DCIM"
def get_folder_shell_from_str(display_names):
    current_shell_folder = get_desktop_shell_folder()
    folders = display_names.split("\\")
    for folder in folders:
        try:
            current_shell_folder = get_child_shell_folder_with_display_name(current_shell_folder, folder)
        except BaseException as exception:
            raise Exception(f"Cannot get shell folder for {display_names} (at '{folder}')") from exception
    return current_shell_folder


# returns a shell item for a string path, e.g. "This PC\Apple iPhone\Internal Storage\DCIM\IMG_0091.JPG"
def get_shell_item_from_path(path):
    try:
        return shell.SHCreateItemFromParsingName(path, None, shell.IID_IShellItem)
    except BaseException as exception:
        raise Exception(f"Cannot get shell item for {path}") from exception


# returns a dictionary of (file name -> shell item) of files in shell folder
def get_files_dict_from_shell(shell_folder):
    result = {}

    for folder_pidl in shell_folder.EnumObjects(0, shellcon.SHCONTF_FOLDERS):
        child_shell_folder = shell_folder.BindToObject(folder_pidl, None, shell.IID_IShellFolder)
        name = shell_folder.GetDisplayNameOf(folder_pidl, shellcon.SHGDN_FORADDRESSBAR)
        print(f"Listing folder '{name}'")
        result |= get_files_dict_from_shell(child_shell_folder)

    for file_pidl in shell_folder.EnumObjects(0, shellcon.SHCONTF_NONFOLDERS):
        sourcefolder_pidl = shell.SHGetIDListFromObject(shell_folder)
        source_file_shell = shell.SHCreateShellItem(sourcefolder_pidl, None, file_pidl)
        sourcefile_name = get_file_full_path(source_file_shell)
        result[sourcefile_name] = source_file_shell

    return result


def copy_single_file(source_file_shell, destination_folder_shell_item, desination_file_name):
    print(
        f"Copying '{get_file_full_path(source_file_shell)}' to '{get_file_full_path(destination_folder_shell_item)}'")

    pfo = pythoncom.CoCreateInstance(shell.CLSID_FileOperation,
                                     None,
                                     pythoncom.CLSCTX_ALL,
                                     shell.IID_IFileOperation)
    pfo.CopyParams(source_file_shell, destination_folder_shell_item, desination_file_name)
    pfo.PerformOperations()


def copy_multiple_files(copy_params_list: list[CopyParams]):
    fileOperationObject = pythoncom.CoCreateInstance(shell.CLSID_FileOperation,
                                                     None,
                                                     pythoncom.CLSCTX_ALL,
                                                     shell.IID_IFileOperation)
    for copy_params in copy_params_list:
        src_str = get_file_full_path(copy_params.source_file_shell)
        dst_str = get_file_full_path(copy_params.destination_folder_shell)
        print(f"Queuing copying '{src_str}' to '{dst_str}'")
        fileOperationObject.CopyItem(copy_params.source_file_shell, copy_params.destination_folder_shell,
                                     copy_params.desination_file_name)
    print(f"Running copy operations...")
    fileOperationObject.PerformOperations()


def get_file_full_path(shell_item):
    return shell_item.GetDisplayName(shellcon.SIGDN_DESKTOPABSOLUTEEDITING)


def get_diplay_name(shell_item):
    return shell_item.GetDisplayName(shellcon.SIGDN_NORMALDISPLAY)
