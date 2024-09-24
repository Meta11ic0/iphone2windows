
def copy_single_file(source_file_shell, destination_folder_shell_item, desination_file_name):
    print(
        f"Copying '{get_file_full_path(source_file_shell)}' to '{get_file_full_path(destination_folder_shell_item)}'")

    pfo = pythoncom.CoCreateInstance(shell.CLSID_FileOperation,
                                     None,
                                     pythoncom.CLSCTX_ALL,
                                     shell.IID_IFileOperation)
    pfo.CopyParams(source_file_shell, destination_folder_shell_item, desination_file_name)
    pfo.PerformOperations()


