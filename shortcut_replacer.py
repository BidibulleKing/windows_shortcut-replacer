import os
import winshell
import win32com.client


def main():
    # Get the current working directory
    current_directory = os.getcwd()

    # Get the list of files in the current working directory
    files = os.listdir(current_directory)

    # Get the list of shortcut files in the current working directory
    shortcut_files = [file for file in files if file.endswith(".lnk")]

    # Replace the shortcut files
    for shortcut_file in shortcut_files:
        replace_shortcut_file(shortcut_file)

    print("Shortcut files replaced successfully.")


def replace_shortcut_file(file_path):
    # Get the target of the shortcut
    shortcut = winshell.shortcut(file_path)
    target_path = shortcut.path

    # Check if target_path is still None
    if target_path is None:
        print(f"Could not find target for shortcut: {file_path}")
        return

    # Remove the original shortcut file
    os.remove(file_path)

    # Create a new shortcut file of the target
    shell = win32com.client.Dispatch("WScript.Shell")
    shortcut = shell.CreateShortCut(file_path)
    shortcut.Targetpath = target_path
    shortcut.WindowStyle = 3 # 7 - Minimized, 3 - Maximized, 1 - Normal
    shortcut.save()


if __name__ == "__main__":
    main()
