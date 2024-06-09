# Shortcut Replacer

This Python script replaces all shortcut files (`.lnk`) in the current working directory with new ones that have the same target.

## Dependencies

This script depends on the following Python packages:

- `os`
- `winshell`
- `win32com.client`

You can install these dependencies by running:

```bash
pip install -r requirements.txt
```

## Usage

To use this script, navigate to the directory containing the shortcut files you want to replace and run the following command:

```bash
python shortcut_replacer.py
```

This will replace all shortcut files in the current directory.

## Note

This script only works on Windows.
You can download the executable file of this script in the release tab.
