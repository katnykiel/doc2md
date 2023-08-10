# doc2md.py
# katnykiel
# script to enable bidirectional conversion of docx <-> markdown

import os
import json
import subprocess
import tkinter as tk
from tkinter import filedialog


def get_md_path():
    """Enable a file-picker GUI from within VSCode for md file

    Returns:
        path (str): absolute path to original md file
    """
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename()
    return file_path


def get_docx_path():
    """Enable a file-picker GUI from within VSCode for docx file

    Returns:
        path (str): absolute path to docx file with changes
    """
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename()
    return file_path


def get_pandoc_extra_args(
    settings_path=os.path.expanduser(
        "~/Library/Application Support/Code/User/settings.json"
    ),
):
    """Get pandoc extra args string from the settings within vscode-pandoc

    Args:
        settings_path (str): path to your vscode settings file, default is for mac

    Returns:
        extra_args (str): extra pandoc args
    """

    # Initialize empty extra args
    extra_args = ""

    # Get args from vscode-pandoc extension

    with open(settings_path, "r") as f:
        settings = json.load(f)

    # Get the settings for the Python extension
    extra_args = settings.get("pandoc.docxOptString", {})

    # Print the Python extension settings
    return extra_args


def docx_to_md(md_path, docx_path, extra_args=""):
    """Convert from docx back to md with pandoc, using extra args from the user

    Args:
        md_path (str): path to original md file
        docx_path (str): path to docx file
        extra_args (str): extra pandoc args
    """

    md_dir = os.path.dirname(md_path)
    export_md_path = os.path.join(md_dir, "import.md")

    # Construct the Pandoc command
    command = [
        "pandoc",
        "-f",
        "docx",
        "-t",
        "markdown",
        "-o",
        export_md_path,
        docx_path,
    ]
    if extra_args:
        command += extra_args.split()

    # Run the Pandoc command
    result = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

    # Check if the command succeeded
    if result.returncode != 0:
        print(result.stderr.decode("utf-8"))
    else:
        print("Conversion successful")


def compare_changes(md_path):
    # Define the paths of the two files to compare
    md_dir = os.path.dirname(md_path)
    export_md_path = os.path.join(md_dir, "import.md")

    # Execute the "Compare Active File With..." command with the two file paths as arguments
    command = f"code --diff {export_md_path} {md_path} "
    subprocess.run(command, shell=True)


# Run script
md_path = get_md_path()
docx_path = get_docx_path()
# extra_args = get_pandoc_extra_args()
docx_to_md(md_path, docx_path)
compare_changes(md_path)
