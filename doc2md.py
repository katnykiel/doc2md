# doc2md.py
# katnykiel
# script to enable bidirectional conversion of docx <-> markdown

# TODO
# fix in-text citation refs
# fix figure/table/equation cross-refs
# fix bold/italicized words
# fix quotes/backticks
# fix latex differences
# fix table formatting
# fix captions (too long?)
# then you're good! :) perfect docx <-> md compatibility with this pandoc workflow

import re
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
    export_md_path = os.path.join(md_dir, "export.md")

    # Remove export, if it already exists
    # if export file already exists, run the next command
    if os.path.exists(export_md_path):
        subprocess.run(["rm", "export.md"])
    else:
        print("export.md does not exist")
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
        "--wrap=preserve",
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
    export_md_path = os.path.join(md_dir, "export.md")

    md_path = md_path.replace(" ", "\\ ")

    # Execute the "Compare Active File With..." command with the two file paths as arguments
    command = f"code --diff {export_md_path} {md_path}"
    subprocess.run(command, shell=True)


def fix_citations(lines, export_lines):
    # # Get all words at any point in all lines that start with @ but not @Fig, @Tbl or @Eq
    # pattern = r"(?<![@\w])@(?!(Fig|Tbl|Eq))[\w-]+"

    # words = []
    # for line in lines:
    #     match = re.search(pattern, line)
    #     while match:
    #         words.append(match.group())
    #         line = line[match.end() :]
    #         match = re.search(pattern, line)

    # # words = sorted(list(set(words)))
    # # Find all hashes

    # pattern = r"\B#\w+[-\w]*"

    # hash_words = []
    # for line in export_lines:
    #     matches = re.findall(pattern, line)
    #     hash_words.extend(matches)

    # for i in range(24):
    #     print(words[i], hash_words[i])

    # # Replace those hases with the corresponding words
    # print(len(hash_words), len(words))
    # if len(hash_words) == len(words):
    #     for i, line in enumerate(export_lines):
    #         for j, hash_word in enumerate(hash_words):
    #             if hash_word in line:
    #                 export_lines[i] = line.replace(hash_word, words[j])
    # else:
    #     print(
    #         "There was an issue swapping refs to citekeys. Do you have extra @ or # symbols?"
    #     )

    # Define the regex patterns for identifying citekeys and cite_refs
    # citekey_pattern = r"(?<![@\w])@(?!(Eq|Tbl|Fig))[\w-]+"
    # cite_ref_pattern = r"\B#\w+[-\w]*"

    # # Find all citekeys in the lines list
    # citekeys = []
    # for line in lines:
    #     match = re.search(citekey_pattern, line)
    #     while match:
    #         citekeys.append(match.group())
    #         line = line[match.end() :]
    #         match = re.search(citekey_pattern, line)

    # # Find all cite_refs in the export_lines list
    # cite_refs = []
    # for line in export_lines:
    #     matches = re.search(cite_ref_pattern, line)
    #     for match in matches:
    #         if match not in cite_refs:
    #             cite_refs.append(match)

    # print(citekeys, len(citekeys))
    # print(cite_refs, len(cite_refs))

    return export_lines


def fix_cross_references(lines, export_lines):
    # Find all instances where a line starts with ![caption]
    for i, line in enumerate(lines):
        if line.startswith("!["):
            # Extract the caption from the line
            caption = line.split("[")[1].split("]")[0]

            # Find the corresponding line in the export file
            for j, export_line in enumerate(export_lines):
                if export_line.startswith("![") and caption in export_line:
                    export_lines[j] = lines[i]

                    # if the next n lines contain the caption, delete those lines
                    if export_line.count(caption) != 2:
                        k = 0
                        while (
                            "".join(export_lines[j : j + k])
                            .replace("\n", "")
                            .count(caption)
                            != 2
                        ):
                            k += 1
                        # remove k-1 lines
                        for l in range(k - 1):
                            del export_lines[j + 1]

    return export_lines


def fix_metadata(lines, export_lines):
    # Copy all lines between "---" and "---" at the start of the file
    metadata = []
    for i, line in enumerate(lines):
        if line.startswith("---"):
            metadata.append(line)
            for j in range(i + 1, len(lines)):
                if lines[j].startswith("---"):
                    metadata.append(lines[j])
                    metadata.append("\n")
                    break
                else:
                    metadata.append(lines[j])
            break

    # replace all lines after ## References in export_lines with all lines after ## References in lines
    for i, line in enumerate(lines):
        references_found = False
        if line.startswith("## References"):
            for j, export_line in enumerate(export_lines):
                if export_line.startswith("## References"):
                    references_found = True
                    export_lines = export_lines[:j] + lines[i:]
                    break
            break
        if not references_found:
            # Starting from the last line, go backward until you find "1. " and remove all of those lines
            for k in range(len(export_lines) - 1, -1, -1):
                if "1. " in export_lines[k]:
                    del export_lines[k - 1 :]
                    break

    export_lines = metadata + export_lines
    return export_lines


def fix_import(md_path):
    """Fix the inconsistencies introduced by pandoc so the workflow can be bidirectional

    Args:
        md_path (str): path to manuscript markdown file
    """

    md_dir = os.path.dirname(md_path)
    export_md_path = os.path.join(md_dir, "export.md")

    # Get lines from each of the files
    with open(md_path, "r") as f:
        lines = f.readlines()

    with open(export_md_path, "r") as f:
        export_lines = f.readlines()

    # Fix pandoc issues
    export_lines = fix_citations(lines, export_lines)
    export_lines = fix_cross_references(lines, export_lines)
    export_lines = fix_metadata(lines, export_lines)

    # Rewrite the export.md file
    with open(export_md_path, "w") as f:
        f.writelines(export_lines)


# test = "test_manuscript"
# test = "High Throughput DTM MXene DFT"
# Run script for testing
md_path = get_md_path()
docx_path = get_docx_path()
# extra_args = get_pandoc_extra_args()
# md_path = f"/Users/kat/Documents/scratch/doc2md/{test}.md"
# docx_path = f"/Users/kat/Documents/scratch/doc2md/{test}.docx"
docx_to_md(md_path, docx_path)
fix_import(md_path)
compare_changes(md_path)
