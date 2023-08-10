# doc2md

## What is it?
*doc2md* is a VSCode plugin to enable bidirectional conversion of docx <-> markdown. 

![doc2md](doc2md.png)
Specifically, it's intended to import changes from an exported docx file into your original markdown file.

## Why?
*doc2md* enables you to write your manuscripts with all the benefits of markdown, and collaborate with others using MS Word.

Pandoc is a wonderful tool, best at moving from simple markup to complicated documents. However, it cannot *cleanly* go back the other way. There are certain pandoc features (citations, cross-references, metadata) which do not convert back from docx->md. This is intended to bridge that gap.

Ok, but why not just use MS Word in the first place? If you want the easy solution, do that. This is only a practical solution if you are **committed** to writing manuscripts in markdown, as I am.  
- markdown is a *minimal* approach to writing
- markdown is *readable in its source code*
- markdown *doesn't rely on proprietary software*

For academic writing, markdown makes it easy to use a journal's LaTeX or MS Word templates to easily export your manuscript to the desired journal's format.

## Who?
I am chiefly making this plugin for myself and my workflow, which I will do my best to document for reproducability. For minor improvements, bug fixes, or suggestions, create an issue in GitHub and I may fix it. For major changes, consider a PR and code it yourself :P I am not a software developer.


## How?
I don't really know. I've never created a VSCode plugin before, but this seems like it may be useful. 

Intended use case:
- manuscript in markdown
- use pandoc to convert to docx with vscode-pandoc
- collaborators make changes in docx file
- accept/reject changes in word
- use *doc2md* to add changes back to markdown file

Code outline:
- get path to docx file
- get pandoc extra args from vscode-pandoc settings
- convert from docx to md with pandoc
- debug: use VSCode compare to select desired changes
- identify the meaningful changes
    - ignore citation changes
    - ignore metadata
    - ignore figure and table refs
- copy the original markdown file and make changes on copy
- save copy as new md file






