# About pdf2docxConverter

Converts pdf files to docx using python's pdf2docx module.  
This script has been created to experiment with the pdf2docx module.


# Acknowledgements
Based on an idea and first version found here https://www.facebook.com/groups/4910929078932537


# Required Python modules
  - [pdf2docx](https://pypi.org/project/pdf2docx/)
  - argparse



# How to execute
``Usage: pdf2docxConverter.py [-p pattern="\.pdf$"] [-s frompagenumber=1] [-e topagenumber=None] [-o outputdir="./"] [-P password=None] [-N] [-G] [source directory="./"]``

If no arguments are given, pdf2docxConverter searches for files with names matching pattern \.(?i)pdf$ in the current working directory. Supported arguments:

``-p pattern`` : regular expression. The pattern the pdf files names that are should be converted to docx must match . Defaults to \.pdf$ .

``-s frompagenumber`` : integer. Page number to start converting pdf from. Will be applied to all files. frompagenumber starts counting from 1. Defaults to 1. If frompagenumber exceeds the number of pages in pdf, nothing is converted.

``-e topagenumber`` : integer. Page number to stop converting pdf. Last page included. Will be applied to all files. topagenumber starts counting from 1. Defaults to None which means up until last page in file.

``-o outputdir`` : string. Path to directory the .docx files should be saved. Defaults to current working directory. docx files will have the same name as source pdf files but with different extension.

``-P password`` : string. password for opening pdf files. The same password is used for all pdf files. Defaults to None.

``source directory``: string. Path to directory containing the pdf files. Defaults to current working directory (./)

``-N`` : Do not delete .docx files if these already exists. By default, if a docx file already exists with the same name, these are deleted before conversion.

``-G`` : Enable debug mode. Not yet implemented/working.

