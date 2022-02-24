<p align="center">
  <img src="image/PZH_Basislogo.svg" width="50%">
</p>

## Introduction
This repository is a graphical user interface (GUI) that converts multiple types of files to pdf. With this tool it is possible to convert multiple pdf's at once. The tool may be run via the .py file or via de executable. The tool is able to:

- convert e-mails (.msg)
- extract and convert e-mail attachments
- combine e-mails and attachments into one pdf file
- convert individual Microsoft Office files
- remove duplicate files
- remove protection on pdf files
- adjust the file name in case of a too long path or a double name
- work directly from a zipped file

## Requirements
Two applications are needed:
- Microsoft Office
- [LibreOffice (opensource officepackage)](https://nl.libreoffice.org/)

The executable ('PDF_conversion_tool.exe') installs all the needed python dependencies to run the tool. If the tool is run from the .py file ('PDF_conversion_tool.py') there are extra files needed (also available in this repository):

- file '_functions/check_length.py'
- file '_functions/combine.py'
- file '_functions/duplicates.py'
- file '_functions/extract_msg.py'
- file '_functions/get_files.py'
- file '_functions/log_table.py'
- file '_functions/print.py'
- file '_functions/unzip_files.py'
- 'logo.ico' (file with a logo)
- [OfficeToPDF.exe](https://github.com/cognidox/OfficeToPDF) 
- packages (see file 'requirements/requirements_gui.txt'):
  - [pikepdf](https://github.com/pikepdf/pikepdf)
  - [pywin32](https://github.com/mhammond/pywin32)
  - [openpyxl](https://openpyxl.readthedocs.io/en/stable/)
  - [pillow](https://github.com/python-pillow/Pillow)

See the user manual for more information on the installation and use of the tool.

## Using the tool
### Run the tool using the python script
- Make a new virtual environment and activate de new environment
- Install the dependencies: 'pip install -r requirements_gui.txt'
- Download the files 'PDF_conversion_tool.py' and 'logo.ico' and the folder '_functions'
- Download OfficeToPDF.exe into the same location as the above files
- Run the file 'PDF_conversion_tool.py' to convert one or more pdf files (see 'PDF-conversion-tool.pdf')

### Run the tool using the executable
- Download the file 'PDF_conversion_tool.exe'
- Read the instructions manual ('PDF-conversion-tool.pdf') to learn how to install and run the tool

### Error detection
Every time the tool is used, information on what happens in every step of the tool is saved in a file (‘Logging_PDF_conversion_tool.txt’), in the same folder als the tool. If errors occur, this file can be used to check which step went wrong.

## Author
Joana Cardoso

## Contact
For questions or suggestions please contact: vdwh@pzh.nl
