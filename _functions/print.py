import logging
import os
import pikepdf
import subprocess
import tempfile
from pikepdf import _cpphelpers #uncomment in py file when making exe with pyinstaller

"""
This file is a python file that converts files to pdf. This file is called from the main file ontgrendel_tool_gui.py, def convert_files.

Author: Joana Cardoso
"""


def print_to_pdf(out_dir, file_log, progress, style):
    """
    Converts different file types to pdf

    Parameters
    ----------
    out_dir: str
        The output directory where converted files will be placed
    file_log: list 
        The logging table
    progress: Progressbar
        The progress bar that is updated while running this method
    style: Style 
        The progress bar style
    """
    libreoffice_convert = []
    ext = ['.xlsx', '.xls', '.pdf', ".htm", ".html"]
    supported_formats = [
        ".pdf", ".doc", ".docx", ".htm", ".html", ".png", ".jpg", ".gif", ".tiff",
        ".csv", ".uos", ".xls", ".xlsx", ".xml", ".xlt", ".dif", ".dbf", ".slk", ".xlsm",
        ".ppt", ".pptx", ".dotx", ".fodp", ".fods", ".fodt",
        ".odb", ".odf", ".odg", ".odm", ".odp", ".ods", ".otg", ".otp", ".ots", ".ott",
        ".oxt", ".psw", ".sda", ".sdc", ".sdd", ".sdp", ".sdw", ".slk", ".smf", ".stc",
        ".std", ".stw", ".sxc", ".sxg", ".sxi", ".sxm", ".sxw", ".uof", ".uop",
        ".uos", ".uot", ".vsd", ".vsdx", ".wdb", ".wps", ".wri", ".tsv", ".txt"
    ]
    for root, dirs, files in os.walk(os.path.join(out_dir)):
        for name in files:
            name_lower = name.lower()
            file_path = os.path.join(root, name)
            if name_lower.endswith(tuple(supported_formats)):
                libreoffice_convert.append(file_path)
            # update log table
            if name_lower.endswith(".xlsx") or name_lower.endswith(".xls") or name_lower.endswith(
                    ".htm") or name_lower.endswith(".html"):
                try:
                    d = next(
                        item for item in file_log if item['File name'] == file_path)
                    d['Check'] = True
                    logging.debug(f'Updating log table for file: {file_path}')
                except:
                    logging.error(
                        f'Failed to update log table for file: {file_path}')

    for file in libreoffice_convert:
        app = 'soffice.exe'
        appPath = os.path.join("C:\Program Files\LibreOffice\program", app)
        commandLine = [app, "--headless", "--convert-to",
                       "pdf", "--outdir", os.path.dirname(file), file]
        # Print to pdf
        if not file.endswith(".pdf"):
            try:
                subprocess.run(commandLine, executable=appPath, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                               stdin=subprocess.PIPE)
                logging.debug(f'Printing file to PDF: {file}')
            except:
                logging.error(f'Failed to print file to PDF: {file}')

            # Update log table
            try:
                d = next(
                    item for item in file_log if item['File name'] == file)
                d['To PDF'] = True
                logging.debug(f'Updating log table for supported file: {file}')
            except Exception as e:
                logging.error(
                    f'Failed to update log table for supported file: {file}')

        # If already PDF, open with and save with pikepdf to get rid of any write protections
        if file.endswith(".pdf"):
            # Copy to tempdir since input file cannot be overwritten
            temp_dir = tempfile.gettempdir()
            cmd = 'copy "%s" "%s"' % (file, temp_dir)
            try:
                subprocess.call(cmd, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                stdin=subprocess.PIPE)
                pdf_file = os.path.join(
                    temp_dir, os.path.basename(file))  # path to tempfile
                pdf = pikepdf.open(pdf_file)  # Open tempfile
                if 'Metadata' in pdf.Root.keys():  # if PDF metadata is present, delete it
                    try:
                        del pdf.Root.Metadata
                        logging.debug(f'Deleting metadata from file: {file}')
                    except:
                        logging.error(
                            f'Failed to delete metadata from file: {file}')
                pdf.save(file)  # Save processed pdf
                logging.debug(f'Resave PDF file: {file}')
            except:
                logging.error(f'Failed to resave PDF file: {file}')

        # Keep Excel and html files since printing these to a suitable layout often requires manual intervention
        if not file.endswith(tuple(ext)):
            logging.info(f'Original file to remove: {file}')
            # Remove input files but keep Excel and html files for manual conversion/check and PDF files
            try:
                os.remove(file)
                logging.debug(f'Removing original file: {file}')
            except:
                logging.error(f'Failed to remove original file: {file}')

        # set progress bar max value
        progress['value'] += 99 / \
            len(libreoffice_convert) if len(libreoffice_convert) > 0 else 99
        progress.update()
        progress.update_idletasks()
        style.configure(
            'text.Horizontal.TProgressbar',
            text='{:.0f} %'.format(progress['value'])
        )
