import datetime
import os

"""
This file is a python file that makes a log table to follow processes and detect errors. This file is called from files _functions.extract_msg and _functions.get_files.

Author: Joana Cardoso
"""


def log_to_table(log_table, filename, file, sourcefile, attachment, attachment_name, review, pdf_parsed,
                 combined, duplicated):
    """
    Write action of the conversion tool pipeline to logging table for quality and audit purposes

    Parameters
    ----------
        log_table: list 
            An existing or empty list of dictionaries to append to
        filename: str
            Filename of file being processed
        file: str
            Full source path of the file being processed, file can be equal to filename or refer to the source email if the file is an attachment
        sourcefile: str
            Full source path of the file in the original folder, is attachment path refers to the main email
        attachment: bool
            Flag for filtering purposes whether file is an email attachment
        attachment_name: str 
            If attachment=True, original filename of the attachment
        review: bool 
            Flag whether the file requires manual inspection
        pdf_parsed: bool 
            Flag whether the file is printed to pdf
        combined: bool 
            Flag whether the file is joined with the parent email as single PDF
        duplicated: bool 
            Flag whether the exact same file already exists in the data seen so far
    """
    log_table.append({"File name": filename,
                      "File": file,
                      "Original file": sourcefile,
                      "Attachment": attachment,
                      "Attachment name": attachment_name,
                      "Check": review,
                      "To PDF": pdf_parsed,
                      "Combined": combined,
                      "Duplicate": duplicated,
                      "User": os.environ.get('USERNAME'),
                      "Date/Time": str(datetime.datetime.now())})
    return log_table
