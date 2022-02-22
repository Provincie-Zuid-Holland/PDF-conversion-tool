import logging
import os
import subprocess
from pathlib import Path
import _functions.extract_msg as msg
import _functions.log_table as log
import _functions.unzip_files as uz

"""
This file is a python file that copies individual files to the output map and extracts emails and attachments. This file is called from the main file PDF_conversion_tool.py, def convert_files.

Author: Joana Cardoso
"""


def get_all_files(files, process_dir, out_dir, progress, style):
    """
    Copies individual files to the output map and if the file is an email it calls the method extract_msg.

    Parameters
    ----------
    files: str
        The files found in the origin(process) directory
    process_dir: str
        The directory of the selected folder
    out_dir: str 
        The output directory where converted files will be placed
    progress: Progressbar
        The progress bar for this method
    style: Style
        The style of the progress bar
    
    Returns
    -------
    msg_list: list
        List with all the emails and respective attachments
    file_log: list 
        The logging table
    """
    msg_list = []
    file_list_msg = []
    file_log = []
    len_files = len(files)

    for file in files:
        file_out = file.replace(process_dir, out_dir)
        logging.info(f'file_out: {file_out}')
        root_dir = os.path.dirname(file_out)
        try:
            Path(root_dir).mkdir(parents=True,
                                 exist_ok=True)  # Create output dir
            logging.debug(f'Creating output directory: {root_dir}')
        except:
            logging.error(f'Failed to create output directory: {root_dir}')

        # add file to log
        try:
            file_log = log.log_to_table(
                log_table=file_log,
                filename=file_out,
                file=file_out,
                sourcefile=file,
                attachment=False,  # may be updated in later step to True
                attachment_name="N/A",  # may be updated in later step to True
                review=False,  # may be updated in later step to True
                pdf_parsed=False,  # may be updated in later step to True
                combined=False,  # may be updated in later step to True
                duplicated=False,  # may be updated in later step to True
            )
            logging.debug(f'Updating log table for file: {file}')
        except:
            logging.error(f'Failed to update log table for file: {file}')

        # in case of emails extract email and attachments
        if file.endswith(".msg"):
            msg_list.append(file)
            main_email = file
            out_email = os.path.splitext(file_out)[0]

            # Extract attachment using win32com
            attachment = msg.extract_msg(file=file, root_dir=root_dir, main_email=main_email,
                                         file_log=file_log, process_dir=process_dir, out_dir=out_dir)
            file_list_msg.append(attachment)
            # unzip extracted zip attachments
            for root, dirs, files in os.walk(out_email):
                for name in files:
                    name_lower = name.lower()
                    if name_lower.endswith('.zip'):
                        logging.info('zip attachment')
                        zip_dir = os.path.join(root, name)
                        attch_process_dir = os.path.abspath(
                            os.path.splitext(zip_dir)[0])
                        uz.unzip_files(zip_dir=zip_dir,
                                       proc_dir=attch_process_dir)

                        try:
                            os.remove(zip_dir)
                            logging.debug(f'Removed zip: {zip_dir}')
                        except:
                            logging.debug(f'Failed to remove zip: {zip_dir}')

                        for root, dirs, files in os.walk(attch_process_dir):
                            for name in files:
                                attachment_name = attachment + '/' + name
                                filename = os.path.join(root, name)
                                try:
                                    file_log = log.log_to_table(
                                        log_table=file_log,
                                        filename=filename,
                                        file=filename,
                                        sourcefile=main_email,
                                        attachment=True,
                                        attachment_name=attachment_name,
                                        review=True,
                                        pdf_parsed=False,  # may be updated in later step to True
                                        combined=True,
                                        duplicated=False,  # may be updated in later step to True
                                    )
                                    logging.debug(
                                        f'Updating log table for file: {file}')
                                except:
                                    logging.error(
                                        f'Failed to update log table for file: {file}')

        else:
            try:
                cmd = 'copy "%s" "%s"' % (file, root_dir)
                subprocess.call(cmd, shell=True, stdout=subprocess.PIPE,
                                stderr=subprocess.PIPE, stdin=subprocess.PIPE)
                logging.debug(f'Copying file to output directory: {file}')
            except:
                logging.error(
                    f'Failed to copy file to output directory: {file}')

        # set progress bar max value
        progress['value'] += 99 / len_files if len_files > 0 else 99
        progress.update()
        progress.update_idletasks()
        style.configure(
            'text.Horizontal.TProgressbar',
            text='{:.0f} %'.format(progress['value'])
        )

    subprocess.call(["taskkill", "/f", "/im", "outlook.exe"], stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE, stdin=subprocess.PIPE, shell=True)

    # clean up extracted msg files (delete original)
    outlook_toremove = []
    for path in Path(out_dir).rglob('*.msg'):
        outlook_toremove.append(path)
    for msg_remove in outlook_toremove:
        try:
            os.remove(msg_remove)
            logging.debug(f'Removing original email: {msg_remove}')
        except:
            logging.error(f'Failed to remove original email: {msg_remove}')

    return msg_list, file_log
