import logging
import os
# import subprocess
import win32com.client
from pathlib import Path
import _functions.log_table as log
import csv

"""
This file is a python file that processes emails. This file is called from the main file PDF_conversion_tool.py, def convert_files.

Author: Joana Cardoso
"""


def extract_msg(file, root_dir, main_email, file_log, out_dir, process_dir, nested=False, index=0):
    """
    Converts email bodies to pdf and extracts attachments for (nested) outlook MSG files.

    Parameters
    ----------
    file: str
        The path to the MSG file to process
    root_dir: str 
        Directory to write the output of the specific MSG file
    main_email: str 
        The path to the main email
    file_log: list 
        The log table to append to
    out_dir: str 
        The output directory where converted files will be placed
    process_dir: str 
        The path to the initial directory to convert
    nested: bool 
        Checks whether the MSG file is the root email or a nested MSG file as attachment. Default is False
    index: int 
        The number given to each attachment so that they are combined as 1 pdf in a specific order. Default is 0

    Returns
    -------
    attachment: str
        The attachment name
    """
    if not nested:
        logging.info('Not nested email')
        new_dir = os.path.join(
            root_dir, (os.path.basename(file).rsplit('.', 1)[0])).replace('\\','/')
        logging.info(f'New_dir: {new_dir}')
    else:
        logging.info("Nested email")
        new_dir = root_dir.replace('\\','/')
        logging.info(f'New_dir: {new_dir}')

    try:
        Path(new_dir).mkdir(parents=True, exist_ok=True)
        logging.debug(f'Creating new directory: {new_dir}')
    except:
        logging.error(f'Failed to create new directory: {new_dir}')

    if not nested:
        file_out = file.replace(process_dir, out_dir)
    if nested:
        file_out = file

    # fileName, fileExtension = os.path.splitext(file)
    # logging.info(f'File: {file}')
    # new_name = fileName.replace(fileName, str(index)) + fileExtension
    # out_pdf = os.path.join(new_dir, new_name).replace(".msg", ".pdf")
    # logging.info(f'PDF: {out_pdf}')

    # try:
    #     subprocess.run(['OfficeToPDF.exe', file, out_pdf], shell=True, stdout=subprocess.PIPE,
    #                    stderr=subprocess.PIPE,
    #                    stdin=subprocess.PIPE)
    #     logging.debug(f'Converting file to PDF: {file}')
    # except:
    #     logging.error(f'Failed to convert file to PDF: {file}')
    # index += 1

    # # update log table
    # logging.info(f'Updating log table')
    # if not nested:
    #     file_out = file.replace(process_dir, out_dir)
    # if nested:
    #     file_out = file

    # try:
    #     d = next(item for item in file_log if item['File name'] == file_out)
    #     d['To PDF'] = True
    #     logging.debug(f'Updating log table for msg: {file_out}')
    # except:
    #     logging.error(f'Failed to update log table for msg: {file_out}')

    outlook = win32com.client.Dispatch(
        "Outlook.Application").GetNamespace("MAPI")
    msg = outlook.OpenSharedItem(os.path.abspath(file))
    count_attachments = msg.Attachments.Count

    fileName, fileExtension = os.path.splitext(file)
    logging.info(f'File: {file}')
    new_name = fileName.replace(fileName, str(index)) + fileExtension
    out_txt = os.path.join(new_dir, new_name).replace(".msg", ".txt")
    logging.info(f'TXT: {out_txt}')

    try:
        with open(out_txt, mode='w', newline='', encoding="utf-8") as csvfile:
            fieldnames = ['text']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames, quoting=csv.QUOTE_MINIMAL)
            # writer.writeheader()
            msg_sender = msg.SenderName
            msg_date = msg.SentOn
            msg_receiver = msg.To
            msg_cc = msg.CC
            msg_subj = msg.Subject
            msg_message = msg.Body

            writer.writerow({'text': 'Van:          ' + msg_sender})
            writer.writerow({'text': 'Verzonden:    ' + str(msg_date)})
            writer.writerow({'text': 'Aan:          ' + msg_receiver})
            writer.writerow({'text': 'CC:           ' + msg_cc})
            writer.writerow({'text': 'Onderwerp:    ' + msg_subj})
            writer.writerow({'text': '\n' + msg_message})
            logging.debug(f'Converting file to TXT: {file}')
    except:
            logging.error(f'Failed to convert file to TXT: {file}')
            d = next(
            item for item in file_log if item['Bestand naam'] == file_out)
            d['Check'] = True
    # except Exception as e:
    #         logging.error(f'e: {e}')
            
    #         print(f'Failed to convert file to TXT: {file}')
    #         d = next(
    #             item for item in file_log if item['Bestand naam'] == file_out)
    #         d['Check'] = True
    
    index += 1

    # If there are attachments, extract those
    if count_attachments > 0:
        try:
            d = next(
                item for item in file_log if item['File name'] == file_out)
            d['Combined'] = True
            logging.debug(f'Updating log table for msg: {file_out}')
        except:
            logging.error(f'Failed to update log table for msg: {file_out}')

        logging.info(f'File {file} has {count_attachments} attachments')
        for item in range(1, count_attachments + 1):
            try:
                att = msg.Attachments.Item(item)
                attachment = att.filename
                logging.info(f'Attachment: {att.filename}')
                fileName, fileExtension = os.path.splitext(att.filename)
                att_name = fileName.replace(
                    fileName, str(index)) + fileExtension
                att_out = os.path.join(os.path.abspath(new_dir), att_name)
                logging.info(f'Output directory attachment: {att_out}')
            except Exception as e:
                logging.error(f'{e}')
                continue
            try:
                att.SaveAsFile(att_out)
                index += 1
                logging.debug(f'Saving attachment to out dir: {att.filename}')
            except:
                logging.error(f'Failed to save attachment: {att.filename}')

            # update log table
            if not att.filename.endswith('.zip'):
                try:
                    file_log = log.log_to_table(
                        log_table=file_log,
                        filename=att_out,
                        file=att_out,
                        sourcefile=main_email,
                        attachment=True,
                        attachment_name=att.Filename,
                        review=False,
                        pdf_parsed=False,
                        combined=True,
                        # Set combined here since file extension will change later (e.g. from .png to .pdf)
                        duplicated=False
                    )
                    logging.debug(
                        f'Updating log table for attachment: {att_out}')
                except Exception as e:
                    logging.error(
                        f'Failed to update log table for attachment: {att_out}')

            if att.Filename.endswith(".msg"):
                logging.info(f'Attached message: {att.filename}')
                nested_dir = os.path.join(
                    new_dir, (os.path.basename(att_out).rsplit('.', 1)[0]))
                logging.info(f'Nested directory: {nested_dir}')
                try:
                    Path(nested_dir).mkdir(parents=True, exist_ok=True)
                    logging.debug(f'Creating nested directory: {nested_dir}')
                except:
                    logging.error(
                        f'Failed to create nested directory: {nested_dir}')
                # Call function recursively untill all nested MSG files are processed
                extract_msg(
                    file=att_out, root_dir=nested_dir, out_dir=out_dir,
                    main_email=main_email, file_log=file_log, process_dir=process_dir,
                    nested=True, index=index + 1
                )

        return attachment
