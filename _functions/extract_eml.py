import os
import logging
from pathlib import Path
from email import message_from_file
from email import policy
from email.parser import BytesParser
import office2pdf_gui_debug as gui

# logging.basicConfig(level=logging.DEBUG,
#                     format='%(asctime)s %(levelname)s %(funcName)s %(message)s',
#                     filename='Logging_SelfServiceToPDF.log'
#                     )  # to see log in console remove filename


def extract_eml (file,root_dir,main_email,file_log,index=0):

    new_dir = os.path.join(root_dir, (os.path.basename(file).rsplit('.', 1)[0]))

    try:
        Path(new_dir).mkdir(parents=True, exist_ok=True)
        # logging.debug(f'Creating new directory: {new_dir}')
    except:
        # logging.error(f'Failed to create new directory: {new_dir}')
        print(f'Failed to create new directory: {new_dir}')
    
    ## get email body
    fileName, fileExtension = os.path.splitext(file)
    new_name = fileName.replace(fileName, str(index)) + '.txt'
    out_txt = os.path.join(new_dir, new_name)
    logging.info(f'File: {file}')
    logging.info(f'TXT: {out_txt}')

    try:
        with open(file, 'rb') as fp:
            msg = BytesParser(policy=policy.default).parse(fp)
            txt = msg.get_body(preferencelist=('plain')).get_content()

        Date = msg["date"]
        From = msg["from"].strip()
        To = msg["to"].strip()
        Subject = msg["subject"].strip()

        textFile = ""
        textFile += "From: " + From + "\n"
        textFile += "To: " + To + "\n"
        textFile += "Subject: " + Subject + "\n"
        textFile += "Date: " + Date + "\n\n"
        textFile += txt

        # logging.debug(f'Reading eml: {file}')
        print(f'Reading eml: {file}')
    except:
        # logging.error(f'Failed to read eml: {file}')
        print(f'Failed to read eml: {file}')

    # save eml body as txt file    
    try:
        with open(out_txt, 'w',encoding="utf-8") as f:
            f.write(textFile)
        # logging.debug(f'Saving eml {file} as {fnm}')
        print(f'Saving eml {file} as {new_name}')
    except:
        # logging.error(f'Failed to save eml {file} as {fnm}')
        print(f'Failed to save eml {file} as {new_name}')
    index += 1

    # extract attachments
    eml = message_from_file(open(file))
    attachments = eml.get_payload()
    print(f'Found {len(attachments)} attachments')
    for attachment in attachments:
        fnam = attachment.get_filename()
        if fnam is not None:
            print(fnam)
            print(len(attachment.get_payload()))
            fileName, fileExtension = os.path.splitext(fnam)
            att_name = fileName.replace(fileName, str(index)) + fileExtension
            att_out = os.path.join(os.path.abspath(new_dir), att_name)
            print(att_out)

            if not att_name.endswith(".eml"):
                try:
                    print(attachment.get_charset())
                    f = open(att_out, 'wb').write(attachment.get_payload(decode=True,))
                    index += 1
                # logging.debug(f'Saving attachment {fnam}')
                    print(f'Saving attachment {fnam} as {att_name}')
                except:
                    print(f'Failed to save attachment {fnam} as {att_name}')
            # else
            # update log table
            try:
                file_log = gui.log_to_table(
                    log_table=file_log,
                    filename=att_out,
                    file=att_out,
                    sourcefile=main_email,
                    attachment=True,
                    attachment_name=fnam,
                    review=True,
                    pdf_parsed=False,
                    combined=True,
                    ## Set combined here since file extension will change later (e.g. from .png to .pdf)
                    duplicated=False
                )
                logging.debug(f'Updating log table for attachment: {att_out}')
            except:
                logging.error(f'Failed to update log table for attachment: {att_out}')