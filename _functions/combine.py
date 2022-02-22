import logging
import os
import pikepdf
import shutil

"""
This file is a python file that combines pdf files. This file is called from the main file PDF_conversion_tool.py, def convert_files.

Author: Joana Cardoso
"""

def combine_pdf(process_dir, out_dir, msg_list, progress, style):
    """
    Combines emails and attachements into 1 pdf

    Parameters
    ----------
    process_dir: str
        The path to the initial directory to convert
    out_dir: str
        The output directory where converted files will be placed
    msg_list: list
        List with all the emails and respective attachments
    progress: Progressbar
        The progress bar that is updated while running this file
    style: Style
        The progress bar style
    """
    # combine into 1 pdf and delete individual pdf's
    msg_dir = []
    for msg in msg_list:
        msg_dir.append(os.path.splitext(msg.replace(process_dir, out_dir))[0])
    logging.info(f'Number of head emails to combine: {len(msg_dir)}')
    for directory in msg_dir:
        logging.info(f'Combined PDF directory: {directory}')
        pdf = pikepdf.Pdf.new()
        for root, dirs, files in os.walk(directory):
            for name in files:
                name_lower = name.lower()
                if name_lower.endswith('.pdf'):
                    try:
                        file = os.path.join(root, name_lower)
                        src = pikepdf.Pdf.open(file)
                        pdf.pages.extend(src.pages)
                        logging.debug(f'Adding file to combined PDF: {file}')
                    except:
                        logging.error(
                            f'Failed to add file to combined pdf: {file}')
                    try:
                        pdf_combined = str(directory) + '.pdf'
                        src.close()
                        pdf.save(pdf_combined)
                        logging.debug(f'Saving PDF combined: {pdf_combined}')
                        try:
                            os.remove(file)
                            logging.debug(
                                f'Removing original PDF file: {file}')
                        except:
                            logging.error(
                                f'Failed to remove original PDF file: {file}')
                    except:
                        logging.error(
                            f'Failed to save PDF combined: {pdf_combined}')

        # set progress bar max value
        progress['value'] += 99 / len(msg_dir) if len(msg_dir) > 0 else 99
        progress.update()
        progress.update_idletasks()
        style.configure(
            'text.Horizontal.TProgressbar',
            text='{:.0f} %'.format(progress['value'])
        )

    # delete empty folders
    removed = set()
    for dirpath, dirnames, filenames in os.walk(out_dir, topdown=False):
        dirs = [dir for dir in dirnames if os.path.join(
            dirpath, dir) not in removed]
        contents = dirs+filenames
        if len(contents) == 0:
            try:
                shutil.rmtree(dirpath)
                removed.add(dirpath)
                logging.debug(f'Removing {dirpath}')
            except:
                logging.error(f'Failed to remove {dirpath}')
