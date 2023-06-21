import hashlib
import logging
import os, re
from pathlib import Path

"""
This file is a python file that processes duplicate files. This file is called from the main file PDF_conversion_tool.py, def convert_files.

Author: Joana Cardoso
"""


def remove_duplicates(process_dir, out_dir, progress, style, file_log, remove_duplicates):
    """
    Identifies duplicates and if set_remove_duplicates is set to True it moves duplicates to another folder.

    Parameters
    ----------
    process_dir: str 
        The path to the initial directory to convert
    out_dir: str 
        The output directory where converted files will be placed
    progress: Progressbar 
        The progress bar that is updated while running this file
    style: Style
        The progress bar style
    file_log: list 
        The logging table
    remove_duplicates: bool
        Remove_duplicates is set to True or False, depending whether duplicates are to be removed or not

    Returns
    -------
    out_duplicates: str
        The path to the duplicates directory
    duplicates: set 
        The duplicate files
    """

    md5_files = []

    for root, dirs, files in os.walk(out_dir):
        for name in files:
            if os.path.getsize(os.path.join(root, name)) > 0:
                # print(f'size:{os.path.getsize(os.path.join(root, name))}')
                md5_files.append(os.path.join(root, name))
            else:
                    logging.error(f'Document is empty: {os.path.join(root, name)}')
                    # print(f'Document is empty: {os.path.join(root, name)}')

    # List duplicates
    md5_dict = {}
    for f in md5_files:
        md5_dict[f] = hashlib.md5(Path(f).read_bytes()).hexdigest()

    result = {}
    for key, value in md5_dict.items():
        if value not in result.values():
            result[key] = value
    keys_a = set(md5_dict.keys())
    keys_b = set(result.keys())
    duplicates = keys_a ^ keys_b

    # Remove duplicates
    out_duplicates = os.path.join(os.path.dirname(
                        process_dir), "Dubbelingen_" + os.path.basename(process_dir))

    if len(duplicates) > 0:
        for rmfile in duplicates:
            src = rmfile
            target = rmfile.replace(out_dir, out_duplicates)
            base_dir = os.path.dirname(target)

            if remove_duplicates == True:
                logging.info(f'Duplicates directory: {out_duplicates}')
                logging.info(f'Duplicate files to remove: {len(duplicates)}')
                
                try:
                    Path(base_dir).mkdir(parents=True, exist_ok=True)
                    logging.debug(f'Creating duplicates directory: {base_dir}')
                except:
                    logging.error(
                        f'Failed to create duplicates directory: {base_dir}')
                try:
                    os.rename(src, target)
                    logging.debug(
                        f'Moving duplicate file to duplicates directory: {rmfile}')
                except:
                    logging.error(
                        f'Failed to move duplicate file to duplicates directory: {rmfile}')

                # set progress bar max value
                progress['value'] += 99 / \
                    len(duplicates) if len(duplicates) > 0 else 99
                progress.update()
                progress.update_idletasks()
                style.configure(
                    'text.Horizontal.TProgressbar',
                    text='{:.0f} %'.format(progress['value'])
                )

            # update log table
            if re.search('\d+.txt$', rmfile):
                file_msg= os.path.dirname(rmfile) + '.msg'            
                try:
                    d = next(item for item in file_log if
                            os.path.join(os.path.dirname(item['File']), item['File name']) == file_msg)
                    d['Duplicate'] = True
                    d['Combined'] = False
                    logging.debug(
                        f'Updating log table for duplicate: {file_msg}')
                except:
                    logging.error(
                        f'Failed to update log table for duplicate: {file_msg}')
            else:           
                try:
                    d = next(item for item in file_log if
                            os.path.join(os.path.dirname(item['File']), item['File name']) == rmfile)
                    d['Duplicate'] = True
                    d['Combined'] = False
                    logging.debug(
                        f'Updating log table for duplicate: {rmfile}')
                    
                except:
                        logging.error(
                            f'Failed to update log table for duplicate: {rmfile}')

    else:
        progress['value'] += 99
        progress.update()
        progress.update_idletasks()
        style.configure(
            'text.Horizontal.TProgressbar',
            text='{:.0f} %'.format(progress['value'])
        )

    return out_duplicates, duplicates
