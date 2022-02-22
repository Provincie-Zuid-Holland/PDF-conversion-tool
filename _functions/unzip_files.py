import logging
import os
import zipfile
from pathlib import Path
import _functions.check_length as cl

"""
    This file is called from the main file PDF_conversion_tool.py. It is used to unzip files and save status in a log file.
    
    Author: Joana Cardoso
"""


def unzip_files(zip_dir, proc_dir):
    """
    Unzips files and saves the unzipped files in a new directory

    Parameters
    ----------
    zip_dir: str
        The path to the selected zip file
    proc_dir: str 
        The directory of the unzipped folder
    """
    
    logging.info(f'Started unzipping folder: {zip_dir}')
    logging.info(f'Zip directory: {zip_dir}')
    logging.info(f'Process directory: {proc_dir}')

    with zipfile.ZipFile(zip_dir, 'r') as zip_file:
        for file in zip_file.namelist():
            filename = os.path.basename(file)
            logging.info(f'Found file {file} in zip directory')
            des_dir = os.path.join(proc_dir, os.path.dirname(file))
            logging.info(f'Unzipped directory: {des_dir}')
            try:
                # Create extract dir
                Path(des_dir).mkdir(parents=True, exist_ok=True)
                logging.debug(f'Creating unzipped directory: {des_dir}')
            except:
                logging.error(f'Failed to create unzip directory: {des_dir}')
            if not filename:
                continue
            else:
                file_name = cl.check_length(des_dir=des_dir, file=filename)[0]
                try:
                    data = zip_file.read(file)
                    # exporting to given location one by one
                    output = open(file_name, 'wb')
                    output.write(data)
                    output.close()
                    logging.debug(f'Unzipping file: {filename}')
                except:
                    logging.error(f'Failed to unzip: {filename}')

                # Call function recursively untill all zip files are extracted
                if filename.endswith('.zip'):
                    logging.info(f'Sub-zip filename: {filename}')
                    new_zip_dir = os.path.join(des_dir, filename)
                    new_proc_dir = os.path.abspath(
                        os.path.splitext(new_zip_dir)[0])
                    logging.info(f'new_zip_dir: {new_zip_dir}')
                    logging.info(f'new_proc_dir: {new_proc_dir}')

                    logging.info('Recursive')
                    unzip_files(zip_dir=new_zip_dir, proc_dir=new_proc_dir)
                    try:
                        os.remove(new_zip_dir)
                        logging.debug(f'Removed zip: {new_zip_dir}')
                    except:
                        logging.debug(f'Failed to remove zip: {new_zip_dir}')
    zip_file.close()
