import os
from shutil import copyfile
import zipfile

file_name = ""
extract_to = ""
ACCEPTED_EXTENSION_LIST = ['.docx']


def get_user_input():
    global file_name
    global extract_to

    input_valid = False

    while not input_valid:
        file_name = input("Please input the valid document file name: ")

        if not is_valid_path(file_name):
            input_valid = False
            continue
        else:
            input_valid = True

        if not is_file_extension_valid(file_name):
            input_valid = False
            continue
        else:
            input_valid = True

    input_valid = False

    while not input_valid:
        extract_to = input("Please input the directory to extract the embeddings to: ")

        if not is_valid_path(file_name):
            input_valid = False
            continue
        else:
            file_name_copy = os.path.splitext(os.path.basename(file_name))[0]
            extract_to += file_name_copy + "/"
            input_valid = True


def copy_file_and_rename(file, new_file_name):
    copyfile(file, new_file_name)


def extract_embeddings(archive):

    zip_archive = zipfile.ZipFile(archive)

    zip_archive.extractall(extract_to)

    print('Extracted content to: ' + extract_to)


def is_valid_path(path):
    if os.path.isfile(path):
        return True
    else:
        print("Please specify a valid file and directory")
        return False


def is_file_extension_valid(path):
    extension = os.path.splitext(path)[1]
    if extension in ACCEPTED_EXTENSION_LIST:
        return True
    else:
        print("File name must be any of the following:\n " + ''.join(str(item) for item in ACCEPTED_EXTENSION_LIST))
        return False


def clean_up(file_to_cleanup):
    os.remove(file_to_cleanup)


if __name__ == '__main__':
    get_user_input()
    copy_file_and_rename(file_name, file_name + '.zip')
    extract_embeddings(file_name + ".zip")
    clean_up(file_name + ".zip")

