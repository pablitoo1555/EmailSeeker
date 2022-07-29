import os


def create_dir(dir_name):
    """
    Creates directory if one does not already exist
    :param dir_name: path or name of dir
    """
    if not os.path.isdir(dir_name):
        os.makedirs(dir_name)
        print(f'created directory: "{dir_name}"')
    else:
        print(f'"{dir_name}" already exists')