import os
import subprocess
import shutil
import string
import random

def random_string(size=32, chars=string.ascii_letters + string.digits):
    return ''.join(random.choice(chars) for _ in range(size))

def base_dir(path_of_file):
    """
    Extract the previous or base directory name.
    """
    base_path = path_of_file.split("/")
    del base_path[0]
    del base_path[-1]
    conc_base = []
    for i in base_path:
        conc_base.append("/" + i)
    conc_base = "".join(conc_base)
    return conc_base

def search_file(path, ext):
    """
    Search in `path' the file with determinate `ext'.
    """
    to_search = []
    for local_files in os.listdir(path):
        if local_files.endswith(ext):
            to_search.append(local_files)
    return to_search

def create_zip(images_path):
    """
    Create a zip file, this contain the diapositive sequences in images
    """
    shutil.make_archive(images_path, 'zip', images_path)

def ppt_to_zip(orig_ppt_file):
    """
    Apply the transformation pptx in images. Return a zip with the images and the
    path of it.
    """
    ppt_file = orig_ppt_file.replace(" ", "_")
    random_name = random_string()
    base_directory = base_dir(ppt_file)
    process_pdf = base_directory + "/" + random_name + "/.libreoffice-headless/"
    image_dir = base_directory + "/" + random_name + "/" + "images"
    os.system("mkdir -p %s" % (process_pdf))
    os.system("mkdir -p %s" % (image_dir))
    command1 = "libreoffice --headless --convert-to pdf -env:UserInstallation=file:///%s %s" % (process_pdf, ppt_file)
    subprocess.call(command1, shell=True)
    pdf_file = os.getcwd() + "/" + search_file(base_directory, ".pdf")[0]
    command2 = "cd %s; convert -density 72 -resize 1024^ %s Outputfile.jpg" % (image_dir, pdf_file)
    subprocess.call(command2, shell=True)
    create_zip(image_dir)
    os.system("rm %s" % (pdf_file))
    return base_directory + "/" + random_name + "/" + "images.zip"
