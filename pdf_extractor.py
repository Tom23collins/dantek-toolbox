import os, shutil

def extract_pdfs(directory):
    pdfs_found = 0
    def moveFile(fileName, newpath):
        try:
            shutil.copy(fileName, newpath)
        except shutil.SameFileError:
            pass
        else:
            pass

    newpath = (os.path.join(os.path.join(os.environ['USERPROFILE']), r'Desktop\PDF Folder'))
    if not os.path.exists(newpath):
        os.makedirs(newpath)

    for root, dirs, files in os.walk(directory):
        for file in files:
            if file.endswith(".pdf"):
                pdfs_found += 1
                fileName = (os.path.join(root, file))
                fileName = fileName.replace('\\','/')
                moveFile(fileName, newpath)

    os.system(f'start {os.path.realpath(os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop'))}')
    return pdfs_found
