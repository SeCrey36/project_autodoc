import os
dirname = '/WORK/project/programs/files'
dirfiles = os.listdir(dirname)

fullpaths = map(lambda name: os.path.join(dirname, name), dirfiles)

dirs = []
files = []

for file in fullpaths:
    if os.path.isdir(file): dirs.append(file)
    if os.path.isfile(file): files.append(file)

for i in files:
    print(i)
    
print(dirs)