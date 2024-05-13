import datetime
import os
import glob
import shutil

section1 = '\n\n' + '***********************************************************' + '\n\n'

##########################################################################
#########################                        #########################
######################### Author: Muneeb Qureshi #########################
#########################                        #########################
##########################################################################

def FileSplit(category, file, extension):

    count168 = 0
    count169 = 0
    sum168 = 0
    sum169 = 0
    Lines168 = []
    Lines169 = []

    # Reads file

    with open(file, 'r') as file:

        for line in file:

            if line[:3] == '168':

                Lines168.append(line[4:])
                count168 += 1

                if category == 2: sum168 += float(line.split('|')[-1])
                
            elif line[:3] == '169':

                Lines169.append(line[4:])
                count169 += 1

                if category == 2: sum169 += float(line.split('|')[-1])

    if category == 1:

        Lines168.append(f'ATRL|168|UNIFIN|{count168}|{date}|')
        Lines169.append(f'ATRL|169|UNIFIN|{count169}|{date}|')

    elif category == 2:

        sum168 = round(sum168, 2)
        sum169 = round(sum169, 2)
        
        Lines168.append(f'ATRL|168|UNIFIN|{count168}|{sum168}|{date}|')
        Lines169.append(f'ATRL|169|UNIFIN|{count169}|{sum169}|{date}|')
    
    return Lines168, Lines169


def CreateFiles(file168, file169, list168, list169):

    with open(file168, 'w') as file:

        for item in list168:
            
            file.write(item)

    with open(file169, 'w') as file:

        for item in list169:

            file.write(item)


def MoveFiles(file, path, ext, date, name168, name169):

    cwd = path + f'\\{ext}\\{date}\\'

    if not os.path.exists(cwd): os.makedirs(cwd)

    old168 = path + name168
    new168 = cwd + name168

    old169 = path + name169
    new169 = cwd + name169

    newFile = cwd + f'Dont send.{ext}'

    shutil.move(file, newFile)
    shutil.move(old168, new168)
    shutil.move(old169, new169)


# --------------------------------------------------------------------------------------- #

# Current date

date = datetime.date.today().strftime('%Y%m%d')

# Folder name

folder = datetime.date.today().strftime('%m-%d-%Y')

# Current working directory

path = os.getcwd()
os.chdir(path)

# --------------------------------------------------------------------------------------- #

pattern = 'AIM*.????'

# Gets all AIM files

AIMFiles = glob.glob(pattern)

for name in AIMFiles:

    fileName = str(name)

    fileExt = fileName.split('.')[1]

    if fileExt == 'APDT': category = 2
    else: category = 1

    name168 = fileName.split('_')[0] + f'_168.{fileExt}'
    name169 = fileName.split('_')[0] + f'_169.{fileExt}'

    List168, List169 = FileSplit(category, fileName, fileExt)
    CreateFiles(path + name168, path + name169, List168, List169)

    print(section1)
    print(f'{fileExt} File has been splitted successfully.'.center(60))

    MoveFiles(name, path, fileExt, folder, name168, name169)

    print(f'Files are moved to {fileExt} folder.'.center(60))






