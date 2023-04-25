# Organise image filenames after insertions or deletions
import os

def renamer(file_start=0, file_end=0, test=False): # [!] to add mod, folder, ext
    # prompt folder
    # prompt extension
    folder = input("ENTER FOLDER TO IMAGES: ")
    ext = ".JPG"
    
    # prompt file_start, file_end
    if file_start == 0 and file_end == 0:
        file_start = int(input("From File (Start): "))
        file_end = int(input("To File (End): "))
    else:
        file_start = file_start
        file_end = file_end

    # prompt what type of rename modification: shift
    mod = "shift-up"

    # Assign paths to old file names and new file names
    old_names, new_names = mod_files(file_start, file_end, mod)
    old_path, new_path = path(old_names, new_names, folder=folder, ext=ext)
    
    # Confirm rename


    # If file exist, start rename
    if os.path.exists(new_path[0]):
        print("Existing file conflict! File: ", new_path[0])
        print("Exiting program.")
        return 300
    for old_file, new_file in zip(old_path, new_path):
        if os.path.exists(old_file):
            if test == False:
                os.rename(old_file, new_file)
            print("Old File: ", old_file)
            print("New File: ", new_file)
        else:
            print("File Not Found: ", old_file)
            print("Exiting program.")
            return 404

def mod_files(file_start, file_end, mod):
    # Modifies file name with modification method selected

    # Shift
    if mod == "shift-up": # Shifts file name up by "shift"
        # prompt shift amount
        shift = 1
        old_names = [str(x) for x in range(file_end, file_start-1, -1)]
        new_names = [str(int(x)+shift) for x in old_names]
    if mod == "shift-down": # Shifts file name down by "shift"
        # prompt shift amount
        shift = 1
        old_names = [str(x) for x in range(file_start, file_end+1)]
        new_names = [str(int(x)-shift) for x in old_names]
        
    return(old_names, new_names)

def path(*names, folder, ext):
    # Returns file name with full folder path and extension
    path =[]
    for i in names:
        files = []
        for file in i:
            files += [folder + file + ext]
        path += [files]
    return path

# Menu

# Can take arguments (file_start, file_end), otherwise asks for input
renamer(test=True)

# Unix:
# use: "touch {1..30}.JPG"
# for example to create files for testing, make new folder with mkdir