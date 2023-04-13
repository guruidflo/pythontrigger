import glob, os, os.path

filelist = glob.glob(os.path.join("*.xlsx"))
for f in filelist:
    os.remove(f)