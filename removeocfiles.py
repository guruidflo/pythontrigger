import os
import glob

dir = "C:\\Users\\Gurumurthy\\Documents\\Overview_Trigger\\oc_detils\\"
filelist = glob.glob(os.path.join(dir, "*"))
for f in filelist:
  os.remove(f)