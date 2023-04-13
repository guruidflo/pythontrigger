import subprocess

def lambda_handler(event, context):
        print("This is Event {}".format(event))

filepaths = [
             "finalexcel.py",
             "excelmerge2.py"]

for filepath in filepaths:
    subprocess.call(["python", filepath])