import subprocess

def lambda_handler(event, context):
        print("This is Event {}".format(event))

filepaths = ["removeocfiles.py",
             "OCIDD1.py",
             "OCIDU1.py",
             "OCIDU3.py",
             "OCIDU4.py",
             "OCIDU5.py",
             "OCIDU8.py",
             "OCIDU9P1.py",
             "OCNJU1P1.py",
             "OCNJU1P2.py",
             "OCNJU2.py",
             "OCNJU3.py",
             "excelmerge.py"
]

for filepath in filepaths:
    subprocess.call(["python", filepath])