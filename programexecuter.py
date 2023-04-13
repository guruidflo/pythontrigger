import subprocess

def lambda_handler(event, context):
        print("This is Event {}".format(event))

filepaths = ["firebasetokengenerator.py",
            "excelcreator.py",
            "formula.py",
             "finalexcel.py",
             "excelmerge2.py",
            "emailsender.py"
]

for filepath in filepaths:
    subprocess.call(["python", filepath])