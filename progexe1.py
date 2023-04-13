import subprocess

def lambda_handler(event, context):
        print("This is Event {}".format(event))

filepaths = ["removejson.py",
        ]

for filepath in filepaths:
    subprocess.call(["python", filepath])