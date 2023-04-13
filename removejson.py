import os

folder_path = "C:\\Users\\Gurumurthy\\Documents\\Overview_Trigger\\"
keep_file = 'fabtrakr-gsheet-token-firebase-adminsdk-jcky7-dc35d5ecac.json'

files = os.listdir(folder_path)

# Loop through each file in the folder
for file in files:
    # If the file is a JSON file and is not the file to keep
    if file.endswith('.json') and file != keep_file:
        # Construct the full path to the file
        file_path = os.path.join(folder_path, file)
        # Remove the file
        os.remove(file_path)
