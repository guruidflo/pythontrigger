import firebase_admin
from firebase_admin import credentials, firestore
import re
import json

cred = credentials.Certificate("fabtrakr-gsheet-token-firebase-adminsdk-jcky7-dc35d5ecac.json")
firebase_admin.initialize_app(cred)

db = firestore.client()  
collection = db.collection('tokens') 
docs = collection.get()
token = docs[0].to_dict()['token']
print(token)
