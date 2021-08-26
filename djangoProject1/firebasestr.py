import pyrebase
from datetime import date
from pathlib import Path
import os

config = {
    "databaseURL": "https://console.firebase.google.com/project/sgs-create/storage/sgs-create.appspot.com",
    "apiKey": "AIzaSyD5vI24SEl-D9vowI74LF8_TPHGZrsN4og",
    "authDomain": "sgs-create.firebaseapp.com",
    "projectId": "sgs-create",
    "storageBucket": "sgs-create.appspot.com",
    "messagingSenderId": "735404252029",
    "appId": "1:735404252029:web:f867f6db3853882d7d8a14",
    "measurementId": "G-46PJJXYGH1"
}

f = pyrebase.initialize_app(config=config)


# pos = "{}.sqlite3".format(db)
# poc = "{}/{}/{}.sqlite3".format(db, date.today(), db)
def upload(poc, pos):
    storage = f.storage()
    storage.child(poc).put(pos)


def download(poc, pos):
    storage = f.storage()
    storage.child(poc).download(pos)


def get(db):
    storage = f.storage()
    dbs = [str(i).split("/")[1] for i in storage.child(db).list_files() if str(i).split("/")[2].startswith(db)]
    return dbs


# upload("file/doc2.docx", "doc.docx")
