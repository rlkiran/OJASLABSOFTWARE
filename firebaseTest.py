import firebase_admin
from firebase_admin import credentials
from firebase_admin import db


def getCollectionList():
    if len(firebase_admin._apps) < 2:
        cred = credentials.Certificate('ojas-key.json')
        default_app = firebase_admin.initialize_app(cred, {
            'databaseURL': 'https://ojas-c3eba.firebaseio.com/'
        })
    collectionList = []
    ref = db.reference('Tests')
    snapshot = ref.order_by_key().get()
    for key, val in snapshot.items():
        collectionList.append(key)
        # print('{0} => {1}'.format(key, val))
    return collectionList
