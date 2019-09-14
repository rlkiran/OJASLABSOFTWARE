
from firebase_admin import firestore

TestData = {}


def getDocuments(testname):
    temp = []
    store = firestore.client()
    docs = store.collection(testname).stream()
    for doc in docs:
        temp.append(doc.to_dict())
        # print(u'{} => {}'.format(doc.id, doc.to_dict()))
    return temp