import os
import sys
from datetime import date
from PyQt5 import QtWidgets, uic
from PyQt5.QtWidgets import QMessageBox, QListWidgetItem
from PyQt5.QtCore import Qt
import firebase_admin
from firebase_admin import credentials
from firebase_admin import db
from firebase_admin import firestore
import mixedTemplateScript
import printerHelper

today = date.today()
paramData = []
paramList = []
AllParamDetails = []
AllTestDetails = []
tmpdict = {}

template = "Template002.docx"

PatientName = ''
PatientAge = ''
PatientSex = ''
PatientId = ''
PatientPhone = ''
SpecimenID = ''
DateCollected = ''
DateReceived = ''
DateOfReport = ''
SampleType = ''
TestName = ''
Parameter = ''
result = ''
units = ''
referenceRange = ''
Doctor = ''
CCName = ''
CCCode = ''
Location = ''
Phone = ''
filename = ''
pwd = ''
docLoc = ''
wordFields = {}


class Ui(QtWidgets.QMainWindow):
    def __init__(self):
        super(Ui, self).__init__()
        uic.loadUi('MainWindow.ui', self)
        global pwd
        global docLoc
        pwd = os.getcwd()
        docLoc = pwd + "\\Docs"
        # self.statusBar().showMessage("WARNING DON'T PRESS SAVE BEFORE ADDING PARAMETERS")
        self.savebutton = self.findChild(QtWidgets.QPushButton, 'saveButton')
        self.printbutton = self.findChild(QtWidgets.QPushButton, 'printButton')
        self.exitbutton = self.findChild(QtWidgets.QPushButton, 'ExitButton')
        self.AddParambutton = self.findChild(QtWidgets.QPushButton, 'addParamButton')
        self.AddTestButton = self.findChild(QtWidgets.QPushButton, 'AddTestButton')

        self.addedParamList = self.findChild(QtWidgets.QListWidget, 'paramAddedList')

        self.checkBox = self.findChild(QtWidgets.QCheckBox, 'checkBox')

        self.checkBox.stateChanged.connect(self.addLogo)

        self.exitbutton.clicked.connect(self.quitApp)
        self.AddParambutton.clicked.connect(self.AddParamsList)
        self.AddTestButton.clicked.connect(self.AddTestsList)
        self.savebutton.clicked.connect(self.saveDocument)
        self.printbutton.clicked.connect(self.printFile)

        self.unitsLabel = self.findChild(QtWidgets.QLabel, 'unitsLabel')
        self.rr = self.findChild(QtWidgets.QLabel, 'RRLabel')

        # Patient Details
        self.ptname = self.findChild(QtWidgets.QLineEdit, 'ptName_ip')
        self.ptage = self.findChild(QtWidgets.QLineEdit, 'ptAge_ip')
        self.ptId = self.findChild(QtWidgets.QLineEdit, 'ptId_ip')
        self.ptphone = self.findChild(QtWidgets.QLineEdit, 'ptPhone_ip')
        self.ptRb1 = self.findChild(QtWidgets.QRadioButton, 'maleRb')
        self.ptRb2 = self.findChild(QtWidgets.QRadioButton, 'femaleRb')
        # Specimen Details
        self.spcID = self.findChild(QtWidgets.QLineEdit, 'sdId_ip')
        self.spcDateC = self.findChild(QtWidgets.QLineEdit, 'sdDc_ip')
        self.spcDateR = self.findChild(QtWidgets.QLineEdit, 'sdDR_ip')
        self.spcDateRp = self.findChild(QtWidgets.QLineEdit, 'sdDoR_ip')
        self.spcType = self.findChild(QtWidgets.QLineEdit, 'sdSt_ip')
        # Doc and CC Details
        self.DocName = self.findChild(QtWidgets.QLineEdit, 'DDName_ip')
        self.cc_code = self.findChild(QtWidgets.QLineEdit, 'DCCcode_ip').setText("PGTPT01")
        self.cc_name = self.findChild(QtWidgets.QLineEdit, 'DCCName_ip').setText("OJAS HEALTH CARE")
        self.cc_loc = self.findChild(QtWidgets.QLineEdit, 'DCCLOC_ip').setText("TIRUPATI")
        self.cc_phone = self.findChild(QtWidgets.QLineEdit, 'DCCPhone_ip')
        # Test Details
        self.resultLabel = self.findChild(QtWidgets.QLineEdit, 'result')
        self.TestsList = self.findChild(QtWidgets.QComboBox, 'TestList_cb')
        self.ParameterList = self.findChild(QtWidgets.QComboBox, 'parameter_cb')
        self.technologyList = self.findChild(QtWidgets.QComboBox, 'technology')
        self.setWindowTitle("Lab Software (Document Saver)")
        self.show()
        self.setTestData()
        self.TestsList.currentTextChanged.connect(self.setParams)
        self.ParameterList.currentTextChanged.connect(self.UAndRR)

    def quitApp(self):
        sys.exit(0)

    def addLogo(self, state):
        global template
        if state == Qt.Checked:
            template = "Template002_logo.docx"
        else:
            template = "Template002.docx"

    def setTestData(self):
        temp1, temp2 = self.getCollectionList()
        self.TestsList.addItems(temp1.copy())
        self.technologyList.addItems(temp2.copy())

    def setParams(self):
        global paramData
        global paramList
        paramData.clear()
        paramList.clear()
        self.ParameterList.clear()
        paramData.append(self.getDocuments(self.TestsList.currentText()))
        for pD in paramData:
            for p in pD:
                paramList.append(p['Parameter'])
        self.ParameterList.addItems(paramList)

    def UAndRR(self):
        global paramData
        for p in paramData:
            for urr in p:
                if self.ParameterList.currentText() in urr['Parameter']:
                    # print(urr['Units'])
                    self.unitsLabel.setText(urr['Units'])
                    self.rr.setText(urr['RefRange'])

    def AddParamsList(self):
        global AllParamDetails
        global tmpdict
        tmpdict['TestName'] = self.TestsList.currentText()
        tmpdict['Parameter'] = self.ParameterList.currentText()
        tmpdict['Technology'] = self.technologyList.currentText()
        tmpdict['Units'] = self.unitsLabel.text()
        tmpdict['RefRange'] = self.rr.text()
        tmpdict['result'] = self.resultLabel.text()
        AllParamDetails.append(tmpdict.copy())
        item = QListWidgetItem(tmpdict['Parameter'])
        self.addedParamList.addItem(item)
        QMessageBox.about(self, "Done", "Parameter added Successfully")

    def AddTestsList(self):
        global AllParamDetails
        global AllTestDetails
        AllTestDetails.append(AllParamDetails.copy())
        AllParamDetails.clear()
        self.addedParamList.clear()
        QMessageBox.about(self, "Done", "Test added Successfully")

    def sendUserDataToDoc(self):
        global wordFields
        global PatientName
        global PatientAge
        global PatientSex
        global PatientId
        global PatientPhone
        global SpecimenID
        global DateCollected
        global DateReceived
        global DateOfReport
        global SampleType
        global Doctor
        global CCName
        global CCCode
        global Location
        global Phone
        global TestName
        global pwd
        global docLoc

        PatientName = self.ptname.text()
        PatientAge = self.ptage.text()
        PatientSex = 'male' if self.ptRb1.isChecked() else 'female'
        PatientId = self.ptId.text()
        PatientPhone = self.ptphone.text()
        SpecimenID = self.spcID.text()
        DateCollected = self.spcDateC.text()
        DateReceived = self.spcDateR.text()
        DateOfReport = self.spcDateRp.text()
        SampleType = self.spcType.text()
        TestName = str(self.TestsList.currentText())
        Doctor = self.DocName.text()
        CCName = self.DCCName_ip.text()
        CCCode = self.DCCcode_ip.text()
        Location = self.DCCLOC_ip.text()
        Phone = self.cc_phone.text()

        wordFields['PatientName'] = PatientName
        wordFields['PatientAge'] = PatientAge
        wordFields['PatientSex'] = PatientSex
        wordFields['PatientId'] = PatientId
        wordFields['PatientPhone'] = PatientPhone
        wordFields['SpecimenID'] = SpecimenID
        wordFields['DateCollected'] = DateCollected
        wordFields['DateReceived'] = DateReceived
        wordFields['DateofReport'] = DateOfReport
        wordFields['SampleType'] = SampleType
        wordFields['Doctor'] = Doctor
        wordFields['Phone'] = Phone

        filename = PatientName + "--" + str(today.strftime("%d-%m-%Y")) + '.docx'
        mixedTemplateScript.addPatientDetails(template, docLoc, filename, **wordFields)
        # QMessageBox.about(self, "Done", "Details added Successfully")
        os.chdir(pwd)

    def saveDocument(self):
        global AllTestDetails
        global docLoc
        global PatientName
        self.sendUserDataToDoc()
        filename = PatientName + "--" + str(today.strftime("%d-%m-%Y")) + '.docx'
        for test in AllTestDetails:
            mixedTemplateScript.addTestTables(docLoc, filename, *test)
        os.chdir(pwd)

    def printFile(self):
        global docLoc
        global PatientName
        # self.statusBar().showMessage('No Printer Added')
        filename = PatientName + "--" + str(today.strftime("%d-%m-%Y")) + '.docx'
        os.chdir(docLoc)
        if not printerHelper.printFile(filename):
            self.statusBar().showMessage('Failed To Print File')

    def getCollectionList(self):
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
        technologyList = []
        ref = db.reference('Technology')
        snapshot = ref.order_by_key().get()
        for key, val in snapshot.items():
            technologyList.append(key)

        return collectionList, technologyList

    def getTechnologyList(self):
        if len(firebase_admin._apps) < 2:
            cred = credentials.Certificate('ojas-key.json')
            default_app = firebase_admin.initialize_app(cred, {
                'databaseURL': 'https://ojas-c3eba.firebaseio.com/'
            })
        collectionList = []
        ref = db.reference('Technology')
        snapshot = ref.order_by_key().get()
        for key, val in snapshot.items():
            collectionList.append(key)
        # print('{0} => {1}'.format(key, val))
        return collectionList

    def getDocuments(self, testname):
        temp = []
        store = firestore.client()
        docs = store.collection(testname).stream()
        for doc in docs:
            temp.append(doc.to_dict())
            # print(u'{} => {}'.format(doc.id, doc.to_dict()))
        return temp


app = QtWidgets.QApplication(sys.argv)
window = Ui()
app.exec_()
