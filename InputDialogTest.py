from PyQt5.QtWidgets import QDialog, QLineEdit, QDialogButtonBox, QFormLayout, QApplication


class InputDialogTest(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.first = QLineEdit(self)
        self.second = QLineEdit(self)
        self.third = QLineEdit(self)
        buttonBox = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel, self)

        layout = QFormLayout(self)
        layout.addRow("Enter Test Name :", self.first)
        layout.addRow("Enter Parameter :", self.second)
        layout.addRow("Enter Unit :", self.third)
        layout.addWidget(buttonBox)

        buttonBox.accepted.connect(self.accept)
        buttonBox.rejected.connect(self.reject)
        self.setWindowTitle("Edit Tests Data")

    def getInputs(self):
        return self.first.text(), self.second.text(), self.third.text()


if __name__ == '__main__':

    import sys

    app = QApplication(sys.argv)
    dialog = InputDialogTest()
    if dialog.exec():
        print(dialog.getInputs(), type(dialog.getInputs()))
        print(dialog.getInputs()[0])
    exit(0)
