import sys
from PyQt5.QtWidgets import QApplication
from mainWindow import SalesforceQueryApp

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = SalesforceQueryApp()
    window.show()
    
    sys.exit(app.exec_())