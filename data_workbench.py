    def error(self, exc):
        QMessageBox.critical(self, "FrMat", str(exc))
