from PyQt5.QtCore import QFileSystemWatcher, QCoreApplication
import sys


class ExcelFileWatcher:

    def __init__(self, excel_file_path):
        self.app = QCoreApplication(sys.argv)
        self.excel_file_path = excel_file_path
        self.watcher = QFileSystemWatcher()
        self.watcher.addPath(self.excel_file_path)
        self.watcher.fileChanged.connect(self.on_excel_file_changed)

    def on_excel_file_changed(self, path):
        # This function is called when the Excel file is saved
        print(f"Excel file {path} has been modified.")
        # Call your desired function here
        self.handle_excel_save()

    def handle_excel_save(self):
        # Your function logic here
        print("Handling Excel save event...")
        # For example, reload the Excel file or update the UI
        # ...

    def start(self):
        sys.exit(self.app.exec_())


if __name__ == "__main__":
    excel_file_path = "/Users/remillieux/Desktop/testV36.xlsm"
    watcher = ExcelFileWatcher(excel_file_path)
    watcher.start()
