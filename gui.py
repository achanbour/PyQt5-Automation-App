import generate_excel
import send_emails
import update_db
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QLabel
from PyQt5.QtCore import QObject, pyqtSignal

"""
The create_log_files() method creates two log files every time the program is run:

    - The app log file which logs every step run by the user in the current program session
    - The error log file which logs any error(s) encountered in the current program session

Both the app and error log files are reinitialised every time the program is run.
"""
def create_log_files():
    try:
        # Create and open app_log.txt
        with open('app_log.txt', 'w') as app_log:
            app_log.write('App log file\n')

        # Create and open error_log.txt
        with open('error_log.txt', 'w') as error_log:
            error_log.write('App error file\n')

        print("Log files created successfully.")
    except Exception as e:
        print(f"An error occurred: {e}")

"""
Define the MainWorker class as a subclass of the QObject class
The QObject class is used in this context to leverage its event handling features (the signal/slot mechanism)

The MainWorker is reponsible of executing the two main tasks of the KRI automation process: The Excel generation and email sending
Each task is defined a separate function calling methods from generate_excel.py and send_emails.py
Signals are defined to handle the progress and termination of each task
"""
class MainWorker(QObject):
    # Define the signal indicating completion after all KRI Excel reports have been generated and archived
    all_excels_generated = pyqtSignal()
    # Define the signal indicating completion after all KRI emails have been sent and archived
    all_emails_sent = pyqtSignal()
    # Define the signal that communicates the progress of the KRI Excel generation process for display in the progress bar
    number_of_excels_generated = pyqtSignal(int)
    # Define the signal that communicates the progress of the KRI email sending process for display in the progress bar
    number_of_emails_sent = pyqtSignal(int)

    def generate_excel(self):
        # Called when the user clicks on the "Generate Excel" button
        # Calls the generate_kri_report() method from generate_excel.py
        for prog in generate_excel.generate_kri_report:
            self.number_of_excels_generated.emit(prog)
        # Emits terminating signal upon completion
        self.all_excels_generated.emit()

    def send_emails(self):
        # Called when the user clicks on the "Send emails" button
        # Calls the send_automatic_emails() method from send_emails.py
        for prog in send_emails.send_automatic_emails:
            self.number_of_emails_sent.emit(prog)
        # Emits terminating signal upon completion
        self.all_emails_sent.emit()

"""
Define the UpdaterWorker class as a subclass of the QObject class
The UpdaterWorker is reponsible of executing the tasks related to updating the KRI table with the received KRI results
"""
class UpdaterWorker(QObject):
    # Define the signal indicating completion after the KRI table has been updated with the KRI results
    kri_table_updated = pyqtSignal()

    def update_kri_table(self):
        # Called when the user clicks on the "Update DB" button
        # Calls the update_kri_table() method from update_db.py
        update_db.update_kri_table()
        # Emits terminating signal upon completion
        self.kri_table_updated.emit()

