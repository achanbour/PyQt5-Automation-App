import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QLabel, QLineEdit, QComboBox, QCalendarWidget, QProgressBar
    QPushButton, QVBoxLayout, QHBoxLayout, QWidget
)
from PyQt5.QtCore import QObject, pyqtSignal
import generate_excel
import send_emails
import update_db

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

"""
Define the Window class as a subclass of QMainWindow.
The Window class defines the application's graphical elements and manages the actions triggered by user interactions with these components.
"""
class Window(QMainWindow):
    def __init__(self):
        super().__init__()
        # Set the window's title, geometry and size
        self.setWindowTitle("KRI Automation Main Window")
        self.setGeometry(100, 100, 600, 400)

        # Initialise all user-collected data (collected via the UI widgets)
        self.user_entered_year = None
        self.user_selected_month = None
        self.user_selected_deadline = None

        # Create central widget and layout
        central_widget = QWidget(self)
        self.setCentralWidget(central_widget)
        self.layout = QVBoxLayout(central_widget)

        # Create and set up a progress bar (hidden initially)
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setFixedHeight(30)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.hide()
        self.layout.addWidget(self.progress_bar)

        # Add a spacer item to push the rest of the UI down
        self.layout.addStretch()

        # Define the form layout where labels and input widgets will be contained
        self.form_layout = QHBoxLayout()

        # Call the UiComponents() and UiWidgets() methods to configure the window's graphical elements
        self.UiComponents()
        self.UiWidgets()

        # Add the form layout to the main layout
        self.layout.addLayout(self.form_layout)

    # The UiComponents() method defines the window's components: Labels and buttons
    def UiComponents(self):
        # Define a label layout
        label_layout = QVBoxLayout()

        # Create a left column with labels
        self.label_year = QLabel("Enter year", self)
        self.label_month = QLabel("Select month", self)
        self.label_deadline = QLabel("Select deadline", self)

        ## Add all labels to the label layout
        label_layout.addWidget(self.label_year)
        label_layout.addWidget(self.label_month)
        label_layout.addWidget(self.label_deadline)

        # Add the label layout to the form layout
        self.form_layout.addLayout(label_layout)

        # Create a button layout
        button_layout = QVBoxLayout()

        # Create a row of buttons
        self.btn_generate_excel = QPushButton("Generate Excel", self)
        self.btn_send_emails = QPushButton("Send emails", self)
        self.btn_send_emails.setEnabled(False)  # Disable the "Send emails" button initially
        self.btn_update_db = QPushButton("Update DB", self)

        # Add all buttons to the button layout
        button_layout.addWidget(self.btn_generate_excel)
        button_layout.addWidget(self.btn_send_emails)
        button_layout.addWidget(self.btn_update_db)

        # Add the button layout to the form layout
        self.form_layout.addLayout(button_layout)

        # Define the action on each button
        self.btn_generate_excel.clicked.connect(self.start_generate_excel)
        self.btn_send_emails.clicked.connect(self.start_send_emails)
        self.btn_update_db.clicked.connect()

    # The UiWidgets() method defines the window's widgets
    def UiWidgets(self):
        # Define an input layout
        input_layout = QVBoxLayout()

        # Create a right column with inputs
        self.input_year = QLineEdit(self)
        self.combo_month = QComboBox(self)
        self.combo_month.addItems(["January", "February", "March", "April", "May", "June",
                                   "July", "August", "September", "October", "November", "December"])
        self.calendar_deadline = QCalendarWidget(self)

        # Add all inputs to the input layout
        input_layout.addWidget(self.input_year)
        input_layout.addWidget(self.combo_month)
        input_layout.addWidget(self.calendar_deadline)

        # Add the input layout to the form layout
        self.form_layout.addLayout(input_layout)

    # Define a thread for generating the Excel reports and sending emails (both tasks are intentionally included in the same thread
    # as they are performed sequentially)
    def setUpMainThread(self):
        self.worker_object = MainWorker()
        self.worker_thread = QThread()
        self.worker_object.moveToThread(self.worker_thread)
        self.worker_thread.started.connect(self.worker_object.generate_excel)

        # Connect signals to slots
        self.worker_object.number_of_excels_generated.connect(self.update_progress_bar)
        self.worker_object.all_excels_generated.connect(self.display_all_excels_generated_confirmation_message)
        self.worker_object.all_excels_generated.connect(self.enable_send_emails_button)
        self.worker_object.number_of_emails_sent.connect(self.update_progress_bar)
        self.worker_object.all_emails_sent.connect(self.display_all_emails_sent_confirmation_message)

    """
    Define all slot functions
    """

    # Define a method to enable the "Send emails" button (another slot function to the all_excels_generated signal)
    def enable_send_emails_button(self):
        self.btn_send_emails.setEnabled(True)

    def start_generate_excel(self):
        # Show the progress bar before starting
        self.show_progress_bar()
        # Start the worker thread
        self.worker_thread.start()

    def start_send_emails(self):
        # Note: Calling the send_emails() method on the worker assumes that the worker thread is already running
        ## Since the "Send emails" step is executed right after "Generate Excel" in the same program session, it is safe to execute this step
        ## in the same thread without opening a new one. It is for this same reason that the "Send emails" button is initially disabled and
        ## only made available to the user after the "Generate Excel" step successfully completes
        self.worker_object.send_emails()  # Call send_emails directly

    # The method below shows the progress bar
    def show_progress_bar(self):
        self.progress_bar.show()
        self.progress_bar.setValue(0)

    def update_progress_bar(self, value):
        QApplication.processEvents()
        self.progress_bar.setValue(value)

    # The method below is a slot function called upon termination of the excel generation task
    def display_all_excels_generated_confirmation_message(self):
        print("All KRI Excel reports have been successfully generated.")
        # Reset and hide progress bar
        self.reset_and_hide_progress_bar()

    # The method below is a slot function called upon termination of the email sending task
    def display_all_emails_sent_confirmation_message(self):
        print("All KRI emails have been successfully sent.")
        # Reset and hide progress bar
        self.reset_and_hide_progress_bar()

    # The method below resets the progress bar and hides it
    def reset_and_hide_progress_bar(self):
        self.progress_bar.setValue(0)
        self.progress_bar.hide()

# Main execution
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = Window()
    window.show()
    sys.exit(app.exec_())
