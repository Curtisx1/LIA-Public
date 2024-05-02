from datetime import datetime, time, date
from dataclasses import asdict
import time as tm
import glob
import json
import logging
import os
import pythoncom
import pywintypes
import re
import requests
import soundfile
import speech_recognition as sr
import sys
import threading
import urllib3
import win32com.client as win32
from collections import Counter
from dataclasses import dataclass
from dotenv import load_dotenv
from jira import JIRA
from jira.exceptions import JIRAError
from PyQt5.QtCore import Qt, QTimer, QThread, pyqtSignal, pyqtSlot, QUrl, QSize, QTimer, QTime, QPropertyAnimation, QRect, QPoint, QEasingCurve
from PyQt5.QtGui import QFont, QColor, QPainter, QImage, QPalette, QIcon, QPixmap, QMovie
from PyQt5.QtWebEngineWidgets import QWebEngineView
from PyQt5.QtWidgets import (QApplication, QAction, QFrame, QGridLayout, QHBoxLayout, QLabel, QMainWindow, QMessageBox, 
                             QSizePolicy, QSpacerItem, QPushButton, QTextEdit, QToolBar, QVBoxLayout, QWidget, QLineEdit, QPushButton, QTableWidget, QTableWidgetItem, QAbstractItemView, QScrollArea, QDialog)
from requests.auth import HTTPBasicAuth
import xml.etree.ElementTree as ET

# Setting up the enviroment variables
dir_path = os.path.dirname(os.path.realpath(__file__))
config_path = os.path.join(dir_path, 'Config')
dotenv_files = glob.glob(os.path.join(config_path, '*.env'))
if dotenv_files:
    load_dotenv(dotenv_files[0])
else:
    print("Unable to load .env file.")

# Disabling insecure warning message - Do not try at home
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

#Create log file
log_dir = os.getenv('LOGS_PATH')
log_file = os.path.join(log_dir, 'errors.log')
logging.basicConfig(filename=log_file, level=logging.DEBUG)

@dataclass
class User:
    id: str
    jira_id: str
    email: str
    state: bool
    ticket_count: int
    weight: int

    def update_state(self, active_users):
        self.state = self.id in active_users

    def assign_ticket(self, ticket_id, users, user_to_assign):
        # Check if user_to_assign's state is True before assigning ticket
        if not user_to_assign.state:
            return False

        ticket_url = os.getenv('UNASSIGN_URL')
        url = f"{ticket_url}{ticket_id}/assignee"
        headers = {
            "Accept": "application/json",
            "Content-Type": "application/json"
        }
        payload = json.dumps({"accountId": user_to_assign.jira_id})
        response = requests.put(url, headers=headers, data=payload, auth=Application.jira_oauth(), verify=False)
        output = f'{ticket_id} assigned to {user_to_assign.id}'
        user_to_assign.weight += 1
        Application.save_weights(users)
        return output

    @staticmethod
    def get_next_user_id():
        username = os.getenv("FINESSE_USERNAME")
        password = os.getenv("FINESSE_PASS")
        url = os.getenv("FINESSE_API")
        path = os.getenv('CONFIG_PATH')
        response_path = os.path.join(path, 'response.xml')
        excluded_users_path = os.path.join(path, 'excluded_users.json')

        # Load the excluded users
        with open(excluded_users_path, 'r') as f:
            excluded_users = json.load(f)
        excluded_usernames = [user['user'] for user in excluded_users]


        response = requests.get(url, auth=HTTPBasicAuth(username, password), verify=False)
        if response.status_code == 200:
            user_list = []
            with open(response_path, 'w') as f:
                f.write(response.text)
            root = ET.fromstring(response.text)
            for user in root.findall('User'):
                state = user.find('state').text
                label_value = None
                loginId = user.find('loginId').text
                reasonCode = user.find('reasonCode')
                if reasonCode is not None:
                    label = reasonCode.find('label')
                    if label is not None:
                        label_value = label.text
                # Checks and makes sure the user is not logged out of Finesse and is not at lunch
                if state != "LOGOUT" and label_value != "Lunch" and loginId not in excluded_usernames:
                    user_list.append(loginId)
            return set(user_list)
        else:
            print(f'Error: {response.status_code}')

path = os.getenv('CONFIG_PATH')
users_path = os.path.join(path, 'users.json')
with open(users_path, 'r') as f:
    users_data = json.load(f)

users = [User(**user) for user in users_data]

def save_users(users):
    users_data = [asdict(user) for user in users]
    path = os.getenv('CONFIG_PATH')
    users_path = os.path.join(path, 'users.json')
    with open(users_path, 'r') as f:
        json.dump(users_data, f)

def update_user_states(users):
    while True:
        active_users = User.get_next_user_id()
        for user in users:
            user.update_state(active_users)

threading.Thread(target=update_user_states, args=(users,), daemon=True).start()

class ErrorTicker(QScrollArea):
    def __init__(self, text):
        super().__init__()
        self.setFrameShape(QFrame.NoFrame)
        self.setWidgetResizable(True)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        content = QWidget(self)
        self.setWidget(content)
        layout = QVBoxLayout(content)

        self.label = QLabel(text, self)
        self.label.setFont(QFont("Arial", 12))
        self.label.setStyleSheet("color: red")
        layout.addWidget(self.label)

        self.animation = QPropertyAnimation(self.label, b"pos")
        self.animation.setDuration(30000)

        # Adjust the start and end values of the animation
        self.animation.setStartValue(QPoint(-self.label.width(), 0))
        self.animation.setEndValue(QPoint(self.width() + self.label.width() + 500, 0))
        self.animation.setEasingCurve(QEasingCurve.Linear)
        self.animation.setLoopCount(-1)
        self.animation.start()

    def set_text(self, text):
        self.label.setText(text)
        self.label.adjustSize()
        self.animation.setStartValue(QPoint(-self.label.width(), 0))
        self.animation.setEndValue(QPoint(self.width() + self.label.width(), 0))

class ErrorTextDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Update Error Text")

        # Disable the WindowContextHelpButtonHint flag
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)

        self.text_field = QLineEdit(self)
        self.ok_button = QPushButton("OK", self)
        self.ok_button.clicked.connect(self.accept)

        layout = QVBoxLayout(self)
        layout.addWidget(self.text_field)
        layout.addWidget(self.ok_button)

    def get_text(self):
        return self.text_field.text()

class AboutSection(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("About")
        icon_path = os.getenv('ICON_PATH')
        self.setWindowIcon(QIcon(icon_path))
        self.createMenuBar()
        self.resize(830, 645)

        # Create a QVBoxLayout
        layout = QVBoxLayout()

        # Create a QLabel for the text
        text_label = QLabel(self)
        text_label.setWordWrap(True)

        # Set the QLabel text
        text_label.setText("""
        <h1>About LIA</h1>
        <h2>This program was written to solve the voicemail ticket creation and unassigned tickets issue. It has grown into a GUI that displays techincal support data/information on a TV.</h2>
        <h2>It pulls live data from Jira and Cisco Finesse. The primary functions are the Auto Unassign and Auto FLS.</h2>
                           
        <h3>Auto Unassign</h3>
        <ul>
            <li>Auto Unassign
            <ul>
                <li>Sub-bullet 1</li>
                <li>Sub-bullet 2</li>
            </ul>              
            </li>
            
        </ul>
                           
        <h3>Auto FLS</h3>
        <ul>
            <li>Auto FLS
            <ul>
                <li>Sub-bullet 1</li>
                <li>Sub-bullet 2</li>
            </ul>               
            </li>
        </ul>
        """)

        # Add the QLabel to the layout
        layout.addWidget(text_label)

        # Create a QWidget and set the layout to it
        widget = QWidget()
        widget.setLayout(layout)

        # Create a QScrollArea and set the widget to it
        scroll_area = QScrollArea()
        scroll_area.setWidget(widget)
        scroll_area.setWidgetResizable(True)
        scroll_area.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        # Set the QScrollArea as the central widget
        self.setCentralWidget(scroll_area)

    def overview(self):
        # Create a QVBoxLayout
        layout = QVBoxLayout()

        # Create a QLabel for the text
        text_label = QLabel(self)

        # Set the QLabel text
        text_label.setText("""
        <h1>Directions</h1>
        <h3>How to obtain Jira ID</h3>
        <ol>
            <li>Assign the new user a ticket or find one they are already assigned to.</li>
            <li>Click on the ellipsis in the top right hand corner. (Three dots)</li>
            <li>Click on "Export XML".</li>
            <li>In the window that opens, find the "assignee accountid" field. This string of numbers and letters is their Jira ID.</li>
        </ol>
        <h3>Reference screenshots below for assistance</h3>
        """)

        # Add the QLabel to the layout
        layout.addWidget(text_label)

        # Create a QWidget and set the layout to it
        widget = QWidget()
        widget.setLayout(layout)

        # Create a QScrollArea and set the widget to it
        scroll_area = QScrollArea()
        scroll_area.setWidget(widget)
        scroll_area.setWidgetResizable(True)
        scroll_area.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        # Set the QScrollArea as the central widget
        self.setCentralWidget(scroll_area)

    def usage(self):
        # Create a QVBoxLayout
        layout = QVBoxLayout()

        # Create a QLabel for the text
        text_label = QLabel(self)

        # Set the QLabel text
        text_label.setText("""
        <h1>Directions</h1>
        <h3>How to obtain Jira ID</h3>
        <ol>
            <li>Assign the new user a ticket or find one they are already assigned to.</li>
            <li>Click on the ellipsis in the top right hand corner. (Three dots)</li>
            <li>Click on "Export XML".</li>
            <li>In the window that opens, find the "assignee accountid" field. This string of numbers and letters is their Jira ID.</li>
        </ol>
        <h3>Reference screenshots below for assistance</h3>
        """)

        # Add the QLabel to the layout
        layout.addWidget(text_label)

        # Create a QWidget and set the layout to it
        widget = QWidget()
        widget.setLayout(layout)

        # Create a QScrollArea and set the widget to it
        scroll_area = QScrollArea()
        scroll_area.setWidget(widget)
        scroll_area.setWidgetResizable(True)
        scroll_area.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        # Set the QScrollArea as the central widget
        self.setCentralWidget(scroll_area)

    def troubleshooting(self):
        # Create a QVBoxLayout
        layout = QVBoxLayout()

        # Create a QLabel for the text
        text_label = QLabel(self)

        # Set the QLabel text
        text_label.setText("""
        <h1>Directions</h1>
        <h3>How to obtain Jira ID</h3>
        <ol>
            <li>Assign the new user a ticket or find one they are already assigned to.</li>
            <li>Click on the ellipsis in the top right hand corner. (Three dots)</li>
            <li>Click on "Export XML".</li>
            <li>In the window that opens, find the "assignee accountid" field. This string of numbers and letters is their Jira ID.</li>
        </ol>
        <h3>Reference screenshots below for assistance</h3>
        """)

        # Add the QLabel to the layout
        layout.addWidget(text_label)

        # Create a QWidget and set the layout to it
        widget = QWidget()
        widget.setLayout(layout)

        # Create a QScrollArea and set the widget to it
        scroll_area = QScrollArea()
        scroll_area.setWidget(widget)
        scroll_area.setWidgetResizable(True)
        scroll_area.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        # Set the QScrollArea as the central widget
        self.setCentralWidget(scroll_area)

    def createMenuBar(self):
        menuBar = self.menuBar()
        self.setMenuBar(menuBar)

        usageAction = QAction("Usage", self)
        usageAction.triggered.connect(self.usage)
        menuBar.addAction(usageAction)

        overviewAction = QAction("Overview", self)
        overviewAction.triggered.connect(self.overview)
        menuBar.addAction(overviewAction)

        troubleshootAction = QAction("Troubleshooting", self)
        troubleshootAction.triggered.connect(self.troubleshooting)
        menuBar.addAction(troubleshootAction)


class UserManagementWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Manage Users")
        icon_path = os.getenv('ICON_PATH')
        self.setWindowIcon(QIcon(icon_path))
        self.createMenuBar()
        self.resize(830, 645)


    def directions(self):
        # Create a QVBoxLayout
        layout = QVBoxLayout()

        # Create a QLabel for the text
        text_label = QLabel(self)

        # Set the QLabel text
        text_label.setText("""
        <h1>Directions</h1>
        <h3>How to obtain Jira ID</h3>
        <ol>
            <li>Assign the new user a ticket or find one they are already assigned to.</li>
            <li>Click on the ellipsis in the top right hand corner. (Three dots)</li>
            <li>Click on "Export XML".</li>
            <li>In the window that opens, find the "assignee accountid" field. This string of numbers and letters is their Jira ID.</li>
        </ol>
        <h3>Reference screenshots below for assistance</h3>
        """)

        # Add the QLabel to the layout
        layout.addWidget(text_label)

        # List of image paths
        config_path = os.getenv("CONFIG_PATH")
        image_paths = [os.path.join(config_path, "step1.png"), 
                       os.path.join(config_path, "step2.png"), 
                       os.path.join(config_path, "step3.png")]

        # Loop over the image paths and create QLabel and QPixmap for each
        for image_path in image_paths:
            image_label = QLabel(self)
            pixmap = QPixmap(image_path)
            image_label.setPixmap(pixmap)
            image_label.setAlignment(Qt.AlignCenter)
            layout.addWidget(image_label)

        # Create a QWidget and set the layout to it
        widget = QWidget()
        widget.setLayout(layout)

        # Create a QScrollArea and set the widget to it
        scroll_area = QScrollArea()
        scroll_area.setWidget(widget)
        scroll_area.setWidgetResizable(True)
        scroll_area.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        # Set the QScrollArea as the central widget
        self.setCentralWidget(scroll_area)

    @staticmethod
    def is_valid_email(email):
        return re.match(r"[^@]+@mhc\.com", email)

    def add_submit(self):
        # Get the text from each QLineEdit
        first_name = self.first_name_field.text()
        last_name = self.last_name_field.text()
        name = first_name + " " + last_name
        id = self.id_field.text()
        jira_id = self.jira_id_field.text()
        email = self.email_field.text()

        # Check if any fields are empty
        if not name.strip() or not id.strip() or not jira_id.strip() or not email.strip():
            error_box = QMessageBox()
            error_box.setIcon(QMessageBox.Warning)
            error_box.setWindowTitle("Empty Field")
            error_box.setText("One or more fields are empty. Please fill in all fields.")
            error_box.setStandardButtons(QMessageBox.Ok)
            error_box.exec_()
            return
        
        # Check if first name and last name fields contain only letters
        if not first_name.isalpha() or not last_name.isalpha():
            error_box = QMessageBox()
            error_box.setIcon(QMessageBox.Warning)
            error_box.setWindowTitle("Invalid Name")
            error_box.setText("First name and last name can only contain letters. Please try again.")
            error_box.setStandardButtons(QMessageBox.Ok)
            error_box.exec_()
            return

        # Email verification
        if not self.is_valid_email(self.email_field.text()):
            error_box = QMessageBox()
            error_box.setIcon(QMessageBox.Warning)
            error_box.setWindowTitle("Invalid Email")
            error_box.setText("The entered email address is not valid. Please try again.")
            error_box.setStandardButtons(QMessageBox.Ok)
            error_box.exec_()
            return

        # Create a QMessageBox for the confirmation dialog
        confirmation_box = QMessageBox()
        confirmation_box.setIcon(QMessageBox.Question)
        confirmation_box.setWindowTitle("Confirm Add")
        confirmation_box.setText(f"Are you sure you want to add user {id}?")
        confirmation_box.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
        confirmation_box.setDefaultButton(QMessageBox.No)

        # Show the confirmation dialog and get the user's response
        response = confirmation_box.exec_()

        # Check the user's response
        if response == QMessageBox.Yes:

            # Loading users.json
            path = os.getenv('CONFIG_PATH')
            user_path = os.path.join(path, 'users.json')
            with open(user_path, 'r') as f:
                data = json.load(f)

            # Check if the user already exists
            if any(user['id'] == id for user in data):
                # Show a QMessageBox to indicate that the user already exists
                error_box = QMessageBox()
                error_box.setIcon(QMessageBox.Warning)
                error_box.setWindowTitle("User Exists")
                error_box.setText(f"User {id} already exists.")
                error_box.setStandardButtons(QMessageBox.Ok)
                error_box.exec_()
            else:
                # Adding user to user_tickets.json file
                path = os.getenv('LOGS_PATH')
                user_tickets_path = os.path.join(path, 'user_tickets.json')
                with open(user_tickets_path, 'r') as f:
                    data = json.load(f)

                data[id] = {"day": 0, "total": 0}

                with open(user_tickets_path, 'w') as f:
                    json.dump(data, f)

                # Adding user to user_names.json
                path = os.getenv('CONFIG_PATH')
                user_names_path = os.path.join(path, 'user_names.json')
                with open(user_names_path, 'r') as f:
                    data = json.load(f)

                data[id] = name

                with open(user_names_path, 'w') as f:
                    json.dump(data, f)

                # Adding user to weights.json
                path = os.getenv('LOGS_PATH')
                user_weights_path = os.path.join(path, 'weights.json')
                with open(user_weights_path, 'r') as f:
                    data = json.load(f)

                data[id] = 0

                with open(user_weights_path, 'w') as f:
                    json.dump(data, f)

                # Adding user to users.json
                path = os.getenv('CONFIG_PATH')
                users_path = os.path.join(path, 'users.json')
                with open(users_path, 'r') as f:
                    data = json.load(f)

                # Create a new user dictionary
                new_user = {"id": id, "jira_id": jira_id, "email": email, "state": False, "ticket_count": 0, "weight": 0}

                # Add the new user to the data
                data.append(new_user)

                with open(users_path, 'w') as f:
                    json.dump(data, f)

                # Show a QMessageBox to confirm that the user has been added
                confirmation_box = QMessageBox()
                confirmation_box.setIcon(QMessageBox.Information)
                confirmation_box.setWindowTitle("User Added")
                confirmation_box.setText(f"Added user {id}")
                confirmation_box.setStandardButtons(QMessageBox.Ok)
                confirmation_box.exec_()
                # Clear the QLineEdit fields
                self.first_name_field.clear()
                self.last_name_field.clear()
                self.id_field.clear()
                self.jira_id_field.clear()
                self.email_field.clear()

        else:
            error_box = QMessageBox()
            error_box.setIcon(QMessageBox.Warning)
            error_box.setWindowTitle("Cancelled")
            error_box.setText("Cancelled User Add")
            error_box.setStandardButtons(QMessageBox.Ok)
            error_box.exec_()
            
            # Clear the QLineEdit fields
            self.first_name_field.clear()
            self.last_name_field.clear()
            self.id_field.clear()
            self.jira_id_field.clear()
            self.email_field.clear()

    def delete_submit(self):
        # Get the text from QLineEdit
        id = self.id_field.text()

        # Check if any fields are empty
        if not id.strip():
            error_box = QMessageBox()
            error_box.setIcon(QMessageBox.Warning)
            error_box.setWindowTitle("Empty Field")
            error_box.setText("ID field is empty. Please enter a valid User ID.")
            error_box.setStandardButtons(QMessageBox.Ok)
            error_box.exec_()
            return

        # Loading users.json
        path = os.getenv('CONFIG_PATH')
        user_path = os.path.join(path, 'users.json')
        with open(user_path, 'r') as f:
            data = json.load(f)

        if any(user['id'] == id for user in data):
            # Show a QMessageBox to indicate that the user already exists
            error_box = QMessageBox()
            error_box.setIcon(QMessageBox.Warning)
            error_box.setWindowTitle("User Exists")
            error_box.setText(f"User {id} exists. Click Ok to proceed.")
            error_box.exec_()

            # Create a QMessageBox for the confirmation dialog
            confirmation_box = QMessageBox()
            confirmation_box.setIcon(QMessageBox.Question)
            confirmation_box.setWindowTitle("Confirm Delete")
            confirmation_box.setText(f"Are you sure you want to delete user {id}?")
            confirmation_box.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
            confirmation_box.setDefaultButton(QMessageBox.No)

            # Show the confirmation dialog and get the user's response
            response = confirmation_box.exec_()

            # Check the user's response 
            if response == QMessageBox.No:
                error_box = QMessageBox()
                error_box.setIcon(QMessageBox.Warning)
                error_box.setWindowTitle("Not Deleted")
                error_box.setText(f"User {id} not deleted.")
                error_box.setStandardButtons(QMessageBox.Ok)
                error_box.exec_()
                self.id_field.clear()

            # Check the user's response
            if response == QMessageBox.Yes:
                config_path = os.getenv('CONFIG_PATH')
                user_path = os.path.join(config_path, 'users.json')
                user_names_path = os.path.join(config_path, 'user_names.json')
                excluded_path = os.path.join(config_path, 'excluded_users.json')

                logs_path = os.getenv('LOGS_PATH')
                weights_path = os.path.join(logs_path, 'weights.json')
                user_tickets_path = os.path.join(logs_path, 'user_tickets.json')

                # Define the JSON files to modify
                json_files = [user_tickets_path, weights_path, user_names_path, user_path]

                # Remove the user from user_tickets.json, weights.json, and user_names.json
                for file in json_files:
                    with open(file, 'r') as f:
                        data = json.load(f)

                    if id in data:
                        del data[id]

                    with open(file, 'w') as f:
                        json.dump(data, f)

                # Remove the user from users.json
                with open(users_path, 'r') as f:
                    data = json.load(f)

                data = [user for user in data if user['id'] != id]

                with open(users_path, 'w') as f:
                    json.dump(data, f)

                # Add the deleted user to the excluded_users.json file
                with open(excluded_path, 'r') as f:
                    data = json.load(f)

                # Create a new user dictionary
                new_user = {"user": id}

                # Add the new user to the data
                data.append(new_user)

                # Write the data back to the file
                with open(excluded_path, 'w') as f:
                    json.dump(data, f)

                error_box = QMessageBox()
                error_box.setIcon(QMessageBox.Warning)
                error_box.setWindowTitle("Deleted")
                error_box.setText(f"User {id} deleted from the system.")
                error_box.setStandardButtons(QMessageBox.Ok)
                error_box.exec_()
                self.id_field.clear()

        else:
            # If the user does not exist, prompt for a new ID
            error_box = QMessageBox()
            error_box.setIcon(QMessageBox.Warning)
            error_box.setWindowTitle("User Does Not Exist")
            error_box.setText(f"User {id} does not exist. Please enter a valid User ID.")
            error_box.setStandardButtons(QMessageBox.Ok)
            error_box.exec_()
            self.id_field.clear()
            return
        

    def createMenuBar(self):
        menuBar = self.menuBar()
        self.setMenuBar(menuBar)

        directionsAction = QAction("Directions", self)
        directionsAction.triggered.connect(self.directions)
        menuBar.addAction(directionsAction)

        # Manage Users
        manageUsersAction = QAction("View/Edit Users", self)
        manageUsersAction.triggered.connect(self.manage_users)
        menuBar.addAction(manageUsersAction)

        # Add and remove options under Add/Remove Users
        addUserAction = QAction("Add User", self)
        addUserAction.triggered.connect(self.add_user)
        menuBar.addAction(addUserAction)

        delUserAction = QAction("Delete User", self)
        delUserAction.triggered.connect(self.delete_user)
        menuBar.addAction(delUserAction)

    def add_user(self):
        self.first_name_label = QLabel("First Name:")
        self.first_name_field = QLineEdit()
        self.first_name_field.setFixedWidth(250)

        self.last_name_label = QLabel("Last Name:")
        self.last_name_field = QLineEdit()
        self.last_name_field.setFixedWidth(250)

        self.id_label = QLabel("Username: (Active Directory Username)")
        self.id_field = QLineEdit()
        self.id_field.setFixedWidth(250)

        self.jira_id_label = QLabel("JIRA ID: (See directions to obtain)")
        self.jira_id_field = QLineEdit()
        self.jira_id_field.setFixedWidth(250)

        self.email_label = QLabel("Email: (Active Directory email)")
        self.email_field = QLineEdit()
        self.email_field.setFixedWidth(250)

        # Create a QPushButton for the submit action
        self.submit_button = QPushButton("Submit")
        self.submit_button.setFixedWidth(250)
        self.submit_button.clicked.connect(self.add_submit)


        # Arrange widgets in a vertical layout
        self.layout = QVBoxLayout()
        self.layout.addWidget(self.first_name_label)
        self.layout.addWidget(self.first_name_field)
        self.layout.addWidget(self.last_name_label)
        self.layout.addWidget(self.last_name_field)
        self.layout.addWidget(self.id_label)
        self.layout.addWidget(self.id_field)
        self.layout.addWidget(self.jira_id_label)
        self.layout.addWidget(self.jira_id_field)
        self.layout.addWidget(self.email_label)
        self.layout.addWidget(self.email_field)
        self.layout.addWidget(self.submit_button)
        self.layout.setAlignment(Qt.AlignCenter)

        # Create a central widget for the window and set the layout
        self.central_widget = QWidget()
        self.central_widget.setLayout(self.layout)
        self.setCentralWidget(self.central_widget)

        spacer = QWidget()
        self.layout.addWidget(spacer)
        self.layout.setStretchFactor(spacer, 1)

    def delete_user(self):
        self.id_label = QLabel("Enter username to delete: (Active Directory username)")
        self.id_field = QLineEdit()
        self.id_field.setFixedWidth(250)

        # Create a QPushButton for the submit action
        self.submit_button = QPushButton("Submit")
        self.submit_button.setFixedWidth(250)
        self.submit_button.clicked.connect(self.delete_submit)

        # Arrange widgets in a vertical layout
        self.layout = QVBoxLayout()
        self.layout.addWidget(self.id_label)
        self.layout.addWidget(self.id_field)
        self.layout.addWidget(self.submit_button)
        self.layout.setAlignment(Qt.AlignCenter)

        # Create a central widget for the window and set the layout
        self.central_widget = QWidget()
        self.central_widget.setLayout(self.layout)
        self.setCentralWidget(self.central_widget)

        spacer = QWidget()
        self.layout.addWidget(spacer)
        self.layout.setStretchFactor(spacer, 1)

    def manage_users(self):
        path = os.getenv('CONFIG_PATH')
        user_path = os.path.join(path, 'users.json')
        with open(user_path, 'r') as f:
            data = json.load(f)

        # Load the id-to-name mapping from a JSON file
        name_path = os.path.join(path, 'user_names.json')
        with open(name_path, 'r') as f:
            id_to_name = json.load(f)

        # Create a QTableWidget and set its column count
        self.table = QTableWidget()

        # Set the size policy of the table to be expanding
        self.table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        # Enable selection of rows and set edit triggers
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.DoubleClicked | QAbstractItemView.SelectedClicked)

        # Exclude specific keys
        keys = [key for key in data[0].keys() if key not in ['state', 'ticket_count', 'weight']]
        keys = ['Full Name'] + keys
        self.table.setRowCount(len(data))
        self.table.setColumnCount(len(keys))
        self.table.setHorizontalHeaderLabels(keys)

        # Add the data to the table
        for i, user in enumerate(data):
            # Add data to the extra column based on the id
            extra_data = id_to_name.get(user['id'], 'No name found')
            item = QTableWidgetItem(extra_data)
            item.setData(Qt.UserRole, user)
            self.table.setItem(i, 0, item)
            for j, key in enumerate(keys[1:]):
                self.table.setItem(i, j + 1, QTableWidgetItem(str(user.get(key, ""))))

        # Resize the columns to fit their content
        self.table.resizeColumnsToContents()

        layout = QGridLayout()
        layout.addWidget(self.table, 1, 1)
        layout.setRowStretch(1, 2)
        layout.setColumnStretch(1, 6)
        layout.setRowStretch(0, 1)
        layout.setRowStretch(2, 1)
        layout.setColumnStretch(0, 1)
        layout.setColumnStretch(2, 1)

        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

class JiraThread(QThread):
    total_signal = pyqtSignal(int)
    resolved_signal = pyqtSignal(int)
    type_signal = pyqtSignal(str)
    percent_signal = pyqtSignal(int)
    percent_color_signal = pyqtSignal(str)

    def __init__(self, app, parent=None):
        super(JiraThread, self).__init__(parent)
        self.app = app

    def run(self):
        self.timer = QTimer()
        self.timer.timeout.connect(self.app.start_jira_thread)
        self.timer.start(100)
        total = self.app.jira_tickets_total()
        self.total_signal.emit(total)
        resolved = self.app.jira_tickets_resolved_fn()
        self.resolved_signal.emit(resolved)
        ticket_type_str = self.app.jira_tickets_type_fn()
        self.type_signal.emit(str(ticket_type_str))
        percent_resolved = self.app.jira_percent_fn()
        color_ranges = [
            (95, "#008000"),
            (90, "#32CD32"),
            (85, "#9ACD32"),
            (80, "#FFD700"),
            (75, "#FFA500"),
            (70, "#FF4500"),
            (0, "#FF0000")
        ]
        color = next((color for limit, color in color_ranges if percent_resolved >= limit), "#FF0000")
        self.percent_signal.emit(percent_resolved)
        self.percent_color_signal.emit(color)

class NewDayThread(QThread):
    update_date_time_signal = pyqtSignal()
    reset_user_uq_day_signal = pyqtSignal()
    reset_uq_day_signal = pyqtSignal()
    reset_weights_signal = pyqtSignal()
    reset_vm_count_signal =pyqtSignal()
    update_vm_counter_signal = pyqtSignal()
    update_ticket_count_signal = pyqtSignal()

    def __init__(self, parent=None):
        super(NewDayThread, self).__init__(parent)

    def run(self):
        self.update_date_time_signal.emit()
        self.reset_user_uq_day_signal.emit()
        self.reset_uq_day_signal.emit()
        self.reset_weights_signal.emit()
        self.reset_vm_count_signal.emit()
        self.update_vm_counter_signal.emit()
        self.update_ticket_count_signal.emit()

class NewWeekThread(QThread):
    reset_week_signal = pyqtSignal()
    reset_vm_week_signal = pyqtSignal()

    def __init__(self, parent=None):
        super(NewWeekThread, self).__init__(parent)

    def run(self):
        self.reset_week_signal.emit()
        self.reset_vm_week_signal.emit()

class NewMonthThread(QThread):
    reset_month_signal = pyqtSignal()
    reset_vm_month_signal = pyqtSignal()

    def __init__(self, parent=None):
        super(NewMonthThread, self).__init__(parent)

    def run(self):
        self.reset_month_signal.emit()
        self.reset_vm_month_signal.emit()

class Worker(QThread):
    update_status_signal = pyqtSignal(str)
    update_color_signal = pyqtSignal(str)
    update_ticket_count_signal = pyqtSignal(list)
    update_ticket_counts_day_signal = pyqtSignal(list)
    update_total_unassigned_signal = pyqtSignal(list)
    update_total_unassigned_day_signal = pyqtSignal(list)
    update_total_unassigned_week_signal = pyqtSignal(list)
    update_total_unassigned_month_signal = pyqtSignal(list)
    update_last_assigned_signal = pyqtSignal(str)
    update_next_assignee_signal = pyqtSignal(str)
    update_vm_ticket_signal = pyqtSignal(str)
    update_vm_signal = pyqtSignal(str)
    update_vm_counts_signal = pyqtSignal(dict)
    vm_status_color_signal = pyqtSignal(str)
    
    def __init__(self, func):
        QThread.__init__(self)
        self.func = func

    def run(self):
        self.func()

    def update_status(self, text):
        self.update_status_signal.emit(text)

    def update_color(self, color):
        self.update_color_signal.emit(color)

    # User Unassigned count functions
    def update_ticket_count(self, user_counts):
        self.update_ticket_count_signal.emit(user_counts)
    
    def update_ticket_counts_day(self, user_counts_day):
        self.update_ticket_counts_day_signal.emit(user_counts_day)

    def update_ticket_counts_week(self, user_counts_week):
        self.update_ticket_counts_day_signal.emit(user_counts_week)

    def update_ticket_counts_month(self, user_counts_month):
        self.update_ticket_counts_day_signal.emit(user_counts_month)

    # Total unassigned count functions
    def update_total_unassigned(self, ticket_counts):
        self.update_total_unassigned_signal.emit(ticket_counts)
    
    def update_total_unassigned_day(self, counts_day):
        self.update_total_unassigned_day_signal.emit(counts_day)

    def update_total_unassigned_week(self, counts_week):
        self.update_total_unassigned_week_signal.emit(counts_week)
    
    def update_total_unassigned_month(self, counts_month):
        self.update_total_unassigned_month_signal.emit(counts_month)

    def update_vm_counts(self, counts):
        self.update_vm_tickets_signal.emit(counts)
    
    def update_vm_status_color(self, color):
        self.vm_status_color_signal.emit(color)
    
class Application(QMainWindow):

    # Auto FLS slots
    @pyqtSlot(dict)
    def update_vm_counts(self, counts):
        text = f"Day: {counts['Day']}\nWeek: {counts['Week']}\nMonth: {counts['Month']}\nTotal: {counts['Total']}"
        self.vm_counts.clear()
        self.vm_counts.setText(text)

    @pyqtSlot(str)
    def update_vm_ticket(self, text):
        self.vm_ticket.clear()
        self.vm_ticket.setText(text)

    @pyqtSlot(str)
    def update_vm(self, text):
        self.voicemail.clear()
        self.voicemail.setText(text)

    # Unassigned queue slots
    @pyqtSlot(str)
    def update_next_assignee(self, assignee):
        self.next_assignee.clear()
        self.next_assignee.setText(assignee)

    @pyqtSlot(str)
    def update_last_assigned(self, text):
        self.last_assigned.clear()
        self.last_assigned.setText(text)

    @pyqtSlot(list)
    # User level unassigned queue counts
    def update_ticket_counts(self, ticket_counts):
        self.ticket_counts_total.clear()
        self.ticket_counts_total.setText('\n'.join(ticket_counts))

    @pyqtSlot(list)
    # User level unassigned queue counts
    def update_ticket_counts_day(self, ticket_counts_day):
        self.ticket_counts_day.clear()
        self.ticket_counts_day.setText('\n'.join(ticket_counts_day))

    # User level unassigned queue counts
    @pyqtSlot(list)
    def update_ticket_counts_week(self, user_counts_week):
        self.week_box.clear()
        self.week_box.setText('\n'.join(user_counts_week))

    # User level unassigned queue counts
    @pyqtSlot(list)
    def update_ticket_counts_month(self, user_counts_month):
        self.month_box.clear()
        self.month_box.setText('\n'.join(user_counts_month))
    
    @pyqtSlot(list)
    # Total level unassigned queue counts-TOTAL
    def update_total_unassigned(self, unassign_total):
        self.total_box.clear()
        self.total_box.setText('\n'.join(unassign_total))

    # Total level unassigned queue counts- DAY
    @pyqtSlot(list)
    def update_total_unassigned_day(self, unassigned_day):
        self.day_box.clear()
        self.day_box.setText('\n'.join(unassigned_day))

    # Total level unassigned queue counts- WEEK
    @pyqtSlot(list)
    def update_total_unassigned_week(self, unassigned_week):
        self.week_box.clear()
        self.week_box.setText('\n'.join(unassigned_week))

    # Total level unassigned queue counts- MONTH
    @pyqtSlot(list)
    def update_total_unassigned_month(self, unassigned_month):
        self.month_box.clear()
        self.month_box.setText('\n'.join(unassigned_month))

    # Jira slots
    @pyqtSlot(int)
    def update_total(self, total):
        self.jira_tickets_created.setText(str(total))

    @pyqtSlot(int)
    def update_resolved(self, resolved):
        self.jira_tickets_resolved.setText(str(resolved))

    @pyqtSlot(str)
    def update_type(self, ticket_type_str):
        self.jira_tickets_type.setText(str(ticket_type_str))

    @pyqtSlot(int)
    def update_percent(self, percent):
        self.jira_percent.setText(str(percent) + "%")

    @pyqtSlot(str)
    def update_percent_color(self, color):
        self.jira_percent.setStyleSheet(f"background-color: {color};")

    @pyqtSlot(str)
    def update_vm_status_color(self, color):
        self.voicemail.setStyleSheet(f"background-color: {color};")

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Labor Distribution Engine")
        icon_path = os.getenv('ICON_PATH')
        self.setWindowIcon(QIcon(icon_path))
        self.widget = QFrame(self)
        self.setCentralWidget(self.widget)
        self.main_layout = QVBoxLayout(self.widget)
        self.layout = QHBoxLayout()
        self.main_layout.addLayout(self.layout)
        self.create_widgets()
        self.createMenuBar()
        self.uq_stopped = True
        self.fls_stopped = True
        self.start_siren_loop()
        self.start_fire_loop()
        self.refresh_finesse_status()
        self.refresh_user_queue()
        self.update_vm_count()
        self.update_ticket_count()
        self.start_jira_thread()
        self.loop_new_week_check()
        self.loop_new_month_check()
        self.loop_new_day_check()
        self.showMaximized()

        self.new_day_timer = QTimer()
        self.new_day_timer.timeout.connect(self.loop_new_day_check)
        self.new_day_timer.start(600000)

        self.jira_timer = QTimer()
        self.jira_timer.timeout.connect(self.start_jira_thread)
        self.jira_timer.start(300000)

        # Unassigned queue worker function
        def run_uq():
            while not self.uq_stopped:
                self.unassign_loop(users)

        # Creating the unassign queue worker fucntion and attaching the signals
        self.uq_worker = Worker(run_uq)
        self.uq_worker.update_status_signal.connect(self.update_uq_status)
        self.uq_worker.update_color_signal.connect(self.uq_update_color)
        self.uq_worker.update_ticket_count_signal.connect(self.update_ticket_counts)
        self.uq_worker.update_ticket_counts_day_signal.connect(self.update_ticket_counts_day)              
        self.uq_worker.update_total_unassigned_signal.connect(self.update_total_unassigned)
        self.uq_worker.update_total_unassigned_day_signal.connect(self.update_total_unassigned_day)
        self.uq_worker.update_total_unassigned_week_signal.connect(self.update_total_unassigned_week)
        self.uq_worker.update_total_unassigned_month_signal.connect(self.update_total_unassigned_month)
        self.uq_worker.update_last_assigned_signal.connect(self.update_last_assigned)
        self.uq_worker.update_next_assignee_signal.connect(self.update_next_assignee)

        # Define the worker function for auto_fls
        def run_fls():
            while not self.fls_stopped:
                self.auto_fls()

        # Create a Worker object for auto_fls
        self.fls_worker = Worker(run_fls)
        # Connect any necessary signals
        self.fls_worker.update_status_signal.connect(self.update_fls_status)
        self.fls_worker.update_color_signal.connect(self.fls_update_color)
        self.fls_worker.update_vm_ticket_signal.connect(self.update_vm_ticket)
        self.fls_worker.update_vm_signal.connect(self.update_vm)
        self.fls_worker.update_vm_counts_signal.connect(self.update_vm_counts)
        self.fls_worker.vm_status_color_signal.connect(self.update_vm_status_color)


        # Create a QToolBar
        self.toolbar = QToolBar("Main toolbar")
        self.addToolBar(Qt.TopToolBarArea, self.toolbar)
        # Set the font size
        font = QFont()
        font.setPointSize(13)

        # Unassign status label
        self.uq_status_label = QLabel()
        self.uq_status_label.setText("Unassign Status: ")
        self.toolbar.addWidget(self.uq_status_label)
        self.uq_status_label.setFont(font)

        # Add a spacer
        spacer1 = QWidget()
        spacer1.setFixedWidth(5)
        self.toolbar.addWidget(spacer1)

        self.uq_status = QLabel()
        self.uq_status.setText("Stopped")
        self.uq_status.setStyleSheet("background-color: red;")
        self.toolbar.addWidget(self.uq_status)
        self.uq_status.setFont(font)

        # Add a spacer
        spacer2 = QWidget()
        spacer2.setFixedWidth(10)
        self.toolbar.addWidget(spacer2)

        # FLS status label
        self.fls_status_label = QLabel()
        self.fls_status_label.setText("FLS Status: ")
        self.toolbar.addWidget(self.fls_status_label)
        self.fls_status_label.setFont(font)

        # Add a spacer
        spacer3 = QWidget()
        spacer3.setFixedWidth(5)
        self.toolbar.addWidget(spacer3)

        self.fls_status = QLabel()
        self.fls_status.setText("Stopped")
        self.fls_status.setStyleSheet("background-color: red;")
        self.toolbar.addWidget(self.fls_status)
        self.fls_status.setFont(font)

        # Add a stretchable spacer to push everything else to the right
        spacer4 = QWidget()
        spacer4.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.toolbar.addWidget(spacer4)

        # Create a QLabel for the date
        self.date_label = QLabel()
        self.date_label.setText(self.get_current_date().strftime("%B %d, %Y"))

        # Create a QLabel object
        self.time_label = QLabel()

        # Create a QTimer object
        timer = QTimer(self)
        timer.timeout.connect(self.showTime)
        timer.start(1000)

        # Set the font size
        font = QFont()
        font.setPointSize(20)  # or whatever size you prefer
        self.date_label.setFont(font)
        self.time_label.setFont(font)

        # Add the date QLabel to the toolbar
        self.toolbar.addWidget(self.date_label)

        # Add a spacer
        spacer5 = QWidget()
        spacer5.setFixedWidth(20)
        self.toolbar.addWidget(spacer5)

        # Add the time QLabel to the toolbar
        self.toolbar.addWidget(self.time_label)

        # Add a spacer
        spacer6 = QWidget()
        spacer6.setFixedWidth(30)
        self.toolbar.addWidget(spacer6)

        # MHC background logo
        background_image = os.getenv("BACKGROUND")
        pixmap = QPixmap(background_image)
        image = QImage(pixmap.size(), QImage.Format_ARGB32)
        image.fill(QColor(0, 0, 0, 0))
        painter = QPainter(image)
        painter.setOpacity(0.5)
        painter.drawPixmap(0, 0, pixmap)
        painter.end()
        pixmap = QPixmap.fromImage(image)
        self.pixmap = pixmap.scaled(400, 400, Qt.KeepAspectRatio)

    def start_new_day_thread(self):
        self.new_day_thread = NewDayThread(self)
        self.new_day_thread.update_date_time_signal.connect(self.date_time_updater)
        self.new_day_thread.reset_user_uq_day_signal.connect(self.user_day_reset)
        self.new_day_thread.reset_uq_day_signal.connect(self.uq_day_reset)
        self.new_day_thread.reset_weights_signal.connect(self.weights_reset)
        self.new_day_thread.reset_vm_count_signal.connect(self.vm_counter_day_reset)
        self.new_day_thread.update_vm_counter_signal.connect(self.update_vm_count)
        self.new_day_thread.update_ticket_count_signal.connect(self.update_ticket_count)
        self.new_day_thread.start()

    def start_new_week_thread(self):
        self.new_week_thread = NewWeekThread(self)
        self.new_week_thread.reset_week_signal.connect(self.uq_week_reset)
        self.new_week_thread.reset_vm_week_signal.connect(self.vm_counter_week_reset)
        self.new_week_thread.start()

    def start_new_month_thread(self):
        self.new_month_thread = NewMonthThread(self)
        self.new_month_thread.reset_month_signal.connect(self.uq_month_reset)
        self.new_month_thread.reset_vm_month_signal.connect(self.vm_counter_month_reset)
        self.new_month_thread.start()

    # Create a new JiraThread
    def start_jira_thread(self):
        self.jira_thread = JiraThread(self)
        # Connect the signals to slots
        self.jira_thread.total_signal.connect(self.update_total)
        self.jira_thread.resolved_signal.connect(self.update_resolved)
        self.jira_thread.type_signal.connect(self.update_type)
        self.jira_thread.percent_signal.connect(self.update_percent)
        self.jira_thread.percent_color_signal.connect(self.update_percent_color)
        # Start the thread
        self.jira_thread.start()

    def createMenuBar(self):
        menuBar = self.menuBar()
        self.setMenuBar(menuBar)

        # Auto Unassign Menu
        autoUnassignMenu = menuBar.addMenu("Auto Unassign")
        startUnassignAction = QAction("Start", self)
        startUnassignAction.triggered.connect(self.start_unassign)
        autoUnassignMenu.addAction(startUnassignAction)

        stopUnassignAction = QAction("Stop", self)
        stopUnassignAction.triggered.connect(self.stop_uq)
        autoUnassignMenu.addAction(stopUnassignAction)

        # Auto FLS Menu
        autoFLSMenu = menuBar.addMenu("Auto FLS")
        startFLSAction = QAction("Start", self)
        startFLSAction.triggered.connect(self.start_fls)
        autoFLSMenu.addAction(startFLSAction)

        stopFLSAction = QAction("Stop", self)
        stopFLSAction.triggered.connect(self.stop_fls)
        autoFLSMenu.addAction(stopFLSAction)

        reset_options = menuBar.addMenu("Reset Options")

        tech_counts_menu = reset_options.addMenu("Unassigned per Tech")
        total_counts_menu = reset_options.addMenu("Unassigned Totals")
        voicemail_counts_menu = reset_options.addMenu("Voicemail")
        weights_counts_menu = reset_options.addMenu("Weights")

        # Add actions to Tech Counts submenu
        user_day_action = QAction("Reset Day", self)
        user_day_action.triggered.connect(self.confirm_user_day_reset)
        tech_counts_menu.addAction(user_day_action)

        user_total_action = QAction("Reset Total", self)
        user_total_action.triggered.connect(self.confirm_user_total_reset)
        tech_counts_menu.addAction(user_total_action)

        user_all_action = QAction("Reset ALL", self)
        user_all_action.triggered.connect(self.confirm_user_all_reset)
        tech_counts_menu.addAction(user_all_action)

        # Add actions to Total Counts submenu
        unassign_day_action = QAction("Reset Day", self)
        unassign_day_action.triggered.connect(self.confirm_uq_day_reset)
        total_counts_menu.addAction(unassign_day_action)

        unassign_day_action = QAction("Reset Week", self)
        unassign_day_action.triggered.connect(self.confirm_uq_week_reset)
        total_counts_menu.addAction(unassign_day_action)

        unassign_day_action = QAction("Reset Month", self)
        unassign_day_action.triggered.connect(self.confirm_uq_month_reset)
        total_counts_menu.addAction(unassign_day_action)

        unassign_total_action = QAction("Reset Total", self)
        unassign_total_action.triggered.connect(self.confirm_uq_total_reset)
        total_counts_menu.addAction(unassign_total_action)

        unassign_all_action = QAction("Reset ALL", self)
        unassign_all_action.triggered.connect(self.confirm_uq_all_reset)
        total_counts_menu.addAction(unassign_all_action)

        #add actions to Weights submenu
        reset_weights_action = QAction("Reset Weights", self)
        reset_weights_action.triggered.connect(self.confirm_weights_reset)
        weights_counts_menu.addAction(reset_weights_action)

        # Add actions to Voicemail Counts submenu
        reset_vm_day = QAction("Reset Day", self)
        reset_vm_day.triggered.connect(self.confirm_vm_reset)
        voicemail_counts_menu.addAction(reset_vm_day)

        reset_vm_week = QAction("Reset Week", self)
        reset_vm_week.triggered.connect(self.confirm_vm_week_reset)
        voicemail_counts_menu.addAction(reset_vm_week)

        reset_vm_month = QAction("Reset Month", self)
        reset_vm_month.triggered.connect(self.confirm_vm_month_reset)
        voicemail_counts_menu.addAction(reset_vm_month)

        reset_vm_all = QAction("Reset All", self)
        reset_vm_all.triggered.connect(self.confirm_vm_all_reset)
        voicemail_counts_menu.addAction(reset_vm_all)

        # User management option
        manage_users = QAction("Manage Users", self)
        manage_users.triggered.connect(self.open_user_management)
        menuBar.addAction(manage_users)

        # About LIA Section
        manage_users = QAction("About", self)
        manage_users.triggered.connect(self.open_about_section)
        menuBar.addAction(manage_users)

        # Refresh users icon
        manage_users = QAction("Error Text", self)
        manage_users.triggered.connect(self.update_error_text)
        menuBar.addAction(manage_users)
        
        # Refresh users icon
        refresh = QAction(self)
        refresh_icon = os.getenv('REFRESH_ICON')
        refresh.setIcon(QIcon(refresh_icon))
        refresh.triggered.connect(self.update_ticket_count)
        menuBar.addAction(refresh)

    def update_error_text(self):
        dialog = ErrorTextDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            error_text = dialog.get_text()
            self.error_ticker.set_text(error_text)

    def open_user_management(self):
        self.second_window = UserManagementWindow()
        self.second_window.show()

    def open_about_section(self):
        self.about_window = AboutSection()
        self.about_window.show()

    def paintEvent(self, event):
        painter = QPainter(self)
        # Calculate the center point of the window
        center_x = (self.width() - self.pixmap.width()) // 2
        center_y = (self.height() - self.pixmap.height()) // 2
        # Shift the image up by subtracting from center_y
        shift_up = 75  # Change this value to shift the image more or less
        center_y -= shift_up
        painter.drawPixmap(center_x, center_y, self.pixmap)

    def create_widgets(self):
        # Font Size
        font = QFont()
        font.setPointSize(14)
        # Create the web_frame
        self.error_frame = QFrame()
        self.error_frame.setFrameShape(QFrame.NoFrame)
        self.error_frame.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.error_frame.setFixedHeight(40)
        self.error_frame.setFixedWidth(1000)

        self.error_ticker = ErrorTicker("")

        # Create a layout for the web_frame and add the QWebEngineView to it
        error_layout = QVBoxLayout(self.error_frame)
        error_layout.addWidget(self.error_ticker)

        # Add the web_frame to the QVBoxLayout
        self.main_layout.insertWidget(0, self.error_frame)
        self.main_layout.setAlignment(self.error_frame, Qt.AlignCenter)

        # Create the web_frame
        self.web_frame = QFrame()
        self.web_frame.setFrameShape(QFrame.StyledPanel)
        self.web_frame.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.web_frame.setFixedHeight(275)
        
        # Creating the web object
        self.web = QWebEngineView(self.web_frame)
        web_url = os.getenv("FINESSE_API3")
        self.web.load(QUrl(web_url))

        # Create a layout for the web_frame and add the QWebEngineView to it
        web_layout = QVBoxLayout(self.web_frame)
        web_layout.addWidget(self.web)

        # Add the web_frame to the QVBoxLayout
        self.main_layout.insertWidget(0, self.web_frame)

        # Create the siren frame
        siren_gif = os.getenv("SIREN")
        self.siren_frame = QFrame()
        self.siren_frame.setFrameShape(QFrame.NoFrame)
        self.siren_frame.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.siren_frame.setFixedHeight(150)

        # Create QLabel and QMovie objects for left and right GIFs
        self.siren_left = QLabel(self.siren_frame)
        self.movie_left = QMovie(siren_gif)
        self.siren_left.setMovie(self.movie_left)
        self.movie_left.start()

        self.siren_right = QLabel(self.siren_frame)
        self.movie_right = QMovie(siren_gif)
        self.siren_right.setMovie(self.movie_right)
        self.movie_right.start()

        # Create a QHBoxLayout
        siren_layout = QHBoxLayout(self.siren_frame)

        # Add QLabel to the QHBoxLayout
        siren_layout.addWidget(self.siren_left)
        siren_layout.addStretch()
        siren_layout.addWidget(self.siren_right)

        # Add the siren_frame to the main layout
        self.main_layout.addWidget(self.siren_frame)
        self.main_layout.insertWidget(1, self.siren_frame)

        # Voicemail frame Configuration
        self.vm_frame = QFrame()
        self.vm_frame.setFrameShape(QFrame.StyledPanel)
        self.layout.addWidget(self.vm_frame)
        vm_layout = QGridLayout(self.vm_frame)

        # Frame background color
        palette = QPalette()
        palette.setColor(QPalette.Background, QColor("#d8e3e6"))
        self.vm_frame.setPalette(palette)
        self.vm_frame.setAutoFillBackground(True)

        # Setting size and alignment
        self.vm_frame.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.layout.addWidget(self.vm_frame)
        self.vm_frame.setFixedHeight(500)
        self.layout.setAlignment(self.vm_frame, Qt.AlignBottom | Qt.AlignRight)

        # Voicemail frame Widgets
        self.voicemail_label = QLabel("Voicemail Queue Status")
        self.voicemail_label.setAlignment(Qt.AlignCenter)
        vm_layout.addWidget(self.voicemail_label)
        self.voicemail_label.setFont(font)
        
        self.voicemail = QTextEdit()
        self.voicemail.setFixedHeight(35)
        vm_layout.addWidget(self.voicemail)
        self.voicemail.setFont(font)

        self.vm_ticket_label = QLabel("Last Ticket Created")
        self.vm_ticket_label.setAlignment(Qt.AlignCenter)
        vm_layout.addWidget(self.vm_ticket_label)
        self.vm_ticket_label.setFont(font)

        self.vm_ticket = QTextEdit()
        self.vm_ticket.setFixedHeight(35)
        vm_layout.addWidget(self.vm_ticket)
        self.vm_ticket.setFont(font)

        # Font Size
        vm_font = QFont()
        vm_font.setPointSize(14)

        self.vm_counts_label = QLabel("Voicemail Tickets Created")
        self.vm_counts_label.setAlignment(Qt.AlignCenter)
        vm_layout.addWidget(self.vm_counts_label)
        self.vm_counts_label.setFont(font)

        self.vm_counts = QTextEdit()
        vm_layout.addWidget(self.vm_counts)
        self.vm_counts.setFixedHeight(110)
        self.vm_counts.setFont(vm_font)

        # Create a vertical spacer to push the widgets to the top of the frame
        spacer = QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding)
        vm_layout.addItem(spacer, 8, 0, 1, 2)

        # Status frame Configuration
        self.status_frame = QFrame()
        self.status_frame.setFrameShape(QFrame.StyledPanel)
        self.layout.addWidget(self.status_frame)
        status_layout = QGridLayout(self.status_frame)
        
        # Status Frame background color
        palette = QPalette()
        palette.setColor(QPalette.Background, QColor("#d8e3e6"))
        self.status_frame.setPalette(palette)
        self.status_frame.setAutoFillBackground(True)

        # Setting size and alignment
        self.status_frame.setFixedHeight(500)
        self.status_frame.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.layout.addWidget(self.status_frame)
        self.layout.setAlignment(self.status_frame, Qt.AlignBottom | Qt.AlignRight)
        
        # Status frame widgets
        self.current_queue_label = QLabel("Users in Queue")
        self.current_queue_label.setAlignment(Qt.AlignCenter)
        status_layout.addWidget(self.current_queue_label)
        self.current_queue_label.setFont(font)

        self.current_queue = QTextEdit()
        self.current_queue.setFixedHeight(225)
        status_layout.addWidget(self.current_queue)
        self.current_queue.setFont(font)

        self.last_assigned_label = QLabel("Last Assigned")
        self.last_assigned_label.setAlignment(Qt.AlignCenter)
        status_layout.addWidget(self.last_assigned_label)
        self.last_assigned_label.setFont(font)

        self.last_assigned = QTextEdit()
        self.last_assigned.setFixedHeight(55)
        status_layout.addWidget(self.last_assigned)
        self.last_assigned.setFont(font)

        self.next_assignee_label = QLabel("Next Assignee")
        self.next_assignee_label.setAlignment(Qt.AlignCenter)
        status_layout.addWidget(self.next_assignee_label)
        self.next_assignee_label.setFont(font)

        self.next_assignee = QTextEdit()
        self.next_assignee.setFixedHeight(35)
        status_layout.addWidget(self.next_assignee)
        self.next_assignee.setFont(font)

        # Create a vertical spacer to push the widgets to the top of the frame
        spacer = QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding)
        status_layout.addItem(spacer, 8, 0, 1, 2)

        # Finesse Status Frame Configuration
        self.finesse_frame = QFrame()
        self.finesse_frame.setFrameShape(QFrame.StyledPanel)
        self.layout.addWidget(self.finesse_frame)

        # Finesse background color
        palette = QPalette()
        palette.setColor(QPalette.Background, QColor("#d8e3e6"))
        self.finesse_frame.setPalette(palette)
        self.finesse_frame.setAutoFillBackground(True)

        # Finesse Frame Font Size
        f_font = QFont()
        f_font.setPointSize(16)

        finesse_layout = QGridLayout(self.finesse_frame)

        # Set size policy
        self.finesse_frame.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.finesse_frame.setFixedHeight(500)
        self.finesse_frame.setMaximumWidth(800)

        # Add the frame to the main layout
        self.layout.addWidget(self.finesse_frame)

        # Set alignment of the main layout to top left
        self.layout.setAlignment(self.finesse_frame, Qt.AlignBottom | Qt.AlignLeft)

        #Finesse Status Widgets
        self.tech_label = QLabel("Tech")
        self.tech_label.setAlignment(Qt.AlignCenter)
        finesse_layout.addWidget(self.tech_label, 1, 0)
        self.tech_label.setFont(f_font)
        
        self.tech1 = QTextEdit()
        finesse_layout.addWidget(self.tech1, 2, 0)
        self.tech1.setFont(f_font)

        self.tech2 = QTextEdit()
        finesse_layout.addWidget(self.tech2, 3, 0)
        self.tech2.setFont(f_font)

        self.tech3 = QTextEdit()
        finesse_layout.addWidget(self.tech3, 4, 0)
        self.tech3.setFont(f_font)

        self.tech4 = QTextEdit()
        finesse_layout.addWidget(self.tech4, 5, 0)
        self.tech4.setFont(f_font)

        self.tech5 = QTextEdit()
        finesse_layout.addWidget(self.tech5, 6, 0)
        self.tech5.setFont(f_font)

        self.tech6 = QTextEdit()
        finesse_layout.addWidget(self.tech6, 7, 0)
        self.tech6.setFont(f_font)

        self.tech7 = QTextEdit()
        finesse_layout.addWidget(self.tech7, 8, 0)
        self.tech7.setFont(f_font)

        self.tech8 = QTextEdit()
        finesse_layout.addWidget(self.tech8, 9, 0)
        self.tech8.setFont(f_font)

        self.tech9 = QTextEdit()
        finesse_layout.addWidget(self.tech9, 10, 0)
        self.tech9.setFont(f_font)

        self.status_label = QLabel("Status")
        self.status_label.setAlignment(Qt.AlignCenter)
        finesse_layout.addWidget(self.status_label, 1, 1)
        self.status_label.setFont(f_font)

        self.finesse1 = QTextEdit()
        finesse_layout.addWidget(self.finesse1, 2, 1)
        self.finesse1.setFont(f_font)

        self.finesse2 = QTextEdit()
        finesse_layout.addWidget(self.finesse2, 3, 1)
        self.finesse2.setFont(f_font)

        self.finesse3 = QTextEdit()
        finesse_layout.addWidget(self.finesse3, 4, 1)
        self.finesse3.setFont(f_font)

        self.finesse4 = QTextEdit()
        finesse_layout.addWidget(self.finesse4, 5, 1)
        self.finesse4.setFont(f_font)

        self.finesse5 = QTextEdit()
        finesse_layout.addWidget(self.finesse5, 6, 1)
        self.finesse5.setFont(f_font)

        self.finesse6 = QTextEdit()
        finesse_layout.addWidget(self.finesse6, 7, 1)
        self.finesse6.setFont(f_font)

        self.finesse7 = QTextEdit()
        finesse_layout.addWidget(self.finesse7, 8, 1)
        self.finesse7.setFont(f_font)

        self.finesse8 = QTextEdit()
        finesse_layout.addWidget(self.finesse8, 9, 1)
        self.finesse8.setFont(f_font)

        self.finesse9 = QTextEdit()
        finesse_layout.addWidget(self.finesse9, 10, 1)
        self.finesse9.setFont(f_font)

        self.reason_label = QLabel("Reason")
        self.reason_label.setAlignment(Qt.AlignCenter)
        finesse_layout.addWidget(self.reason_label, 1, 2)
        self.reason_label.setFont(f_font)

        self.reason1 = QTextEdit()
        finesse_layout.addWidget(self.reason1, 2, 2)
        self.reason1.setFont(f_font)

        self.reason2 = QTextEdit()
        finesse_layout.addWidget(self.reason2, 3, 2)
        self.reason2.setFont(f_font)

        self.reason3 = QTextEdit()
        finesse_layout.addWidget(self.reason3, 4, 2)
        self.reason3.setFont(f_font)
        
        self.reason4 = QTextEdit()
        finesse_layout.addWidget(self.reason4, 5, 2)
        self.reason4.setFont(f_font)

        self.reason5 = QTextEdit()
        finesse_layout.addWidget(self.reason5, 6, 2)
        self.reason5.setFont(f_font)

        self.reason6 = QTextEdit()
        finesse_layout.addWidget(self.reason6, 7, 2)
        self.reason6.setFont(f_font)

        self.reason7 = QTextEdit()
        finesse_layout.addWidget(self.reason7, 8, 2)
        self.reason7.setFont(f_font)

        self.reason8 = QTextEdit()
        finesse_layout.addWidget(self.reason8, 9, 2)
        self.reason8.setFont(f_font)

        self.reason9 = QTextEdit()
        finesse_layout.addWidget(self.reason9, 10, 2)
        self.reason9.setFont(f_font)

        # Jira Frame
        self.jira_frame = QFrame()
        self.jira_frame.setFrameShape(QFrame.StyledPanel)
        self.layout.addWidget(self.jira_frame)

        jira_layout = QGridLayout(self.jira_frame)

        # Jira frame background color
        palette = QPalette()
        palette.setColor(QPalette.Background, QColor("#d8e3e6"))
        self.jira_frame.setPalette(palette)
        self.jira_frame.setAutoFillBackground(True)

        # Set Max height of the frame
        self.jira_frame.setFixedHeight(500)
        self.jira_frame.setMaximumWidth(350)

        # Set size policy
        self.jira_frame.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        # Set alignment of the main layout to top left
        self.layout.setAlignment(self.jira_frame, Qt.AlignBottom | Qt.AlignLeft)

        # Adding the flame background when a certain amount of tickets is reached
        self.fire_label = QLabel(self.jira_frame)
        self.fire_label.setGeometry(0, 150, self.jira_frame.width(), self.jira_frame.height())
        gif_path = os.getenv("FIRE")
        fire_movie = QMovie(gif_path)
        self.fire_label.setMovie(fire_movie)
        fire_movie.start()
        self.fire_label.hide()

        self.jira_tickets_created_label = QLabel("Jira Tickets Created")
        self.jira_tickets_created_label.setAlignment(Qt.AlignCenter)
        jira_layout.addWidget(self.jira_tickets_created_label)
        self.jira_tickets_created_label.setFont(font)

        self.jira_tickets_created = QTextEdit()
        jira_layout.addWidget(self.jira_tickets_created)
        self.jira_tickets_created.setFixedHeight(35)
        self.jira_tickets_created.setFont(font)

        self.jira_tickets_resolved_label = QLabel("Jira Tickets Resolved")
        self.jira_tickets_resolved_label.setAlignment(Qt.AlignCenter)
        jira_layout.addWidget(self.jira_tickets_resolved_label)
        self.jira_tickets_resolved_label.setFont(font)

        self.jira_tickets_resolved = QTextEdit()
        jira_layout.addWidget(self.jira_tickets_resolved)
        self.jira_tickets_resolved.setFixedHeight(35)
        self.jira_tickets_resolved.setFont(font)

        self.jira_percent_label = QLabel("Percent Resolved")
        self.jira_percent_label.setAlignment(Qt.AlignCenter)
        jira_layout.addWidget(self.jira_percent_label)
        self.jira_percent_label.setFont(font)

        self.jira_percent = QTextEdit()
        jira_layout.addWidget(self.jira_percent)
        self.jira_percent.setFixedHeight(35)
        self.jira_percent.setFont(font)

        self.jira_tickets_type_label = QLabel("Ticket Types")
        self.jira_tickets_type_label.setAlignment(Qt.AlignCenter)
        jira_layout.addWidget(self.jira_tickets_type_label)
        self.jira_tickets_type_label.setFont(font)

        self.jira_tickets_type = QTextEdit()
        jira_layout.addWidget(self.jira_tickets_type)
        self.jira_tickets_type.setFixedHeight(80)
        self.jira_tickets_type.setFont(font)

        # Create a vertical spacer to push the widgets to the top of the frame
        spacer = QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding)
        jira_layout.addItem(spacer, 9, 0, 1, 2)

        # Information Frame
        self.info_frame = QFrame()
        self.info_frame.setFrameShape(QFrame.StyledPanel)
        self.layout.addWidget(self.info_frame)

        info_layout = QGridLayout(self.info_frame)

        # Information fram background color
        palette = QPalette()
        palette.setColor(QPalette.Background, QColor("#d8e3e6"))
        self.info_frame.setPalette(palette)
        self.info_frame.setAutoFillBackground(True)

        # Set Max height of the frame
        self.info_frame.setFixedHeight(500)
        self.info_frame.setMaximumWidth(350)

        # Font Size
        uq_font = QFont()
        uq_font.setPointSize(14)

        # Font Size sub categories
        sub_font = QFont()
        sub_font.setPointSize(12)

        # Set size policy
        self.info_frame.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.layout.setAlignment(self.info_frame, Qt.AlignBottom | Qt.AlignLeft)

        self.main_unassign_label = QLabel("Unassigned Ticket Counts")
        self.main_unassign_label.setAlignment(Qt.AlignCenter)
        info_layout.addWidget(self.main_unassign_label, 0, 0, 1, -1)
        self.main_unassign_label.setFont(font)

        self.main_unassign_label = QLabel("---Per Tech---")
        self.main_unassign_label.setAlignment(Qt.AlignCenter)
        info_layout.addWidget(self.main_unassign_label, 1, 0, 1, -1)
        self.main_unassign_label.setFont(font)
   
        self.ticket_counts_day_label = QLabel("Day")
        self.ticket_counts_day_label.setAlignment(Qt.AlignCenter)
        info_layout.addWidget(self.ticket_counts_day_label, 2, 0)
        self.ticket_counts_day_label.setFont(sub_font)

        self.ticket_counts_day = QTextEdit()
        info_layout.addWidget(self.ticket_counts_day, 3, 0)
        self.ticket_counts_day.setFixedHeight(225)
        self.ticket_counts_day.setFont(uq_font)

        self.ticket_counts_label = QLabel("Total")
        self.ticket_counts_label.setAlignment(Qt.AlignCenter)
        info_layout.addWidget(self.ticket_counts_label, 2, 1)
        self.ticket_counts_label.setFont(sub_font)

        self.ticket_counts_total = QTextEdit()
        info_layout.addWidget(self.ticket_counts_total, 3, 1)
        info_layout.setAlignment(self.ticket_counts_total, Qt.AlignTop)
        self.ticket_counts_total.setFixedHeight(225)
        self.ticket_counts_total.setFont(uq_font)

        self.totals_label = QLabel("---Totals---")
        self.totals_label.setAlignment(Qt.AlignCenter)
        info_layout.addWidget(self.totals_label, 5, 0, 1, -1)
        self.totals_label.setFont(font)

        self.day_label = QLabel("Day")
        self.day_label.setAlignment(Qt.AlignCenter)
        info_layout.addWidget(self.day_label, 6, 0)
        self.day_label.setFont(sub_font)

        # Unassigned total "day" box
        self.day_box = QTextEdit()
        info_layout.addWidget(self.day_box, 7, 0)
        self.day_box.setFixedHeight(35)
        self.day_box.setFont(uq_font)

        self.week_label = QLabel("Week")
        self.week_label.setAlignment(Qt.AlignCenter)
        info_layout.addWidget(self.week_label, 6, 1)
        self.week_label.setFont(sub_font)

        # Unassigned total "week" box
        self.week_box = QTextEdit()
        info_layout.addWidget(self.week_box, 7, 1)
        self.week_box.setFixedHeight(35)
        self.week_box.setFont(uq_font)

        self.month_label = QLabel("Month")
        self.month_label.setAlignment(Qt.AlignCenter)
        info_layout.addWidget(self.month_label, 8, 0)
        self.month_label.setFont(sub_font)

        # Unassigned total "month" box
        self.month_box = QTextEdit()
        info_layout.addWidget(self.month_box, 9, 0)
        self.month_box.setFixedHeight(35)
        self.month_box.setFont(uq_font)

        self.total_label = QLabel("Total")
        self.total_label.setAlignment(Qt.AlignCenter)
        info_layout.addWidget(self.total_label, 8, 1)
        self.total_label.setFont(sub_font)

        # Unassigned total "total" box
        self.total_box = QTextEdit()
        info_layout.addWidget(self.total_box, 9, 1)
        self.total_box.setFixedHeight(35)
        self.total_box.setFont(uq_font)

        # Create a vertical spacer to push the unassign widgets to the top of the frame
        spacer = QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding)
        info_layout.addItem(spacer, 11, 0, 1, 2)

    def showTime(self):
        current_time = datetime.now()
        label_time = current_time.strftime('%I:%M:%S')
        self.time_label.setText(label_time)
            
    def date_time_updater(self):
        current_date = self.get_current_date()
        current_date_str = current_date.strftime("%Y-%m-%d")
        formatted_date = current_date.strftime("%b %d, %Y")

        current_time = self.get_current_time()
        current_time_str = current_time.strftime("%H:%M")

        path = os.getenv('LOGS_PATH')
        json_files = ['date_time.json', 'weights.json', 'user_tickets.json']

        for json_file in json_files:
                json_path = os.path.join(path, json_file)
                with open(json_path, 'r') as f:
                    data = json.load(f)
                    
                for key in data:
                    if key == "Date":
                        data[key] = current_date_str

                for key in data:
                    if key == "Time":
                        data[key] = current_time_str

                with open(json_path, 'w') as f:
                    json.dump(data, f)

        self.date_label.setText(formatted_date)
            
    def get_current_time(self):
        return datetime.now().time()
    
    def get_current_date(self):
        return datetime.now().date()
    
    def start_siren_loop(self):
        threading.Thread(target=self.loop_set_siren_visibility).start()

    def loop_set_siren_visibility(self):
        while True:
            self.set_siren_visibility()
            tm.sleep(5)

    def set_siren_visibility(self):
        time_now = self.get_current_time()
        user_dict = self.get_users_state()
        # Count how many users are in 'READY', 'TALKING', or 'WORK' status
        count = sum(1 for details in user_dict.values() if details['State'] in {'READY', 'TALKING', 'WORK', 'RESERVED'})
        # If the current time is between 7:30 and 17:30 and there is exactly one such user, turn on the sirens, otherwise turn them off
        if time(7, 00) <= time_now <= time(17, 30) and count == 1:
            self.siren_left.show()
            self.siren_right.show()
        else:
            self.siren_left.hide()
            self.siren_right.hide()

    def start_fire_loop(self):
        threading.Thread(target=self.loop_fire_visibility).start()

    def loop_fire_visibility(self):
        while True:
            self.set_fire_visibility()
            tm.sleep(5)

    def set_fire_visibility(self):
        ticket_count = self.jira_tickets_total()
        if ticket_count >= 200:
            self.fire_label.show()
        else:
            self.fire_label.hide()

    def confirm_user_day_reset(self):
        reply = QMessageBox.question(self, 'Confirmation',
                                    'Are you sure you want to reset user day counts to 0?',
                                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.user_day_reset()

    def user_day_reset(self):
        path = os.getenv('LOGS_PATH')
        user_tickets_path = os.path.join(path, 'user_tickets.json')
        with open(user_tickets_path, 'r') as f:
            data = json.load(f)
            
            for key in data:
                if key != "Date":
                    data[key]['day'] = 0

            with open(user_tickets_path, 'w') as f:
                json.dump(data, f)

        ticket_counts = []
        ticket_counts_day = []
        for key, value in data.items():
            if key != "Date" and key != "Total Unassigned":
                ticket_counts.append(f'{key}: {value["total"]}')
                ticket_counts_day.append(f'{key}: {value["day"]}')
        # Set the new status
        self.ticket_counts_total.setText('\n'.join(ticket_counts))
        self.ticket_counts_day.setText('\n'.join(ticket_counts_day))

    def confirm_user_total_reset(self):
        reply = QMessageBox.question(self, 'Confirmation',
                                    'Are you sure you want to reset user total counts to 0?',
                                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.user_total_reset()

    def user_total_reset(self):
        path = os.getenv('LOGS_PATH')
        user_tickets_path = os.path.join(path, 'user_tickets.json')
        with open(user_tickets_path, 'r') as f:
            data = json.load(f)
            
            for key in data:
                if key != "Date":
                    data[key]['total'] = 0

            with open(user_tickets_path, 'w') as f:
                json.dump(data, f)

        ticket_counts = []
        ticket_counts_day = []
        for key, value in data.items():
            if key != "Date" and key != "Total Unassigned":
                ticket_counts.append(f'{key}: {value["total"]}')
                ticket_counts_day.append(f'{key}: {value["day"]}')
        # Set the new status
        self.ticket_counts_total.setText('\n'.join(ticket_counts))
        self.ticket_counts_day.setText('\n'.join(ticket_counts_day))

    def confirm_user_all_reset(self):
        reply = QMessageBox.question(self, 'Confirmation',
                                    'Are you sure you want to reset ALL user counts to 0?',
                                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.user_all_reset()
    
    def user_all_reset(self):
        path = os.getenv('LOGS_PATH')
        user_tickets_path = os.path.join(path, 'user_tickets.json')
        with open(user_tickets_path, 'r') as f:
            data = json.load(f)
            
            for key in data:
                if key != "Date":
                    data[key]['day'] = 0
                    data[key]['total'] = 0

            with open(user_tickets_path, 'w') as f:
                json.dump(data, f)

        ticket_counts = []
        ticket_counts_day = []
        for key, value in data.items():
            if key != "Date" and key != "Total Unassigned":
                ticket_counts.append(f'{key}: {value["total"]}')
                ticket_counts_day.append(f'{key}: {value["day"]}')
        # Set the new status
        self.ticket_counts_total.setText('\n'.join(ticket_counts))
        self.ticket_counts_day.setText('\n'.join(ticket_counts_day))
        
    def confirm_uq_day_reset(self):
        reply = QMessageBox.question(self, 'Confirmation',
                                    'Are you sure you want to reset unassign day counts to 0?',
                                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
                self.uq_day_reset()

    def uq_day_reset(self):
        path = os.getenv('LOGS_PATH')
        user_tickets_path = os.path.join(path, 'user_tickets.json')
        with open(user_tickets_path, 'r') as f:
            data = json.load(f)
            
            for key in data:
                if key == "Total Unassigned":
                    data[key]['day'] = 0

            with open(user_tickets_path, 'w') as f:
                json.dump(data, f)

        unassign_total_day = []
        for key, value in data.items():
            if key == "Total Unassigned":
                unassign_total_day.append(str(value["day"])) 

        self.day_box.setText('\n'.join(unassign_total_day))

    def confirm_uq_week_reset(self):
        reply = QMessageBox.question(self, 'Confirmation',
                                    'Are you sure you want to reset unassign week counts to 0?',
                                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
                self.uq_week_reset()

    def uq_week_reset(self):
        path = os.getenv('LOGS_PATH')
        user_tickets_path = os.path.join(path, 'user_tickets.json')
        with open(user_tickets_path, 'r') as f:
            data = json.load(f)
            
            for key in data:
                if key == "Total Unassigned":
                    data[key]['week'] = 0

            with open(user_tickets_path, 'w') as f:
                json.dump(data, f)

        unassign_total = []
        for key, value in data.items():
            if key == "Total Unassigned":
                unassign_total.append(str(value["week"]))

        self.week_box.setText('\n'.join(unassign_total))

    def confirm_uq_month_reset(self):
        reply = QMessageBox.question(self, 'Confirmation',
                                    'Are you sure you want to reset unassign month counts to 0?',
                                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
                self.uq_month_reset()

    def uq_month_reset(self):
        path = os.getenv('LOGS_PATH')
        user_tickets_path = os.path.join(path, 'user_tickets.json')
        with open(user_tickets_path, 'r') as f:
            data = json.load(f)
            
            for key in data:
                if key == "Total Unassigned":
                    data[key]['month'] = 0

            with open(user_tickets_path, 'w') as f:
                json.dump(data, f)

        unassign_total = []
        for key, value in data.items():
            if key == "Total Unassigned":
                unassign_total.append(str(value["month"]))

        self.month_box.setText('\n'.join(unassign_total))

    def confirm_uq_total_reset(self):
        reply = QMessageBox.question(self, 'Confirmation',
                                    'Are you sure you want to reset unassign total counts to 0?',
                                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
                self.uq_total_reset()

    def uq_total_reset(self):
        path = os.getenv('LOGS_PATH')
        user_tickets_path = os.path.join(path, 'user_tickets.json')
        with open(user_tickets_path, 'r') as f:
            data = json.load(f)
            
            for key in data:
                if key == "Total Unassigned":
                    data[key]['total'] = 0

            with open(user_tickets_path, 'w') as f:
                json.dump(data, f)

        unassign_total = []
        for key, value in data.items():
            if key == "Total Unassigned":
                unassign_total.append(str(value["total"]))

        self.total_box.setText('\n'.join(unassign_total))

    def confirm_uq_all_reset(self):
        reply = QMessageBox.question(self, 'Confirmation',
                                    'Are you sure you want to reset ALL unassign counts to 0?',
                                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
                self.uq_all_reset()

    def uq_all_reset(self):
        path = os.getenv('LOGS_PATH')
        user_tickets_path = os.path.join(path, 'user_tickets.json')
        with open(user_tickets_path, 'r') as f:
            data = json.load(f)
            
            for key in data:
                if key == "Total Unassigned":
                    data[key]['total'] = 0
                    data[key]['day'] = 0
                    data[key]['week'] = 0
                    data[key]['month'] = 0

            with open(user_tickets_path, 'w') as f:
                json.dump(data, f)

        unassign_total = []
        unassign_day = []
        unassign_week = []
        unassign_month = []
        for key, value in data.items():
            if key == "Total Unassigned":
                unassign_total.append(str(value["total"]))
                unassign_day.append(str(value["day"]))
                unassign_week.append(str(value["week"]))
                unassign_month.append(str(value["month"]))

        self.total_box.setText('\n'.join(unassign_total))
        self.day_box.setText('\n'.join(unassign_total))
        self.week_box.setText('\n'.join(unassign_total))
        self.month_box.setText('\n'.join(unassign_total))

    def confirm_weights_reset(self):
        reply = QMessageBox.question(self, 'Confirmation',
                            'Are you sure you want to reset all users weights to 0?',
                            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
                self.weights_reset()

    def weights_reset(self):
        path = os.getenv('LOGS_PATH')
        user_tickets_path = os.path.join(path, 'weights.json')
        with open(user_tickets_path, 'r') as f:
            data = json.load(f)

            for key in data:
                if key != 'Date':
                    data[key] = 0

        with open(user_tickets_path, 'w') as f:
            json.dump(data, f)

    def confirm_vm_reset(self):
        reply = QMessageBox.question(self, 'Confirmation',
                            'Are you sure you want to reset Voicemail-Day to 0?',
                            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
                self.vm_counter_day_reset()

    def vm_counter_day_reset(self):
        path = os.getenv('LOGS_PATH')
        file_path = os.path.join(path, 'vm_ticket_count.json')
        with open(file_path, 'r') as f:
            vm_count = json.load(f)
            
        vm_count['Day'] = 0
            
        with open(file_path, 'w') as f:
            json.dump(vm_count, f)

        self.update_vm_count()
    
    def confirm_vm_week_reset(self):
        reply = QMessageBox.question(self, 'Confirmation',
                            'Are you sure you want to reset Voicemail-Week to 0?',
                            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
                self.vm_counter_week_reset()

    def vm_counter_week_reset(self):
        path = os.getenv('LOGS_PATH')
        file_path = os.path.join(path, 'vm_ticket_count.json')
        with open(file_path, 'r') as f:
            vm_count = json.load(f)
            
        vm_count['Week'] = 0
            
        with open(file_path, 'w') as f:
            json.dump(vm_count, f)

        self.update_vm_count()

    def confirm_vm_month_reset(self):
        reply = QMessageBox.question(self, 'Confirmation',
                            'Are you sure you want to reset Voicemail-Month to 0?',
                            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
                self.vm_counter_month_reset()

    def vm_counter_month_reset(self):
        path = os.getenv('LOGS_PATH')
        file_path = os.path.join(path, 'vm_ticket_count.json')
        with open(file_path, 'r') as f:
            vm_count = json.load(f)
            
        vm_count['Month'] = 0
            
        with open(file_path, 'w') as f:
            json.dump(vm_count, f)

        self.update_vm_count()

    def confirm_vm_all_reset(self):
        reply = QMessageBox.question(self, 'Confirmation',
                            'Are you sure you want to reset the voicemail-Day, Week, Month to 0?',
                            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
                self.vm_counter_all_reset()

    def vm_counter_all_reset(self):
        path = os.getenv('LOGS_PATH')
        file_path = os.path.join(path, 'vm_ticket_count.json')
        with open(file_path, 'r') as f:
            vm_count = json.load(f)
            
        vm_count['Month'] = 0
        vm_count['Week'] = 0
        vm_count['Day'] = 0
            
        with open(file_path, 'w') as f:
            json.dump(vm_count, f)

        self.update_vm_count()

    def stop_uq(self):
        self.uq_stopped = True
        self.uq_status.setText("Stopped")
        self.uq_status.setStyleSheet("background-color: red;")

    def stop_fls(self):
        self.fls_stopped = True
        self.fls_status.setText("Stopped")
        self.fls_status.setStyleSheet("background-color: red;")

    def start_unassign(self):
        self.uq_stopped = False
        # Start the worker thread
        self.uq_worker.start()
        self.start_uq_time_check()

    def start_fls(self):
        self.fls_stopped = False
        # Start the worker thread
        self.fls_worker.start()

    def uq_update_color(self, color):
        self.uq_status.setStyleSheet(f"background-color: {color};")
    
    def fls_update_color(self, color):
        self.fls_status.setStyleSheet(f"background-color: {color};")

    def update_uq_status(self, text):
        self.uq_status.setText(text)

    def update_fls_status(self, text):
        self.fls_status.setText(text)

    def start_uq_time_check(self):
        threading.Thread(target=self.loop_uq_after_hours_stop).start()

    def loop_uq_after_hours_stop(self):
        while True:
            self.uq_after_hours_stop()
            tm.sleep(60)

    def uq_after_hours_stop(self):
        time_now = self.get_current_time()
        if time_now < time(7, 0) or time(18, 0) <= time_now:
            self.uq_stopped = True
            self.uq_worker.update_color("red")
            self.uq_worker.update_status("Stopped")

        else:
            if self.uq_stopped:
                self.uq_stopped = False
                self.uq_worker.start()
                self.uq_worker.update_color("green")
                self.uq_worker.update_status("Running")

    def start_fls_check(self):
        threading.Thread(target=self.loop_fls_check).start()

    def loop_fls_check(self):
        while True:
            self.fls_check()
            tm.sleep(60)

    def fls_check(self):
        time_now = self.get_current_time()
        if time_now < time(6, 30) or time_now > time(18, 0):
            self.fls_stopped = False

        else:
            if self.fls_stopped:  # if it was stopped, start it
                self.start_fls()

    def update_ticket_count(self):
        # Read JSON data
        path = os.getenv('LOGS_PATH')
        file_path = os.path.join(path, 'user_tickets.json')
        with open(file_path, 'r') as f:
            data = json.load(f)
        # Prepare the new status
        ticket_counts_total = []
        ticket_counts_day = []
        for key, value in data.items():
            if key != "Date" and key != "Total Unassigned":
                ticket_counts_total.append(f'{key}: {value["total"]}')
                ticket_counts_day.append(f'{key}: {value["day"]}')

        # Set the new status
        self.ticket_counts_total.setText('\n'.join(ticket_counts_total))
        self.ticket_counts_day.setText('\n'.join(ticket_counts_day))

        unassign_total = []
        unassign_total_day = []
        unassign_total_week = []
        unassign_total_month = []
        for key, value in data.items():
            if key == "Total Unassigned":
                unassign_total.append(str(value["total"]))
                unassign_total_day.append(str(value["day"])) 
                unassign_total_week.append(str(value["week"]))
                unassign_total_month.append(str(value["month"]))

        self.total_box.setText('\n'.join(unassign_total))
        self.day_box.setText('\n'.join(unassign_total_day))
        self.week_box.setText('\n'.join(unassign_total_week))
        self.month_box.setText('\n'.join(unassign_total_month))

    def update_vm_count(self):
        path = os.getenv('LOGS_PATH')
        file_path = os.path.join(path, 'vm_ticket_count.json')
        with open(file_path, 'r') as f:
            data = json.load(f)
        text = "\n".join(f"{key}: {value}" for key, value in data.items() if key != 'Date')

        self.vm_counts.clear()
        self.vm_counts.setText(text)

    def auto_fls(self):
        """Creates Jira ticket from Outlook email"""
        path = os.getenv('LOGS_PATH')
        vm_file_path = os.path.join(path, 'vm_ticket_count.json')
        try:
            with open(vm_file_path, 'r') as f:
                vm_count = json.load(f)
        except FileNotFoundError as e:
            print(e)
            print("Check the vm_ticket_count.json file.")

        count = 0
        while not self.fls_stopped:
            if self.fls_stopped:
                return
            try:
                path = os.getenv('CONFIG_PATH')
                file_path = os.path.join(path, 'voicemessage.wav')
                self.fls_worker.update_status("Running")
                self.fls_worker.update_color("green")
                self.start_fls_check()
                issue_dict = {}
                pythoncom.CoInitialize()
                outlook = win32.Dispatch('Outlook.Application')
                namespace = outlook.GetNamespace("MAPI")
                # Real email removed for privacy on public version
                inbox = namespace.Folders['VOICEMAIL_EMAIL'].Folders['Inbox']
                # Real email removed for privacy on public version
                archive = namespace.Folders['VOICEMAIL_EMAIL'].Folders['Archive']
                messages = inbox.Items
                message = messages.GetLast()
                sender = message.SenderName
                if sender == "Cisco Unity Connection Messaging System":
                    sender = "FLS Voicemail Inbox"
                creation_time = message.CreationTime.strftime(format="%H:%M-%b %d")
                subject = message.Subject
                attachment = message.Attachments
                attachment = attachment.Item(1)
                file_name = str(attachment).lower()
                attachment.SaveASFile(f'{path}/{file_name}')
                if subject.startswith('Message from B'):
                    subject_cleaned = re.split('\s+', subject)
                    branch = (f'Branch Number: {subject_cleaned[2]}')
                    call_back = (f'Call back number/EXT: {subject_cleaned[-1]}')
                    caller = (f'Caller Name:')
                else:
                    subject_cleaned = re.split('\s+', subject)
                    caller = (f'Caller Name: {subject_cleaned[2]}')
                    call_back = (f'Call back number/EXT: {subject_cleaned[-1]}')
                    branch = ("No branch given.")
                label = "Voicemail"
                issue_dict.update({
                    'project': {'key': 'ITDESK'},
                    'summary': f'VM @ {creation_time} | From: {sender}',
                    'description': f'TIME CREATED:{creation_time}\nSENT FROM: {sender}\nCALLER INFO: {subject}\n{branch}\n{call_back}\n{caller}\nGoogle voice to text:\n{self.wav_text()}',
                    'issuetype': {'name': 'Service Request'},
                    'labels': [label],
                })
                new_issue = self.jira_connect().create_issue(fields=issue_dict)
                url = f'{os.getenv("DOMAIN")}/rest/api/3/issue/{new_issue}/attachments'
                headers = {
                    "X-Atlassian-Token": "no-check"
                }
                with open(file_path, "rb") as vm_file:
                    file_data = vm_file.read()
                files = {
                    "file": ("voicemessage.wav", file_data)
                }
                response = requests.post(url, headers=headers, files=files, auth=Application.jira_oauth(), verify=False)
                self.fls_worker.update_vm_ticket_signal.emit(str(new_issue))
                if message.UnRead:
                    message.UnRead = False
                message.Move(archive)
                count += 1

                vm_count['Day'] = vm_count.get('Day', 0) + 1
                vm_count['Week'] = vm_count.get('Week', 0) + 1
                vm_count['Month'] = vm_count.get('Month', 0) + 1
                vm_count['Total'] = vm_count.get('Total', 0) + 1
                with open(vm_file_path, 'w') as f:
                    json.dump(vm_count, f)
                counts = {'Day': vm_count['Day'], 'Week': vm_count['Week'], 'Month': vm_count['Month'], 'Total': vm_count['Total']}
                self.fls_worker.update_vm_counts_signal.emit(counts)

            except pywintypes.com_error:
                if message.UnRead:
                    message.UnRead = False
                message.Move(archive)
                continue

            except AttributeError:
                self.fls_worker.update_vm_signal.emit(str("Voicemail Inbox Cleared"))
                self.fls_worker.vm_status_color_signal.emit("Green")
                # Delete the voicemail file
                if os.path.exists(file_path):
                    os.remove(file_path)
                tm.sleep(20)
                continue

    def wav_text(self):
        path = os.getenv('CONFIG_PATH')
        file_path = os.path.join(path, 'voicemessage.wav')
        new_file_path = os.path.join(path, 'new.wav')
        data, samplerate = soundfile.read(file_path)
        soundfile.write(new_file_path, data, samplerate, subtype='PCM_16')
        r = sr.Recognizer()
        hellow = sr.AudioFile(new_file_path)
        with hellow as source:
            audio = r.record(source)
        try:
            s = r.recognize_google(audio, show_all = True, )
            results = s['alternative'][0]
            transcript = results['transcript']
        except Exception as e:
            self.error_ticker.set_text(f"An error occurred: {e}. Thrown from function: wav_text")
            transcript = None

        return transcript

    def jira_oauth():
        jira_connection = (os.getenv("JIRA_LOGIN"), os.getenv("API_KEY"))
        return jira_connection

    def jira_connect(self):
        options = {
            'server': os.getenv("DOMAIN"),
            'verify': False
            }
        jira = JIRA(options=options, basic_auth=(os.getenv("JIRA_LOGIN"), os.getenv("API_KEY")))
        return jira

    def unassigned_queue(self):
        """Searches Jira for all unassigned tickets, returns list of ticket numbers"""
        try:
            jira_connection = self.jira_connect()
            jira_search_query = 'assignee in (EMPTY) AND reporter not in (qm:251901c3-3c6e-422d-894c-674ab8eb81e4:5b1eeb44882031170e5e24c5, qm:251901c3-3c6e-422d-894c-674ab8eb81e4:0519053f-359f-4a98-a0f8-d79d2b6b2dcd, qm:251901c3-3c6e-422d-894c-674ab8eb81e4:fa91fa70-098e-4f30-bd27-27e0ef471df5, qm:251901c3-3c6e-422d-894c-674ab8eb81e4:cf86f794-5b7a-4ac6-bc84-394decaad2a1, qm:251901c3-3c6e-422d-894c-674ab8eb81e4:aa622cb4-d28e-4f45-a7bf-716741b318ba, qm:251901c3-3c6e-422d-894c-674ab8eb81e4:9a32cb47-7551-4ef0-96bd-d124094a285f, qm:251901c3-3c6e-422d-894c-674ab8eb81e4:1000f0b5-4989-4e4a-b341-88e368ae1c32) AND project = ITDESK AND issuetype not in (Task, subTaskIssueTypes(), "Purchase Request", Purchase) AND status in (Open, Reopened, "Waiting for support") order by created DESC'
            jira_issues = jira_connection.search_issues(jira_search_query)
            unassigned_list = [issue.key for issue in jira_issues]
            while not unassigned_list:
                jira_issues = jira_connection.search_issues(jira_search_query)
                unassigned_list = [issue.key for issue in jira_issues]
        except JIRAError as e:
            self.error_ticker.set_text(f"An error occurred: {e}. Thrown from function: unassigned_queue")
            self.unassigned_queue()
        return unassigned_list[0]
    
    def minutes_since_six_thirty(self):
        now = datetime.now()
        six_thirty = now.replace(hour=6, minute=30, second=0, microsecond=0)
        
        if now < six_thirty:
            # If it's before 6:30, return None
            return None
            
        diff = now - six_thirty
        minutes = diff.total_seconds() / 60
        rounded_minutes = round(minutes)
        
        return rounded_minutes
    
    def jira_tickets_total(self):
        jira_connection = self.jira_connect()
        minutes = self.minutes_since_six_thirty()
        if minutes == None:
            return 0
        jira_search_query = f'project = ITDESK AND issuetype in (Incident, Problem, "Service Request") AND created >= -{minutes}m order by created DESC'
        startAt = 0
        maxResults = 100
        unassigned_list = []
        
        while True:
            jira_issues = jira_connection.search_issues(jira_search_query, startAt=startAt, maxResults=maxResults)
            if len(jira_issues) == 0:
                # No more issues were found, so exit the loop
                break
            unassigned_list.extend(issue.key for issue in jira_issues)
            startAt += maxResults

        return int(len(unassigned_list))
    
    def jira_tickets_resolved_fn(self):
        jira_connection = self.jira_connect()
        minutes = self.minutes_since_six_thirty()
        if minutes == None:
            return 0
        jira_search_query = f'project = ITDESK AND issuetype in (Incident, Problem, "Service Request") AND created >= -{minutes}m AND status = Resolved order by created DESC'        
        startAt = 0
        maxResults = 100
        resolved_list = []
        
        while True:
            jira_issues = jira_connection.search_issues(jira_search_query, startAt=startAt, maxResults=maxResults)
            if len(jira_issues) == 0:
                break
            resolved_list.extend(issue.key for issue in jira_issues)
            startAt += maxResults

        return int(len(resolved_list))

    def jira_tickets_type_fn(self):
        jira_connection = self.jira_connect()
        minutes = self.minutes_since_six_thirty()
        if minutes == None:
            return 0
        jira_search_query = f'project = ITDESK AND issuetype in (Incident, Problem, "Service Request") AND created >= -{minutes}m order by created DESC'
        startAt = 0
        maxResults = 100
        type_list = []

        while True:
            jira_issues = jira_connection.search_issues(jira_search_query, startAt=startAt, maxResults=maxResults)
            if len(jira_issues) == 0:
                # No more issues were found, so exit the loop
                break
            type_list.extend(issue.fields.issuetype.name for issue in jira_issues if issue.fields.issuetype.name not in ['Purchase Request', 'Task'])
            startAt += maxResults
        
        ticket_type_counts = Counter(type_list)

        ticket_type_dict = {
        'Service Request': ticket_type_counts.get('Service Request', 0),
        'Incident': ticket_type_counts.get('Incident', 0),
        'Problem': ticket_type_counts.get('Problem', 0)
        }
        
        ticket_type_str = ""
        for key, value in ticket_type_dict.items():
            ticket_type_str += f"{key}: {value}\n"

        return str(ticket_type_str)

    def jira_percent_fn(self):
        total = self.jira_tickets_total()
        resolved = self.jira_tickets_resolved_fn()
        try:
            percent = resolved / total * 100
            percent = round(percent, 2)
        except ZeroDivisionError:
            percent = 0
        return int(percent)
    
    def send_email(self, recipient, ticket):
        # Sender email changed to preserve name privacy. Pulls from secrets.env normally
        email_address = os.getenv('SENDER_EMAIL')
        ticket_url = os.getenv('TICKET_URL')
        pythoncom.CoInitialize()
        outlook = win32.Dispatch('Outlook.Application')
        namespace = outlook.GetNamespace("MAPI")
        account = next(acc for acc in namespace.Accounts if acc.SmtpAddress == email_address)
        mail = outlook.CreateItem(0)
        mail._oleobj_.Invoke(*(64209, 0, 8, 0, account))  # 64209 corresponds to the SendUsingAccount property
        mail.To = recipient
        mail.Subject = 'Unassigned Ticket Queue Notification'
        html_body = f"You have been assigned an unassigned ticket. {ticket_url}{ticket}"
        mail.HTMLBody = html_body
        # Send the email
        mail.Send()

    def get_users_state(self):
        username = os.getenv("FINESSE_USERNAME")
        password = os.getenv("FINESSE_PASS")
        url = os.getenv("FINESSE_API2")
        # Users to exclude
        path = os.getenv("CONFIG_PATH")
        exclude_path = os.path.join(path, 'excluded_users.json')
        # Load the excluded users from a JSON file
        with open(exclude_path, 'r') as f:
            exclude_users_data = json.load(f)
        # Extract the 'user' value from each dictionary and assign the values to the exclude_users list
        exclude_users = [user_dict['user'] for user_dict in exclude_users_data]
        response = requests.get(url, auth=HTTPBasicAuth(username, password), verify=False)
        user_dict = {}
        if response.status_code == 200:
            root = ET.fromstring(response.text)
            users = root.find('users')
            if users is not None:
                for user in users.findall('User'):
                    state_text = ''
                    loginid_text = ''
                    label_text = ''
                    loginId = user.find('loginId')
                    if loginId is not None:
                        loginid_text = loginId.text
                    if loginid_text in exclude_users:
                        continue
                    state = user.find('state')
                    if state is not None:
                        state_text = state.text
                    reasonCode = user.find('reasonCode')
                    if reasonCode is not None:
                        label = reasonCode.find('label')
                        if label is not None:
                            label_text = label.text
                    user_dict[loginid_text] = {'State': state_text, 'Reason': label_text}
        else:
            print(f'Error: {response.status_code}')
        return user_dict
    
    def finesse_status_gui(self):
        # Call the function to get the user data
        users = self.get_users_state()
        # Create a list of your QTextEdit widgets
        reason_boxes = [self.reason1, self.reason2, self.reason3, self.reason4, self.reason5, self.reason6, self.reason7, self.reason8, self.reason9]
        status_boxes = [self.finesse1, self.finesse2, self.finesse3, self.finesse4, self.finesse5, self.finesse6, self.finesse7, self.finesse8, self.finesse9]
        tech_boxes = [self.tech1, self.tech2, self.tech3, self.tech4, self.tech5, self.tech6, self.tech7, self.tech8, self.tech9]
        # Iterate over the users and the QTextEdit widgets together
        for (user, details), reason_box, status_box, tech_box in zip(users.items(), reason_boxes, status_boxes, tech_boxes):
            # Create the text to display
            reason_text = details["Reason"]
            status_text = details["State"]
            tech_text = user
            # Clear the QTextEdit widgets
            reason_box.clear()
            status_box.clear()
            tech_box.clear()
            # Insert the text into the appropriate QTextEdit widget
            reason_box.append(reason_text)
            status_box.append(status_text)
            tech_box.append(tech_text)
            # Change the background color of the status box based on the status
            if status_text == 'READY':
                status_box.setStyleSheet("background-color: green")

            elif status_text == 'WORK':
                status_box.setStyleSheet("background-color: orange")

            elif status_text == 'TALKING':
                status_box.setStyleSheet("background-color: yellow")

            elif status_text == 'RESERVED':
                status_box.setStyleSheet("background-color: yellow")

            elif status_text == 'NOT_READY':
                status_box.setStyleSheet("background-color: red")

            elif status_text == 'LOGOUT':
                status_box.setStyleSheet("background-color: black")

    def refresh_finesse_status(self):
        self.finesse_status_gui()
        QTimer.singleShot(5000, self.refresh_finesse_status)

    def refresh_user_queue(self):
        queue = User.get_next_user_id()
        self.current_queue.clear()
        for i, item in enumerate(queue, start=1):
            self.current_queue.append(f"{i}. {item}")
        QTimer.singleShot(5000, self.refresh_user_queue)

    def loop_new_day_check(self):
        current_date = self.get_current_date()
        path = os.getenv('LOGS_PATH')
        date_time_path = os.path.join(path, 'date_time.json')
        with open(date_time_path, 'r') as f:
            data = json.load(f)
            for key in data:
                if key == "Date":
                    date_variable = datetime.strptime(data[key], "%Y-%m-%d").date()
                    if date_variable != current_date:
                        self.start_new_day_thread()

    def loop_new_week_check(self):
        current_date = self.get_current_date()
        path = os.getenv('LOGS_PATH')
        date_time_path = os.path.join(path, 'date_time.json')
        with open(date_time_path, 'r') as f:
            data = json.load(f)
            for key in data:
                if key == "Date":
                    date_variable = datetime.strptime(data[key], "%Y-%m-%d").date()
                    if date_variable.isocalendar()[1] != current_date.isocalendar()[1]:
                        self.start_new_week_thread()

    def loop_new_month_check(self):
        current_date = self.get_current_date()
        path = os.getenv('LOGS_PATH')
        date_time_path = os.path.join(path, 'date_time.json')
        with open(date_time_path, 'r') as f:
            data = json.load(f)
            for key in data:
                if key == "Date":
                    date_variable = datetime.strptime(data[key], "%Y-%m-%d").date()
                    if date_variable.month != current_date.month:
                        self.start_new_month_thread()

    def load_user_tickets(self):
        try:
            path = os.getenv('LOGS_PATH')
            user_tickets_path = os.path.join(path, 'user_tickets.json')
            with open(user_tickets_path, 'r') as f:
                user_tickets = json.load(f)
        except FileNotFoundError:
            user_tickets = {}
        return user_tickets

    def update_user_tickets(self, user_tickets, user_id, ticket, next_assignee):
        assignment_path = os.getenv('LOGS_PATH')
        user_assignment_path = os.path.join(assignment_path, 'user_tickets.json')
        logs_path = os.getenv('LOGS_PATH')
        user_tickets_path = os.path.join(logs_path, 'user_tickets.json')
        named_tuple = tm.localtime()
        time_string = tm.strftime("%Y-%m-%d %H:%M", named_tuple)
        user_tickets[user_id]["day"] += 1
        user_tickets[user_id]["total"] += 1
        user_tickets["Total Unassigned"]["day"] += 1
        user_tickets["Total Unassigned"]["week"] += 1
        user_tickets["Total Unassigned"]["month"] += 1
        user_tickets["Total Unassigned"]["total"] += 1
        assignment = f'{time_string}: {user_id} assigned {ticket}'
        with open(user_assignment_path, 'a') as f:
            f.write(assignment + '\n')
        with open(user_tickets_path, 'w') as f:
            json.dump(user_tickets, f)
        with open(user_tickets_path, 'r') as f:
            user_tickets = json.load(f)
        ticket_counts = []
        ticket_counts_day = []
        # Insert new status
        for key, value in user_tickets.items():
            if key != "Date" and key != "Total Unassigned":
                ticket_counts.append(f'{key}: {value["total"]}')
                ticket_counts_day.append(f'{key}: {value["day"]}')

        self.uq_worker.update_ticket_count(ticket_counts)
        self.uq_worker.update_ticket_counts_day(ticket_counts_day)

        unassign_total = []
        unassign_total_day = []
        unassign_total_week = []
        unassign_total_month = []
        for key, value in user_tickets.items():
            if key == "Total Unassigned":
                unassign_total.append(f'{value["total"]}')
                unassign_total_day.append(f'{value["day"]}')
                unassign_total_week.append(f'{value["week"]}')
                unassign_total_month.append(f'{value["month"]}')
        
        self.uq_worker.update_total_unassigned(unassign_total)
        self.uq_worker.update_total_unassigned_day(unassign_total_day)
        self.uq_worker.update_total_unassigned_week(unassign_total_week)
        self.uq_worker.update_total_unassigned_month(unassign_total_month)
        self.uq_worker.update_next_assignee_signal.emit(next_assignee)
        
    def get_next_assignee(self, users, last_assigned_user):
        self.load_weights(users)
        # Sort users by weight and find the first active user who is not the last assigned user
        users.sort(key=lambda user: user.weight)
        active_users = [user for user in users if user.state]
        for user in active_users:
            if len(active_users) > 1 and user == last_assigned_user:
                continue  # Skip to the next user if this user was the last assigned user and there is more than one active user
            return user
        return None  # return None if no suitable user is found

    @staticmethod
    def save_weights(users):
        path = os.getenv('LOGS_PATH')
        file_path = os.path.join(path, 'weights.json')
        with open(file_path, 'r') as f:
            data = json.load(f)
        for user in users:
            data[user.id] = user.weight
        with open(file_path, 'w') as f:
            json.dump(data, f)

    @staticmethod
    def load_weights(users):
        path = os.getenv('LOGS_PATH')
        file_path = os.path.join(path, 'weights.json')
        if os.path.exists(file_path):
            with open(file_path, 'r') as f:
                weights = json.load(f)
            for user in users:
                if user.id in weights:
                    user.weight = weights[user.id]

    def unassign_loop(self, users):
        self.load_weights(users)
        user_tickets = self.load_user_tickets()
        total_tickets_assigned = 0
        daily_tickets_assigned = 0
        next_user_index = 0  # variable to keep track of the next user to assign

        # Load the JSON file and create a list of user names in the desired order
        logs_path = os.getenv('LOGS_PATH')
        assignment_path = os.path.join(logs_path, 'weights.json')
        with open(assignment_path, 'r') as f:
            data = json.load(f)
            user_order = [key for key in data.keys() if key != "Date"]

        while not self.uq_stopped:
            tm.sleep(1)
            if self.uq_stopped:
                return
            try:
                self.uq_worker.update_status("Running")
                self.uq_worker.update_color("green")
                ticket = self.unassigned_queue()
                if ticket is not None:
                    # Sort users by weight and then by their order in the user_order list
                    users.sort(key=lambda user: (user.weight, user_order.index(user.id)))
                    active_users = [user for user in users if user.state]
                    if active_users:
                        # Find the next user in user_order who is also in active_users
                        user = None
                        while user is None and next_user_index < len(user_order):
                            if user_order[next_user_index] in [u.id for u in active_users]:
                                user = next((u for u in active_users if u.id == user_order[next_user_index]))
                            else:
                                next_user_index += 1
                                if next_user_index == len(user_order):  # if next_user_index has reached the end of user_order
                                    next_user_index = 0  # reset next_user_index
                        if user is None:  # If no suitable user was found, all users are inactive
                            print("No active users found.")
                            continue

                        if not self.uq_stopped:
                            output = user.assign_ticket(ticket, users, user)
                            total_tickets_assigned += 1
                            daily_tickets_assigned += 1
                        else:
                            break
                        if output is not False:
                            next_assignee = self.get_next_assignee(users, user)
                            self.uq_worker.update_last_assigned_signal.emit(output)
                            self.save_weights(users)
                            if next_assignee is not None:
                                self.update_user_tickets(user_tickets, user.id, ticket, next_assignee.id)
                            else:
                                print("No suitable next assignee found.")
                        if user_order.index(user.id) == len(user_order) - 1:  # if the user is the last one in the list
                            next_user_index = 0  # reset the next user index
                        else:
                            next_user_index = user_order.index(user.id) + 1  # move to the next user
                    else:
                        print("No active users found.")

                if user is not None and user.id != None:
                    named_tuple = tm.localtime()
                    time_string = tm.strftime("%m-%d-%y %H:%M", named_tuple)
                    assignment = f'{time_string}: {user.id} assigned {ticket}'
                    logs_path = os.getenv('LOGS_PATH')
                    assignment_path = os.path.join(logs_path, 'assignments.txt')
                    with open(assignment_path, 'a') as f:
                        f.write(assignment + '\n')
                else:
                    break

            except IndexError:
                output = "Unassign queue is empty."

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = Application()
    window.show()
    sys.exit(app.exec_())