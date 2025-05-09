from PyQt5.QtWidgets import QFrame, QGridLayout, QLabel, QLineEdit, QPushButton
from src.gui.components.styles import style_button

def create_file_section():
    frame = QFrame()
    layout = QGridLayout()
    frame.setLayout(layout)
    frame.setStyleSheet("""
        QFrame {
            border: 1px solid #cccccc;
            border-radius: 8px;
            background-color: #f9f9f9;
            padding: 12px;
        }
        QLabel {
            font-weight: bold;
            font-size: 11pt;
        }
    """)

    label = QLabel("File to create a report for")
    fileText = QLineEdit()
    fileButton = QPushButton("Select File")
    fileButton.setFixedSize(100, 32)
    style_button(fileButton)

    layout.addWidget(label, 0, 0, 1, 2)
    layout.addWidget(fileText, 1, 0)
    layout.addWidget(fileButton, 1, 1)
    layout.setColumnStretch(0, 3)
    layout.setColumnStretch(1, 1)

    return frame, fileText, fileButton

def create_feature_section():
    frame = QFrame()
    layout = QGridLayout()
    frame.setLayout(layout)
    frame.setStyleSheet("""
        QFrame {
            border: 1px solid #cccccc;
            border-radius: 8px;
            background-color: #f9f9f9;
            padding: 12px;
        }
    """)

    aiButton = QPushButton("AI Analysis")
    reportButton = QPushButton("Generate Report")
    saveButton = QPushButton("Save")

    for btn in [aiButton, reportButton, saveButton]:
        btn.setFixedSize(140, 32)
        style_button(btn)

    layout.addWidget(aiButton, 0, 0)
    layout.addWidget(reportButton, 1, 0)
    layout.addWidget(saveButton, 1, 1)
    layout.setColumnStretch(0, 1)
    layout.setColumnStretch(1, 1)

    return frame, {"ai": aiButton, "report": reportButton, "save": saveButton}