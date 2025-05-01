from PyQt5.QtWidgets import (QWidget, QPushButton, QGridLayout, QFrame, QScrollArea, 
                             QLabel, QHBoxLayout, QCheckBox, QSizePolicy, QVBoxLayout)
from PyQt5.QtCore import Qt

class AIAnalysis:
    def __init__(self, outputs: list, frame):
        self.root = QWidget()
        self.root.setWindowTitle("AI Analysis")
        self.root.setMaximumWidth(640)
        self.selectedItems = ["" for item in outputs for _ in (item if isinstance(item, list) else [item])]
        self.mainFrame = frame

        # Layouts
        self.mainGrid = QGridLayout()
        self.aiGrid = QGridLayout()
        self.buttonGrid = QHBoxLayout()

        # Header labels
        # Header row inside scroll layout
        header_layout = QGridLayout()
        self.includeLabel = QLabel("Include?")
        self.aILabel = QLabel("AI Analysis")
        self.format_title(self.includeLabel)
        self.format_title(self.aILabel)
        header_layout.addWidget(self.includeLabel, 0, 0)
        header_layout.addWidget(self.aILabel, 0, 1)
        header_layout.setColumnStretch(0, 1)
        header_layout.setColumnStretch(1, 9)

        header_widget = QWidget()
        header_widget.setLayout(header_layout)


        # Section titles
        self.titles = [
            "Deflection", "Operating Force", "Air Infiltration",
            "Water Penetration", "Ultimate Strength"
        ]

        # Scrollable Container
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)

        scroll_content = QWidget()
        scroll_layout = QVBoxLayout(scroll_content)
        scroll_layout.addWidget(header_widget)

        # Add content
        # widgetIdx = 1
        pos = 0
        for idx, section_outputs in enumerate(outputs):
            section_frame, count = self.format_analysis_frame(self.titles[idx], section_outputs, pos)
            scroll_layout.addWidget(section_frame)
            pos += count

        scroll_area.setWidget(scroll_content)
        self.mainGrid.addWidget(scroll_area, 1, 0)

        # Finish Button
        self.finishButton = QPushButton("Finish")
        self.finishButton.clicked.connect(self.finish_button_action)
        self.style_button(self.finishButton)
        self.buttonGrid.addWidget(self.finishButton)
        self.buttonGrid.setAlignment(Qt.AlignCenter)

        # Set layout
        self.mainGrid.addLayout(self.aiGrid, 0, 0)
        self.mainGrid.addLayout(self.buttonGrid, 2, 0)
        self.root.setLayout(self.mainGrid)

    def format_analysis_frame(self, title: str, ai_outputs: list, base_idx: int):
        frame = QFrame()
        frame.setStyleSheet("""
            QFrame {
                border: 1px solid #cccccc;
                border-radius: 8px;
                background-color: #fdfdfd;
                margin-top: 10px;
                padding: 8px;
            }
            QLabel[role='title'] {
                font-weight: bold;
                font-size: 12pt;
                background-color: #f7f779;
                padding: 4px 6px;
                border-radius: 4px;
            }
            QLabel {
                font-size: 10.5pt;
            }
        """)

        layout = QVBoxLayout()
        titleLabel = QLabel(title)
        titleLabel.setProperty("role", "title")
        layout.addWidget(titleLabel)

        for i, text in enumerate(ai_outputs):
            row = QHBoxLayout() # row for current ai output
            checkbox = QCheckBox()
            checkbox.setMaximumWidth(24)
            label = QLabel(text)
            label.setWordWrap(True)
            label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
            checkbox.stateChanged.connect(
                lambda state, lbl=label, idx=base_idx + i: self.handle_checkbox_change(state, lbl, idx)
            )
            row.addWidget(checkbox)
            row.addWidget(label)
            layout.addLayout(row)

        frame.setLayout(layout)
        return frame, len(ai_outputs)

    def handle_checkbox_change(self, state, label, idx):
        self.selectedItems[idx] = label.text() if state == Qt.Checked else ""

    def finish_button_action(self):
        print("Selected Items:", [item for item in self.selectedItems if item])
        self.mainFrame.setCurrentIndex(0)

    def run(self):
        self.root.show()

    def style_button(self, button):
        button.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                font-weight: bold;
                font-size: 11pt;
                padding: 10px;
                border-radius: 6px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
        """)
    
    def __call__(self, *args, **kwds):
        return self.root
    
    def get_window(self):
        return self.root

    def format_title(self, label, font_size=12):
        label.setStyleSheet(f"""
            QLabel {{
                font-weight: bold;
                font-size: {font_size}pt;
                background-color: #f7f779;
                padding: 4px 6px;
                border-radius: 4px;
            }}
        """)

