from PyQt5.QtWidgets import QWidget, QVBoxLayout, QGridLayout, QPushButton, QScrollArea, QHBoxLayout, QLabel
from PyQt5.QtCore import Qt
from src.gui.components.checkbox_list_section import create_checkbox_section
from src.gui.components.styles import style_button, format_header_label

class AIAnalysis:
    def __init__(self, outputs: list, frame, report_frame):
        self.root = QWidget()
        self.root.setWindowTitle("AI Analysis")
        self.root.setMaximumWidth(640)
        self.mainFrame = frame
        self.mainFrame.addWidget(self.root)
        self.reportFrame = report_frame

        self.titles = ["Deflection", "Operating Force", "Air Infiltration", "Water Penetration", "Ultimate Strength"]
        self.selectedItems = [""] * len(outputs)


        self.mainGrid = QGridLayout()
        self._build_scrollable_layout(outputs)
        self._build_button_layout()
        self.root.setLayout(self.mainGrid)

    def _build_scrollable_layout(self, outputs):
        header_layout = QGridLayout()
        label_include = QLabel("Include?")
        label_ai = QLabel("AI Analysis")
        format_header_label(label_include)
        format_header_label(label_ai)
        header_layout.addWidget(label_include, 0, 0)
        header_layout.addWidget(label_ai, 0, 1)
        header_layout.setColumnStretch(0, 1)
        header_layout.setColumnStretch(1, 9)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_content = QWidget()
        scroll_layout = QVBoxLayout(scroll_content)
        scroll_layout.addLayout(header_layout)

        for idx, section_outputs in enumerate(outputs):
            section_frame = create_checkbox_section(
                self.titles[idx], section_outputs, idx, self._on_checkbox_state_changed
            )
            scroll_layout.addWidget(section_frame)


        scroll_area.setWidget(scroll_content)
        self.mainGrid.addWidget(scroll_area, 1, 0)

    def _build_button_layout(self):
        button_layout = QHBoxLayout()
        self.finishButton = QPushButton("Finish")
        self.finishButton.clicked.connect(self._on_finish)
        style_button(self.finishButton)
        button_layout.addWidget(self.finishButton)
        button_layout.setAlignment(Qt.AlignCenter)
        self.mainGrid.addLayout(button_layout, 2, 0)

    def _on_checkbox_state_changed(self, state, label, section_idx, local_state, i):
        local_state[i] = label.text() if state == Qt.Checked else ""
        self.selectedItems[section_idx] = "\n".join(filter(None, local_state))

    def _on_finish(self):
        self.mainFrame.setCurrentIndex(0)

    def get_window(self):
        return self.root

    def run(self):
        self.root.show()
