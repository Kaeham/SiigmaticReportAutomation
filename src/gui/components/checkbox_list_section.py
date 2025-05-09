from PyQt5.QtWidgets import QFrame, QVBoxLayout, QLabel, QCheckBox, QHBoxLayout, QSizePolicy

def create_checkbox_section(title, items, section_idx, on_check_changed):

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
    """)  # Keep your existing stylesheet

    layout = QVBoxLayout()
    titleLabel = QLabel(title)
    titleLabel.setProperty("role", "title")
    layout.addWidget(titleLabel)

    local_selection = [""] * len(items)
    
    items = items if isinstance(items, list) else [items]
    for i, text in enumerate(items):
        row = QHBoxLayout()
        checkbox = QCheckBox()
        checkbox.setMaximumWidth(24)
        label = QLabel(text)
        label.setWordWrap(True)
        label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)

        checkbox.stateChanged.connect(
            lambda state, lbl=label, i=i: on_check_changed(state, lbl, section_idx, local_selection, i)
        )

        row.addWidget(checkbox)
        row.addWidget(label)
        layout.addLayout(row)

    frame.setLayout(layout)
    return frame
