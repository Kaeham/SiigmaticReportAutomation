def style_button(button):
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

def format_header_label(label, font_size=12):
    label.setStyleSheet(f"""
        QLabel {{
            font-weight: bold;
            font-size: {font_size}pt;
            background-color: #f7f779;
            padding: 4px 6px;
            border-radius: 4px;
        }}
    """)
