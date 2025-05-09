def format_button(widget):
    widget.setStyleSheet("""
            QPushButton {
                background-color: #e3d917;
                color: black;
                padding: 6px 12px;
                font-weight: bold;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #f5f1a4;
            }""")

def format_label_text(widget, font_size=12):
    widget.setStyleSheet(f"""
            QLabel {{
                    font-size: {font_size}px;
                    a
                    b
                    c
                    d
                    
            }}
""")
    
def format_title(widget, font_size=14):
    widget.setStyleSheet(f"""
            QLabel {{
                font-weight: bold;
                font-size: {font_size}px;
                background-color: #e3d917;
                border-radius: 6px;
                    }}""")