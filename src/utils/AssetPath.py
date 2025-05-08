import os
import sys

def get_asset_path(filename):
    if getattr(sys, 'frozen', False):
        base = sys._MEIPASS
    else:
        base = os.path.dirname(__file__)
    
    return os.path.join(base, filename)