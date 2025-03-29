# Utility functions that are useful within the tooling.

import pyfiglet
from pyfiglet import FigletString

def print_title(title_string: str):
    "Create a nice pretty ascii art image to populate into CLI"
    
    font = pyfiglet.Figlet()
    
    ascii_art = font.renderText(title_string)
    print(ascii_art)
    