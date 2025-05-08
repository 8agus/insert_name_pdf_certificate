# User-configurable variables for positioning
selected_position = "right_higher"  # Change this to "right_higher"/"default" to use the desired option
up = 30    # Move the position up by this many units
down = 0  # Move the position down by this many units
left = 0  # Move the position left by this many units
right = 0 # Move the position right by this many units

# This module handles the positioning logic for text placement on certificates.

def get_positioning_options():
    """Return available positioning options."""
    return {
        "default": {
            "description": "Centered horizontally and slightly below the middle vertically.",
            "x_center_factor": 0.5,
            "y_position": 320
        },
        "right_higher": {
            "description": "Shifted slightly to the right and 20 units higher.",
            "x_center_factor": 0.675,  # Between the previous values of 0.65 and 0.7
            "y_position": 260          # 20 units higher
        }
    }

# Corrected import for stringWidth
from reportlab.pdfbase.pdfmetrics import stringWidth

def calculate_position(text, font_name, font_size, page_width, selected_position):
    """Calculate the x_center and y_position for text placement."""
    positions = get_positioning_options()
    if selected_position not in positions:
        raise ValueError(f"Invalid position selected: {selected_position}")

    position = positions[selected_position]
    text_width = stringWidth(text, font_name, font_size)
    x_center = page_width * position["x_center_factor"] - (text_width / 2) + right - left
    y_position = position["y_position"] + up - down

    return x_center, y_position