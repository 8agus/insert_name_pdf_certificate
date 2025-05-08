# Certificate Generator

This project generates PDF certificates from a template by adding participant names from an Excel file.

## Setup

### Virtual Environment

```bash
# Create virtual environment
python -m venv env

# Activate virtual environment
# On Windows:
env\Scripts\activate
# On macOS/Linux:
source env/bin/activate

# Install dependencies
pip install -r requirements.txt
```

## Configuration

Edit these variables in `pdf_certificate.py` to customize the application:

```python
# File paths
excel_path = "IC-Juniors-Participants.xlsx"  # Path to Excel file with participant data
template_pdf_path = "Imagine Cup Junior Participation-2025.pdf"  # Path to certificate template
output_dir = 'certificates'  # Directory where certificates will be saved
```

### Text Positioning

Adjust these variables to position the text correctly on your certificate:

```python
# In the create_certificate_pdf function
x_center = page_width * 0.5 - (text_width / 2)  # Horizontal centering
y_position = 280  # Vertical position (lower value = lower on page)
```

## Excel Format Requirements

The Excel file should contain columns with these headers (case-insensitive):
- ID or Participant ID or Student ID
- Name (or any column with "name" in its title)
- Email (optional, for future email functionality)

## Usage

Run the script:
```bash
python3 pdf_certificate.py
```

You can choose to run in test mode first (processes only first 5 entries).

## Output

- Generated certificates are saved to the output directory
- A log file is generated at `output_dir/pdf_certificate_generation.log`

## New Features

### Modular Text Positioning

The text positioning logic has been moved to a separate file, `positioning.py`, for easier customization. Users can now adjust the positioning options and fine-tune the placement of names on certificates by modifying the following variables in `positioning.py`:

```python
# User-configurable variables for positioning
selected_position = "right_higher"  # Options: "default", "right_higher"
up = 0    # Move the position up by this many units
down = 0  # Move the position down by this many units
left = 0  # Move the position left by this many units
right = 0 # Move the position right by this many units
```

### Positioning Options

The available positioning options are defined in `positioning.py`:

- **default**: Centered horizontally and slightly below the middle vertically.
- **right_higher**: Shifted slightly to the right and 20 units higher.

### How to Use

1. Open `positioning.py`.
2. Modify the `selected_position` variable to choose a predefined option.
3. Adjust the `up`, `down`, `left`, and `right` variables to fine-tune the position.

The main script, `pdf_certificate.py`, now automatically uses these settings when generating certificates.
