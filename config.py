"""
Configuration file for Witness Evaluation Tool
Customize colors, styles, and settings here
"""

# Document Settings
DOCUMENT_CONFIG = {
    'margins': {
        'top': 0.5,      # inches
        'bottom': 0.5,
        'left': 0.75,
        'right': 0.75
    },
    'title_font_size': 14,  # points
    'body_font_size': 11,
    'table_style': 'Light Grid Accent 1'  # Word table style name
}

# Chart Settings
CHART_CONFIG = {
    'figure_size': (9, 6),  # width, height in inches - WIDER for better label spacing
    'dpi': 300,
    'colors': {
        'plaintiff': '#DC143C',   # Crimson Red
        'defendant': '#4169E1',   # Royal Blue
        'other': '#808080'        # Grey
    },
    'font_sizes': {
        'labels': 12,
        'ylabel': 13,
        'percentage': 14
    }
}

# Value Mapping
VALUE_LABELS = {
    1: 'Not at all',
    2: 'Not very',
    3: 'Somewhat',
    4: 'Very',
    5: 'Extremely'
}

# Text Templates
TEXT_TEMPLATES = {
    'task1_intro': 'For each characteristic, please circle the number that best expresses your opinion about this witness. The witness, {witness_name}, was:',
    'task2_intro': "Which side did {witness_name}'s testimony help the most?",
    'sample_size': '(n = {n})'
}

# Characteristic Parsing
CHARACTERISTIC_PARSER = {
    'delimiter': ': -',  # Delimiter to split on
}

# Color Keywords (case-insensitive matching)
SIDE_KEYWORDS = {
    'plaintiff': ['plaintiff', "plaintiff's"],
    'defendant': ['defendant', 'defendants', "defendant's", "defendants'"]
}

# File Upload Settings
UPLOAD_CONFIG = {
    'max_file_size_mb': 50,
    'allowed_extensions': ['.sav'],
    'upload_folder': '/tmp/uploads'
}

# Table Column Headers
TABLE_HEADERS = ['Mean', '', 'Not at all', 'Not very', 'Somewhat', 'Very', 'Extremely']

# Date Format
DATE_FORMAT = '%B %d, %Y'  # e.g., "November 21, 2024"