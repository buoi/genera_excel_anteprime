# styles.py

# Drop area styles
DROP_AREA_NORMAL = """
    background-color: #F7F7FF;
    border: 2px dashed #B8C4FF;
    border-radius: 8px;
    padding: 20px;
    font-size: 18px;
    color: #666677;
"""

DROP_AREA_SUCCESS = """
    background-color: #F0FFF0; 
    border: 2px dashed #90EE90;
    border-radius: 8px;
    padding: 20px;
    font-size: 18px;
    color: #666677;
"""

DROP_AREA_ERROR = """
    background-color: #FFF0F0;
    border: 2px dashed #FFB6C1;
    border-radius: 8px;
    padding: 20px;
    font-size: 18px;
    color: #666677;
"""

# Button styles
BROWSE_BUTTON = """
    background-color: #E0E4FF;
    color: #444455;
    border: none;
    border-radius: 4px;
    padding: 8px 15px;
    font-size: 14px;
"""

GENERATE_BUTTON = """
    background-color: #B8C4FF;
    color: #444455;
    border: none;
    border-radius: 4px;
    padding: 10px 20px;
    font-size: 14px;
    font-weight: bold;
"""

# Input field styles
PATH_INPUT = """
    border: 1px solid #B8C4FF;
    border-radius: 4px;
    padding: 8px;
    font-size: 14px;
"""

# Text styles

TITLE_TEXT = """
    color: #444455;
    font-weight: bold;
    font-size: 24pt;
    text-align: center;
"""

DESCRIPTION_TEXT = """
    color: #666677;
    font-size: 14px;
    margin-bottom: 5px;
    text-align: center;
"""

FORMATTED_INFO = """
    color: #444455;
    font-size: 14px;
    margin-bottom: 5px;
"""

SUPPORT_INFO = """
    color: #888899;
    font-size: 12px;
    font-style: italic;
    margin-bottom: 10px;
"""

FOOTER_TEXT = """
    color: #888899;
    font-style: italic;
"""

INFO_TEXT = """
    color: #888899;
    font-size: 12px;
    font-style: italic;
    margin-top: 5px;
    margin-bottom: 15px;
"""

# Status text area
STATUS_TEXT = """
    background-color: #FFFFFF;
    border: 1px solid #B8C4FF;
    border-radius: 4px;
    padding: 5px;
    font-size: 12px;
    color: #444455;
"""

STATUS_TEXT_ERROR = """
    background-color: #FFF0F0;
    border: 1px solid #FFCCCC;
    border-radius: 4px;
    padding: 5px;
    font-size: 12px;
    color: #CC0000;
"""

# Progress bar
PROGRESS_BAR = """
    QProgressBar {
        border: 1px solid #B8C4FF;
        border-radius: 4px;
        text-align: center;
        padding: 2px;
        background-color: #FFFFFF;
    }
    
    QProgressBar::chunk {
        background-color: #B8C4FF;
        width: 10px;
        margin: 0px;
    }
"""

# Status indicators
STATUS_INDICATOR_WAITING = """
    background-color: #F0F0F0;
    color: #888888;
    border: none;
    border-radius: 4px;
    padding: 5px 10px;
    font-size: 12px;
"""

STATUS_INDICATOR_SUCCESS = """
    background-color: #E0FFE0;
    color: #006600;
    border: none;
    border-radius: 4px;
    padding: 5px 10px;
    font-size: 12px;
"""

STATUS_INDICATOR_ERROR = """
    background-color: #FFE0E0;
    color: #660000;
    border: none;
    border-radius: 4px;
    padding: 5px 10px;
    font-size: 12px;
"""