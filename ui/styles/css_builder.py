def get_custom_css(theme):
    """Generates dynamic, premium glassmorphic styling based on theme."""
    if theme == "dark":
        bg_color = "#0f172a"
        text_color = "#f8fafc"
        card_bg = "rgba(30, 41, 59, 0.7)"
        card_border = "rgba(51, 65, 85, 0.8)"
        subtitle_color = "#94a3b8"
        input_border = "#334155"
        shadow_color = "rgba(0, 0, 0, 0.3)"
        popover_bg = "#1e293b"
        placeholder_color = "#94a3b8"
        alert_bg = "#1e293b"
        alert_border = "#334155"
    else:
        bg_color = "#f8fafc"
        text_color = "#0f172a"
        card_bg = "rgba(255, 255, 255, 0.8)"
        card_border = "rgba(226, 232, 240, 0.8)"
        subtitle_color = "#64748b"
        input_border = "#e2e8f0"
        shadow_color = "rgba(31, 38, 135, 0.05)"
        popover_bg = "#ffffff"
        placeholder_color = "#64748b"
        alert_bg = "rgba(241, 245, 249, 0.8)"
        alert_border = "rgba(226, 232, 240, 0.8)"

    css = f"""
<style>
    /* Google Fonts */
    @import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;700&display=swap');
    
    :root {{
        --background-color: {bg_color};
        --text-color: {text_color};
        --secondary-background-color: {popover_bg};
    }}
    
    /* App global background */
    .stApp {{
        background-color: {bg_color} !important;
    }}
    
    html, body {{
        background-color: {bg_color} !important;
        color: {text_color} !important;
        transition: background-color 0.3s ease, color 0.3s ease;
    }}
    
    /* Hide Streamlit header and footer & remove top spacing */
    [data-testid="stHeader"], footer {{
        display: none !important;
    }}
    
    [data-testid="stMainBlockContainer"] {{
        margin-top: 0px !important;
    }}

    .block-container {{
        padding-top: 3rem !important;
    }}
    
    html, body, [class*="css"] {{
        font-family: 'Outfit', sans-serif;
    }}
    
    h1:not(.main-title), h2, h3, h4, h5, h6, p, span, label, li {{
        color: {text_color} !important;
    }}
    
    code, pre {{
        font-family: 'JetBrains Mono', monospace !important;
        background-color: {popover_bg} !important;
        color: {text_color} !important;
        border: 1px solid {card_border} !important;
        border-radius: 6px !important;
        padding: 0.2rem 0.4rem !important;
    }}
    
    pre {{
        padding: 1rem !important;
        overflow-x: auto !important;
        background-color: {popover_bg} !important;
    }}
    
    pre code {{
        background-color: transparent !important;
        border: none !important;
        padding: 0 !important;
    }}
    
    /* Elegant Title and Badges */
    .main-title {{
        background: linear-gradient(135deg, #4A90E2 0%, #50E3C2 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-size: 3rem;
        font-weight: 700;
        margin-bottom: 0.2rem;
        letter-spacing: -0.5px;
    }}
    /* Premium styled sidebar */
    [data-testid="stSidebar"] {{
        background-color: {popover_bg} !important;
        border-right: 1px solid {card_border} !important;
    }}
    
    [data-testid="stSidebar"] .st-emotion-cache-16txtl3,
    [data-testid="stSidebar"] [data-testid="stSidebarNav"] {{
        background-color: transparent !important;
    }}
    
    /* Hide sidebar completely when collapsed in Streamlit */
    [data-testid="stSidebar"][data-collapsed="true"],
    [data-testid="stSidebar"][aria-expanded="false"] {{
        display: none !important;
        width: 0px !important;
    }}
    
    [data-testid="stSidebar"] > div, [data-testid="stSidebarUserContent"] {{
        background-color: {popover_bg} !important;
        overflow-y: auto !important;
    }}
    
    /* Main section styling */
    section[data-testid="stMain"] {{
        background-color: transparent !important;
    }}
    
    /* Force resizer / drag handle to sit on top of the sidebar and be clickable */
    [data-testid="stSidebarResizer"], [data-testid="stSidebarDragHandle"] {{
        z-index: 10000 !important; /* Higher than stSidebar */
    }}
    
    /* Dialog/Modal Styling for Dark Theme */
    div[role="dialog"], [data-testid="stModal"], div[data-testid="stDialog"] {{
        background-color: {popover_bg} !important;
        color: {text_color} !important;
        border: 1px solid {card_border} !important;
        border-radius: 16px !important;
    }}
    
    div[role="dialog"] > div, [data-testid="stDialog"] > div {{
        background-color: transparent !important;
    }}

    [data-testid="stSidebarResizer"]:hover, [data-testid="stSidebarDragHandle"]:hover {{
        background-color: #4A90E2 !important;
    }}
    
    /* Hide the native Streamlit expand button (when collapsed) */
    [data-testid="collapsedControl"] {{
        display: none !important;
    }}
    
    /* Hide the native Streamlit collapse button (when expanded) */
    [data-testid="stSidebarCollapseButton"] {{
        display: none !important;
    }}
    .subtitle {{
        color: {subtitle_color} !important;
        font-size: 1.15rem;
        margin-bottom: 2rem;
        font-weight: 300;
    }}
    
    /* Premium Styled Card Container */
    .card {{
        background: {card_bg};
        backdrop-filter: blur(10px);
        -webkit-backdrop-filter: blur(10px);
        border-radius: 16px;
        box-shadow: 0 8px 32px 0 {shadow_color};
        padding: 1.8rem;
        margin-bottom: 1.8rem;
        border: 1px solid {card_border};
        color: {text_color};
        transition: all 0.3s ease;
    }}
    
    /* Styled widgets & alerts */
    .stAlert, div[data-testid="stAlert"] {{
        border-radius: 12px !important;
        border: 1px solid {alert_border} !important;
        background-color: {alert_bg} !important;
        color: {text_color} !important;
        box-shadow: 0 4px 12px rgba(0,0,0,0.02) !important;
    }}
    
    /* Button custom hover effects */
    button[kind="primary"] {{
        background: linear-gradient(135deg, #4A90E2 0%, #357ABD 100%) !important;
        border: none !important;
        color: white !important;
        box-shadow: 0 4px 15px rgba(74, 144, 226, 0.3) !important;
        transition: all 0.25s ease-in-out !important;
    }}
    button[kind="primary"]:hover {{
        transform: translateY(-2px) !important;
        box-shadow: 0 6px 20px rgba(74, 144, 226, 0.4) !important;
    }}
    
    button[kind="secondary"] {{
        border-radius: 8px !important;
        border: 1px solid {input_border} !important;
        background-color: {popover_bg} !important;
        color: {text_color} !important;
        transition: all 0.2s ease !important;
    }}
    button[kind="secondary"]:hover {{
        border-color: #4A90E2 !important;
        color: #4A90E2 !important;
        background-color: rgba(74, 144, 226, 0.03) !important;
    }}
    
    /* Code block copy toolbar buttons styled */
    button[kind="elementToolbar"], button[data-testid="stBaseButton-elementToolbar"] {{
        background-color: {popover_bg} !important;
        color: {text_color} !important;
        border: 1px solid {card_border} !important;
        border-radius: 6px !important;
    }}
    button[kind="elementToolbar"] svg, button[data-testid="stBaseButton-elementToolbar"] svg {{
        fill: {text_color} !important;
        color: {text_color} !important;
    }}
    button[kind="elementToolbar"]:hover, button[data-testid="stBaseButton-elementToolbar"]:hover {{
        background-color: #4A90E2 !important;
        color: white !important;
        border-color: #4A90E2 !important;
    }}
    button[kind="elementToolbar"]:hover svg, button[data-testid="stBaseButton-elementToolbar"]:hover svg {{
        fill: white !important;
        color: white !important;
    }}
    
    /* Tooltip popup styled */
    div[data-baseweb="tooltip"],
    div[data-baseweb="tooltip"] *,
    div[data-testid="stTooltipContent"],
    div[data-testid="stTooltipContent"] * {{
        background-color: {popover_bg} !important;
        color: {text_color} !important;
    }}
    div[data-baseweb="tooltip"] {{
        border: 1px solid {card_border} !important;
        border-radius: 8px !important;
        box-shadow: 0 4px 12px {shadow_color} !important;
    }}
    
    .badge-icon {{
        font-size: 1.5rem;
        margin-right: 0.5rem;
    }}
    
    /* Input fields and Selectboxes custom styling */
    .stTextInput input,
    .stSelectbox div[data-baseweb="select"],
    .stSelectbox div[data-baseweb="select"] > div,
    .stSelectbox div[role="combobox"],
    .stSelectbox div[role="combobox"] > div,
    .stNumberInput input,
    .stTextArea textarea,
    .stMultiSelect div[data-baseweb="select"],
    .stMultiSelect div[data-baseweb="select"] > div {{
        border-color: {input_border} !important;
        background-color: {popover_bg} !important;
        color: {text_color} !important;
    }}
    
    /* Input selected text color and control elements */
    .stSelectbox div[data-baseweb="select"] *, 
    .stSelectbox div[role="combobox"] *,
    .stMultiSelect div[data-baseweb="select"] * {{
        background-color: {popover_bg} !important;
        color: {text_color} !important;
    }}
    
    /* Input field placeholders styling */
    input::placeholder, textarea::placeholder {{
        color: {placeholder_color} !important;
        opacity: 0.75 !important;
    }}
    input::-webkit-input-placeholder, textarea::-webkit-input-placeholder {{
        color: {placeholder_color} !important;
        opacity: 0.75 !important;
    }}
    
    /* baseweb popover dropdown lists styled */
    div[data-baseweb="popover"], div[data-baseweb="menu"], ul[role="listbox"], li[role="option"] {{
        background-color: {popover_bg} !important;
        color: {text_color} !important;
    }}
    
    div[data-baseweb="popover"] ul, div[data-baseweb="popover"] li, div[data-baseweb="popover"] span {{
        background-color: {popover_bg} !important;
        color: {text_color} !important;
    }}
    
    /* baseweb hover highlights */
    div[data-baseweb="popover"] li:hover,
    div[data-baseweb="popover"] li[aria-selected="true"],
    div[data-baseweb="popover"] li:hover * {{
        background-color: #4A90E2 !important;
        color: #ffffff !important;
    }}
    
    /* Document Preview Container (Sheet/Word) Overrides */
    div.document-preview-container,
    div.document-preview-container *,
    div.word-preview-container,
    div.word-preview-container *,
    .document-preview-container p,
    .document-preview-container span,
    .document-preview-container td,
    .document-preview-container th,
    .document-preview-container tr,
    .document-preview-container table,
    .document-preview-container h1,
    .document-preview-container h2,
    .document-preview-container h3,
    .document-preview-container h4,
    .document-preview-container h5,
    .document-preview-container h6,
    .document-preview-container li,
    .document-preview-container ul,
    .word-preview-container p,
    .word-preview-container span,
    .word-preview-container td,
    .word-preview-container th,
    .word-preview-container tr,
    .word-preview-container table,
    .word-preview-container h1,
    .word-preview-container h2,
    .word-preview-container h3,
    .word-preview-container h4,
    .word-preview-container h5,
    .word-preview-container h6,
    .word-preview-container li,
    .word-preview-container ul {{
        color: #1a202c !important;
    }}
    div.document-preview-container,
    div.word-preview-container,
    div.document-preview-container div,
    div.word-preview-container div {{
        background-color: #ffffff !important;
    }}
    
    /* Details & Expanders */
    div[data-testid="stExpander"] {{
        background-color: {card_bg} !important;
        border: 1px solid {card_border} !important;
        border-radius: 8px !important;
        margin-bottom: 0.5rem !important;
    }}
    
    /* Force details and open details to be transparent to show the wrapper card background */
    div[data-testid="stExpander"] details,
    div[data-testid="stExpander"] details[open],
    div[data-testid="stExpander"] summary,
    div[data-testid="stExpander"] div[role="region"],
    div[data-testid="stExpander"] [data-testid="stExpanderDetails"],
    div[data-testid="stExpander"] .streamlit-expanderContent,
    div[data-testid="stExpander"] div[role="region"] [data-testid="stVerticalBlock"] {{
        background-color: transparent !important;
        background: transparent !important;
        border: none !important;
    }}
    
    /* Make sure all descendant container divs inside the expander are transparent,
       excluding actual widgets like input fields, selectboxes, buttons, alerts and nested expanders */
    div[data-testid="stExpander"] div[role="region"] div:not(.stTextInput):not(.stSelectbox):not(.stTextArea):not(.stButton):not(.stNumberInput):not(.stMultiSelect):not([data-testid="stExpander"]):not(.stAlert):not([data-testid="stAlert"]) {{
        background-color: transparent !important;
        background: transparent !important;
    }}
    
    /* Summary headers inside expanders */
    div[data-testid="stExpander"] > details > summary,
    div[data-testid="stExpander"] > details > summary * {{
        background-color: transparent !important;
        color: {text_color} !important;
    }}
    
    /* Make sure nested expanders restore their background and don't remain transparent */
    div[data-testid="stExpander"] div[role="region"] div[data-testid="stExpander"] {{
        background-color: {card_bg} !important;
        border: 1px solid {card_border} !important;
    }}
    
    /* Scoped card styling for the help page content container */
    div[data-testid="stVerticalBlockBorderWrapper"]:has(.help-page-marker) {{
        background: {card_bg} !important;
        backdrop-filter: blur(10px) !important;
        -webkit-backdrop-filter: blur(10px) !important;
        border: 1px solid {card_border} !important;
        border-radius: 16px !important;
        box-shadow: 0 8px 32px 0 {shadow_color} !important;
        border-radius: 8px !important;
        border: 1px solid {input_border} !important;
        background-color: {popover_bg} !important;
        color: {text_color} !important;
        transition: all 0.2s ease !important;
    }}
    button[kind="secondary"]:hover {{
        border-color: #4A90E2 !important;
        color: #4A90E2 !important;
        background-color: rgba(74, 144, 226, 0.03) !important;
    }}
    
    /* Code block copy toolbar buttons styled */
    button[kind="elementToolbar"], button[data-testid="stBaseButton-elementToolbar"] {{
        background-color: {popover_bg} !important;
        color: {text_color} !important;
        border: 1px solid {card_border} !important;
        border-radius: 6px !important;
    }}
    button[kind="elementToolbar"] svg, button[data-testid="stBaseButton-elementToolbar"] svg {{
        fill: {text_color} !important;
        color: {text_color} !important;
    }}
    button[kind="elementToolbar"]:hover, button[data-testid="stBaseButton-elementToolbar"]:hover {{
        background-color: #4A90E2 !important;
        color: white !important;
        border-color: #4A90E2 !important;
    }}
    button[kind="elementToolbar"]:hover svg, button[data-testid="stBaseButton-elementToolbar"]:hover svg {{
        fill: white !important;
        color: white !important;
    }}
    
    /* Tooltip popup styled */
    div[data-baseweb="tooltip"],
    div[data-baseweb="tooltip"] *,
    div[data-testid="stTooltipContent"],
    div[data-testid="stTooltipContent"] * {{
        background-color: {popover_bg} !important;
        color: {text_color} !important;
    }}
    div[data-baseweb="tooltip"] {{
        border: 1px solid {card_border} !important;
        border-radius: 8px !important;
        box-shadow: 0 4px 12px {shadow_color} !important;
    }}
    
    .badge-icon {{
        font-size: 1.5rem;
        margin-right: 0.5rem;
    }}
    
    /* Input fields and Selectboxes custom styling */
    .stTextInput input,
    .stSelectbox div[data-baseweb="select"],
    .stSelectbox div[data-baseweb="select"] > div,
    .stSelectbox div[role="combobox"],
    .stSelectbox div[role="combobox"] > div,
    .stNumberInput input,
    .stTextArea textarea,
    .stMultiSelect div[data-baseweb="select"],
    .stMultiSelect div[data-baseweb="select"] > div {{
        border-color: {input_border} !important;
        background-color: {popover_bg} !important;
        color: {text_color} !important;
    }}
    
    /* Input selected text color and control elements */
    .stSelectbox div[data-baseweb="select"] *, 
    .stSelectbox div[role="combobox"] *,
    .stMultiSelect div[data-baseweb="select"] * {{
        background-color: {popover_bg} !important;
        color: {text_color} !important;
    }}
    
    /* Input field placeholders styling */
    input::placeholder, textarea::placeholder {{
        color: {placeholder_color} !important;
        opacity: 0.75 !important;
    }}
    input::-webkit-input-placeholder, textarea::-webkit-input-placeholder {{
        color: {placeholder_color} !important;
        opacity: 0.75 !important;
    }}
    
    /* baseweb popover dropdown lists styled */
    div[data-baseweb="popover"], div[data-baseweb="menu"], ul[role="listbox"], li[role="option"] {{
        background-color: {popover_bg} !important;
        color: {text_color} !important;
    }}
    
    div[data-baseweb="popover"] ul, div[data-baseweb="popover"] li, div[data-baseweb="popover"] span {{
        background-color: {popover_bg} !important;
        color: {text_color} !important;
    }}
    
    /* Fix buttons inside popovers having white background under text */
    div[data-baseweb="popover"] button span,
    div[data-baseweb="popover"] button div,
    div[data-baseweb="popover"] button p,
    div[data-baseweb="popover"] button * {{
        background-color: transparent !important;
        color: inherit !important;
    }}
    
    /* baseweb hover highlights */
    div[data-baseweb="popover"] li:hover,
    div[data-baseweb="popover"] li[aria-selected="true"],
    div[data-baseweb="popover"] li:hover * {{
        background-color: #4A90E2 !important;
        color: #ffffff !important;
    }}
    
    /* Document Preview Container (Sheet/Word) Overrides */
    div.document-preview-container,
    div.document-preview-container *,
    div.word-preview-container,
    div.word-preview-container *,
    .document-preview-container p,
    .document-preview-container span,
    .document-preview-container td,
    .document-preview-container th,
    .document-preview-container tr,
    .document-preview-container table,
    .document-preview-container h1,
    .document-preview-container h2,
    .document-preview-container h3,
    .document-preview-container h4,
    .document-preview-container h5,
    .document-preview-container h6,
    .document-preview-container li,
    .document-preview-container ul,
    .word-preview-container p,
    .word-preview-container span,
    .word-preview-container td,
    .word-preview-container th,
    .word-preview-container tr,
    .word-preview-container table,
    .word-preview-container h1,
    .word-preview-container h2,
    .word-preview-container h3,
    .word-preview-container h4,
    .word-preview-container h5,
    .word-preview-container h6,
    .word-preview-container li,
    .word-preview-container ul {{
        color: #1a202c !important;
    }}
    div.document-preview-container,
    div.word-preview-container,
    div.document-preview-container div,
    div.word-preview-container div {{
        background-color: #ffffff !important;
    }}
    
    /* Details & Expanders */
    div[data-testid="stExpander"] {{
        background-color: {card_bg} !important;
        border: 1px solid {card_border} !important;
        border-radius: 8px !important;
        margin-bottom: 0.5rem !important;
    }}
    
    /* Force details and open details to be transparent to show the wrapper card background */
    div[data-testid="stExpander"] details,
    div[data-testid="stExpander"] details[open],
    div[data-testid="stExpander"] summary,
    div[data-testid="stExpander"] div[role="region"],
    div[data-testid="stExpander"] [data-testid="stExpanderDetails"],
    div[data-testid="stExpander"] .streamlit-expanderContent,
    div[data-testid="stExpander"] div[role="region"] [data-testid="stVerticalBlock"] {{
        background-color: transparent !important;
        background: transparent !important;
        border: none !important;
    }}
    
    /* Make sure all descendant container divs inside the expander are transparent,
       excluding actual widgets like input fields, selectboxes, buttons, alerts and nested expanders */
    div[data-testid="stExpander"] div[role="region"] div:not(.stTextInput):not(.stSelectbox):not(.stTextArea):not(.stButton):not(.stNumberInput):not(.stMultiSelect):not([data-testid="stExpander"]):not(.stAlert):not([data-testid="stAlert"]) {{
        background-color: transparent !important;
        background: transparent !important;
    }}
    
    /* Summary headers inside expanders */
    div[data-testid="stExpander"] > details > summary,
    div[data-testid="stExpander"] > details > summary * {{
        background-color: transparent !important;
        color: {text_color} !important;
    }}
    
    /* Make sure nested expanders restore their background and don't remain transparent */
    div[data-testid="stExpander"] div[role="region"] div[data-testid="stExpander"] {{
        background-color: {card_bg} !important;
        border: 1px solid {card_border} !important;
    }}
    
    /* Scoped card styling for the help page content container */
    div[data-testid="stVerticalBlockBorderWrapper"]:has(.help-page-marker) {{
        background: {card_bg} !important;
        backdrop-filter: blur(10px) !important;
        -webkit-backdrop-filter: blur(10px) !important;
        border: 1px solid {card_border} !important;
        border-radius: 16px !important;
        box-shadow: 0 8px 32px 0 {shadow_color} !important;
        padding: 3rem !important;
        margin-top: 1.5rem !important;
        margin-bottom: 2.5rem !important;
        color: {text_color} !important;
    }}
    /* Hide the pop-up toolbar on tables and charts */
    [data-testid="stElementToolbar"] {{
        display: none !important;
    }}

</style>
"""
    return css
