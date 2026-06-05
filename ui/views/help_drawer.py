import streamlit as st
import os

def get_formatted_documentation_markdown():
    """Returns the full technical guide content formatted as premium Markdown, loaded dynamically from README.md."""
    try:
        filepath = "README.md"
        if os.path.exists(filepath):
            with open(filepath, "r", encoding="utf-8") as f:
                return f.read()
    except Exception as e:
        st.warning(f"Не вдалося завантажити файл довідки: {e}")
    
    return "### 📖 Документація\nНе вдалося завантажити файл довідки `README.md`. Будь ласка, переконайтеся, що файл існує у робочій папці."

def render_help_view():
    """Renders the help and user guide view wrapped in a clean, non-shifting card."""
    st.header("📖 Повний посібник користувача")
    st.write("Детальний опис можливостей та посібник роботи комбайна (завантажено з README.md).")
    
    st.markdown("---")
    
    doc_markdown = get_formatted_documentation_markdown()
    
    # Render inside a styled glassmorphic card container without affecting the main page header
    st.markdown(f"""
<div class="card" style="padding: 3rem; margin-top: 1.5rem; margin-bottom: 2.5rem;">

{doc_markdown}

</div>
""", unsafe_allow_html=True)
