import streamlit as st

def configure_page():
    st.set_page_config(page_title="Control Center", layout="wide")

def load_css(path):
    try:
        with open(path) as f:
            st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)
    except:
        pass

def render_topbar():
    st.markdown("## 🏭 Control Center")
