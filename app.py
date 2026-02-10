import streamlit as st

st.set_page_config(
    page_title="Test Streamlit",
    layout="centered"
)

st.title("✅ Streamlit fonctionne")
st.write("Si tu vois ce texte, le serveur Streamlit est OK.")

st.button("Bouton test")
