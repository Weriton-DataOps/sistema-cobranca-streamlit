import streamlit as st


# CONFIG
st.set_page_config(page_title="Dashboard FA", page_icon="🏖️", layout="wide")

# MENU
menu = st.sidebar.selectbox(
    "Navegação",
    ["Página Inicial", "Relatório Geral", "Gestão de Lote", "Acionamento"]
)

# PÁGINA INICIAL
if menu == "Página Inicial":
    st.title("🏠 Página Inicial")
    st.write("Use o menu lateral para navegar entre:")
    st.markdown("- Relatório Geral")
    st.markdown("- Gestão de Lote")
    st.markdown("- Acionamento")
