import streamlit as st
# IMPORTANTE: IMPORTS DAS SUAS PÁGINAS
from pages.relatorioGeral import relatorio_geral
from pages.gestaoLote import show_lote
from pages.acionamento import show_acionamento

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

# ROTAS
elif menu == "Relatório Geral":
    relatorio_geral()

elif menu == "Gestão de Lote":
    show_lote()

elif menu == "Acionamento":
    show_acionamento()