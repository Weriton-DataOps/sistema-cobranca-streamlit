import pandas as pd
from datetime import datetime 
from streamlit_tags import st_tags
import os
import win32com.client as win32
import shutil
import numpy as np
from pathlib import Path
# streamlit_app_title: 🧾 Gestão de Lote
import streamlit as st


st.set_page_config(
    page_title="Relatório Geral Fã", 
    page_icon="🏖️",                 
    layout="wide"
)

def Menu():
    st.sidebar.title("MENU")
    if st.sidebar.button("🔁 Atualizar Dados"):
        st.session_state['atualizar_dados'] = True

    if st.session_state.get('atualizar_dados', False):
        st.cache_data.clear()
        st.session_state['atualizar_dados'] = False
        st.rerun()


@st.cache_data
def dados_lote():
    pasta_base = "data/LOTES"
    subpastas = ["ATIVOS", "EXPIRADOS"]
    dataframes = []

    for subpasta in subpastas:
        pasta = os.path.join(pasta_base, subpasta)
        for arquivo in os.listdir(pasta):
            if arquivo.endswith('.xlsx'):
                caminho_arquivo = os.path.join(pasta, arquivo)
                try:
                    df = pd.read_excel(caminho_arquivo, sheet_name="Lote")
                    df.columns = df.columns.str.strip()

                    partes = arquivo.replace(".xlsx", "").split("_")
                    mes_ano = partes[1] if len(partes) > 1 else ""
                    status_lote = partes[-1] if len(partes) > 2 else ""

                    df['arquivo_origem'] = arquivo
                    df['Mês_Ano'] = mes_ano
                    df['Status_Lote'] = status_lote

                    dataframes.append(df)
                except Exception as e:
                    print(f"Erro ao ler {arquivo}: {e}")

    if not dataframes:
        st.warning("⚠️ Nenhum lote foi encontrado nas pastas especificadas.")
        return pd.DataFrame()

    df_lote = pd.concat(dataframes, ignore_index=True)

    colunas_essenciais = [
        "Passaporte", "Fornecedor", "Id", "Vencimento", "Faixa", "Tipo", "Valor", "Meta", "Status",
        "Consultor", "StatusAc.", "Data Rec.", "Valor Rec.", "MeioPag.", "Valor EmDia", "Observação", "ValorBaixado",
        "arquivo_origem", "Mês_Ano", "Status_Lote"
    ]

    colunas_presentes = [col for col in colunas_essenciais if col in df_lote.columns]
    df_lote = df_lote[colunas_presentes]

    if "Faixa" in df_lote.columns:
        df_lote = df_lote[df_lote["Faixa"] != "MODELO"]

    return df_lote





def filtros(df_lote):
    # ✅ Garante que as colunas existam
    colunas_a_garantir = [
        "StatusAc.", "Data Rec.", "Valor Rec.", "MeioPag.",
        "Valor EmDia", "Observação", "ValorBaixado"
    ]
    for col in colunas_a_garantir:
        if col not in df_lote.columns:
            df_lote[col] = np.nan

    col1,col2 = st.columns([1,7])
    with col1:
        status_lote = status_lote =sorted(df_lote['Status_Lote'].unique())
        status_lote_opcoes = ["Todos"] + status_lote
        index_ativos = status_lote_opcoes.index("ATIVO") if "ATIVO" in status_lote_opcoes else 0
        filtro_status_lote = st.selectbox("✅ Status Lote", status_lote_opcoes, index=index_ativos)
        if filtro_status_lote != "Todos":
            df_lote = df_lote[df_lote["Status_Lote"] == filtro_status_lote]
    
    
    col1, col2, col3 = st.columns([1,7,1.3])
    
    mes = mes = sorted(df_lote['Mês_Ano'].unique())
    faixa = faixa = sorted(df_lote['Faixa'].unique())
    consultor = consultor = sorted(df_lote['Consultor'].unique())

    with col1:
        filtro_mes = st.selectbox("🗓️ Mes_Ano", ["Todos"] + mes)
        
    with col2:
        filtro_faixa = st.multiselect("🔠 Faixa", options=faixa, default=faixa)
    with col3:
        filtro_consultor = st.selectbox("👨‍💼 Consultores", ["Todos"] + consultor)
    
    df = df_lote.copy()
    if filtro_status_lote != "Todos":
        df = df[df["Status_Lote"] == filtro_status_lote]

    if filtro_mes != "Todos":
        df = df[df["Mês_Ano"] == filtro_mes]

    if filtro_faixa:
        df = df[df["Faixa"].isin(filtro_faixa)]

    if filtro_consultor != "Todos":
        df = df[df["Consultor"] == filtro_consultor]

    return df

    
def formatar_tabela_personalizada(df):
    for col in df.columns:
        for row in df.index:
            valor = df.loc[row, col]
            if pd.isna(valor):
                continue

            if row in ['Valor_Base', 'Meta', 'Valor_Baixado', 'Falta_para_Meta']:
                df.loc[row, col] = f"R$ {float(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            elif row in ['Qtd_Contratos', 'Qtde_Baixado']:
                df.loc[row, col] = f"{int(valor):,}".replace(",", ".")
            else:
                pass
    return df


def relatorio_geral_por_faixa(df):
    st.markdown("### 📊 Relatório Geral por Faixa")

    df['Valor'] = pd.to_numeric(df['Valor'], errors='coerce')
    df['Meta'] = pd.to_numeric(df['Meta'], errors='coerce')
    df['ValorBaixado'] = pd.to_numeric(df['ValorBaixado'], errors='coerce')

    base = df.groupby('Faixa').agg(
        Qtd_Contratos=('Passaporte', 'nunique'),
        Valor_Base=('Valor', 'sum'),
        Meta=('Meta', 'sum'),
        Qtde_Baixado=('ValorBaixado', lambda x: x.gt(0).sum()),
        Valor_Baixado=('ValorBaixado', 'sum'),
    )

    base.loc['Total Geral'] = base.sum(numeric_only=True)
    base.loc['Total Geral', 'Qtd_Contratos'] = df['Passaporte'].nunique()

    base['Falta_para_Meta'] = base['Meta'] - base['Valor_Baixado']
    base['%_Meta_Alcançada'] = (base['Valor_Baixado'] / base['Meta']).fillna(0)

    final = base.T
    percentual_formatado = pd.to_numeric(final.loc['%_Meta_Alcançada'], errors='coerce').map("{:.1%}".format)
    final = final.astype(object)  # <- converte todo o DataFrame para tipo objeto
    final.loc['%_Meta_Alcançada'] = percentual_formatado
    final = formatar_tabela_personalizada(final)
    st.dataframe(final, use_container_width=True)

def relatorio_valor_recebido_manual(df):
    st.markdown("### 📑 Relatório Recebido Manual (Valor Rec.) por Faixa")

    df['Valor Rec.'] = pd.to_numeric(df['Valor Rec.'], errors='coerce')
    df['ValorBaixado'] = pd.to_numeric(df['ValorBaixado'], errors='coerce')

    tabela = df.groupby("Faixa").agg(
        Valor_Recebido_Manual=('Valor Rec.', 'sum'),
        Valor_Baixado=('ValorBaixado', 'sum'),
    )
    tabela['Diferença'] = tabela['Valor_Recebido_Manual'] - tabela['Valor_Baixado']

    tabela.loc['Total Geral'] = tabela.sum(numeric_only=True)

    tabela_formatada = tabela.T
    tabela_formatada = formatar_tabela_personalizada(tabela.T)
    st.dataframe(tabela_formatada, use_container_width=True)


def show_relatorio():
    Menu()
    df_lote = dados_lote()
    df = filtros(df_lote)
    relatorio_geral_por_faixa(df)
    relatorio_valor_recebido_manual(df)

if __name__ == "__main__":
    show_relatorio()