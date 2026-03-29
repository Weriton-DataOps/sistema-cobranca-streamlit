import pandas as pd
from datetime import datetime 
from streamlit_tags import st_tags
import os
import shutil
# streamlit_app_title: 📊 Relatório Geral
import streamlit as st


st.set_page_config(
    page_title="Gestão Lote Fã", 
    page_icon="🏖️",                 
    layout="wide"
)

def Menu():
    st.sidebar.title("MENU")
    if st.sidebar.button("🔁 Atualizar Dados"):
        banco.clear()
        st.session_state["force_reload"] = True
        st.session_state["forcar_query"] = "tudo"  # Atualiza todas as queries
        st.rerun()

@st.cache_data
def banco(atualizar_queries=False, atualizar_somente_lotes=False):
    caminho = "data/BASE FA.xlsx"

    if atualizar_queries or atualizar_somente_lotes:
        

        try:
            if atualizar_queries:
                st.warning("Função desabilitada na versão web")
            elif atualizar_somente_lotes:
                st.warning("Função desabilitada na versão web")

            st.warning("Função desabilitada na versão web")
        finally:
            st.warning("Função desabilitada na versão web")

    # Sempre lê apenas a aba ReceberRecebidas
    df_receberRecebida = pd.read_excel(caminho, sheet_name="ReceberRecebidas")
    return df_receberRecebida


def data_receberRecebidas(df_receberRecebida):
    #Resumo dataframe
    hoje = pd.to_datetime(datetime.today())
    resumo_receberRecebidas = df_receberRecebida[['Passaporte','Fornecedor','Id','Vencimento','Tipo','Valor','Status','TiposBaixa','Dias Venc. Ant.','Status Lote']].copy()
    resumo_receberRecebidas["Vencimento"] = pd.to_datetime(resumo_receberRecebidas["Vencimento"], format="%d/%m/%Y")
    resumo_receberRecebidas = resumo_receberRecebidas.sort_values(by=['Vencimento'])
    resumo_receberRecebidas = resumo_receberRecebidas[
        (resumo_receberRecebidas["Dias Venc. Ant."].notna()) &
        (resumo_receberRecebidas["Status"] == "P") &
        (resumo_receberRecebidas["Vencimento"] < hoje) &
        (resumo_receberRecebidas["Status Lote"] != "ATIVO")
    ]
    return resumo_receberRecebidas

def filtros(resumo_receberRecebidas):
    st.divider()
    st.markdown("#### 🔍 Filtros")
    colf1, colf2, colf3, colf4 = st.columns([2,1,0.3,0.3])

    # Filtro por Tipo (multiselect)
    tipos_disponiveis = sorted(resumo_receberRecebidas['Tipo'].dropna().unique())
    tipos_selecionados = colf1.multiselect("Meio de Pagamento", options=tipos_disponiveis, default=[])
    if tipos_selecionados:
        resumo_receberRecebidas = resumo_receberRecebidas[resumo_receberRecebidas['Tipo'].isin(tipos_selecionados)]

    # Filtro por TiposBaixa
    tipos_baixa_disponiveis = sorted(resumo_receberRecebidas['TiposBaixa'].dropna().unique())
    tipos_baixa_excluir = colf2.multiselect("Excluir Tipos de Baixa", options=tipos_baixa_disponiveis, default=[])

    if tipos_baixa_excluir:
        resumo_receberRecebidas = resumo_receberRecebidas[~resumo_receberRecebidas['TiposBaixa'].isin(tipos_baixa_excluir)]

    # Filtro por intervalo de Dias Venc. Ant.
    min_dias = int(resumo_receberRecebidas["Dias Venc. Ant."].min())
    max_dias = int(resumo_receberRecebidas["Dias Venc. Ant."].max())
    dias_min = colf3.number_input("Venc. Mín.", min_value=min_dias, max_value=max_dias, value=min_dias)
    dias_max = colf4.number_input("Venc. Máx.", min_value=min_dias, max_value=max_dias, value=max_dias)
    resumo_receberRecebidas = resumo_receberRecebidas[
        (resumo_receberRecebidas["Dias Venc. Ant."] >= dias_min) &
        (resumo_receberRecebidas["Dias Venc. Ant."] <= dias_max)    
    ]
    return resumo_receberRecebidas

def resumoGeral(resumo_receberRecebidas):
# Métricas
    Total_passaportes = f"{resumo_receberRecebidas['Passaporte'].nunique():,}".replace(",", ".")
    Total_valor = resumo_receberRecebidas['Valor'].sum()

    st.markdown("#### 📥 Total Pendente a Designar")
    col1, col2 = st.columns(2)
    col1.metric(label="Total de Passaportes Únicos", value=Total_passaportes)
    col2.metric(label="Soma Total dos Valores", value=f"R$ {Total_valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))


def gerarLote(col_esquerda):
    #Caixa de Faixa de INAD
    with col_esquerda:
        st.markdown("#### 📑 Gerar Lote de Cobrança")
        #Caixa de Faixa de INAD
        faixa_selecionada = st.selectbox(
            "🧮 Faixa Lote",
            options=[""]+["FLASH 05", "FLASH 10", "FLASH 15", "FLASH 20", "FLASH 25", "INAD 1-30", "INAD 31-90", "INAD 91-180", "INAD 181+"],
            index=0
        )
        colaboradores = st_tags(
            label="👥 Consultores",
            text="Digite e pressione Enter",
            value=[],        # começa vazio
            suggestions=[],  # sem lista fixa de sugestões
            maxtags=20
        )
        meta_percentual = st.number_input(
            "🎯 Meta de Cobrança (%)",
            min_value=0.0,
            max_value=100.0,
            step=0.5
        )
        
        return faixa_selecionada,colaboradores,meta_percentual
    

def data_lote(colaboradores,faixa_selecionada,resumo_receberRecebidas,meta_percentual):
    
    num_consultores = len(colaboradores)
    resumo_receberRecebidas = resumo_receberRecebidas.copy()

    # Agrupar por Fornecedor e somar os valores
    agrupado = resumo_receberRecebidas.groupby("Fornecedor").agg(
        TotalValor=("Valor", "sum"),
        NumParcelas=("Passaporte", "count")
    ).reset_index()

    # Ordenar para distribuir proporcionalmente
    agrupado = agrupado.sort_values(by="TotalValor", ascending=False).reset_index(drop=True)

    # Distribuição dos fornecedores para os colaboradores
    agrupado["Consultor"] = [colaboradores[i % len(colaboradores)] for i in range(len(agrupado))]

    # Juntar com o dataframe original para gerar o lote
    df_lote = resumo_receberRecebidas.merge(agrupado[["Fornecedor", "Consultor"]], on="Fornecedor", how="left")
    
    hoje = pd.to_datetime(datetime.today())
    # Nome do lote
    mes_ano = hoje.strftime("%B-%Y").upper().replace("Ç", "C").replace("É", "E")
    data_hoje = datetime.today().strftime("%d-%m-%Y")
    nome_lote = f"{faixa_selecionada}_{mes_ano}_{data_hoje}_ATIVO"
    df_lote["Meta"] = df_lote["Valor"] * (meta_percentual / 100)

    return df_lote,nome_lote


def tabela_distribuir(df_lote):
    # Tabela de distribuição por consultor com valores formatados
    resumo_lote = (
        df_lote.groupby("Consultor")
        .agg(Qtd_Passaportes=("Passaporte", lambda x: x.nunique()), Valor_Total=("Valor", "sum"))
        .reset_index()
        .sort_values(by="Valor_Total", ascending=False)
    )
    resumo_lote["Valor_Total"] = resumo_lote["Valor_Total"].apply(
        lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    )
    st.session_state['resumo_lote'] = resumo_lote


def bot_distribuir(resumo_receberRecebidas,col_esquerda,col_direita,faixa_selecionada,colaboradores,meta_percentual):
        with col_esquerda:
            # 🔄 Distribuir clientes proporcionalmente entre os consultores
            distribuir = st.button("📦 Distribuir Lote de Cobrança")
        if distribuir and colaboradores and not resumo_receberRecebidas.empty and faixa_selecionada:
            df_lote, nome_lote = data_lote(colaboradores, faixa_selecionada, resumo_receberRecebidas, meta_percentual)
            df_lote['Faixa'] = faixa_selecionada
            st.session_state['df_lote'] = df_lote
            st.session_state['nome_lote'] = nome_lote
            st.session_state['lote_distribuido'] = True
            with col_direita:
                tabela_distribuir(df_lote)        
            loteDetalhado(df_lote)


def bot_gerarLote(col_direita):
    with col_direita:
        if st.button("📁 Gerar Lote"):
            if 'df_lote' in st.session_state and 'nome_lote' in st.session_state:
                df_lote = st.session_state['df_lote']
                nome_lote = st.session_state['nome_lote']
                
                diretorio_destino = "data/LOTES/ATIVOS"
                caminho_arquivo = os.path.join(diretorio_destino, f"{nome_lote}.xlsx")

                try:
                    st.success("Lote gerado (simulação para portfólio)")

                except Exception as e:
                    st.error(f"Erro ao salvar o arquivo Excel: {e}")
            else:
                st.warning("⚠️ Dados do lote não encontrados.")



def loteDetalhado(df_lote):
    #tipos auterados
    df_lote = st.session_state['df_lote']
    #df_lote["Valor"] = df_lote["Valor"].apply(lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    df_lote["Id"] = df_lote["Id"].astype(str)
    df_lote["Dias Venc. Ant."] = df_lote["Dias Venc. Ant."].astype(int).replace(",",".")
    st.session_state['df_lote'] = df_lote



def painel_lotes_form(diretorio_base):
    st.markdown("#### 🗂️ Gerenciar Lotes Existentes")

    # Definir subpastas
    pasta_ativos = os.path.join(diretorio_base, "ATIVOS")
    pasta_expirados = os.path.join(diretorio_base, "EXPIRADOS")

    # Buscar arquivos nas duas pastas
    arquivos = []
    for pasta in [pasta_ativos, pasta_expirados]:
        arquivos += [os.path.join(pasta, f) for f in os.listdir(pasta) if f.endswith(".xlsx")]

    if not arquivos:
        st.info("Nenhum lote encontrado.")
        return

    lista_infos = []
    for caminho_completo in arquivos:
        nome = os.path.basename(caminho_completo)
        partes = nome.replace(".xlsx", "").rsplit("_", 3)
        if len(partes) == 4:
            faixa_raw, mes_ano, data_criacao, status = partes
            faixa = faixa_raw.replace("+", "").strip()
            lista_infos.append((caminho_completo, nome, faixa, mes_ano, data_criacao, status))

    faixas = sorted(set(x[2] for x in lista_infos))
    meses = sorted(set(x[3] for x in lista_infos))
    status_set = sorted(set(x[5] for x in lista_infos))

    col1, col2, col3 = st.columns(3)
    with col1:
        filtro_faixa = st.selectbox("🔠 Faixa", ["Todos"] + faixas)
    with col2:
        filtro_mes_ano = st.selectbox("🗓️ Mês/Ano", ["Todos"] + meses)
    with col3:
        filtro_status = st.selectbox("🎯 Status", ["ATIVO"] + status_set)

    lotes_filtrados = []
    for caminho, nome, faixa, mes_ano, data, status in lista_infos:
        if filtro_faixa != "Todos" and faixa != filtro_faixa:
            continue
        if filtro_mes_ano != "Todos" and mes_ano != filtro_mes_ano:
            continue
        if filtro_status != "Todos" and status != filtro_status:
            continue
        lotes_filtrados.append((caminho, nome, faixa, mes_ano, data, status))

    with st.form("form_lotes"):
        col4, col5 = st.columns([2, 1])
        with col4:
            nova_situacao = st.selectbox("📌 Nova Situação", ["EXPIRADO", "EXCLUIR", "ATIVO"])
        with col5:
            confirmar = st.form_submit_button("🔄 Atualizar Selecionados")

        st.markdown("Selecione os lotes para alterar status:")
        selecionados = []
        for caminho, nome, faixa, mes_ano, data, status in sorted(lotes_filtrados, key=lambda x: x[1], reverse=True):
            if st.checkbox(f"📄 {nome}", key=f"ck_{nome}"):
                selecionados.append((caminho, nome))

    if confirmar and selecionados:
        for caminho_antigo, nome_antigo in selecionados:
            partes = nome_antigo.replace(".xlsx", "").rsplit("_", 3)
            if len(partes) == 4:
                faixa_raw, mes_ano, data_criacao, _ = partes
                if nova_situacao == "EXCLUIR":
                    os.remove(caminho_antigo)
                else:
                    novo_nome = f"{faixa_raw}_{mes_ano}_{data_criacao}_{nova_situacao}.xlsx"
                    novo_caminho = os.path.join(pasta_expirados if nova_situacao == "EXPIRADO" else pasta_ativos, novo_nome)
                    shutil.move(caminho_antigo, novo_caminho)
        st.session_state["force_reload"] = True 
        st.session_state["forcar_query"] = "lotes"  # Apenas a query dos LOTES
        st.success("✅ Status atualizado com sucesso.")
        st.rerun()


def show_lote():
    
    Menu()

    if st.session_state.pop("force_reload", False):
        banco.clear()

    query_flag = st.session_state.pop("forcar_query", False)

    if query_flag == "tudo":
        df_receberRecebida = banco(atualizar_queries=True)
    elif query_flag == "lotes":
        df_receberRecebida = banco(atualizar_somente_lotes=True)
    else:
        df_receberRecebida = banco()


    painel_lotes_form("data/...")
    resumo_receberRecebidas = data_receberRecebidas(df_receberRecebida)
    if resumo_receberRecebidas.empty:
        st.info("✅ Nenhuma parcela pendente de designação encontrada.")
        return
    resumo_receberRecebidas = filtros(resumo_receberRecebidas)
    st.divider()
    resumoGeral(resumo_receberRecebidas)
    st.divider()
    col_esquerda, col_centro, col_direita = st.columns([3,1,5])
    faixa_selecionada,colaboradores,meta_percentual = gerarLote(col_esquerda)
    bot_distribuir(resumo_receberRecebidas,col_esquerda,col_direita,faixa_selecionada,colaboradores,meta_percentual)
    
    
    with col_direita:
        if st.session_state.get('lote_distribuido'):
            st.markdown(f"#### 🧾 Lote: `{st.session_state.get('nome_lote')}`")
            st.dataframe(st.session_state['resumo_lote'], use_container_width=True)
        else:
            st.markdown("#### 🧾 Lote: ")
            st.info("⚠️ Selecione uma faixa de lote e ao menos um consultor para visualizar a divisão.")
    
    if st.session_state.get('lote_distribuido'):
        bot_gerarLote(col_direita)
    
    
    if st.session_state.get('lote_distribuido'):
        st.divider()
        st.markdown("#### 📋 Detalhamento das Parcelas do Lote")
        st.dataframe(st.session_state['df_lote'], use_container_width=True)

    


if __name__ == "__main__":
    show_lote()
