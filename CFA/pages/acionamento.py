import pandas as pd
import streamlit as st
from st_aggrid import AgGrid, GridOptionsBuilder
from datetime import datetime
import win32com.client as win32
import os
import xlwings as xw
import pythoncom
import time
import win32com.client
import io

st.set_page_config(
    page_title="Acionamento Fã", 
    page_icon="🏖️",                 
    layout="wide"
)

@st.cache_data(show_spinner=False)
def dados_lote(atualizar=False):  # Parâmetro dummy só para invalidar cache
    pasta = "data/LOTES/ATIVOS"
    dataframes = []
    arquivos_validos = [arq for arq in os.listdir(pasta) if arq.endswith('.xlsx')]

    for arquivo in arquivos_validos:
        caminho_arquivo = os.path.join(pasta, arquivo)
        try:
            df = pd.read_excel(caminho_arquivo, sheet_name="Lote")
            df.columns = df.columns.str.strip()
            df['arquivo_origem'] = arquivo
            dataframes.append(df)
        except Exception as e:
            print(f"Erro ao ler {arquivo}: {e}")

    if not dataframes:
        st.warning("⚠️ Nenhum lote ativo foi encontrado.")
        return pd.DataFrame()

    return pd.concat(dataframes, ignore_index=True)

pd.read_excel

@st.cache_data(show_spinner=False)
def carregar_base_contrato(atualizar=False):
    return pd.read_excel("data/BASE FA.xlsx", sheet_name="Contratos")

def carregar_dados(df_lote, consultor):
    chave = st.session_state.get("chave_atualizacao", False)
    df_contrato = carregar_base_contrato(atualizar=chave)

    colunas_essenciais = [
        "Passaporte", "Fornecedor", "Id", "Vencimento", "Faixa", "Tipo", "Valor", "Meta", "Status",
        "Consultor", "arquivo_origem", "StatusAc.", "Data Rec.", "Valor Rec.", "MeioPag.",
        "Valor EmDia", "Observação", "ValorBaixado"
    ]

    # Garante que só as colunas disponíveis sejam mantidas
    df_lote = df_lote[[col for col in colunas_essenciais if col in df_lote.columns]]
    df_lote = df_lote[df_lote["Faixa"] != "MODELO"]

    # Junta com e-mails e telefones
    df_contrato_filtrado = df_contrato[["Numero", "Email", "Telefone"]].rename(columns={"Numero": "Passaporte"})
    df_merged = pd.merge(df_lote, df_contrato_filtrado, on="Passaporte", how="left")
    # Garante que colunas numéricas estejam presentes
    for col in ['ValorBaixado', 'Valor Rec.']:
        if col not in df_merged.columns:
            df_merged[col] = 0.0

    # Filtro por consultor
    df_merged = df_merged[df_merged["Consultor"] == consultor]



    # Medidas por passaporte
    medidas = (
        df_merged.groupby("Passaporte")
        .agg(
            Qnt_Parc=('Passaporte', 'count'),
            Total_Devido=('Valor', 'sum'),
            Total_Meta=('Meta', 'sum'),
            Total_Baixado=('ValorBaixado', 'sum')
        )
        .reset_index()
    )

    # Junta medidas ao dataframe original
    df_merged = pd.merge(df_merged, medidas, on="Passaporte", how="left")

    return df_merged




def formatar_moeda(df, colunas):
    for col in colunas:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            df[col] = df[col].apply(lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    return df



def gerar_tabela_grid(df):
    gb = GridOptionsBuilder.from_dataframe(df)
    gb.configure_default_column(filterable=True, sortable=True, resizable=True)
    gb.configure_selection(selection_mode="single", use_checkbox=False)
    gb.configure_column("Fornecedor", filter="agTextColumnFilter")
    gb.configure_column("Passaporte", filter="agTextColumnFilter")
    gb.configure_grid_options(enableRowGroup=True, groupDisplayType="multipleColumns")
    grid_response = AgGrid(df, gridOptions=gb.build(), update_mode='SELECTION_CHANGED', theme='streamlit', height=480)
    return grid_response


def preparar_editor_parcelas(df, passaporte):
    colunas_exibidas = ['Id', 'Vencimento', 'Tipo', 'Valor', 'ValorBaixado', 
                        'StatusAc.', 'Data Rec.', 'Valor Rec.', 'MeioPag.', 
                        'Valor EmDia', 'Observação']
    detalhes = df[df['Passaporte'] == passaporte][colunas_exibidas].copy()

    detalhes['StatusAc.'] = detalhes['StatusAc.'].astype(str)
    detalhes['MeioPag.'] = detalhes['MeioPag.'].astype(str)
    detalhes['Data Rec.'] = pd.to_datetime(detalhes['Data Rec.'], errors='coerce')
    detalhes['Valor Rec.'] = pd.to_numeric(detalhes['Valor Rec.'], errors='coerce')
    detalhes['Valor EmDia'] = pd.to_numeric(detalhes['Valor EmDia'], errors='coerce')
    detalhes['Observação'] = detalhes['Observação'].astype(str)

    return detalhes



def arquivo_em_uso(caminho):
    try:
        with open(caminho, 'a'):
            return False
    except:
        return True

def identificar_usuario_arquivo(caminho):
    try:
        network = win32com.client.Dispatch("WScript.Network")
        excel_app = win32com.client.Dispatch("Excel.Application")
        for wb in excel_app.Workbooks:
            if caminho.lower() in wb.FullName.lower():
                return network.UserName
    except:
        return None


def salvar_detalhes_editados(detalhes_editados, caminho_arquivo,consultor):
    pythoncom.CoInitialize()
    tempo_max_espera = 60
    tempo_esperado = 0
    intervalo_espera = 2
    mensagem_status = st.empty()

    while arquivo_em_uso(caminho_arquivo):
        usuario = identificar_usuario_arquivo(caminho_arquivo)
        mensagem_status.warning(f"⏳ Aguardando liberação... Arquivo em uso por: **{usuario or 'usuário desconhecido'}**.")
        time.sleep(intervalo_espera)
        tempo_esperado += intervalo_espera
        if tempo_esperado >= tempo_max_espera:
            mensagem_status.error(f"❌ Tempo excedido. Ainda em uso por **{usuario or 'usuário desconhecido'}**.")
            return

    try:
        app = xw.App(visible=False)
        wb = app.books.open(caminho_arquivo)
        ws = wb.sheets['Lote']

        header = ws.range("A1").expand("right").value
        col_map = {col.strip(): idx for idx, col in enumerate(header)}
        novas_colunas = ['StatusAc.', 'Data Rec.', 'Valor Rec.', 'MeioPag.', 'Valor EmDia', 'Observação']

        for col in novas_colunas:
            if col not in col_map:
                nova_col_idx = len(header)
                ws.range((1, nova_col_idx + 1)).value = col
                col_map[col] = nova_col_idx
                header.append(col)

        valores_ws = ws.range("A2").expand("table").value

        # Garante que cada linha tenha o mesmo número de colunas do header
        valores_ws_corrigido = [linha + [None] * (len(header) - len(linha)) if len(linha) < len(header) else linha[:len(header)] for linha in valores_ws]

        df_excel = pd.DataFrame(valores_ws_corrigido, columns=header)


        atualizados = 0
        for _, row in detalhes_editados.iterrows():
            id_parcela = str(row['Id']).strip()
            idx = df_excel[df_excel['Id'].astype(str).str.strip() == id_parcela].index
            if not idx.empty:
                linha_idx = idx[0] + 2
                for col in novas_colunas:
                    ws.range((linha_idx, col_map[col] + 1)).value = row[col]
                atualizados += 1

        wb.save()
        wb.close()
        app.quit()
        st.session_state['chave_atualizacao'] = time.time()
        df_lote = dados_lote(atualizar=time.time())  # Lê arquivos da pasta "LOTES/ATIVOS"
        df_merged = carregar_dados(df_lote, consultor)
        if atualizados:
            mensagem_status.success("✅ Dados salvos com sucesso!")
            st.rerun()
        else:
            mensagem_status.warning("⚠️ Nenhuma linha foi atualizada. IDs podem estar incorretos.")

    except Exception as e:
        mensagem_status.error(f"❌ Erro ao salvar: {e}")
    finally:
        try: wb.close()
        except: pass
        try: app.quit()
        except: pass


def tabelaPrincipal(df_merged,consultor):
    colfiltro1, colfiltro2, colatu = st.columns([1,6,1])
    with colfiltro1:
        consultor_selecionado = st.session_state["consultor_logado"]
        st.text_input("👨‍💼 Consultor", value=consultor_selecionado, disabled=True)
    with colfiltro2:
        faixas = sorted(df_merged['Faixa'].dropna().unique())
        faixas_selecionadas = st.multiselect("💸 Faixa", options=faixas, default=faixas)
    #df_merged = df_merged[df_merged['Consultor'] == consultor_selecionado]


    # Garante que colunas de edição existam no df_merged
    colunas_adicionais = {
        'StatusAc.': '',
        'Data Rec.': pd.NaT,
        'Valor Rec.': 0.0,
        'MeioPag.': '',
        'Valor EmDia': 0.0,
        'Observação': '',
        'ValorBaixado': 0.0
    }

    if 'ValorBaixado' not in df_merged.columns:
        df_merged['ValorBaixado'] = 0.

    for col, valor_padrao in colunas_adicionais.items():
        if col not in df_merged.columns:
            df_merged[col] = valor_padrao


    if faixas_selecionadas != "Todos":
        df_merged = df_merged[df_merged['Faixa'].isin(faixas_selecionadas)]
    with colatu:
        if st.button("🔄 Atualizar Dados"):
            with st.spinner("🔄 Lendo planilhas de lotes e atualizando..."):
                df_lote = dados_lote(atualizar=time.time())  # Lê arquivos da pasta "LOTES/ATIVOS"
                df_merged = carregar_dados(df_lote, consultor)  # Junta com base de contratos
                st.session_state['df_lote'] = df_lote
                st.session_state['df_merged'] = df_merged
                st.success("✅ Dados atualizados com sucesso!")
                st.rerun()
    
    colunas_moeda = ['Total_Devido', 'Valor', 'Total_Meta', 'Meta', 'Total_Baixado', 'ValorBaixado']
    df_merged = formatar_moeda(df_merged, colunas_moeda)


    col1, col2 = st.columns([2,2])
    with col1:
        tab_principal = df_merged[["Passaporte", "Fornecedor", "Faixa", "Qnt_Parc", "Total_Devido", "Total_Baixado", "Email", "Telefone"]].drop_duplicates()
        
        colbot1,colbot2 = st.columns([3,2])
        with colbot1:
            st.markdown("### 📋 Inadimplentes")
        with colbot2:
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine="xlsxwriter") as writer:
                df_exportar = df_merged.copy()
                df_exportar.to_excel(writer, index=False, sheet_name="Dados Filtrados")

            st.download_button(
                label="⬇️ Baixar Tabela como Excel",
                data=excel_buffer.getvalue(),
                file_name=f"Inadimplentes_{datetime.today().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        grid_response = gerar_tabela_grid(tab_principal)
    
    with col2:

        selected = grid_response.get('selected_rows')       
        
        if selected is not None and len(selected) > 0:
            row = pd.DataFrame(selected).iloc[0]
            passaporte = row['Passaporte']
            st.markdown(f"### 🧾 Parcelas Vencidas: `{passaporte}`")
            detalhes = preparar_editor_parcelas(df_merged, passaporte)          
            detalhes_editados = st.data_editor(
                detalhes,
                column_config={
                    "StatusAc.": st.column_config.SelectboxColumn("StatusAc.",options=["ACIONADO","NÃO ACIONADO","STATUS COB."]),
                    "Data Rec.": st.column_config.DateColumn("Data Rec."),
                    "Valor Rec.": st.column_config.NumberColumn("Valor Rec."),
                    "MeioPag.": st.column_config.SelectboxColumn("MeioPag.",options=["PIX","BOLETO","CARTÃO"]),
                    "Valor EmDia": st.column_config.NumberColumn("Valor EmDia"),
                    "Observação": st.column_config.TextColumn("Observação")
                },
                disabled=['Vencimento','Tipo','Valor','ValorBaixado'],
                use_container_width=True,
                hide_index=True
            )
            if st.button("💾 Salvar Atualizações"):
                arquivo_origem = df_merged[df_merged['Passaporte'] == passaporte]['arquivo_origem'].values[0]
                caminho_arquivo = os.path.join(
                    "data/...",
                    arquivo_origem
                )
                salvar_detalhes_editados(detalhes_editados, caminho_arquivo,consultor)


def atualizar_pagamentos():

    atualizar_query_excel()
    # Inicializa o COM para trabalhar com Excel
    pythoncom.CoInitialize()

    # Carrega os dados da aba ReceberRecebidas
    try:
        df_recebidas = pd.read_excel("data/Recebimento.xlsx", sheet_name="ReceberRecebidas")
        df_recebidas = df_recebidas[["Id", "ValorBaixado"]].dropna()
    except Exception as e:
        st.error(f"❌ Erro ao carregar aba ReceberRecebidas: {e}")
        st.stop()

    pasta = "data/..."
    arquivos = [arq for arq in os.listdir(pasta) if arq.endswith(".xlsx")]

    for arquivo in arquivos:
        caminho = os.path.join(pasta, arquivo)

        def arquivo_em_uso(caminho):
            try:
                with open(caminho, 'a'):
                    return False
            except:
                return True

        tempo_max_espera = 60
        tempo_esperado = 0
        while arquivo_em_uso(caminho):
            time.sleep(2)
            tempo_esperado += 2
            if tempo_esperado >= tempo_max_espera:
                st.warning(f"⚠️ Arquivo em uso: {arquivo}. Pulando...")
                continue

        try:
            app = xw.App(visible=False)
            wb = app.books.open(caminho)
            ws = wb.sheets['Lote']

            # Mapeia colunas
            header = ws.range("A1").expand("right").value
            col_map = {col.strip(): idx for idx, col in enumerate(header)}
            if 'ValorBaixado' not in col_map:
                nova_coluna_idx = len(header)
                ws.range((1, nova_coluna_idx + 1)).value = 'ValorBaixado'
                col_map['ValorBaixado'] = nova_coluna_idx
                header.append('ValorBaixado')
            if 'Id' not in col_map or 'ValorBaixado' not in col_map:
                st.warning(f"❗ Colunas 'Id' ou 'ValorBaixado' não encontradas em {arquivo}.")
                wb.close()
                continue

            dados = ws.range("A2").expand("table").value
            linhas_alteradas = 0
            df_recebidas['Id'] = df_recebidas['Id'].astype(str).str.strip()
            for i, linha in enumerate(dados):
                id_linha = str(linha[col_map['Id']]).strip()
                id_linha = linha[col_map['Id']]
                if id_linha in df_recebidas['Id'].values:
                    novo_valor = df_recebidas[df_recebidas['Id'] == id_linha]['ValorBaixado'].values[0]
                    col_idx_excel = col_map['ValorBaixado'] + 1
                    ws.range((i + 2, col_idx_excel)).value = novo_valor
                    linhas_alteradas += 1

            wb.save()
            wb.close()
            app.quit()

            st.success(f"✅ {linhas_alteradas} valores atualizados em {arquivo}")

        except Exception as e:
            st.error(f"❌ Erro ao processar {arquivo}: {e}")
            try: wb.close()
            except: pass
            try: app.quit()
            except: pass


def atualizar_query_excel():
    pythoncom.CoInitialize()
    try:
        excel = win32.gencache.EnsureDispatch("Excel.Application")
        wb = excel.Workbooks.Open(os.path.abspath("Recebimento.xlsx"))
        excel.Visible = False

        atualizou = False
        for connection in wb.Connections:
            # Ajuste esse nome para o exato nome da consulta ligada à aba ReceberRecebidas
            if "ReceberRecebidas" in connection.Name:
                connection.Refresh()
                atualizou = True
                break

        if not atualizou:
            st.warning("⚠️ Nenhuma conexão correspondente à aba 'ReceberRecebidas' foi encontrada.")

        excel.CalculateUntilAsyncQueriesDone()
        wb.Save()
        wb.Close(False)
        excel.Quit()

        if atualizou:
            st.info("🔄 Query atualizada com sucesso!")

    except Exception as e:
        st.error(f"❌ Falha ao atualizar a query: {e}")
    finally:
        try:
            wb.close()
        except:
            pass
        try:
            excel.quit()
        except:
            pass


def formatar_tabela_personalizada(df):
    df = df.astype(object)
    for col in df.columns:
        for row in df.index:
            valor = df.loc[row, col]
            if pd.isna(valor):
                continue

            if row in ['Valor_Base', 'Meta', 'Valor_Baixado', 'Falta_para_Meta','Valor Rec.','Total Geral']:
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
    df['Valor Rec.'] = pd.to_numeric(df['Valor Rec.'], errors='coerce')

    base = df.groupby('Faixa').agg(
        Qtd_Contratos=('Passaporte', 'nunique'),
        Valor_Base=('Valor', 'sum'),
        Meta=('Meta', 'sum'),
        # Qtde_Baixado=('ValorBaixado', lambda x: x.gt(0).sum()),
        # Valor_Baixado=('ValorBaixado', 'sum'),
        Qtde_Recebido=('Valor Rec.', lambda x: x.gt(0).sum()),
        Valor_Recebido=('Valor Rec.', 'sum'),
    )

    base.loc['Total Geral'] = base.sum(numeric_only=True)
    base.loc['Total Geral', 'Qtd_Contratos'] = df['Passaporte'].nunique()

    base['Falta_para_Meta'] = base['Meta'] - base['Valor_Recebido']
    base['%_Meta_Alcançada'] = (base['Valor_Recebido']/ base['Meta']).fillna(0)

    final = base.T
    final.loc['%_Meta_Alcançada'] = pd.to_numeric(final.loc['%_Meta_Alcançada'], errors='coerce').map("{:.1%}".format)
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


def relatorio_qtd_valor_por_faixa_meiopag(df):
    st.markdown("### 🧾 Relatório de Quantidade e Valor Recebido por Faixa e Meio de Pagamento")

    df = df.copy()
    if 'Valor Rec.' not in df.columns:
        df['Valor Rec.'] = 0.0
    if 'MeioPag.' not in df.columns:
        df['MeioPag.'] = 0.0
    df['Valor Rec.'] = pd.to_numeric(df['Valor Rec.'], errors='coerce').fillna(0)
    df['Vencimento'] = pd.to_datetime(df['Vencimento'], errors='coerce')
    df['MeioPag.'] = df['MeioPag.'].fillna('(vazio)').astype(str)

    df['Faixa'] = df['Faixa'].astype(str)

    df_recebido = df[df['Valor Rec.'] > 0].copy()

    # -------- TABELA 1: QUANTIDADE --------
    qtd = df_recebido.groupby(['MeioPag.', 'Faixa']).size().unstack(fill_value=0)
    qtd['Total Geral'] = qtd.sum(axis=1)
    qtd.loc['Total Geral'] = qtd.sum()

    st.markdown("#### 📌 Quantidade de Recebimentos")
    st.dataframe(qtd, use_container_width=True)

    # -------- TABELA 2: VALOR TOTAL --------
    valor = df_recebido.groupby(['MeioPag.', 'Faixa'])['Valor Rec.'].sum().unstack(fill_value=0)
    valor['Total Geral'] = valor.sum(axis=1)
    valor.loc['Total Geral'] = valor.sum()

    # Formatação R$
    valor_formatado = formatar_tabela_personalizada(valor.T).T

    st.markdown("#### 💰 Valor Total Recebido")
    st.dataframe(valor_formatado, use_container_width=True)


def show_acionamento():
    df_temp = dados_lote()
    consultores_disponiveis = sorted(df_temp['Consultor'].dropna().unique()) if not df_temp.empty else []

    # Apenas seleciona se ainda não foi feito
    if "consultor_logado" not in st.session_state:
        consultores_opcoes = ["⤵️"] + consultores_disponiveis
        consultor = st.selectbox("👤 Selecione seu nome para continuar", consultores_opcoes,index=0, key="selecao_consultor")
        

        # Apenas armazena após interação explícita
        if st.button("👉 Confirmar"):
            st.session_state["consultor_logado"] = consultor
            st.session_state.pop("df_lote", None)
            st.session_state.pop("df_merged", None)
            st.rerun()
        st.stop()

    consultor = st.session_state["consultor_logado"]

    if 'df_lote' not in st.session_state or 'df_merged' not in st.session_state:
        chave = st.session_state.get("chave_atualizacao", False)
        df_lote = dados_lote(atualizar=chave)
        df_merged = carregar_dados(df_lote, consultor)
        st.session_state['df_lote'] = df_lote
        st.session_state['df_merged'] = df_merged
    else:
        df_lote = st.session_state['df_lote']
        df_merged = st.session_state['df_merged']

    tabelaPrincipal(df_merged,consultor)

    if st.button("📥 Atualizar Pagamentos"):
        with st.spinner("Atualizando ..."):
            atualizar_pagamentos()
            chave = time.time()  # força atualização do cache
            df_lote = dados_lote(atualizar=chave)
            df_merged = carregar_dados(df_lote, consultor)
            st.session_state['df_lote'] = df_lote
            st.session_state['df_merged'] = df_merged
            st.session_state['chave_atualizacao'] = chave
            st.success("✅ Pagamentos atualizados com sucesso!")
            st.rerun()
    relatorio_geral_por_faixa(df_merged)
    relatorio_valor_recebido_manual(df_merged)
    relatorio_qtd_valor_por_faixa_meiopag(df_merged)


show_acionamento()


