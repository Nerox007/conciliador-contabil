import streamlit as st
import pandas as pd
import io

# Configuração da página do site
st.set_page_config(page_title="Conciliador Bancário Inteligente", page_icon="⚡", layout="centered")

# Título principal na tela
st.title("⚡ Conciliador Bancário Inteligente")
st.write("Simplificando o fechamento contábil e financeiro em segundos.")

st.write("---")

# 🔑 SISTEMA DE SEGURANÇA
SENHA_CORRETA = "galvao3dias" 

senha_usuario = st.text_input("Digite sua senha de acesso de teste:", type="password")

if senha_usuario == SENHA_CORRETA:
    st.success("🔓 Acesso liberado com sucesso!")
    st.subheader("📁 Envie as planilhas para cruzamento")

    arquivo_banco = st.file_uploader("1. Escolha o Extrato Bancário (.xlsx)", type=["xlsx"])
    arquivo_sistema = st.file_uploader("2. Escolha o Relatório do Sistema (.xlsx)", type=["xlsx"])

    if arquivo_banco and arquivo_sistema:
        if st.button("🚀 Processar Conciliação Agora", use_container_width=True):
            with st.spinner("O robô está analisando os dados..."):
                try:
                    df_banco = pd.read_excel(arquivo_banco)
                    df_sistema = pd.read_excel(arquivo_sistema)

                    # 🔧 CORREÇÃO AUTOMÁTICA: Remove espaços e ajusta maiúsculas
                    df_banco.columns = df_banco.columns.astype(str).str.strip().str.capitalize()
                    df_sistema.columns = df_sistema.columns.astype(str).str.strip().str.capitalize()

                    # Padronizando as colunas de 'Valor' e 'Data'
                    for col in df_banco.columns:
                        if 'val' in col.lower():
                            df_banco.rename(columns={col: 'Valor'}, inplace=True)
                        if 'data' in col.lower() or 'dt' in col.lower():
                            df_banco.rename(columns={col: 'Data'}, inplace=True)
                            
                    for col in df_sistema.columns:
                        if 'val' in col.lower():
                            df_sistema.rename(columns={col: 'Valor'}, inplace=True)
                        if 'data' in col.lower() or 'dt' in col.lower():
                            df_sistema.rename(columns={col: 'Data'}, inplace=True)

                    # Verificação de segurança das colunas essenciais
                    if 'Valor' not in df_banco.columns or 'Data' not in df_banco.columns:
                        st.error("❌ Erro: Não encontramos as colunas 'Data' e 'Valor' no Extrato.")
                    elif 'Valor' not in df_sistema.columns or 'Data' not in df_sistema.columns:
                        st.error("❌ Erro: Não encontramos as colunas 'Data' e 'Valor' no Sistema.")
                    else:
                        # Convertendo Valor para número e Data para formato padrão
                        df_banco['Valor'] = df_banco['Valor'].astype(float)
                        df_sistema['Valor'] = df_sistema['Valor'].astype(float)
                        
                        df_banco['Data'] = pd.to_datetime(df_banco['Data'], errors='coerce').dt.date
                        df_sistema['Data'] = pd.to_datetime(df_sistema['Data'], errors='coerce').dt.date

                        # 🧠 NOVO CRUZAMENTO INTELIGENTE (Data + Valor)
                        df_conciliado = pd.merge(
                            df_banco, df_sistema, on=['Data', 'Valor'], how='outer', 
                            suffixes=('_Banco', '_Sistema'), indicator=True
                        )

                        # Separando o que bateu e o que deu erro
                        df_batem = df_conciliado[df_conciliado['_merge'] == 'both'].copy()
                        df_batem = df_batem.drop(columns=['_merge'])

                        df_divergencias = df_conciliado[df_conciliado['_merge'] != 'both'].copy()
                        df_divergencias['_merge'] = df_divergencias['_merge'].map({
                            'left_only': 'Apenas no Banco (Falta lançar no Sistema)',
                            'right_only': 'Apenas no Sistema (Não consta no Extrato Bancário)'
                        })
                        df_divergencias = df_divergencias.rename(columns={'_merge': 'Status_da_Divergencia'})

                        # 📊 NOVO: DASHBOARD COM ESTATÍSTICAS
                        st.write("---")
                        st.subheader("📊 Dashboard de Resultados")
                        
                        col1, col2, col3 = st.columns(3)
                        col1.metric("Total no Banco", len(df_banco), "linhas")
                        col2.metric("Conciliados (OK)", len(df_batem), "sucesso")
                        col3.metric("Divergências", len(df_divergencias), "erros", delta_color="inverse")
                        
                        st.write("---")

                        # 💾 PREPARANDO O EXCEL PARA DOWNLOAD
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            df_batem.to_excel(writer, sheet_name='Valores_Batem', index=False)
                            df_divergencias.to_excel(writer, sheet_name='Divergencias', index=False)
                        processed_data = output.getvalue()

                        st.balloons()
                        st.success("🔥 Conciliação finalizada perfeitamente!")
                        
                        st.download_button(
                            label="📥 Baixar Relatório de Conciliação (.xlsx)",
                            data=processed_data,
                            file_name="resultado_conciliacao.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                except Exception as e:
                    st.error(f"Erro ao processar as planilhas: {e}")

elif senha_usuario != "":
    st.error("❌ Senha incorreta ou período de teste expirado. Entre em contato com o suporte.")
else:
    st.info("🔒 Digite a senha fornecida pelo desenvolvedor para acessar o sistema.")
