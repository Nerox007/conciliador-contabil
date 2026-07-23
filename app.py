import streamlit as st
import pandas as pd
import io

# Configuração da página do site
st.set_page_config(page_title="Conciliador Bancário Inteligente", page_icon="⚡", layout="centered")

# Título principal na tela
st.title("⚡ Conciliador Bancário Inteligente")
st.write("Simplificando o fechamento contábil e financeiro em segundos.")

st.write("---")

# 🔑 SISTEMA DE SEGURANÇA (Sua trava de teste)
# Mude essa palavra quando o prazo de 3 dias acabar para bloquear o cliente!
SENHA_CORRETA = "galvao3dias" 

senha_usuario = st.text_input("Digite sua senha de acesso de teste:", type="password")

if senha_usuario == SENHA_CORRETA:
    st.success("🔓 Acesso liberado com sucesso!")
    st.subheader("📁 Envie as planilhas para cruzamento")

    # Botões bonitos para o cliente arrastar os arquivos no site
    arquivo_banco = st.file_uploader("1. Escolha o Extrato Bancário (.xlsx)", type=["xlsx"])
    arquivo_sistema = st.file_uploader("2. Escolha o Relatório do Sistema (.xlsx)", type=["xlsx"])

    if arquivo_banco and arquivo_sistema:
        # Botão gigante para ativar o robô
        if st.button("🚀 Processar Conciliação Agora"):
            with st.spinner("O robô está analisando os dados..."):
                try:
                    # Lendo os arquivos direto do upload do site
                    df_banco = pd.read_excel(arquivo_banco)
                    df_sistema = pd.read_excel(arquivo_sistema)

                    # 🔧 CORREÇÃO AUTOMÁTICA: Remove espaços em branco e ajusta maiúsculas
                    df_banco.columns = df_banco.columns.astype(str).str.strip().str.capitalize()
                    df_sistema.columns = df_sistema.columns.astype(str).str.strip().str.capitalize()

                    for col in df_banco.columns:
                        if 'val' in col.lower():
                            df_banco.rename(columns={col: 'Valor'}, inplace=True)
                            
                    for col in df_sistema.columns:
                        if 'val' in col.lower():
                            df_sistema.rename(columns={col: 'Valor'}, inplace=True)

                    # Verificação de segurança das colunas
                    if 'Valor' not in df_banco.columns or 'Valor' not in df_sistema.columns:
                        st.error("❌ Erro de estrutura: Não encontramos a coluna 'Valor' nas planilhas.")
                    else:
                        # Convertendo para número decimal
                        df_banco['Valor'] = df_banco['Valor'].astype(float)
                        df_sistema['Valor'] = df_sistema['Valor'].astype(float)

                        # Cruzamento inteligente dos dados
                        df_conciliado = pd.merge(
                            df_banco, df_sistema, on='Valor', how='outer', 
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

                        # 💾 PREPARANDO O EXCEL NA MEMÓRIA DO SITE
                        # Como o site roda na nuvem, precisamos salvar o arquivo em formato de "bytes" para o download funcionar
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            df_batem.to_excel(writer, sheet_name='Valores_Batem', index=False)
                            df_divergencias.to_excel(writer, sheet_name='Divergencias', index=False)
                        processed_data = output.getvalue()

                        st.balloons() # Efeito visual de comemoração na tela!
                        st.success("🔥 Conciliação finalizada perfeitamente!")
                        
                        # Botão que entrega o relatório pronto para o cliente baixar
                        st.download_button(
                            label="📥 Baixar Relatório de Conciliação (.xlsx)",
                            data=processed_data,
                            file_name="resultado_conciliacao.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                except Exception as e:
                    st.error(f"Erro ao processar as planilhas: {e}")

elif senha_usuario != "":
    st.error("❌ Senha incorreta ou período de teste expirado. Entre em contato com o desenvolvedor.")
else:
    st.info("🔒 Digite a senha fornecida pelo desenvolvedor para acessar o sistema.")