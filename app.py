import streamlit as st
import json

st.title("🔧 Consertador de Chave do Google")
st.write("Vamos acabar com o erro de JWT Signature de uma vez por todas!")

arquivo = st.file_uploader("Arraste o seu arquivo .json original do Google aqui", type=["json"])

if arquivo is not None:
    try:
        dados = json.load(arquivo)
        
        # O Python extrai e formata as quebras de linha com precisão cirúrgica
        toml_text = f"""[gcp_service_account]
type = "{dados['type']}"
project_id = "{dados['project_id']}"
private_key_id = "{dados['private_key_id']}"
private_key = "{dados['private_key'].replace('\n', '\\n')}"
client_email = "{dados['client_email']}"
client_id = "{dados['client_id']}"
auth_uri = "{dados['auth_uri']}"
token_uri = "{dados['token_uri']}"
auth_provider_x509_cert_url = "{dados['auth_provider_x509_cert_url']}"
client_x509_cert_url = "{dados['client_x509_cert_url']}"
"""
        st.success("✅ Sucesso! Clique no ícone de 'Copiar' no canto superior direito do bloco preto abaixo e cole TUDO lá na tela de Secrets do Streamlit:")
        st.code(toml_text, language="toml")
        
    except Exception as e:
        st.error(f"Ocorreu um erro ao ler o arquivo: {e}")
