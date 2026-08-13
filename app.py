import streamlit as st
import json

st.title("🔧 Tradutor de Credenciais Google")
st.write("1. Arraste o novo arquivo JSON que você acabou de baixar.")
st.write("2. Copie o bloco abaixo e cole no seu Secrets do Streamlit.")

arquivo = st.file_uploader("Arraste o arquivo JSON aqui", type=["json"])

if arquivo is not None:
    try:
        dados = json.load(arquivo)
        # Formata o TOML exatamente como o Streamlit espera
        toml_output = f"""[gcp_service_account]
type = "service_account"
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
        st.code(toml_output, language="toml")
        st.success("Copie o bloco acima, vá nos Secrets do Streamlit, apague tudo e cole lá!")
    except Exception as e:
        st.error(f"Erro ao processar: {e}")
