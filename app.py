import streamlit as st
import pandas as pd
import unicodedata
import os
import smtplib
import pytz
import json
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Shelf Space Royal Canin", page_icon="🐾", layout="centered")

# --- MENU LATERAL DE NAVEGAÇÃO ---
st.sidebar.image("https://upload.wikimedia.org/wikipedia/commons/thumb/1/14/Royal_Canin_logo.svg/2560px-Royal_Canin_logo.svg.png", width=150)
st.sidebar.markdown("---")
menu = st.sidebar.radio("Navegação", ["📝 Área da Promotora", "📊 Dados Benedito (Admin)"])
st.sidebar.markdown("---")

# --- 1. CARREGAR PLANILHA DE CLIENTES ---
@st.cache_data
def load_clients():
    try:
        return pd.read_excel("Cópia de clientes com cnpj corretinho novinho.xlsx")
    except Exception as e:
        st.error(f"Erro ao carregar a planilha de clientes: {e}")
        return pd.DataFrame()

df_clientes = load_clients()

# --- FUNÇÕES DE BANCO DE DADOS E GERAÇÃO DE PDF ---
def salvar_no_google_sheets(dados_auditoria):
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        
        # Leitura da credencial JSON e limpeza de quebras de linha
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
        
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        
        sheet = client.open("Historico_Auditorias_RoyalCanin").sheet1
        sheet.append_row(dados_auditoria)
        return True
    except Exception as e:
        st.error(f"Erro ao salvar no Google Sheets: {e}")
        return False

def gerar_pdf_auditoria(promotora, loja, cidade, endereco, dados_completos, nota_total, frentes_dados):
    loja_limpa = "".join([c for c in loja if c.isalnum() or c in (' ', '_', '-')]).strip().replace(' ', '_')
    arq_simples = f"Auditoria_Simples_{loja_limpa}.pdf"
    arq_completo = f"Auditoria_Completa_{loja_limpa}.pdf"

    estilos = getSampleStyleSheet()
    style_celula = ParagraphStyle('EstiloCelula', parent=estilos['Normal'], fontSize=8, leading=10, textColor=colors.HexColor('#1F2937'))
    style_celula_cab = ParagraphStyle('EstiloCelulaCab', parent=estilos['Normal'], fontSize=8, leading=10, textColor=colors.white, fontName="Helvetica-Bold")
    
    cor_cabecalho_principal = colors.HexColor('#1E3A8A')
    cor_cabecalho_secundario = colors.HexColor('#059669')
    
    fuso_sp = pytz.timezone('America/Sao_Paulo')
    agora = datetime.now(fuso_sp).strftime("%d/%m/%Y %H:%M")

    def criar_pagina_1():
        elem = []
        elem.append(Paragraph("<b>RELATÓRIO EXECUTIVO - AUDITORIA SHELF SPACE</b>", estilos['Title']))
        elem.append(Spacer(1, 5))
        elem.append(Paragraph(f"<b>LOJA:</b> {loja} | <b>CIDADE:</b> {cidade}", estilos['Normal']))
        elem.append(Paragraph(f"<b>ENDEREÇO:</b> {endereco}", estilos['Normal']))
        elem.append(Paragraph(f"<b>PROMOTORA:</b> {promotora} | <b>DATA/HORA:</b> {agora}", estilos['Normal']))
        elem.append(Paragraph(f"<b>NOTA FINAL DO PDV:</b> <font color='#1E3A8A'><b>{nota_total:.2f} / 10.0 pts</b></font>", estilos['Heading2']))
        elem.append(Spacer(1, 10))

        elem.append(Paragraph("<b>1. PERFORMANCE DE SHARE POR PILAR</b>", estilos['Heading3']))
        data_share = [[
            Paragraph("<b>Pilar</b>", style_celula_cab),
            Paragraph("<b>Share</b>", style_celula_cab),
            Paragraph("<b>Meta</b>", style_celula_cab),
            Paragraph("<b>Progresso (Visual Base 10)</b>", style_celula_cab)
        ]]

        shares = [
            ("Linha Cão", dados_completos['share_cao'], 30.0),
            ("Linha Gato", dados_completos['share_gato'], 35.0),
            ("Linha Veterinária", dados_completos['share_vet'], 50.0)
        ]

        for nome_pilar, val, meta in shares:
            blocos_cheios = int(round(min(val, 100) / 10))
            blocos_vazios = 10 - blocos_cheios
            barra_visual = "[" + ("█" * blocos_cheios) + ("░" * blocos_vazios) + "]"
            cor_txt = "#059669" if val >= meta else "#DC2626" 
            
            data_share.append([
                Paragraph(nome_pilar, style_celula),
                Paragraph(f"<font color='{cor_txt}'><b>{val:.1f}%</b></font>", style_celula),
                Paragraph(f"{meta:.1f}%", style_celula),
                Paragraph(f"<font name='Courier' size=11 color='{cor_txt}'>{barra_visual}</font>", style_celula)
            ])

        t_share = Table(data_share, colWidths=[100, 50, 40, 350])
        t_share.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), cor_cabecalho_principal), ('GRID', (0,0), (-1,-1), 0.5, colors.grey), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
        elem.append(t_share)
        elem.append(Spacer(1, 10))

        elem.append(Paragraph("<b>2. CHECKLIST DE EXECUÇÃO</b>", estilos['Heading3']))
        data_exec = [
            [Paragraph("<b>Pilar / Categoria</b>", style_celula_cab), Paragraph("<b>Planograma?</b>", style_celula_cab), Paragraph("<b>Abertura Fluxo?</b>", style_celula_cab), Paragraph("<b>Regra Específica</b>", style_celula_cab)],
            [Paragraph("Linha Cão", style_celula), Paragraph(dados_completos['plano_cao'], style_celula), Paragraph(dados_completos['fluxo_cao'], style_celula), Paragraph("Meta de Frentes >= 30%", style_celula)],
            [Paragraph("Linha Gato", style_celula), Paragraph(dados_completos['plano_gato'], style_celula), Paragraph(dados_completos['fluxo_gato'], style_celula), Paragraph(f"Super Premium Sep: {dados_completos['sep_fhn']}", style_celula)],
            [Paragraph("Linha Veterinária", style_celula), Paragraph(dados_completos['plano_vet'], style_celula), Paragraph(dados_completos['fluxo_vet'], style_celula), Paragraph("Meta de Frentes >= 50%", style_celula)],
        ]
        t_exec = Table(data_exec, colWidths=[110, 80, 95, 255])
        t_exec.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), cor_cabecalho_secundario), ('GRID', (0,0), (-1,-1), 0.5, colors.grey), ('VALIGN', (0,0), (-1,-1), 'MIDDLE')]))
        elem.append(t_exec)
        elem.append(Spacer(1, 10))

        if dados_completos['observacoes']:
            elem.append(Paragraph("<b>3. OBSERVAÇÕES DA PROMOTORA</b>", estilos['Heading3']))
            t_obs = Table([[Paragraph(dados_completos['observacoes'], style_celula)]], colWidths=[540])
            t_obs.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor('#F3F4F6')), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#9CA3AF')), ('VALIGN', (0,0), (-1,-1), 'TOP')]))
            elem.append(t_obs)
        return elem

    def criar_pagina_2():
        elem = []
        elem.append(PageBreak())
        elem.append(Paragraph("<b>ANEXO: MAPEAMENTO DETALHADO DO PDV</b>", estilos['Title']))
        return elem

    doc_simples = SimpleDocTemplate(arq_simples, pagesize=A4, rightMargin=25, leftMargin=25, topMargin=25, bottomMargin=25)
    doc_simples.build(criar_pagina_1())
    doc_completo = SimpleDocTemplate(arq_completo, pagesize=A4, rightMargin=25, leftMargin=25, topMargin=25, bottomMargin=25)
    doc_completo.build(criar_pagina_1() + criar_pagina_2())
    return arq_simples, arq_completo

def enviar_email_auditoria(assunto, pdf_paths):
    remetente = "beneditobandola@gmail.com"
    senha = "kfih ccqx cskn oito"
    destino = "benedito.bandola@minassal.com.br"
    msg = MIMEMultipart()
    msg['From'], msg['To'], msg['Subject'] = remetente, destino, assunto
    try:
        for pdf_path in pdf_paths:
            with open(pdf_path, "rb") as f:
                part = MIMEApplication(f.read(), Name=os.path.basename(pdf_path))
                part.add_header('Content-Disposition', 'attachment', filename=os.path.basename(pdf_path))
                msg.attach(part)
        s = smtplib.SMTP('smtp.gmail.com', 587)
        s.starttls()
        s.login(remetente, senha)
        s.sendmail(remetente, destino, msg.as_string())
        s.quit()
        return True
    except Exception: return False

# ==============================================================
# FLUXO 1: ÁREA DA PROMOTORA
# ==============================================================
if menu == "📝 Área da Promotora":
    st.title("🐾 Shelf Space Royal Canin")
    promotora = st.selectbox("Promotora:", ["Selecione...", "Pamela", "Fernanda", "Madalla"])
    if promotora != "Selecione...":
        # (Restante do fluxo da promotora permanece igual)
        st.success("Continue o preenchimento...")

# ==============================================================
# FLUXO 2: DADOS BENEDITO (ADMIN)
# ==============================================================
elif menu == "📊 Dados Benedito (Admin)":
    st.title("📊 Painel Administrativo")
    senha = st.text_input("Senha:", type="password")
    if senha == "1234":
        st.success("Acesso Liberado!")
        # (Lógica de puxar histórico aqui)
    elif senha != "":
        st.error("Senha incorreta.")
