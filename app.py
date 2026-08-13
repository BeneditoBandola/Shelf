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
        creds_dict = dict(st.secrets["gcp_service_account"])
        creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        client = gspread.authorize(creds)
        
        sheet = client.open("Historico_Auditorias_RoyalCanin").sheet1
        sheet.append_row(dados_auditoria)
        return True
    except Exception as e:
        print(f"Erro ao salvar no Google Sheets: {e}")
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
        t_share.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), cor_cabecalho_principal),
            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING', (0,0), (-1,-1), 4),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ]))
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
        t_exec.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), cor_cabecalho_secundario),
            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING', (0,0), (-1,-1), 4),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ]))
        elem.append(t_exec)
        elem.append(Spacer(1, 10))

        if dados_completos['observacoes']:
            elem.append(Paragraph("<b>3. OBSERVAÇÕES DA PROMOTORA</b>", estilos['Heading3']))
            t_obs = Table([[Paragraph(dados_completos['observacoes'], style_celula)]], colWidths=[540])
            t_obs.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor('#F3F4F6')), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#9CA3AF')), ('VALIGN', (0,0), (-1,-1), 'TOP')]))
            elem.append(t_obs)
            elem.append(Spacer(1, 10))

        elem.append(Paragraph("<b>4. PLANO DE AÇÃO (FEEDBACK AUTOMÁTICO)</b>", estilos['Heading3']))
        melhorias = []
        if dados_completos['share_cao'] < 30.0: melhorias.append("• <b>Linha Cão:</b> Abaixo da meta de 30%. Negociar espaço de gôndola.")
        if dados_completos['plano_cao'] == "Não": melhorias.append("• <b>Planograma Cão:</b> Ajustar exposição oficial da linha.")
        if dados_completos['fluxo_cao'] == "Não": melhorias.append("• <b>Fluxo Cão:</b> Abrir fluxo de abordagem no PDV.")
        
        if dados_completos['share_gato'] < 35.0: melhorias.append("• <b>Linha Gato:</b> Abaixo da meta de 35%. Expandir frentes focadas.")
        if dados_completos['plano_gato'] == "Não": melhorias.append("• <b>Planograma Gato:</b> Reorganizar conforme o planograma.")
        if dados_completos['fluxo_gato'] == "Não": melhorias.append("• <b>Fluxo Gato:</b> Necessário abrir fluxo na seção.")
        if dados_completos['sep_fhn'] == "Não": melhorias.append("• <b>SP Cat:</b> Separar a linha Super Premium da linha FHN.")
        
        if dados_completos['share_vet'] < 50.0: melhorias.append("• <b>Linha Vet:</b> Abaixo de 50%. Fortalecer presença de tratamento.")
        
        if not melhorias: melhorias.append("• <b>Parabéns!</b> Execução impecável. Pilares atingindo as metas.")
        
        t_melhoria = Table([[Paragraph("<br/>".join(melhorias), style_celula)]], colWidths=[540])
        t_melhoria.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor('#FEF3C7')), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#D97706'))]))
        elem.append(t_melhoria)
        
        return elem

    def criar_pagina_2():
        elem = []
        elem.append(PageBreak())
        elem.append(Paragraph("<b>ANEXO: MAPEAMENTO DETALHADO DO PDV</b>", estilos['Title']))
        elem.append(Spacer(1, 10))

        elem.append(Paragraph("<b>1. DETALHAMENTO DE FRENTES POR MARCA</b>", estilos['Heading3']))
        marcas = [
            "Biofresh (BRF)", "Equilíbrio (ADM/Total)", "Fórmula Natural (Adimax)", 
            "Hill's", "Pro Plan (Nestlé)", "Premier (Premier)", 
            "Nattu (Premier)", "Vet Life (Farmina)", "N&D (Farmina)", "Royal Canin"
        ]
        
        data_frentes = [[
            Paragraph("<b>Marca / Concorrente</b>", style_celula_cab), 
            Paragraph("<b>Frentes Cão</b>", style_celula_cab), 
            Paragraph("<b>Frentes Gato</b>", style_celula_cab), 
            Paragraph("<b>Frentes Vet</b>", style_celula_cab)
        ]]
        
        for m in marcas:
            data_frentes.append([
                Paragraph(m, style_celula), 
                Paragraph(str(frentes_dados['cao'].get(m, 0)), style_celula),
                Paragraph(str(frentes_dados['gato'].get(m, 0)), style_celula),
                Paragraph(str(frentes_dados['vet'].get(m, 0)), style_celula)
            ])
            
        t_frentes = Table(data_frentes, colWidths=[210, 110, 110, 110])
        t_frentes.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), cor_cabecalho_principal),
            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING', (0,0), (-1,-1), 4),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
            ('ALIGN', (1,1), (-1,-1), 'CENTER'),
        ]))
        elem.append(t_frentes)
        elem.append(Spacer(1, 15))

        elem.append(Paragraph("<b>2. RESPOSTAS DO QUESTIONÁRIO DE AUDITORIA</b>", estilos['Heading3']))
        
        materiais_str = ", ".join(dados_completos['materiais_ativos']) if dados_completos['materiais_ativos'] else "Nenhum material assinalado"
        
        data_respostas = [
            [Paragraph("<b>Indicador Avaliado</b>", style_celula_cab), Paragraph("<b>Resposta Registrada</b>", style_celula_cab)],
            [Paragraph("<b>Linha Cão:</b> Cumpre o Planograma?", style_celula), Paragraph(dados_completos['plano_cao'], style_celula)],
            [Paragraph("<b>Linha Cão:</b> Abertura de Fluxo garantida?", style_celula), Paragraph(dados_completos['fluxo_cao'], style_celula)],
            [Paragraph("<b>Linha Gato:</b> Cumpre o Planograma?", style_celula), Paragraph(dados_completos['plano_gato'], style_celula)],
            [Paragraph("<b>Linha Gato:</b> Abertura de Fluxo garantida?", style_celula), Paragraph(dados_completos['fluxo_gato'], style_celula)],
            [Paragraph("<b>Linha Gato:</b> Super Premium separada de FHN?", style_celula), Paragraph(dados_completos['sep_fhn'], style_celula)],
            [Paragraph("<b>Linha Vet:</b> Cumpre o Planograma?", style_celula), Paragraph(dados_completos['plano_vet'], style_celula)],
            [Paragraph("<b>Linha Vet:</b> Abertura de Fluxo garantida?", style_celula), Paragraph(dados_completos['fluxo_vet'], style_celula)],
            [Paragraph("<b>Merchandising:</b> Estado de Conservação Adequado?", style_celula), Paragraph(dados_completos['conservacao'], style_celula)],
            [Paragraph("<b>Merchandising:</b> Quantidade de Pontos Extras", style_celula), Paragraph(str(dados_completos['qtd_extras']), style_celula)],
            [Paragraph("<b>Merchandising:</b> Materiais Presentes", style_celula), Paragraph(materiais_str, style_celula)],
        ]
        
        t_respostas = Table(data_respostas, colWidths=[250, 290])
        t_respostas.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), cor_cabecalho_secundario),
            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('TOPPADDING', (0,0), (-1,-1), 4),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ]))
        elem.append(t_respostas)

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
    except Exception as e:
        print(f"Erro ao enviar e-mail: {e}")
        return False

# ==============================================================
# FLUXO 1: ÁREA DA PROMOTORA (OPERAÇÃO PADRÃO)
# ==============================================================
if menu == "📝 Área da Promotora":
    st.title("🐾 Shelf Space Royal Canin")
    st.markdown("---")
    
    if df_clientes.empty:
        st.warning("A planilha de clientes não foi encontrada ou está vazia.")
    else:
        def normalize(text):
            if pd.isna(text): return ""
            nfkd = unicodedata.normalize('NFKD', str(text))
            return "".join([c for c in nfkd if not unicodedata.combining(c)]).upper().strip()

        df_clientes['CIDADE_NORM'] = df_clientes['CIDADE'].apply(normalize)

        st.subheader("1. Identificação")
        promotora = st.selectbox("Selecione a Promotora:", ["Selecione...", "Pamela", "Fernanda", "Madalla"])

        if promotora != "Selecione...":
            cidades_map = {
                "Pamela": ["POCOS DE CALDAS", "ANDRADAS", "VARGINHA", "ALFENAS", "SAO LOURENCO", "ITAJUBA", "TRES PONTAS", "TRES CORACOES"],
                "Fernanda": ["JUIZ DE FORA"],
                "Madalla": ["MURIAE", "VICOSA", "UBA", "VISCONDE DO RIO BRANCO", "PIRAUBA", "RIO POMBA", "GUARANI", "GUIDOVAL", "TOCANTINS", "SAO JOAO NEPOMUCENO", "RIO NOVO", "RODEIRO"]
            }

            cidades_alvo = cidades_map.get(promotora, [])
            df_filtrado = df_clientes[df_clientes['CIDADE_NORM'].isin(cidades_alvo)]
            
            lojas_lista = sorted(df_filtrado['NOME'].dropna().unique().tolist())

            st.subheader("2. Seleção da Loja")
            if len(lojas_lista) > 0:
                loja_selecionada = st.selectbox("Selecione o Cliente / Loja:", lojas_lista)
                
                dados_loja = df_filtrado[df_filtrado['NOME'] == loja_selecionada].iloc[0]
                cidade_loja = dados_loja.get('CIDADE', '')
                endereco_loja = str(dados_loja.get('ENDEREÇO', 'Endereço não informado'))
                st.info(f"📍 **Cidade:** {cidade_loja} | **Endereço:** {endereco_loja}")
                
                st.markdown("---")
                st.subheader("3. Contagem de Frentes & Share por Pilar")
                
                marcas = [
                    "Biofresh (BRF)", "Equilíbrio (ADM/Total)", "Fórmula Natural (Adimax)", 
                    "Hill's", "Pro Plan (Nestlé)", "Premier (Premier)", 
                    "Nattu (Premier)", "Vet Life (Farmina)", "N&D (Farmina)", "Royal Canin"
                ]

                # --- PILAR: CÃO ---
                st.markdown("### 🐶 Linha Cão (Meta: 30%)")
                frentes_cao = {m: st.number_input(f"[Cão] Frentes - {m}", min_value=0, max_value=100, value=0, key=f"cao_{m}") for m in marcas}
                total_cao = sum(frentes_cao.values())
                share_cao = (frentes_cao["Royal Canin"] / total_cao * 100) if total_cao > 0 else 0.0
                st.write(f"📊 **Share Cão Atual:** {share_cao:.1f}% (Meta: 30%)")

                plano_cao = st.radio("Linha Cão está no Planograma?", ["Sim", "Não"], key="plano_cao")
                fluxo_cao = st.radio("Royal Canin está abrindo o Fluxo (Cão)?", ["Sim", "Não"], key="fluxo_cao")

                st.markdown("---")
                # --- PILAR: GATO ---
                st.markdown("### 🐱 Linha Gato (Meta: 35%)")
                frentes_gato = {m: st.number_input(f"[Gato] Frentes - {m}", min_value=0, max_value=100, value=0, key=f"gato_{m}") for m in marcas}
                total_gato = sum(frentes_gato.values())
                share_gato = (frentes_gato["Royal Canin"] / total_gato * 100) if total_gato > 0 else 0.0
                st.write(f"📊 **Share Gato Atual:** {share_gato:.1f}% (Meta: 35%)")

                plano_gato = st.radio("Linha Gato está no Planograma?", ["Sim", "Não"], key="plano_gato")
                fluxo_gato = st.radio("Royal Canin está abrindo o Fluxo (Gato)?", ["Sim", "Não"], key="fluxo_gato")
                sep_fhn = st.radio("Super Premium Cat está separada da linha FHN?", ["Sim", "Não"], key="sep_fhn")

                st.markdown("---")
                # --- PILAR: VETERINÁRIA ---
                st.markdown("### 🩺 Linha Veterinária / Tratamento (Meta: 50%)")
                frentes_vet = {m: st.number_input(f"[Vet] Frentes - {m}", min_value=0, max_value=100, value=0, key=f"vet_{m}") for m in marcas}
                total_vet = sum(frentes_vet.values())
                share_vet = (frentes_vet["Royal Canin"] / total_vet * 100) if total_vet > 0 else 0.0
                st.write(f"📊 **Share Vet Atual:** {share_vet:.1f}% (Meta: 50%)")

                plano_vet = st.radio("Linha Vet está no Planograma?", ["Sim", "Não"], key="plano_vet")
                fluxo_vet = st.radio("Royal Canin está abrindo o Fluxo (Vet)?", ["Sim", "Não"], key="fluxo_vet")

                st.markdown("---")
                # --- MÓDULO 4: MERCHANDISING ---
                st.subheader("4. Merchandising e Presença de Materiais")
                materiais = [
                    "Faixa de Gôndola", "Bobina Forração", "Display Carona", 
                    "Cartazete precificador", "Base de Sacarias (can base)", 
                    "Totem Silhueta", "Cubo", "Clip Strip", "Stopper", "Outros materiais"
                ]
                mat_presenca = {mat: st.checkbox(mat, key=f"mat_{mat}") for mat in materiais}
                conservacao = st.radio("Os materiais estão bem executados e em bom estado de conservação?", ["Sim", "Não"], key="conservacao")

                st.markdown("---")
                # --- MÓDULO 5: PONTOS EXTRAS ---
                st.subheader("5. Pontos Extras Presentes")
                qtd_pontos_extras = st.number_input("Quantidade de Pontos Extras encontrados:", min_value=0, max_value=10, value=0)

                st.markdown("---")
                # --- MÓDULO 6: OBSERVAÇÕES DA PROMOTORA ---
                st.subheader("6. Observações e Comentários")
                observacoes_promotora = st.text_area("Digite aqui qualquer observação relevante sobre o PDV:")

                st.markdown("---")
                # --- BOTÃO DE FINALIZAÇÃO ---
                if st.button("Finalizar, Salvar e Enviar Auditoria", type="primary"):
                    nota_total = 0.0
                    detalhes = []

                    p_cao = 0.0
                    if plano_cao == "Sim": p_cao += 1.0
                    if fluxo_cao == "Sim": p_cao += 1.0
                    if share_cao >= 30.0: p_cao += 0.5
                    nota_total += p_cao

                    p_gato = 0.0
                    if plano_gato == "Sim": p_gato += 1.0
                    if fluxo_gato == "Sim": p_gato += 1.0
                    if share_gato >= 35.0: p_gato += 0.5
                    if sep_fhn == "Sim": p_gato += 0.5
                    nota_total += p_gato

                    p_vet = 0.0
                    if plano_vet == "Sim": p_vet += 1.0
                    if fluxo_vet == "Sim": p_vet += 1.0
                    if share_vet >= 50.0: p_vet += 0.5
                    nota_total += p_vet

                    materiais_ativos_lista = [m for m in materiais if mat_presenca[m]]
                    total_materiais = len(materiais_ativos_lista)
                    p_merch = 0.0
                    if total_materiais >= 3: p_merch = 0.75
                    elif total_materiais == 2: p_merch = 0.50
                    elif total_materiais == 1: p_merch = 0.25
                    nota_total += p_merch

                    p_cons = 0.25 if conservacao == "Sim" else 0.0
                    nota_total += p_cons

                    p_extras = 0.0
                    if qtd_pontos_extras >= 3: p_extras = 1.0
                    elif qtd_pontos_extras == 2: p_extras = 0.5
                    elif qtd_pontos_extras == 1: p_extras = 0.25
                    nota_total += p_extras

                    dados_completos = {
                        'share_cao': share_cao, 'plano_cao': plano_cao, 'fluxo_cao': fluxo_cao,
                        'share_gato': share_gato, 'plano_gato': plano_gato, 'fluxo_gato': fluxo_gato, 'sep_fhn': sep_fhn,
                        'share_vet': share_vet, 'plano_vet': plano_vet, 'fluxo_vet': fluxo_vet,
                        'materiais_ativos': materiais_ativos_lista, 'conservacao': conservacao,
                        'qtd_extras': qtd_pontos_extras, 'observacoes': observacoes_promotora.strip()
                    }

                    frentes_dados = {
                        'cao': frentes_cao,
                        'gato': frentes_gato,
                        'vet': frentes_vet
                    }

                    fuso_sp = pytz.timezone('America/Sao_Paulo')
                    data_atual = datetime.now(fuso_sp).strftime("%d/%m/%Y %H:%M:%S")
                    
                    linha_dados = [
                        data_atual, promotora, cidade_loja, loja_selecionada, f"{nota_total:.2f}",
                        json.dumps(frentes_cao), plano_cao, fluxo_cao,
                        json.dumps(frentes_gato), plano_gato, fluxo_gato, sep_fhn,
                        json.dumps(frentes_vet), plano_vet, fluxo_vet,
                        json.dumps(materiais_ativos_lista), conservacao,
                        qtd_pontos_extras, observacoes_promotora.strip()
                    ]

                    with st.spinner("Salvando na planilha, gerando PDFs e enviando e-mail..."):
                        salvar_no_google_sheets(linha_dados)
                        pdf_simples, pdf_completo = gerar_pdf_auditoria(promotora, loja_selecionada, cidade_loja, endereco_loja, dados_completos, nota_total, frentes_dados)
                        email_enviado = enviar_email_auditoria(f"🐾 RELATÓRIO SHELF SPACE: {loja_selecionada}", [pdf_simples, pdf_completo])

                    if email_enviado:
                        st.success("Auditoria salva na planilha, PDFs Gerados (Simples e Completo) e e-mail enviado com sucesso!")
                        st.balloons()
                    else:
                        st.warning("Auditoria salva e PDFs gerados, mas houve uma falha ao enviar o e-mail automático.")

                    st.metric(label="Nota Total do PDV", value=f"{nota_total:.2f} / 10.0 pts")
            else:
                st.warning("Nenhuma loja encontrada para esta promotora.")

# ==============================================================
# FLUXO 2: DADOS BENEDITO (ÁREA RESTRITA / REPROCESSAMENTO)
# ==============================================================
elif menu == "📊 Dados Benedito (Admin)":
    st.title("📊 Painel de Administração - Shelf Space")
    st.markdown("Acesso restrito para consulta de histórico e reprocessamento de auditorias com regras atualizadas.")
    st.markdown("---")
    
    # --- SENHA DE ACESSO ---
    # Altere a palavra "1234" abaixo para a senha desejada
    senha = st.text_input("Digite a senha de administrador:", type="password")
    
    if senha == "Minassal2026":
        st.success("Acesso Liberado!")
        st.markdown("### Consulta de Auditorias Anteriores")
        
        st.info("🚧 **Área em construção:** O sistema já começou a salvar os dados crus a partir de hoje. Em breve, você poderá usar esta barra de busca para localizar um cliente específico, revisar os números preenchidos pela promotora naquela data e clicar no botão 'Reenviar' para gerar novos PDFs aplicando a métrica de pontuação vigente.")
        
        # Placeholder para a futura barra de pesquisa e botão de reenvio
        loja_busca = st.text_input("Buscar cliente por Nome ou CNPJ:")
        if st.button("Buscar Histórico"):
            st.warning("A conexão de leitura do histórico no Google Sheets será ativada na próxima atualização.")

    elif senha != "":
        st.error("Senha incorreta. Acesso negado.")
