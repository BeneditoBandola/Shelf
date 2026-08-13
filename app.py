import streamlit as st
import pandas as pd
import unicodedata
import os
import smtplib
import pytz
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Configuração da Página
st.set_page_config(page_title="Shelf Space Royal Canin", page_icon="🐾", layout="centered")

st.title("🐾 Shelf Space Royal Canin")
st.markdown("---")

# 1. Carregar Planilha de Clientes Local
@st.cache_data
def load_clients():
    try:
        return pd.read_excel("Cópia de clientes com cnpj corretinho novinho.xlsx")
    except Exception as e:
        st.error(f"Erro ao carregar a planilha de clientes: {e}")
        return pd.DataFrame()

df_clientes = load_clients()

# --- FUNÇÃO PARA SALVAR NO GOOGLE SHEETS ---
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

# --- FUNÇÃO DE GERAÇÃO DE PDF (SIMPLES E COMPLETO) ---
def gerar_pdf_auditoria(promotora, loja, cidade, endereco, dados_completos, nota_total, frentes_dados):
    loja_limpa = "".join([c for c in loja if c.isalnum() or c in (' ', '_', '-')]).strip().replace(' ', '_')
    arq_simples = f"Auditoria_Simples_{loja_limpa}.pdf"
    arq_completo = f"Auditoria_Completa_{loja_limpa}.pdf"

    estilos = getSampleStyleSheet()
    style_celula = ParagraphStyle('EstiloCelula', parent=estilos['Normal'], fontSize=8, leading=10, textColor=colors.HexColor('#1F2937'))
    style_celula_cab = ParagraphStyle('EstiloCelulaCab', parent=estilos['Normal'], fontSize=8, leading=10, textColor=colors.white, fontName="Helvetica-Bold")

    # Ajuste do Fuso Horário e Padrão de Data Brasileiro (DD/MM/AAAA)
    fuso_sp = pytz.timezone('America/Sao_Paulo')
    agora = datetime.now(fuso_sp).strftime("%d/%m/%Y %H:%M")
    
    # ==========================================
    # PÁGINA 1: RELATÓRIO EXECUTIVO (COMUM AOS DOIS)
    # ==========================================
    elem_comuns = []
    elem_comuns.append(Paragraph("<b>RELATÓRIO EXECUTIVO - AUDITORIA SHELF SPACE</b>", estilos['Title']))
    elem_comuns.append(Spacer(1, 5))
    elem_comuns.append(Paragraph(f"<b>LOJA:</b> {loja} | <b>CIDADE:</b> {cidade}", estilos['Normal']))
    elem_comuns.append(Paragraph(f"<b>ENDEREÇO:</b> {endereco}", estilos['Normal']))
    elem_comuns.append(Paragraph(f"<b>PROMOTORA:</b> {promotora} | <b>DATA/HORA:</b> {agora}", estilos['Normal']))
    elem_comuns.append(Paragraph(f"<b>NOTA FINAL DO PDV:</b> <font color='#1E3A8A'><b>{nota_total:.2f} / 10.0 pts</b></font>", estilos['Heading2']))
    elem_comuns.append(Spacer(1, 10))

    # Tabela 1: Resumo de Shares com Gráfico de Barras Colorido
    elem_comuns.append(Paragraph("<b>1. PERFORMANCE DE SHARE POR PILAR</b>", estilos['Heading3']))
    
    data_share = [[
        Paragraph("<b>Pilar</b>", style_celula_cab),
        Paragraph("<b>Share Atual</b>", style_celula_cab),
        Paragraph("<b>Meta</b>", style_celula_cab),
        Paragraph("<b>Progresso Visual (Share vs Meta)</b>", style_celula_cab)
    ]]

    shares = [
        ("Linha Cão", dados_completos['share_cao'], 30.0),
        ("Linha Gato", dados_completos['share_gato'], 35.0),
        ("Linha Veterinária", dados_completos['share_vet'], 50.0)
    ]

    for nome_pilar, val, meta in shares:
        blocos_cheios = int(min(val, 100) / 5)
        blocos_vazios = 20 - blocos_cheios
        barra_visual = "█" * blocos_cheios + "░" * blocos_vazios
        
        # Corrigindo a cor e inserindo na barra
        cor_txt = "#059669" if val >= meta else "#DC2626" # Verde se bater a meta, Vermelho se não bater
        
        data_share.append([
            Paragraph(nome_pilar, style_celula),
            Paragraph(f"<font color='{cor_txt}'><b>{val:.1f}%</b></font>", style_celula),
            Paragraph(f"{meta:.1f}%", style_celula),
            # A barra agora recebe a mesma cor dinamicamente:
            Paragraph(f"<font name='Courier' size=9 color='{cor_txt}'>{barra_visual}</font>", style_celula)
        ])

    t_share = Table(data_share, colWidths=[100, 70, 50, 340])
    t_share.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1E3A8A')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ]))
    elem_comuns.append(t_share)
    elem_comuns.append(Spacer(1, 10))

    # Tabela 2: Execução e Pilares
    elem_comuns.append(Paragraph("<b>2. CHECKLIST DE EXECUÇÃO (PLANOGRAMA & FLUXO)</b>", estilos['Heading3']))
    data_exec = [
        [Paragraph("<b>Pilar / Categoria</b>", style_celula_cab), Paragraph("<b>Planograma?</b>", style_celula_cab), Paragraph("<b>Abertura de Fluxo?</b>", style_celula_cab), Paragraph("<b>Regra Específica</b>", style_celula_cab)],
        [Paragraph("Linha Cão", style_celula), Paragraph(dados_completos['plano_cao'], style_celula), Paragraph(dados_completos['fluxo_cao'], style_celula), Paragraph("Meta de Frentes >= 30%", style_celula)],
        [Paragraph("Linha Gato", style_celula), Paragraph(dados_completos['plano_gato'], style_celula), Paragraph(dados_completos['fluxo_gato'], style_celula), Paragraph(f"Super Premium Separada: {dados_completos['sep_fhn']}", style_celula)],
        [Paragraph("Linha Veterinária", style_celula), Paragraph(dados_completos['plano_vet'], style_celula), Paragraph(dados_completos['fluxo_vet'], style_celula), Paragraph("Meta de Frentes >= 50%", style_celula)],
    ]
    t_exec = Table(data_exec, colWidths=[110, 80, 95, 275])
    t_exec.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#059669')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ]))
    elem_comuns.append(t_exec)
    elem_comuns.append(Spacer(1, 10))

    # Tabela 3: Merchandising e Pontos Extras
    elem_comuns.append(Paragraph("<b>3. MERCHANDISING E PONTOS EXTRAS</b>", estilos['Heading3']))
    materiais_str = ", ".join(dados_completos['materiais_ativos']) if dados_completos['materiais_ativos'] else "Nenhum material encontrado"
    data_merch = [
        [Paragraph("<b>Descrição dos Elementos</b>", style_celula_cab), Paragraph("<b>Status / Quantidade</b>", style_celula_cab)],
        [Paragraph("Materiais de Merchandising Presentes", style_celula), Paragraph(materiais_str, style_celula)],
        [Paragraph("Estado de Conservação dos Materiais", style_celula), Paragraph(dados_completos['conservacao'], style_celula)],
        [Paragraph("Quantidade de Pontos Extras no PDV", style_celula), Paragraph(str(dados_completos['qtd_extras']), style_celula)],
    ]
    t_merch = Table(data_merch, colWidths=[200, 360])
    t_merch.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1E3A8A')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ]))
    elem_comuns.append(t_merch)
    elem_comuns.append(Spacer(1, 10))

    # Observações da Promotora (Se preenchido)
    if dados_completos['observacoes']:
        elem_comuns.append(Paragraph("<b>4. OBSERVAÇÕES / COMENTÁRIOS DA PROMOTORA</b>", estilos['Heading3']))
        data_obs = [[Paragraph(dados_completos['observacoes'], style_celula)]]
        t_obs = Table(data_obs, colWidths=[560])
        t_obs.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#F3F4F6')),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#9CA3AF')),
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('TOPPADDING', (0,0), (-1,-1), 6),
            ('BOTTOMPADDING', (0,0), (-1,-1), 6),
            ('LEFTPADDING', (0,0), (-1,-1), 6),
            ('RIGHTPADDING', (0,0), (-1,-1), 6),
        ]))
        elem_comuns.append(t_obs)
        elem_comuns.append(Spacer(1, 10))

    # Relatório de Oportunidades de Melhoria (Automático)
    elem_comuns.append(Paragraph("<b>5. PLANO DE AÇÃO E OPORTUNIDADES DE MELHORIA</b>", estilos['Heading3']))
    
    melhorias = []
    if dados_completos['share_cao'] < 30.0:
        melhorias.append("• <b>Linha Cão:</b> O share atual está abaixo da meta de 30%. É necessário negociar mais espaço na gôndola e recuperar frentes frente aos concorrentes.")
    if dados_completos['plano_cao'] == "Não":
        melhorias.append("• <b>Planograma Cão:</b> A linha não está respeitando o planograma oficial. Ajustar a disposição dos produtos.")
    if dados_completos['fluxo_cao'] == "Não":
        melhorias.append("• <b>Fluxo Cão:</b> A Royal Canin não está abrindo o fluxo principal. Realizar abordagem com o gerente da loja.")
    
    if dados_completos['share_gato'] < 35.0:
        melhorias.append("• <b>Linha Gato:</b> O share está abaixo de 35%. Expandir frentes de focado em gatos.")
    if dados_completos['plano_gato'] == "Não":
        melhorias.append("• <b>Planograma Gato:</b> Ausente no planograma estabelecido. Necessário reorganizar a seção.")
    if dados_completos['fluxo_gato'] == "Não":
        melhorias.append("• <b>Fluxo Gato:</b> Royal Canin precisa abrir o fluxo de gatos no PDV.")
    if dados_completos['sep_fhn'] == "Não":
        melhorias.append("• <b>Super Premium Cat:</b> A linha Super Premium Cat não está separada da FHN. Executar o bolsão correto.")

    if dados_completos['share_vet'] < 50.0:
        melhorias.append("• <b>Linha Veterinária:</b> O share está abaixo de 50%. Fortalecer a presença de tratamento e prescrição.")
    if dados_completos['plano_vet'] == "Não":
        melhorias.append("• <b>Planograma Vet:</b> Linha Veterinária fora do planograma padrão.")
    if dados_completos['fluxo_vet'] == "Não":
        melhorias.append("• <b>Fluxo Vet:</b> Necessário abrir fluxo para os produtos veterinários.")

    if dados_completos['conservacao'] == "Não":
        melhorias.append("• <b>Conservação:</b> Os materiais de merchandising apresentam problemas de conservação e precisam ser substituídos ou higienizados.")
    if len(dados_completos['materiais_ativos']) < 2:
        melhorias.append("• <b>Merchandising:</b> Baixa presença de materiais de PDV (menos de 2 ativos). Instalar faixas, stoppers ou displays carona.")
    if dados_completos['qtd_extras'] == 0:
        melhorias.append("• <b>Pontos Extras:</b> Nenhum ponto extra localizado na loja. Negociar pontas de gôndola ou ilhas promocionais.")

    if not melhorias:
        melhorias.append("• <b>Parabéns!</b> Excelente execução no PDV. Todos os pilares, shares e materiais estão atingindo ou superando as metas estabelecidas.")

    texto_melhorias = "<br/>".join(melhorias)
    
    data_melhoria = [[Paragraph(texto_melhorias, style_celula)]]
    t_melhoria = Table(data_melhoria, colWidths=[560])
    t_melhoria.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#FEF3C7')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#D97706')),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
    ]))
    elem_comuns.append(t_melhoria)


    # ==========================================
    # PÁGINA 2: DETALHAMENTO (APENAS PDF COMPLETO)
    # ==========================================
    elem_detalhes = []
    elem_detalhes.append(PageBreak())
    elem_detalhes.append(Paragraph("<b>ANEXO: DETALHAMENTO POR MARCA E TIPO</b>", estilos['Title']))
    elem_detalhes.append(Spacer(1, 10))

    marcas = [
        "Biofresh (BRF)", "Equilíbrio (ADM/Total)", "Fórmula Natural (Adimax)", 
        "Hill's", "Pro Plan (Nestlé)", "Premier (Premier)", 
        "Nattu (Premier)", "Vet Life (Farmina)", "N&D (Farmina)", "Royal Canin"
    ]

    estilo_tabela_frentes = TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1E3A8A')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
    ])

    # QUADRO 1: LINHA CÃO
    elem_detalhes.append(Paragraph("<b>Frentes Encontradas - Linha CÃO</b>", estilos['Heading3']))
    data_cao = [[Paragraph("<b>Marca / Concorrente</b>", style_celula_cab), Paragraph("<b>Qtd Frentes</b>", style_celula_cab)]]
    for m in marcas:
        data_cao.append([Paragraph(m, style_celula), Paragraph(str(frentes_dados['cao'].get(m, 0)), style_celula)])
    t_cao = Table(data_cao, colWidths=[200, 100])
    t_cao.setStyle(estilo_tabela_frentes)
    elem_detalhes.append(t_cao)
    elem_detalhes.append(Spacer(1, 15))

    # QUADRO 2: LINHA GATO
    elem_detalhes.append(Paragraph("<b>Frentes Encontradas - Linha GATO</b>", estilos['Heading3']))
    data_gato = [[Paragraph("<b>Marca / Concorrente</b>", style_celula_cab), Paragraph("<b>Qtd Frentes</b>", style_celula_cab)]]
    for m in marcas:
        data_gato.append([Paragraph(m, style_celula), Paragraph(str(frentes_dados['gato'].get(m, 0)), style_celula)])
    t_gato = Table(data_gato, colWidths=[200, 100])
    t_gato.setStyle(estilo_tabela_frentes)
    elem_detalhes.append(t_gato)
    elem_detalhes.append(Spacer(1, 15))

    # QUADRO 3: LINHA VETERINÁRIA
    elem_detalhes.append(Paragraph("<b>Frentes Encontradas - Linha VET</b>", estilos['Heading3']))
    data_vet = [[Paragraph("<b>Marca / Concorrente</b>", style_celula_cab), Paragraph("<b>Qtd Frentes</b>", style_celula_cab)]]
    for m in marcas:
        data_vet.append([Paragraph(m, style_celula), Paragraph(str(frentes_dados['vet'].get(m, 0)), style_celula)])
    t_vet = Table(data_vet, colWidths=[200, 100])
    t_vet.setStyle(estilo_tabela_frentes)
    elem_detalhes.append(t_vet)
    elem_detalhes.append(Spacer(1, 15))

    # QUADRO 4: DETALHAMENTO DE MATERIAIS E PONTOS EXTRAS
    elem_detalhes.append(Paragraph("<b>Detalhamento de Execução e Materiais</b>", estilos['Heading3']))
    materiais_selecionados = ", ".join(dados_completos['materiais_ativos']) if dados_completos['materiais_ativos'] else "Nenhum material assinalado."
    
    elem_detalhes.append(Paragraph(f"<b>Quantidade de Pontos Extras:</b> {dados_completos['qtd_extras']}", estilos['Normal']))
    elem_detalhes.append(Spacer(1, 5))
    elem_detalhes.append(Paragraph(f"<b>Materiais Específicos Encontrados:</b> {materiais_selecionados}", estilos['Normal']))

    # GERANDO OS DOIS ARQUIVOS
    doc_simples = SimpleDocTemplate(arq_simples, pagesize=A4, rightMargin=25, leftMargin=25, topMargin=25, bottomMargin=25)
    doc_simples.build(elem_comuns)

    doc_completo = SimpleDocTemplate(arq_completo, pagesize=A4, rightMargin=25, leftMargin=25, topMargin=25, bottomMargin=25)
    doc_completo.build(elem_comuns + elem_detalhes)

    return arq_simples, arq_completo

# --- FUNÇÃO DE ENVIO DE E-MAIL (AGORA ACEITA MÚLTIPLOS PDFS) ---
def enviar_email_auditoria(assunto, pdf_paths):
    remetente = "beneditobandola@gmail.com"
    senha = "kfih ccqx cskn oito"
    destino = "benedito.bandola@minassal.com.br"

    msg = MIMEMultipart()
    msg['From'], msg['To'], msg['Subject'] = remetente, destino, assunto

    try:
        # Loop para anexar todos os PDFs na lista
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


# --- FLUXO PRINCIPAL DO APP ---
if df_clientes.empty:
    st.warning("A planilha de clientes não foi encontrada ou está vazia.")
else:
    def normalize(text):
        if pd.isna(text):
            return ""
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

        cidades_alvo = cidades_map[promotora]
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
            # --- BOTÃO DE FINALIZAÇÃO, CÁLCULO E ENVIO ---
            if st.button("Finalizar, Salvar e Enviar Auditoria", type="primary"):
                nota_total = 0.0
                detalhes = []

                # Pilar Cão
                p_cao = 0.0
                if plano_cao == "Sim": p_cao += 1.0
                if fluxo_cao == "Sim": p_cao += 1.0
                if share_cao >= 30.0: p_cao += 0.5
                nota_total += p_cao
                detalhes.append(f"Pilar Cão: {p_cao} pts")

                # Pilar Gato
                p_gato = 0.0
                if plano_gato == "Sim": p_gato += 1.0
                if fluxo_gato == "Sim": p_gato += 1.0
                if share_gato >= 35.0: p_gato += 0.5
                if sep_fhn == "Sim": p_gato += 0.5
                nota_total += p_gato
                detalhes.append(f"Pilar Gato: {p_gato} pts")

                # Pilar Vet
                p_vet = 0.0
                if plano_vet == "Sim": p_vet += 1.0
                if fluxo_vet == "Sim": p_vet += 1.0
                if share_vet >= 50.0: p_vet += 0.5
                nota_total += p_vet
                detalhes.append(f"Pilar Vet: {p_vet} pts")

                # Merchandising (Presença)
                materiais_ativos_lista = [m for m in materiais if mat_presenca[m]]
                total_materiais = len(materiais_ativos_lista)
                p_merch = 0.0
                if total_materiais >= 3: p_merch = 0.75
                elif total_materiais == 2: p_merch = 0.50
                elif total_materiais == 1: p_merch = 0.25
                nota_total += p_merch
                detalhes.append(f"Merchandising: {p_merch} pts")

                # Conservação
                p_cons = 0.25 if conservacao == "Sim" else 0.0
                nota_total += p_cons
                detalhes.append(f"Conservação: {p_cons} pts")

                # Pontos Extras
                p_extras = 0.0
                if qtd_pontos_extras >= 3: p_extras = 1.0
                elif qtd_pontos_extras == 2: p_extras = 0.5
                elif qtd_pontos_extras == 1: p_extras = 0.25
                nota_total += p_extras
                detalhes.append(f"Pontos Extras: {p_extras} pts")

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

                # Fuso horário para gravação na Planilha
                fuso_sp = pytz.timezone('America/Sao_Paulo')
                data_atual = datetime.now(fuso_sp).strftime("%d/%m/%Y %H:%M:%S")
                linha_dados = [data_atual, promotora, cidade_loja, loja_selecionada, f"{nota_total:.2f}"]

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
