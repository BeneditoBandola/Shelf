import streamlit as st
import pandas as pd
import unicodedata
import os
import smtplib
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
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
        creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
        client = gspread.authorize(creds)
        
        sheet = client.open("Historico_Auditorias_RoyalCanin").sheet1
        sheet.append_row(dados_auditoria)
        return True
    except Exception as e:
        print(f"Erro ao salvar no Google Sheets: {e}")
        return False

# --- FUNÇÃO DE GERAÇÃO DE PDF ---
def gerar_pdf_auditoria(promotora, loja, cidade, detalhes_auditoria, nota_total):
    loja_limpa = "".join([c for c in loja if c.isalnum() or c in (' ', '_', '-')]).strip().replace(' ', '_')
    nome_arquivo = f"Auditoria_ShelfSpace_{loja_limpa}.pdf"

    doc = SimpleDocTemplate(nome_arquivo, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    elementos, estilos = [], getSampleStyleSheet()
    
    style_celula = ParagraphStyle('EstiloCelula', parent=estilos['Normal'], fontSize=9, leading=12, textColor=colors.HexColor('#1F2937'))
    style_celula_cab = ParagraphStyle('EstiloCelulaCab', parent=estilos['Normal'], fontSize=10, leading=13, textColor=colors.white, fontName="Helvetica-Bold")

    agora = datetime.now().strftime("%d/%m/%Y %H:%M")
    
    elementos.append(Paragraph("<b>RELATÓRIO DE AUDITORIA - SHELF SPACE</b>", estilos['Title']))
    elementos.append(Spacer(1, 10))
    elementos.append(Paragraph(f"<b>LOJA:</b> {loja} | <b>CIDADE:</b> {cidade}", estilos['Normal']))
    elementos.append(Paragraph(f"<b>PROMOTORA:</b> {promotora} | <b>DATA/HORA:</b> {agora}", estilos['Normal']))
    elementos.append(Paragraph(f"<b>NOTA FINAL DO PDV:</b> <font color='#1E3A8A'><b>{nota_total:.2f} / 10.0 pts</b></font>", estilos['Heading2']))
    elementos.append(Spacer(1, 15))

    elementos.append(Paragraph("<b>DETALHAMENTO DA PONTUAÇÃO</b>", estilos['Heading3']))
    
    data_tabela = [[Paragraph("<b>Critério / Pilar</b>", style_celula_cab), Paragraph("<b>Resultado</b>", style_celula_cab)]]
    for item in detalhes_auditoria:
        partes = item.split(":")
        criterio = partes[0]
        pontos = partes[1] if len(partes) > 1 else ""
        data_tabela.append([Paragraph(criterio, style_celula), Paragraph(pontos, style_celula)])

    t1 = Table(data_tabela, colWidths=[350, 150])
    t1.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1E3A8A')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('TOPPADDING', (0,0), (-1,-1), 6),
    ]))
    elementos.append(t1)

    doc.build(elementos)
    return nome_arquivo

# --- FUNÇÃO DE ENVIO DE E-MAIL ---
def enviar_email_auditoria(assunto, pdf_path):
    remetente = "beneditobandola@gmail.com"
    senha = "kfih ccqx cskn oito"
    destino = "benedito.bandola@minassal.com.br"

    msg = MIMEMultipart()
    msg['From'], msg['To'], msg['Subject'] = remetente, destino, assunto

    try:
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
    st.warning("A planilha de clientes não foi encontrada ou está vazia. Certifique-se de que o arquivo 'Cópia de clientes com cnpj corretinho novinho.xlsx' está na mesma pasta.")
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
            st.info(f"📍 **Cidade:** {cidade_loja} | **Endereço:** {dados_loja.get('ENDEREÇO', '')}")
            
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
            st.caption("Display Alimento Seco, Display Alimento Úmido, Ponta de Gôndola, Ilha, Vitrine")
            qtd_pontos_extras = st.number_input("Quantidade de Pontos Extras encontrados:", min_value=0, max_value=10, value=0)

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
                total_materiais = sum(1 for m in materiais if mat_presenca[m])
                p_merch = 0.0
                if total_materiais >= 3: p_merch = 0.75
                elif total_materiais == 2: p_merch = 0.50
                elif total_materiais == 1: p_merch = 0.25
                nota_total += p_merch
                detalhes.append(f"Merchandising (Materiais: {total_materiais}): {p_merch} pts")

                # Conservação
                p_cons = 0.25 if conservacao == "Sim" else 0.0
                nota_total += p_cons
                detalhes.append(f"Conservação dos Materiais: {p_cons} pts")

                # Pontos Extras
                p_extras = 0.0
                if qtd_pontos_extras >= 3: p_extras = 1.0
                elif qtd_pontos_extras == 2: p_extras = 0.5
                elif qtd_pontos_extras == 1: p_extras = 0.25
                nota_total += p_extras
                detalhes.append(f"Pontos Extras ({qtd_pontos_extras} un): {p_extras} pts")

                data_atual = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                linha_dados = [data_atual, promotora, cidade_loja, loja_selecionada, f"{nota_total:.2f}"]

                # Processar salvamento e disparo
                with st.spinner("Salvando na planilha, gerando PDF e enviando e-mail..."):
                    # 1. Salvar no Google Sheets
                    salvar_no_google_sheets(linha_dados)
                    # 2. Gerar PDF
                    pdf_gerado = gerar_pdf_auditoria(promotora, loja_selecionada, cidade_loja, detalhes, nota_total)
                    # 3. Enviar E-mail
                    email_enviado = enviar_email_auditoria(f"🐾 RELATÓRIO SHELF SPACE: {loja_selecionada}", pdf_gerado)

                if email_enviado:
                    st.success("Auditoria salva na planilha, PDF gerado e e-mail enviado com sucesso!")
                    st.balloons()
                else:
                    st.warning("Auditoria salva e PDF gerado, mas houve uma falha ao enviar o e-mail automático.")

                st.metric(label="Nota Total do PDV", value=f"{nota_total:.2f} / 10.0 pts")

                with st.expander("Ver Detalhamento da Pontuação"):
                    for d in detalhes:
                        st.write(f"- {d}")
        else:
            st.warning("Nenhuma loja encontrada para esta promotora.")
