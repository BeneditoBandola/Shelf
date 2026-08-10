import streamlit as st
import pandas as pd
import unicodedata
from datetime import datetime

# Configuração da Página
st.set_page_config(page_title="Auditoria Royal Canin", page_icon="🐾", layout="centered")

st.title("🐾 Sistema de Auditoria PDV - Royal Canin")
st.markdown("---")

# 1. Carregar Planilha de Clientes
@st.cache_data
def load_clients():
    try:
        return pd.read_excel("Cópia de clientes com cnpj corretinho novinho.xlsx")
    except Exception as e:
        st.error(f"Erro ao carregar a planilha de clientes: {e}")
        return pd.DataFrame()

df_clientes = load_clients()

if df_clientes.empty:
    st.warning("A planilha de clientes não foi encontrada ou está vazia. Certifique-se de que o arquivo 'Cópia de clientes com cnpj corretinho novinho.xlsx' está na mesma pasta.")
else:
    # Normalizar cidades para o filtro
    def normalize(text):
        if pd.isna(text):
            return ""
        nfkd = unicodedata.normalize('NFKD', str(text))
        return "".join([c for c in nfkd if not unicodedata.combining(c)]).upper().strip()

    df_clientes['CIDADE_NORM'] = df_clientes['CIDADE'].apply(normalize)

    # 2. Seleção da Promotora
    st.subheader("1. Identificação")
    promotora = st.selectbox("Selecione a Promotora:", ["Selecione...", "Pamela", "Fernanda", "Madalla"])

    if promotora != "Selecione...":
        # Mapeamento de Cidades por Promotora
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
            
            # Recuperar dados da loja selecionada
            dados_loja = df_filtrado[df_filtrado['NOME'] == loja_selecionada].iloc[0]
            st.info(f"📍 **Cidade:** {dados_loja.get('CIDADE', '')} | **Endereço:** {dados_loja.get('ENDEREÇO', '')}")
            
            st.markdown("---")
            st.subheader("3. Contagem de Frentes & Share por Pilar")
            
            marcas = [
                "Biofresh (BRF)",
                "Equilíbrio (ADM/Total)",
                "Fórmula Natural (Adimax)",
                "Hill's",
                "Pro Plan (Nestlé)",
                "Premier (Premier)",
                "Nattu (Premier)",
                "Vet Life (Farmina)",
                "N&D (Farmina)",
                "Royal Canin"
            ]

            # --- PILAR: CÃO ---
            st.markdown("### 🐶 Linha Cão (Meta: 30%)")
            frentes_cao = {}
            total_cao = 0
            for m in marcas:
                frentes_cao[m] = st.number_input(f"[Cão] Frentes - {m}", min_value=0, max_value=100, value=0, key=f"cao_{m}")
                total_cao += frentes_cao[m]
            
            share_cao = (frentes_cao["Royal Canin"] / total_cao * 100) if total_cao > 0 else 0.0
            st.write(f"📊 **Share Cão Atual:** {share_cao:.1f}% (Meta: 30%)")

            plano_cao = st.radio("Linha Cão está no Planograma?", ["Sim", "Não"], key="plano_cao")
            fluxo_cao = st.radio("Royal Canin está abrindo o Fluxo (Cão)?", ["Sim", "Não"], key="fluxo_cao")

            st.markdown("---")
            # --- PILAR: GATO ---
            st.markdown("### 🐱 Linha Gato (Meta: 35%)")
            frentes_gato = {}
            total_gato = 0
            for m in marcas:
                frentes_gato[m] = st.number_input(f"[Gato] Frentes - {m}", min_value=0, max_value=100, value=0, key=f"gato_{m}")
                total_gato += frentes_gato[m]

            share_gato = (frentes_gato["Royal Canin"] / total_gato * 100) if total_gato > 0 else 0.0
            st.write(f"📊 **Share Gato Atual:** {share_gato:.1f}% (Meta: 35%)")

            plano_gato = st.radio("Linha Gato está no Planograma?", ["Sim", "Não"], key="plano_gato")
            fluxo_gato = st.radio("Royal Canin está abrindo o Fluxo (Gato)?", ["Sim", "Não"], key="fluxo_gato")
            sep_fhn = st.radio("Super Premium Cat está separada da linha FHN?", ["Sim", "Não"], key="sep_fhn")

            st.markdown("---")
            # --- PILAR: VETERINÁRIA / TRATAMENTO ---
            st.markdown("### 🩺 Linha Veterinária / Tratamento (Meta: 50%)")
            frentes_vet = {}
            total_vet = 0
            for m in marcas:
                frentes_vet[m] = st.number_input(f"[Vet] Frentes - {m}", min_value=0, max_value=100, value=0, key=f"vet_{m}")
                total_vet += frentes_vet[m]

            share_vet = (frentes_vet["Royal Canin"] / total_vet * 100) if total_vet > 0 else 0.0
            st.write(f"📊 **Share Vet Atual:** {share_vet:.1f}% (Meta: 50%)")

            plano_vet = st.radio("Linha Vet está no Planograma?", ["Sim", "Não"], key="plano_vet")
            fluxo_vet = st.radio("Royal Canin está abrindo o Fluxo (Vet)?", ["Sim", "Não"], key="fluxo_vet")

            st.markdown("---")
            # --- MÓDULO 4: MERCHANDISING E MATERIAIS ---
            st.subheader("4. Merchandising e Presença de Materiais")
            materiais = [
                "Faixa de Gôndola", "Bobina Forração", "Display Carona", 
                "Cartazete precificador", "Base de Sacarias (can base)", 
                "Totem Silhueta", "Cubo", "Clip Strip", "Stopper", "Outros materiais"
            ]
            
            mat_presenca = {}
            for mat in materiais:
                mat_presenca[mat] = st.checkbox(mat, key=f"mat_{mat}")

            conservacao = st.radio("Os materiais estão bem executados e em bom estado de conservação?", ["Sim", "Não"], key="conservacao")

            st.markdown("---")
            # --- MÓDULO 5: PONTOS EXTRAS ---
            st.subheader("5. Pontos Extras Presentes")
            st.caption("Display Alimento Seco, Display Alimento Úmido, Ponta de Gôndola, Ilha, Vitrine")
            qtd_pontos_extras = st.number_input("Quantidade de Pontos Extras encontrados:", min_value=0, max_value=10, value=0)

            st.markdown("---")
            # --- CÁLCULO DA PONTUAÇÃO TOTAL ---
            if st.button("Finalizar e Calcular Auditoria", type="primary"):
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

                st.success("Auditoria Finalizada com Sucesso!")
                st.metric(label="Nota Total do PDV", value=f"{nota_total:.2f} / 10.0 pts")

                with st.expander("Ver Detalhamento da Pontuação"):
                    for d in detalhes:
                        st.write(f"- {d}")
        else:
            st.warning("Nenhuma loja encontrada para esta promotora.")
