"""
APPLICATION STREAMLIT - PLANNING PRODUCTION POMMES DE TERRE
Version complète prête pour déploiement
"""

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import gspread
from google.oauth2.service_account import Credentials
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO

# Configuration page
st.set_page_config(
    page_title="Planning Production - Culture Pom",
    page_icon="🥔",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS personnalisé
st.markdown("""
<style>
    /* Header principal */
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #6B7F3B;
        text-align: center;
        padding: 1rem;
        background: linear-gradient(90deg, #8FA94B 0%, #6B7F3B 100%);
        border-radius: 10px;
        margin-bottom: 2rem;
        color: white;
    }
    
    /* Métriques - même hauteur */
    .stMetric {
        background-color: #f0f2f6;
        padding: 1.5rem 1rem;
        border-radius: 10px;
        border-left: 5px solid #6B7F3B;
        min-height: 120px;
        display: flex;
        flex-direction: column;
        justify-content: center;
    }
    
    /* Sidebar */
    [data-testid="stSidebar"] {
        background-color: #F5F5F0;
    }
    
    /* Boutons */
    .stButton>button {
        background-color: #6B7F3B;
        color: white;
        border-radius: 8px;
        border: none;
        padding: 0.5rem 1rem;
        font-weight: 500;
    }
    
    .stButton>button:hover {
        background-color: #8FA94B;
    }
    
    /* Logo et crédits */
    .footer-credits {
        position: fixed;
        bottom: 10px;
        right: 10px;
        background-color: rgba(255, 255, 255, 0.9);
        padding: 8px 15px;
        border-radius: 8px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        font-size: 0.75rem;
        text-align: right;
        z-index: 999;
    }
    
    .footer-credits img {
        max-height: 30px;
        vertical-align: middle;
        margin-left: 5px;
    }
</style>
""", unsafe_allow_html=True)

# Logo et titre dans la sidebar
with st.sidebar:
    st.image("https://www.culturepom.com/wp-content/uploads/2023/01/LogoCulturePom.webp", 
             use_container_width=True)
    st.markdown("---")

# Footer avec crédits
st.markdown("""
<div class="footer-credits">
    <div style="margin-bottom: 5px;">
        🇫🇷 Réalisé en France avec passion
    </div>
    <div style="font-weight: 600; color: #6B7F3B;">
        3Force Consulting × Culture Pom
    </div>
    <div style="font-size: 0.65rem; color: #888; margin-top: 3px;">
        © 2025 Tous droits réservés
    </div>
</div>
""", unsafe_allow_html=True)

# =============================================================================
# CONNEXION GOOGLE SHEETS
# =============================================================================

@st.cache_resource
def connect_to_sheets():
    """Connexion à Google Sheets"""
    try:
        import os
        import json
        
        # Heroku : variables d'environnement
        if 'GCP_SERVICE_ACCOUNT' in os.environ:
            service_account_info = json.loads(os.environ['GCP_SERVICE_ACCOUNT'])
            creds = Credentials.from_service_account_info(
                service_account_info,
                scopes=[
                    "https://www.googleapis.com/auth/spreadsheets",
                    "https://www.googleapis.com/auth/drive"
                ]
            )
        # Streamlit Cloud : secrets
        elif 'gcp_service_account' in st.secrets:
            creds = Credentials.from_service_account_info(
                st.secrets["gcp_service_account"],
                scopes=[
                    "https://www.googleapis.com/auth/spreadsheets",
                    "https://www.googleapis.com/auth/drive"
                ]
            )
        else:
            # OAuth local
            from google.auth import default
            creds, _ = default()
        
        gc = gspread.authorize(creds)
        return gc
    except Exception as e:
        st.error(f"Erreur connexion : {e}")
        return None

@st.cache_data(ttl=30)
def charger_donnees(_gc, sheet_url):
    """Charge les données depuis Google Sheets"""
    try:
        spreadsheet = _gc.open_by_url(sheet_url)
        
        data = {}
        onglets = [
            'REF_Variétés', 'REF_Lignes', 'Produits', 'Lots', 'Lots_Lavés',
            'Previsions', 'Affectations', 'Planning_Lavage',
            'Planning_Production', 'Alerte_Stocks', 'Parametres'
        ]
        
        for onglet in onglets:
            try:
                worksheet = spreadsheet.worksheet(onglet)
                records = worksheet.get_all_records()
                data[onglet] = pd.DataFrame(records)
            except:
                data[onglet] = pd.DataFrame()
        
        return data, spreadsheet
    except Exception as e:
        st.error(f"Erreur chargement : {e}")
        return None, None

# =============================================================================
# FONCTIONS MÉTIER
# =============================================================================

def calculer_extrapolation(data):
    """Calcule S4-S5"""
    previsions = data['Previsions'].copy()
    prev_saisies = previsions[previsions['Type_Prévision'] == 'SAISIE']
    
    if len(prev_saisies) < 3:
        return pd.DataFrame()
    
    moyennes = prev_saisies.groupby('Code_Produit')['Volume_Prévu_T'].mean().reset_index()
    semaine_max = int(prev_saisies['Semaine_Num'].max())
    
    nouvelles_prev = []
    for _, row in moyennes.iterrows():
        for offset, sem in [(7, semaine_max+1), (14, semaine_max+2)]:
            nouvelles_prev.append({
                'Semaine_Num': sem,
                'Code_Produit': row['Code_Produit'],
                'Volume_Prévu_T': round(row['Volume_Prévu_T'], 2),
                'Type_Prévision': 'EXTRAPOLÉE'
            })
    
    return pd.DataFrame(nouvelles_prev)

def generer_planning_production(data):
    """Génère le planning production"""
    previsions = data['Previsions'].copy()
    produits = data['Produits'].copy()
    lignes = data['REF_Lignes'][data['REF_Lignes']['Type'] == 'Production'].copy()
    
    planning = []
    of_id = 1
    
    for semaine in sorted(previsions['Semaine_Num'].unique()):
        prev_sem = previsions[previsions['Semaine_Num'] == semaine]
        
        for _, prev in prev_sem.iterrows():
            produit = produits[produits['Code_Produit'] == prev['Code_Produit']]
            if len(produit) == 0:
                continue
            
            ligne_aff = produit['Ligne_Affectée'].iloc[0]
            if pd.isna(ligne_aff) or ligne_aff == '':
                continue
            
            ligne = lignes[lignes['Code_Ligne'] == ligne_aff]
            if len(ligne) == 0:
                continue
            
            cadence = ligne['Capacité_T_h'].iloc[0]
            nb_equipes = ligne['Nb_Équipes'].iloc[0]
            
            volume_jour = prev['Volume_Prévu_T'] / 5
            volume_equipe = volume_jour / nb_equipes
            
            for jour_idx in range(5):
                for equipe in range(1, nb_equipes + 1):
                    planning.append({
                        'OF_ID': f'OF_{of_id:03d}',
                        'Semaine': int(semaine),
                        'Jour': jour_idx + 1,
                        'Équipe': f'Équipe_{equipe}' if nb_equipes > 1 else 'Unique',
                        'Ligne': ligne_aff,
                        'Produit': prev['Code_Produit'],
                        'Tonnage': round(volume_equipe, 2)
                    })
                    of_id += 1
    
    return pd.DataFrame(planning)

# =============================================================================
# SIDEBAR NAVIGATION
# =============================================================================

def sidebar_navigation():
    st.sidebar.markdown("## 🥔 PLANNING PRODUCTION")
    st.sidebar.markdown("---")
    
    menu = st.sidebar.radio(
        "Navigation",
        [
            "🏠 Accueil",
            "📊 Données",
            "📈 Prévisions",
            "🎯 Affectations",
            "🧼 Planning Lavage",
            "🏭 Planning Production",
            "⚠️ Alertes Stocks",
            "💾 Export"
        ]
    )
    
    st.sidebar.markdown("---")
    
    return menu

# =============================================================================
# PAGE : ACCUEIL
# =============================================================================

def page_accueil(data):
    st.markdown('<div class="main-header">🥔 PLANNING PRODUCTION - TABLEAU DE BORD</div>', unsafe_allow_html=True)
    
    # KPIs
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        nb_lots = len(data['Lots'][data['Lots']['Statut'] == 'Stock_Brut'])
        tonnage = data['Lots']['Tonnage_Brut_Restant'].sum()
        st.metric("🥔 Lots en stock", nb_lots, f"{tonnage:.0f}T")
    
    with col2:
        nb_produits = len(data['Produits'][data['Produits']['Actif'] == 'OUI'])
        st.metric("📦 Produits actifs", nb_produits)
    
    with col3:
        if len(data['Previsions']) > 0:
            volume = data['Previsions']['Volume_Prévu_T'].sum()
            st.metric("📈 Prévisions", f"{volume:.0f}T")
        else:
            st.metric("📈 Prévisions", "0T")
    
    with col4:
        nb_aff = len(data['Affectations'][data['Affectations']['Statut_Affectation'] == 'Active'])
        st.metric("🎯 Affectations", nb_aff)
    
    st.markdown("---")
    
    # Graphiques
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### 📊 Stocks par variété")
        if len(data['Lots']) > 0:
            stocks = data['Lots'].groupby('Code_Variété')['Tonnage_Brut_Restant'].sum().reset_index()
            fig = px.bar(stocks, x='Code_Variété', y='Tonnage_Brut_Restant',
                        title='Tonnage disponible')
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Aucun lot")
    
    with col2:
        st.markdown("### 📈 Prévisions par semaine")
        if len(data['Previsions']) > 0:
            prev_sem = data['Previsions'].groupby('Semaine_Num')['Volume_Prévu_T'].sum().reset_index()
            fig = px.line(prev_sem, x='Semaine_Num', y='Volume_Prévu_T',
                         markers=True, title='Évolution des volumes')
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Aucune prévision")
    
    # Alertes
    st.markdown("### ⚠️ Alertes")
    alertes = data.get('Alerte_Stocks', pd.DataFrame())
    
    if len(alertes) > 0:
        critiques = alertes[alertes['Statut'].str.contains('MANQUE', na=False)]
        if len(critiques) > 0:
            st.error(f"❌ {len(critiques)} variété(s) en manque")
            st.dataframe(critiques[['Code_Variété', 'Écart_T', 'Action_Recommandée']], 
                        use_container_width=True)
        else:
            st.success("✅ Tous les stocks sont OK")
    else:
        st.info("Aucune alerte générée")

# =============================================================================
# PAGE : DONNÉES
# =============================================================================

def page_donnees(data):
    st.markdown('<div class="main-header">📊 DONNÉES DE BASE</div>', unsafe_allow_html=True)
    
    tab1, tab2, tab3, tab4 = st.tabs(["🌱 Variétés", "🏭 Lignes", "📦 Produits", "🥔 Lots"])
    
    with tab1:
        st.markdown("### Variétés")
        if len(data['REF_Variétés']) > 0:
            st.dataframe(data['REF_Variétés'], use_container_width=True)
        else:
            st.warning("Aucune variété")
    
    with tab2:
        st.markdown("### Lignes")
        if len(data['REF_Lignes']) > 0:
            st.dataframe(data['REF_Lignes'], use_container_width=True)
        else:
            st.warning("Aucune ligne")
    
    with tab3:
        st.markdown("### Produits")
        if len(data['Produits']) > 0:
            st.dataframe(data['Produits'], use_container_width=True)
        else:
            st.warning("Aucun produit")
    
    with tab4:
        st.markdown("### Lots")
        if len(data['Lots']) > 0:
            # Filtres
            col1, col2, col3 = st.columns(3)
            
            with col1:
                varietes = ['Toutes'] + list(data['Lots']['Code_Variété'].unique())
                var_select = st.selectbox("Variété", varietes)
            
            with col2:
                types = ['Tous'] + list(data['Lots']['Type_Lot'].unique())
                type_select = st.selectbox("Type", types)
            
            with col3:
                statuts = ['Tous'] + list(data['Lots']['Statut'].unique())
                statut_select = st.selectbox("Statut", statuts)
            
            lots = data['Lots'].copy()
            
            if var_select != 'Toutes':
                lots = lots[lots['Code_Variété'] == var_select]
            if type_select != 'Tous':
                lots = lots[lots['Type_Lot'] == type_select]
            if statut_select != 'Tous':
                lots = lots[lots['Statut'] == statut_select]
            
            st.dataframe(lots, use_container_width=True)
            
            # Stats
            col1, col2 = st.columns(2)
            with col1:
                st.metric("Lots affichés", len(lots))
            with col2:
                st.metric("Tonnage total", f"{lots['Tonnage_Brut_Restant'].sum():.1f}T")
        else:
            st.warning("Aucun lot")

# =============================================================================
# PAGE : PRÉVISIONS
# =============================================================================

def page_previsions(data, spreadsheet):
    st.markdown('<div class="main-header">📈 PRÉVISIONS & EXTRAPOLATION</div>', unsafe_allow_html=True)
    
    tab1, tab2 = st.tabs(["📊 Prévisions actuelles", "🔮 Extrapoler S4-S5"])
    
    with tab1:
        if len(data['Previsions']) > 0:
            st.dataframe(data['Previsions'], use_container_width=True)
            
            # Graphique
            fig = px.bar(data['Previsions'], x='Semaine_Num', y='Volume_Prévu_T',
                        color='Code_Produit', title='Prévisions par produit')
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.warning("Aucune prévision")
    
    with tab2:
        st.info("📌 Calcule la moyenne S1-S3 pour générer S4-S5")
        
        if st.button("🔮 Calculer extrapolation", type="primary"):
            with st.spinner("Calcul..."):
                df_extrap = calculer_extrapolation(data)
                
                if len(df_extrap) > 0:
                    st.success("✅ Extrapolations calculées")
                    st.dataframe(df_extrap, use_container_width=True)
                    
                    if st.button("✅ Écrire dans Google Sheets"):
                        try:
                            worksheet = spreadsheet.worksheet('Previsions')
                            
                            # Ajouter les extrapolations
                            for _, row in df_extrap.iterrows():
                                nouvelle_ligne = [
                                    int(row['Semaine_Num']),
                                    '', # Date_Debut à compléter
                                    row['Code_Produit'],
                                    float(row['Volume_Prévu_T']),
                                    'EXTRAPOLÉE',
                                    'Prévisionnel',
                                    '', '', '', ''
                                ]
                                worksheet.append_row(nouvelle_ligne)
                            
                            st.success("✅ Extrapolations écrites dans Google Sheets")
                            st.cache_data.clear()
                            st.rerun()
                        except Exception as e:
                            st.error(f"❌ Erreur : {e}")
                else:
                    st.error("Moins de 3 semaines de prévisions")

# =============================================================================
# PAGE : AFFECTATIONS
# =============================================================================

def page_affectations(data, spreadsheet):
    st.markdown('<div class="main-header">🎯 AFFECTATIONS LOTS → PRODUITS</div>', unsafe_allow_html=True)
    
    tab1, tab2 = st.tabs(["➕ Créer", "📋 Voir"])
    
    with tab1:
        st.markdown("### Nouvelle affectation")
        
        col1, col2 = st.columns(2)
        
        with col1:
            produits = data['Produits'][data['Produits']['Actif'] == 'OUI']
            produit = st.selectbox("Produit", produits['Code_Produit'].tolist())
            
            semaine_debut = st.number_input("Semaine début", 1, 53, 47)
            
            epuisement = st.checkbox("Jusqu'à épuisement")
            if epuisement:
                semaine_fin = "*"
            else:
                semaine_fin = st.number_input("Semaine fin", int(semaine_debut), 53, 50)
        
        with col2:
            if produit:
                var_req = produits[produits['Code_Produit'] == produit]['Code_Variété'].iloc[0]
                st.info(f"🌱 Variété requise: {var_req}")
                
                lots_comp = data['Lots'][
                    (data['Lots']['Code_Variété'] == var_req) &
                    (data['Lots']['Tonnage_Brut_Restant'] > 0)
                ]
                
                if len(lots_comp) > 0:
                    lot = st.selectbox("Lot", lots_comp['Lot_ID'].tolist())
                    
                    if st.button("✅ Créer l'affectation", type="primary"):
                        try:
                            # Calculer les données
                            lot_data = data['Lots'][data['Lots']['Lot_ID'] == lot].iloc[0]
                            previsions = data['Previsions'].copy()
                            
                            if semaine_fin == "*":
                                prev_periode = previsions[
                                    (previsions['Code_Produit'] == produit) &
                                    (previsions['Semaine_Num'] >= semaine_debut)
                                ]
                                semaine_fin_texte = "Épuisement"
                            else:
                                prev_periode = previsions[
                                    (previsions['Code_Produit'] == produit) &
                                    (previsions['Semaine_Num'] >= semaine_debut) &
                                    (previsions['Semaine_Num'] <= int(semaine_fin))
                                ]
                                semaine_fin_texte = str(int(semaine_fin))
                            
                            tonnage_net = prev_periode['Volume_Prévu_T'].sum() if len(prev_periode) > 0 else 0
                            taux_dechet = lot_data['Taux_Déchet_Estimé']
                            
                            if taux_dechet > 1:
                                taux_dechet = taux_dechet / 100
                            
                            tonnage_brut = tonnage_net / (1 - taux_dechet) if tonnage_net > 0 else 0
                            tonnage_dispo = lot_data['Tonnage_Brut_Restant']
                            ecart = tonnage_dispo - tonnage_brut
                            
                            # Générer ID
                            affectations = data['Affectations']
                            if len(affectations) == 0:
                                nouvel_id = 'AFF_001'
                            else:
                                dernier_id = affectations['ID_Affectation'].max()
                                if pd.isna(dernier_id) or dernier_id == '':
                                    nouvel_id = 'AFF_001'
                                else:
                                    numero = int(dernier_id.split('_')[1]) + 1
                                    nouvel_id = f'AFF_{numero:03d}'
                            
                            # Écrire dans Google Sheets
                            worksheet = spreadsheet.worksheet('Affectations')
                            nouvelle_ligne = [
                                nouvel_id,
                                datetime.now().strftime('%Y-%m-%d %H:%M'),
                                produit,
                                int(semaine_debut),
                                semaine_fin_texte,
                                lot,
                                float(tonnage_dispo),
                                float(tonnage_brut),
                                float(ecart),
                                'Active',
                                'Streamlit',
                                ''
                            ]
                            worksheet.append_row(nouvelle_ligne)
                            
                            st.success(f"✅ Affectation {nouvel_id} créée dans Google Sheets")
                            st.cache_data.clear()
                            st.rerun()
                            
                        except Exception as e:
                            st.error(f"❌ Erreur : {e}")
                else:
                    st.warning("Aucun lot compatible")
    
    with tab2:
        if len(data['Affectations']) > 0:
            st.dataframe(data['Affectations'], use_container_width=True)
        else:
            st.info("Aucune affectation")

# =============================================================================
# PAGE : PLANNING PRODUCTION
# =============================================================================

def page_planning_production(data):
    st.markdown('<div class="main-header">🏭 PLANNING PRODUCTION</div>', unsafe_allow_html=True)
    
    if len(data['Planning_Production']) > 0:
        planning = data['Planning_Production'].copy()
        
        # Filtres
        col1, col2 = st.columns(2)
        
        with col1:
            semaines = ['Toutes'] + sorted(planning['Semaine_Num'].unique().tolist())
            sem_select = st.selectbox("Semaine", semaines)
        
        with col2:
            lignes = ['Toutes'] + list(planning['Ligne_Prod'].unique())
            ligne_select = st.selectbox("Ligne", lignes)
        
        # Filtrer
        if sem_select != 'Toutes':
            planning = planning[planning['Semaine_Num'] == sem_select]
        if ligne_select != 'Toutes':
            planning = planning[planning['Ligne_Prod'] == ligne_select]
        
        st.dataframe(planning, use_container_width=True)
        
        # Stats
        st.markdown("### Statistiques")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric("OF total", len(planning))
        with col2:
            st.metric("Tonnage total", f"{planning['Tonnage_Planifié'].sum():.0f}T")
        with col3:
            st.metric("Lignes utilisées", planning['Ligne_Prod'].nunique())
        
        # Graphique
        stats_ligne = planning.groupby('Ligne_Prod')['Tonnage_Planifié'].sum().reset_index()
        fig = px.bar(stats_ligne, x='Ligne_Prod', y='Tonnage_Planifié',
                    title='Charge par ligne')
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Aucun planning généré")
        st.info("💡 Créez des affectations et exécutez le workflow Colab")

# =============================================================================
# PAGE : PLANNING LAVAGE
# =============================================================================

def page_planning_lavage(data):
    st.markdown('<div class="main-header">🧼 PLANNING LAVAGE</div>', unsafe_allow_html=True)
    
    if len(data['Planning_Lavage']) > 0:
        planning = data['Planning_Lavage'].copy()
        
        # Filtres
        col1, col2 = st.columns(2)
        
        with col1:
            semaines = ['Toutes'] + sorted(planning['Semaine_Num'].unique().tolist())
            sem_select = st.selectbox("Semaine", semaines)
        
        with col2:
            lignes = ['Toutes'] + list(planning['Ligne_Lavage'].unique())
            ligne_select = st.selectbox("Ligne", lignes)
        
        # Filtrer
        if sem_select != 'Toutes':
            planning = planning[planning['Semaine_Num'] == sem_select]
        if ligne_select != 'Toutes':
            planning = planning[planning['Ligne_Lavage'] == ligne_select]
        
        st.dataframe(planning, use_container_width=True)
        
        # Stats
        st.markdown("### Statistiques")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric("Opérations", len(planning))
        with col2:
            st.metric("Tonnage brut", f"{planning['Tonnage_Brut'].sum():.0f}T")
        with col3:
            st.metric("Lignes utilisées", planning['Ligne_Lavage'].nunique())
        
        # Graphique
        stats_ligne = planning.groupby('Ligne_Lavage')['Tonnage_Brut'].sum().reset_index()
        fig = px.bar(stats_ligne, x='Ligne_Lavage', y='Tonnage_Brut',
                    title='Tonnage par ligne de lavage')
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Aucun planning lavage généré")
        st.info("💡 Créez des affectations et exécutez le workflow Colab")

# =============================================================================
# PAGE : ALERTES STOCKS
# =============================================================================

def page_alertes_stocks(data):
    st.markdown('<div class="main-header">⚠️ ALERTES STOCKS</div>', unsafe_allow_html=True)
    
    if len(data['Alerte_Stocks']) > 0:
        alertes = data['Alerte_Stocks'].copy()
        
        # Compter par statut
        nb_manque = len(alertes[alertes['Statut'].str.contains('MANQUE', na=False)])
        nb_limite = len(alertes[alertes['Statut'].str.contains('LIMITE', na=False)])
        nb_ok = len(alertes[alertes['Statut'].str.contains('OK', na=False)])
        
        # KPIs
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("❌ Manques", nb_manque)
        with col2:
            st.metric("⚠️ Limites", nb_limite)
        with col3:
            st.metric("✅ OK", nb_ok)
        
        st.markdown("---")
        
        # Filtres
        filtre_statut = st.multiselect(
            "Filtrer par statut",
            options=['❌ MANQUE', '⚠️ LIMITE', '✅ OK'],
            default=['❌ MANQUE', '⚠️ LIMITE']
        )
        
        if filtre_statut:
            alertes_filtrees = alertes[alertes['Statut'].isin(filtre_statut)]
        else:
            alertes_filtrees = alertes
        
        st.dataframe(alertes_filtrees, use_container_width=True)
        
        # Graphique
        fig = px.bar(alertes_filtrees, x='Code_Variété', y='Écart_T',
                    color='Statut', title='Écarts de stock par variété')
        st.plotly_chart(fig, use_container_width=True)
        
    else:
        st.info("Aucune alerte générée")
        st.info("💡 Exécutez le workflow Colab pour générer les alertes")

# =============================================================================
# PAGE : EXPORT
# =============================================================================

def page_export(data):
    st.markdown('<div class="main-header">💾 EXPORT DONNÉES</div>', unsafe_allow_html=True)
    
    st.markdown("### Télécharger les données")
    
    # Export Excel
    if st.button("📥 Télécharger Excel complet", type="primary"):
        output = BytesIO()
        
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            for nom, df in data.items():
                if len(df) > 0 and nom != 'Calendrier':
                    df.to_excel(writer, sheet_name=nom, index=False)
        
        output.seek(0)
        
        st.download_button(
            label="📥 Télécharger",
            data=output,
            file_name=f'Planning_Export_{datetime.now().strftime("%Y%m%d")}.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
        st.success("✅ Fichier prêt")

# =============================================================================
# MAIN
# =============================================================================

def main():
    menu = sidebar_navigation()
    
    # URL Google Sheets
    sheet_url = st.sidebar.text_input(
        "URL Google Sheets",
        value="https://docs.google.com/spreadsheets/d/1OEwROl08gdVLBiTpnEhs-IZ7ZenDtW1ODNzOL4QkwBk/edit?usp=sharing"
    )
    
    if st.sidebar.button("🔄 Recharger"):
        st.cache_data.clear()
        st.rerun()
    
    # Connexion
    gc = connect_to_sheets()
    
    if gc is None:
        st.error("Impossible de se connecter")
        return
    
    data, spreadsheet = charger_donnees(gc, sheet_url)
    
    if data is None:
        st.error("Impossible de charger les données")
        st.info("Vérifiez l'URL et le partage")
        return
    
    # Router
    if menu == "🏠 Accueil":
        page_accueil(data)
    elif menu == "📊 Données":
        page_donnees(data)
    elif menu == "📈 Prévisions":
        page_previsions(data, spreadsheet)
    elif menu == "🎯 Affectations":
        page_affectations(data, spreadsheet)
    elif menu == "🧼 Planning Lavage":
        page_planning_lavage(data)
    elif menu == "🏭 Planning Production":
        page_planning_production(data)
    elif menu == "⚠️ Alertes Stocks":
        page_alertes_stocks(data)
    elif menu == "💾 Export":
        page_export(data)
    else:
        st.info("🚧 Page en développement")

if __name__ == "__main__":
    main()
