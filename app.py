"""
APPLICATION STREAMLIT - PLANNING PRODUCTION POMMES DE TERRE
Version MVP avec OF et OL + Génération PDF
Culture Pom × 3Force Consulting - 2025
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
import json
import os

# Tentative d'import du module PDF
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import cm
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_CENTER
    PDF_AVAILABLE = True
except:
    PDF_AVAILABLE = False

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
    
    div[data-testid="stMetric"] {
        background-color: #f0f2f6 !important;
        padding: 1.5rem 1rem !important;
        border-radius: 10px !important;
        border-left: 5px solid #6B7F3B !important;
        min-height: 140px !important;
        height: 140px !important;
    }
    
    div[data-testid="stMetricValue"] {
        font-size: 2rem !important;
    }
    
    [data-testid="stSidebar"] {
        background-color: #F5F5F0;
    }
    
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
    
    .footer-credits {
        position: fixed;
        bottom: 10px;
        right: 10px;
        background-color: rgba(255, 255, 255, 0.95);
        padding: 10px 15px;
        border-radius: 8px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.15);
        font-size: 0.75rem;
        text-align: right;
        z-index: 999;
        border-left: 3px solid #6B7F3B;
    }
    
    .logo-container {
        text-align: center;
        padding: 1rem 0;
        background: white;
        border-radius: 10px;
        margin-bottom: 1rem;
    }
    
    .stat-box {
        background-color: #F5F5F0;
        padding: 10px;
        border-radius: 5px;
        border-left: 3px solid #6B7F3B;
        margin: 5px 0;
    }
</style>
""", unsafe_allow_html=True)

# Logo sidebar
with st.sidebar:
    st.markdown("""
    <div class="logo-container">
        <div style="font-size: 2rem; font-weight: bold; color: #6B7F3B;">
            🥔 Culture Pom
        </div>
        <div style="font-size: 0.9rem; color: #8FA94B; margin-top: 5px;">
            Planning Production
        </div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("---")

# Footer
st.markdown("""
<div class="footer-credits">
    <div style="margin-bottom: 5px;">🇫🇷 Réalisé en France avec passion</div>
    <div style="font-weight: 600; color: #6B7F3B;">3Force Consulting × Culture Pom</div>
    <div style="font-size: 0.65rem; color: #888; margin-top: 3px;">© 2025 Tous droits réservés</div>
</div>
""", unsafe_allow_html=True)

# =============================================================================
# CONNEXION GOOGLE SHEETS
# =============================================================================

@st.cache_resource
def connect_to_sheets():
    """Connexion à Google Sheets"""
    try:
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
            st.error("Aucune authentification trouvée")
            return None
        
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
# FONCTIONS GÉNÉRATION PDF
# =============================================================================

def generer_pdf_of_simple(liste_of):
    """Génère un PDF simple pour OF"""
    if not PDF_AVAILABLE:
        st.error("Module PDF non disponible")
        return None
    
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, topMargin=2*cm, bottomMargin=2*cm,
                           leftMargin=2*cm, rightMargin=2*cm)
    
    story = []
    styles = getSampleStyleSheet()
    
    titre_style = ParagraphStyle('Titre', parent=styles['Heading1'], 
                                 fontSize=20, textColor=colors.HexColor('#6B7F3B'),
                                 alignment=TA_CENTER, spaceAfter=20)
    
    for idx, of in enumerate(liste_of):
        if idx > 0:
            from reportlab.platypus import PageBreak
            story.append(PageBreak())
        
        story.append(Paragraph("🥔 CULTURE POM", titre_style))
        story.append(Paragraph(f"ORDRE DE FABRICATION N° {of.get('OF_ID', 'N/A')}", titre_style))
        story.append(Spacer(1, 0.5*cm))
        
        info_data = [
            ['Ligne', of.get('Ligne_Prod', 'N/A')],
            ['Produit', of.get('Code_Produit', 'N/A')],
            ['Tonnage', f"{of.get('Tonnage_Planifié', 0):.2f} T"],
            ['Heure', f"{of.get('Heure_Début', '')} - {of.get('Heure_Fin', '')}"],
            ['Équipe', of.get('Équipe', 'N/A')],
        ]
        
        info_table = Table(info_data, colWidths=[5*cm, 12*cm])
        info_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#F5F5F0')),
            ('GRID', (0, 0), (-1, -1), 1, colors.grey),
            ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 11),
            ('LEFTPADDING', (0, 0), (-1, -1), 10),
            ('TOPPADDING', (0, 0), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ]))
        
        story.append(info_table)
        story.append(Spacer(1, 1*cm))
        
        realisation_data = [
            ['Tonnage réalisé', '_______ T'],
            ['Heure début', '___:___'],
            ['Heure fin', '___:___'],
            ['Opérateur', '_______________________'],
            ['Signature', '_______________________'],
        ]
        
        realisation_table = Table(realisation_data, colWidths=[5*cm, 12*cm])
        realisation_table.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),
            ('FONTSIZE', (0, 0), (-1, -1), 11),
            ('LEFTPADDING', (0, 0), (-1, -1), 10),
        ]))
        
        story.append(realisation_table)
        
        story.append(Spacer(1, 1*cm))
        footer_style = ParagraphStyle('Footer', parent=styles['Normal'], 
                                     fontSize=8, textColor=colors.grey, alignment=TA_CENTER)
        story.append(Paragraph("🇫🇷 3Force Consulting × Culture Pom © 2025", footer_style))
    
    doc.build(story)
    buffer.seek(0)
    return buffer

def generer_pdf_ol_simple(liste_ol):
    """Génère un PDF simple pour OL"""
    if not PDF_AVAILABLE:
        st.error("Module PDF non disponible")
        return None
    
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, topMargin=2*cm, bottomMargin=2*cm,
                           leftMargin=2*cm, rightMargin=2*cm)
    
    story = []
    styles = getSampleStyleSheet()
    
    titre_style = ParagraphStyle('Titre', parent=styles['Heading1'], 
                                 fontSize=20, textColor=colors.HexColor('#6B7F3B'),
                                 alignment=TA_CENTER, spaceAfter=20)
    
    for idx, ol in enumerate(liste_ol):
        if idx > 0:
            from reportlab.platypus import PageBreak
            story.append(PageBreak())
        
        story.append(Paragraph("🥔 CULTURE POM", titre_style))
        story.append(Paragraph(f"ORDRE DE LAVAGE N° {ol.get('ID_Lavage', 'N/A')}", titre_style))
        story.append(Spacer(1, 0.5*cm))
        
        info_data = [
            ['Ligne lavage', ol.get('Ligne_Lavage', 'N/A')],
            ['Lot ID', ol.get('Lot_ID', 'N/A')],
            ['Variété', ol.get('Code_Variété', 'N/A')],
            ['Tonnage brut', f"{ol.get('Tonnage_Brut', 0):.2f} T"],
            ['Heure', f"{ol.get('Heure_Début', '')} - {ol.get('Heure_Fin', '')}"],
        ]
        
        info_table = Table(info_data, colWidths=[5*cm, 12*cm])
        info_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#F5F5F0')),
            ('GRID', (0, 0), (-1, -1), 1, colors.grey),
            ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 11),
            ('LEFTPADDING', (0, 0), (-1, -1), 10),
            ('TOPPADDING', (0, 0), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ]))
        
        story.append(info_table)
        story.append(Spacer(1, 1*cm))
        
        resultats_data = [
            ['Tonnage net obtenu', '_______ T'],
            ['Taux déchet réel', '_______ %'],
            ['  • Purs', '_______ %'],
            ['  • Grenailles', '_______ %'],
            ['  • Terre', '_______ %'],
            ['Opérateur', '_______________________'],
            ['Signature', '_______________________'],
        ]
        
        resultats_table = Table(resultats_data, colWidths=[5*cm, 12*cm])
        resultats_table.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 0.5, colors.lightgrey),
            ('FONTSIZE', (0, 0), (-1, -1), 11),
            ('LEFTPADDING', (0, 0), (-1, -1), 10),
        ]))
        
        story.append(resultats_table)
        
        story.append(Spacer(1, 1*cm))
        footer_style = ParagraphStyle('Footer', parent=styles['Normal'], 
                                     fontSize=8, textColor=colors.grey, alignment=TA_CENTER)
        story.append(Paragraph("🇫🇷 3Force Consulting × Culture Pom © 2025", footer_style))
    
    doc.build(story)
    buffer.seek(0)
    return buffer

# =============================================================================
# FONCTIONS UTILITAIRES
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


# =============================================================================
# PAGES
# =============================================================================

def sidebar_navigation():
    menu = st.sidebar.radio(
        "Navigation",
        [
            "🏠 Accueil",
            "📊 Données",
            "📈 Prévisions",
            "🎯 Affectations",
            "📋 Ordres de Fabrication",
            "🧼 Ordres de Lavage",
            "🏭 Planning Production",
            "⚠️ Alertes Stocks",
            "💾 Export"
        ]
    )
    return menu

def page_accueil(data):
    st.markdown('<div class="main-header">🥔 PLANNING PRODUCTION - TABLEAU DE BORD</div>', unsafe_allow_html=True)
    
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
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### 📊 Stocks par variété")
        if len(data['Lots']) > 0:
            stocks = data['Lots'].groupby('Code_Variété')['Tonnage_Brut_Restant'].sum().reset_index()
            fig = px.bar(stocks, x='Code_Variété', y='Tonnage_Brut_Restant', title='Tonnage disponible')
            st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        st.markdown("### 📈 Prévisions par semaine")
        if len(data['Previsions']) > 0:
            prev_sem = data['Previsions'].groupby('Semaine_Num')['Volume_Prévu_T'].sum().reset_index()
            fig = px.line(prev_sem, x='Semaine_Num', y='Volume_Prévu_T', markers=True, title='Évolution')
            st.plotly_chart(fig, use_container_width=True)

def page_donnees(data):
    st.markdown('<div class="main-header">📊 DONNÉES DE BASE</div>', unsafe_allow_html=True)
    
    tab1, tab2, tab3, tab4 = st.tabs(["🌱 Variétés", "🏭 Lignes", "📦 Produits", "🥔 Lots"])
    
    with tab1:
        st.markdown("### Variétés")
        if len(data['REF_Variétés']) > 0:
            st.dataframe(data['REF_Variétés'], use_container_width=True)
    
    with tab2:
        st.markdown("### Lignes")
        if len(data['REF_Lignes']) > 0:
            st.dataframe(data['REF_Lignes'], use_container_width=True)
    
    with tab3:
        st.markdown("### Produits")
        if len(data['Produits']) > 0:
            st.dataframe(data['Produits'], use_container_width=True)
    
    with tab4:
        st.markdown("### Lots")
        if len(data['Lots']) > 0:
            st.dataframe(data['Lots'], use_container_width=True)

def page_previsions(data, spreadsheet):
    st.markdown('<div class="main-header">📈 PRÉVISIONS</div>', unsafe_allow_html=True)
    
    if len(data['Previsions']) > 0:
        st.dataframe(data['Previsions'], use_container_width=True)

def page_affectations(data, spreadsheet):
    st.markdown('<div class="main-header">🎯 AFFECTATIONS</div>', unsafe_allow_html=True)
    
    if len(data['Affectations']) > 0:
        st.dataframe(data['Affectations'], use_container_width=True)
    else:
        st.info("Aucune affectation")

# =============================================================================
# PAGE : ORDRES DE FABRICATION
# =============================================================================

def page_ordres_fabrication(data, spreadsheet):
    st.markdown('<div class="main-header">📋 ORDRES DE FABRICATION</div>', unsafe_allow_html=True)
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        date_selectionnee = st.date_input("📅 Sélectionner une date", value=datetime.now())
    
    with col2:
        if st.button("🔍 Charger les OF", type="primary"):
            st.cache_data.clear()
            st.rerun()
    
    if len(data['Planning_Production']) == 0:
        st.warning("Aucun planning de production généré")
        st.info("💡 Créez des affectations et exécutez le workflow Colab")
        return
    
    planning = data['Planning_Production'].copy()
    planning['Date'] = pd.to_datetime(planning['Date'], errors='coerce')
    
    of_jour = planning[planning['Date'].dt.date == date_selectionnee]
    
    if len(of_jour) == 0:
        st.info(f"Aucun OF pour le {date_selectionnee.strftime('%d/%m/%Y')}")
        return
    
    st.markdown(f"### {len(of_jour)} OF pour le {date_selectionnee.strftime('%A %d %B %Y')}")
    
    col1, col2 = st.columns([3, 1])
    
    with col1:
        selection_totale = st.checkbox("☑ Sélectionner tout", key="select_all_of")
    
    with col2:
        st.markdown(f"**Total : {of_jour['Tonnage_Planifié'].sum():.1f}T**")
    
    of_selectionnes = []
    
    for idx, row in of_jour.iterrows():
        col1, col2, col3, col4, col5, col6, col7 = st.columns([0.5, 2, 1.5, 2, 3, 1.5, 1.5])
        
        with col1:
            selected = st.checkbox("", value=selection_totale, key=f"of_{row['OF_ID']}")
        
        with col2:
            st.text(row['OF_ID'])
        
        with col3:
            st.text(f"{row['Heure_Début']}")
        
        with col4:
            st.text(row['Ligne_Prod'])
        
        with col5:
            st.text(row['Code_Produit'])
        
        with col6:
            st.text(f"{row['Tonnage_Planifié']:.2f}T")
        
        with col7:
            statut_color = {"Planifié": "🔵", "En cours": "🟢", "Terminé": "✅", "Annulé": "❌"}
            statut = row.get('Statut', 'Planifié')
            st.text(f"{statut_color.get(statut, '⚪')} {statut}")
        
        if selected:
            of_selectionnes.append(row.to_dict())
    
    if len(of_selectionnes) > 0:
        st.markdown("---")
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            if st.button("▶️ Passer en cours", use_container_width=True):
                try:
                    worksheet = spreadsheet.worksheet('Planning_Production')
                    all_data = worksheet.get_all_values()
                    headers = all_data[0]
                    
                    statut_idx = headers.index('Statut')
                    of_id_idx = headers.index('OF_ID')
                    
                    for of in of_selectionnes:
                        for row_idx, row_data in enumerate(all_data[1:], start=2):
                            if row_data[of_id_idx] == of['OF_ID']:
                                worksheet.update_cell(row_idx, statut_idx + 1, 'En cours')
                    
                    st.success(f"✅ {len(of_selectionnes)} OF passés en cours")
                    st.cache_data.clear()
                    st.rerun()
                except Exception as e:
                    st.error(f"Erreur : {e}")
        
        with col2:
            if st.button("✅ Marquer terminé", use_container_width=True):
                try:
                    worksheet = spreadsheet.worksheet('Planning_Production')
                    all_data = worksheet.get_all_values()
                    headers = all_data[0]
                    
                    statut_idx = headers.index('Statut')
                    of_id_idx = headers.index('OF_ID')
                    
                    for of in of_selectionnes:
                        for row_idx, row_data in enumerate(all_data[1:], start=2):
                            if row_data[of_id_idx] == of['OF_ID']:
                                worksheet.update_cell(row_idx, statut_idx + 1, 'Terminé')
                    
                    st.success(f"✅ {len(of_selectionnes)} OF terminés")
                    st.cache_data.clear()
                    st.rerun()
                except Exception as e:
                    st.error(f"Erreur : {e}")
        
        with col3:
            if PDF_AVAILABLE and st.button("🖨️ Imprimer PDF", use_container_width=True):
                try:
                    pdf_buffer = generer_pdf_of_simple(of_selectionnes)
                    if pdf_buffer:
                        st.download_button(
                            label="📥 Télécharger PDF",
                            data=pdf_buffer,
                            file_name=f"OF_{date_selectionnee.strftime('%Y%m%d')}.pdf",
                            mime="application/pdf",
                            use_container_width=True
                        )
                except Exception as e:
                    st.error(f"Erreur génération PDF : {e}")
            elif not PDF_AVAILABLE:
                st.warning("Module PDF non disponible")
        
        with col4:
            st.markdown(f"**{len(of_selectionnes)} OF sélectionnés**")
            st.markdown(f"**{sum(of['Tonnage_Planifié'] for of in of_selectionnes):.1f}T**")

# =============================================================================
# PAGE : ORDRES DE LAVAGE
# =============================================================================

def page_ordres_lavage(data, spreadsheet):
    st.markdown('<div class="main-header">🧼 ORDRES DE LAVAGE</div>', unsafe_allow_html=True)
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        date_selectionnee = st.date_input("📅 Sélectionner une date", value=datetime.now(), key="date_ol")
    
    with col2:
        if st.button("🔍 Charger les OL", type="primary"):
            st.cache_data.clear()
            st.rerun()
    
    if len(data['Planning_Lavage']) == 0:
        st.warning("Aucun planning de lavage généré")
        st.info("💡 Créez des affectations et exécutez le workflow Colab")
        return
    
    planning = data['Planning_Lavage'].copy()
    planning['Date'] = pd.to_datetime(planning['Date'], errors='coerce')
    
    ol_jour = planning[planning['Date'].dt.date == date_selectionnee]
    
    if len(ol_jour) == 0:
        st.info(f"Aucun OL pour le {date_selectionnee.strftime('%d/%m/%Y')}")
        return
    
    st.markdown(f"### {len(ol_jour)} OL pour le {date_selectionnee.strftime('%A %d %B %Y')}")
    
    col1, col2 = st.columns([3, 1])
    
    with col1:
        selection_totale = st.checkbox("☑ Sélectionner tout", key="select_all_ol")
    
    with col2:
        st.markdown(f"**Total : {ol_jour['Tonnage_Brut'].sum():.1f}T**")
    
    ol_selectionnes = []
    
    for idx, row in ol_jour.iterrows():
        col1, col2, col3, col4, col5, col6, col7 = st.columns([0.5, 2, 1.5, 2, 3, 1.5, 1.5])
        
        with col1:
            selected = st.checkbox("", value=selection_totale, key=f"ol_{row['ID_Lavage']}")
        
        with col2:
            st.text(row['ID_Lavage'])
        
        with col3:
            st.text(f"{row['Heure_Début']}")
        
        with col4:
            st.text(row['Ligne_Lavage'])
        
        with col5:
            st.text(f"{row['Lot_ID']} ({row['Code_Variété']})")
        
        with col6:
            st.text(f"{row['Tonnage_Brut']:.2f}T")
        
        with col7:
            statut_color = {"Planifié": "🔵", "En cours": "🟢", "Terminé": "✅"}
            statut = row.get('Statut', 'Planifié')
            st.text(f"{statut_color.get(statut, '⚪')} {statut}")
        
        if selected:
            ol_selectionnes.append(row.to_dict())
    
    if len(ol_selectionnes) > 0:
        st.markdown("---")
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            if st.button("▶️ Passer en cours", use_container_width=True, key="ol_encours"):
                try:
                    worksheet = spreadsheet.worksheet('Planning_Lavage')
                    all_data = worksheet.get_all_values()
                    headers = all_data[0]
                    
                    statut_idx = headers.index('Statut')
                    id_idx = headers.index('ID_Lavage')
                    
                    for ol in ol_selectionnes:
                        for row_idx, row_data in enumerate(all_data[1:], start=2):
                            if row_data[id_idx] == ol['ID_Lavage']:
                                worksheet.update_cell(row_idx, statut_idx + 1, 'En cours')
                    
                    st.success(f"✅ {len(ol_selectionnes)} OL passés en cours")
                    st.cache_data.clear()
                    st.rerun()
                except Exception as e:
                    st.error(f"Erreur : {e}")
        
        with col2:
            if st.button("✅ Marquer terminé", use_container_width=True, key="ol_termine"):
                st.info("💡 Utilisez le formulaire de saisie des résultats pour terminer un OL")
        
        with col3:
            if PDF_AVAILABLE and st.button("🖨️ Imprimer PDF", use_container_width=True, key="ol_pdf"):
                try:
                    pdf_buffer = generer_pdf_ol_simple(ol_selectionnes)
                    if pdf_buffer:
                        st.download_button(
                            label="📥 Télécharger PDF",
                            data=pdf_buffer,
                            file_name=f"OL_{date_selectionnee.strftime('%Y%m%d')}.pdf",
                            mime="application/pdf",
                            use_container_width=True
                        )
                except Exception as e:
                    st.error(f"Erreur génération PDF : {e}")
        
        with col4:
            st.markdown(f"**{len(ol_selectionnes)} OL sélectionnés**")
            st.markdown(f"**{sum(ol['Tonnage_Brut'] for ol in ol_selectionnes):.1f}T**")

def page_planning_production(data):
    st.markdown('<div class="main-header">🏭 PLANNING PRODUCTION</div>', unsafe_allow_html=True)
    
    if len(data['Planning_Production']) > 0:
        st.dataframe(data['Planning_Production'], use_container_width=True)
    else:
        st.info("Aucun planning généré")

def page_alertes_stocks(data):
    st.markdown('<div class="main-header">⚠️ ALERTES STOCKS</div>', unsafe_allow_html=True)
    
    if len(data['Alerte_Stocks']) > 0:
        st.dataframe(data['Alerte_Stocks'], use_container_width=True)
    else:
        st.info("Aucune alerte")

def page_export(data):
    st.markdown('<div class="main-header">💾 EXPORT</div>', unsafe_allow_html=True)
    
    if st.button("📥 Télécharger Excel", type="primary"):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            for nom, df in data.items():
                if len(df) > 0:
                    df.to_excel(writer, sheet_name=nom, index=False)
        output.seek(0)
        st.download_button(
            label="📥 Télécharger",
            data=output,
            file_name=f'Planning_{datetime.now().strftime("%Y%m%d")}.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

# =============================================================================
# MAIN
# =============================================================================

def main():
    menu = sidebar_navigation()
    
    sheet_url = st.sidebar.text_input(
        "URL Google Sheets",
        value="https://docs.google.com/spreadsheets/d/1OEwROl08gdVLBiTpnEhs-IZ7ZenDtW1ODNzOL4QkwBk/edit?usp=sharing"
    )
    
    if st.sidebar.button("🔄 Recharger"):
        st.cache_data.clear()
        st.rerun()
    
    gc = connect_to_sheets()
    
    if gc is None:
        st.error("Impossible de se connecter")
        return
    
    data, spreadsheet = charger_donnees(gc, sheet_url)
    
    if data is None:
        st.error("Impossible de charger les données")
        return
    
    if menu == "🏠 Accueil":
        page_accueil(data)
    elif menu == "📊 Données":
        page_donnees(data)
    elif menu == "📈 Prévisions":
        page_previsions(data, spreadsheet)
    elif menu == "🎯 Affectations":
        page_affectations(data, spreadsheet)
    elif menu == "📋 Ordres de Fabrication":
        page_ordres_fabrication(data, spreadsheet)
    elif menu == "🧼 Ordres de Lavage":
        page_ordres_lavage(data, spreadsheet)
    elif menu == "🏭 Planning Production":
        page_planning_production(data)
    elif menu == "⚠️ Alertes Stocks":
        page_alertes_stocks(data)
    elif menu == "💾 Export":
        page_export(data)

if __name__ == "__main__":
    main()
