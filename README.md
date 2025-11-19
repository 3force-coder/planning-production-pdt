# 🥔 Planning Production - Application Streamlit

Application de gestion complète du planning de production pour pommes de terre.

## 🚀 Fonctionnalités

- ✅ Tableau de bord avec KPIs
- ✅ Gestion des données (variétés, lignes, produits, lots)
- ✅ Prévisions & extrapolation automatique S4-S5
- ✅ Affectations lots → produits
- ✅ Planning lavage
- ✅ Planning production
- ✅ Alertes stocks (3 semaines)
- ✅ Export Excel complet

## 📋 Prérequis

- Python 3.8+
- Compte Google avec accès au Google Sheet

## 🔧 Installation locale

```bash
pip install -r requirements.txt
streamlit run app.py
```

## ☁️ Déploiement sur Streamlit Cloud (GRATUIT)

### Étape 1 : Préparer le repository GitHub

1. Créer un repository GitHub (public ou privé)
2. Y placer les fichiers :
   - `app.py`
   - `requirements.txt`
   - `.streamlit/config.toml`

### Étape 2 : Créer un Service Account Google

1. Aller sur https://console.cloud.google.com
2. Créer un projet (ou utiliser un existant)
3. Activer l'API Google Sheets
4. Créer un Service Account :
   - IAM & Admin → Service Accounts → Create Service Account
   - Donner un nom (ex: "streamlit-app")
   - Créer une clé JSON
   - Télécharger le fichier JSON

5. Partager le Google Sheet avec l'email du Service Account
   - Copier l'email du service account (quelquechose@PROJECT_ID.iam.gserviceaccount.com)
   - Dans Google Sheets → Partager → Coller l'email → Droits "Éditeur"

### Étape 3 : Déployer sur Streamlit Cloud

1. Aller sur https://share.streamlit.io
2. Se connecter avec GitHub
3. Cliquer "New app"
4. Sélectionner :
   - Repository
   - Branch (main)
   - Main file path (app.py)

5. Ajouter les secrets :
   - Cliquer sur "Advanced settings"
   - Section "Secrets"
   - Coller le contenu du fichier JSON du service account dans ce format :

```toml
[gcp_service_account]
type = "service_account"
project_id = "votre-project-id"
private_key_id = "votre-key-id"
private_key = "-----BEGIN PRIVATE KEY-----\nVOTRE_CLE_PRIVEE\n-----END PRIVATE KEY-----\n"
client_email = "votre-service-account@project.iam.gserviceaccount.com"
client_id = "votre-client-id"
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "votre-cert-url"
```

6. Cliquer "Deploy"

### Étape 4 : Configuration Google Sheets

Dans l'application déployée :
1. Coller l'URL de votre Google Sheet
2. L'app se connecte automatiquement
3. ✅ C'est prêt !

## 🔒 Sécurité

- Les credentials ne sont jamais exposés
- Connexion sécurisée via Service Account
- Google Sheet accessible uniquement via l'app

## 📱 Accès

Une fois déployée, l'app est accessible via une URL du type :
`https://votre-app-name.streamlit.app`

Vous pouvez :
- La partager publiquement
- La mettre en privé (authentification requise)

## 🆘 Support

Pour toute question, consulter la documentation Streamlit :
https://docs.streamlit.io
