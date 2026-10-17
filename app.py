import streamlit as st
import streamlit_authenticator as stauth

st.set_page_config(
    page_title="NAP SARL - Karité",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==================== AUTHENTIFICATION ====================
config = {
    'credentials': {
        'usernames': {
            'admin': {
                'email': 'alliancebodjrenou8@gmail.com',
                'name': 'Admin',
                'password': '69060724@'
            },
            'Manager': {
                'email': 'logistics.napbj@gmail.com',
                'name': 'Deepak',
                'password': 'Deepak2026'
            },
            'Hospice': {
                'email': 'logistics.napbj@gmail.com',
                'name': 'Hospice',
                'password': 'Hospice2026'
            },
            'Alliance': {
                'email': 'alliance@napbj.com',
                'name': 'Alliance',
                'password': 'Nap@2026'
            }
        }
    },
    'cookie': {
        'expiry_days': 30,
        'key': 'nap_secret_key_2026_changez_moi',
        'name': 'nap_auth_cookie'
    },
    'preauthorized': {
        'emails': []
    }
}

authenticator = stauth.Authenticate(
    config['credentials'],
    config['cookie']['name'],
    config['cookie']['key'],
    config['cookie']['expiry_days']
)

authenticator.login(location='main')

if st.session_state.get("authentication_status") is False:
    st.error("Nom d'utilisateur ou mot de passe incorrect")
    st.stop()
elif st.session_state.get("authentication_status") is None:
    st.warning("Veuillez entrer vos identifiants pour accéder à l'application")
    st.stop()

# ==================== UTILISATEUR CONNECTÉ ====================
name = st.session_state.get("name")
username = st.session_state.get("username")

with st.sidebar:
    st.success(f"Connecté : **{name}**")
    st.caption(f"Compte : `{username}`")
    authenticator.logout(location='sidebar')
    st.divider()
    st.markdown("### Navigation")
    st.markdown("Utilisez le menu ci-dessus pour accéder aux pages.")

# ==================== PAGE D'ACCUEIL ====================
st.title("📄 NAP SARL – Gestion Achats (Soja)")
st.markdown("---")

st.markdown("""
### Bienvenue dans l'application de gestion des bons de commande Soja (graines).

**Pages disponibles :**
- **1. Tableau de Bord** → Vue d'ensemble (KPIs, tracker, soldes)
- **2. Bons de Commande** → Créer, générer et télécharger les bons

Utilisez le menu de gauche (ou le sélecteur de pages en haut) pour naviguer.
""")

st.info("Référence auto selon le produit : **SHEA** / **SOYABEAN** / **COTTON** → `XXXX-BC-…/NAP-2026`")