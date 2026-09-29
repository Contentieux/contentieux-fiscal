import sys
import io
import streamlit as st
import pdfplumber
import pandas as pd
import json
from groq import Groq
from docx import Document
from docx.shared import Pt

# Forcer l'encodage UTF-8
if hasattr(sys.stdout, 'reconfigure'): sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stderr, 'reconfigure'): sys.stderr.reconfigure(encoding='utf-8')

# --- FONCTIONS DE LECTURE ---
def extraire_texte(fichier):
    if fichier is None: return ""
    ext = fichier.name.split('.')[-1].lower()
    texte = f"\n--- DOCUMENT : {fichier.name} ---\n"
    try:
        if ext == 'pdf':
            with pdfplumber.open(fichier) as pdf:
                for page in pdf.pages:
                    txt = page.extract_text()
                    if txt: texte += txt + "\n"
        elif ext in ['xlsx', 'xls']:
            df = pd.read_excel(fichier)
            texte += df.to_string()
        elif ext in ['docx', 'doc']:
            doc = Document(fichier)
            for para in doc.paragraphs:
                texte += para.text + "\n"
        elif ext in ['txt']:
            texte += str(fichier.read(), 'utf-8')
    except Exception as e:
        return f"Erreur de lecture : {e}"
    return texte

# --- GÉNÉRATION WORD ---
def generer_word(donnees_json, type_reponse):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Arial'
    style.font.size = Pt(11)

    doc.add_heading('CABINET D\'EXPERTISE COMPTABLE', 0)
    doc.add_paragraph(f"Objet : {type_reponse}")
    doc.add_paragraph("A l'attention de l'Administration Fiscale\n")
    
    doc.add_heading("1. Synthèse du Dossier", level=1)
    doc.add_paragraph(donnees_json.get('synthese', 'Non spécifié.'))
    
    doc.add_heading("2. Argumentaire de Défense", level=1)
    doc.add_paragraph(donnees_json.get('argumentaire', 'Aucun argument généré.'))
    
    doc.add_heading("3. Conclusion & Demande", level=1)
    doc.add_paragraph(donnees_json.get('conclusion', 'Non spécifié.'))
    
    doc.add_paragraph("\nSignature de l'Expert :")
    
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- INTERFACE UTILISATEUR ---
st.set_page_config(page_title="Audit Fiscal Rapide", layout="wide", page_icon="⚡")
st.title("⚡ Système de Contentieux Fiscal (Mode Cloud Rapide)")

# Verrou de sécurité
st.warning("⚠️ **Règle de confidentialité stricte :** Ce mode utilise des serveurs externes pour garantir une vitesse maximale. Vous devez impérativement masquer le NIF, le nom de l'entreprise et l'adresse du client sur vos PDF avant de les importer.")



# Onglets
tab_admin, tab_compta, tab_audit = st.tabs([
    "🏛️ Pièces de l'Administration", 
    "💼 Éléments Comptables", 
    "⚙️ Lancement de l'Audit"
])

texte_admin = ""
texte_compta = ""

with tab_admin:
    f_c4 = st.file_uploader("Notification de redressement (C4)", type=['pdf'])
    f_role = st.file_uploader("Extrait de rôle", type=['pdf'])
    for f in [f_c4, f_role]:
        if f: texte_admin += extraire_texte(f)

with tab_compta:
    f_bilan = st.file_uploader("Bilan / Balance", type=['pdf', 'xlsx'])
    for f in [f_bilan]:
        if f: texte_compta += extraire_texte(f)

with tab_audit:
    type_reponse = st.radio("Type de réponse :", ["Réponse standard (Contestation)", "Recours gracieux"])
    
    confirmation_anonymisation = st.checkbox("✅ **Je confirme sur l'honneur avoir masqué les données d'identification de mon client sur les documents importés.**")

    if st.button("🚀 Lancer l'Audit Croisé Rapide", type="primary"):
        
        if not confirmation_anonymisation:
            st.error("🛑 Le traitement est bloqué. Vous devez cocher la case confirmant l'anonymisation des données.")
            st.stop()
            
        if not texte_admin.strip():
            st.error("⚠️ Fournissez au moins une notification C4.")
            st.stop()

        with st.spinner("⚡ Analyse ultra-rapide en cours via Groq Cloud..."):
            prompt = f"""Tu es un expert fiscaliste algérien. Rédige un {type_reponse} très professionnel.
            
            PIÈCES DE L'ADMINISTRATION :
            {texte_admin}
            
            PIÈCES DU CABINET :
            {texte_compta}
            
            Renvoie UNIQUEMENT un objet JSON valide avec ces clés : "synthese", "argumentaire", "conclusion".
            """

            try:
                # Votre clé est intégrée proprement ici
                client = Groq(api_key="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")
                
                # Utilisation du modèle Llama 3
                chat_completion = client.chat.completions.create(
                    messages=[{"role": "user", "content": prompt}],
                    model="llama3-70b-8192", 
                    response_format={"type": "json_object"}
                )
                
                resultats = json.loads(chat_completion.choices[0].message.content)
                
                st.success("✅ Audit terminé en quelques secondes !")
                
                st.download_button(
                    label="📄 TÉLÉCHARGER LE DOSSIER DE RÉPONSE (Word)",
                    data=generer_word(resultats, type_reponse),
                    file_name="Reponse_Fiscale_Rapide.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
            except Exception as e:
                st.error(f"Erreur de communication Cloud : {e}")