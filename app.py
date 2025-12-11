
import streamlit as st
import subprocess
import pandas as pd
import openpyxl
import os
import io
import importlib
from datetime import datetime
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
import sys

# --- BLOC DE DÉPLOIEMENT SÉCURISÉ ---
def setup_remote_files():
    # 1. Reconstruction des fichiers JSON Google
    # Mapping: "Nom du fichier réel" : "Nom de la clé dans les secrets Streamlit"
    json_files = {
        "credentials.json": "credentials_json",
        "token_calendar.json": "token_calendar_json",
        "token.json": "token_json", # Le token utilisé par agent2.py
        # "token_drive.json": "token_drive_json" # Décommentez si nécessaire
    }
    
    if "google_files" in st.secrets:
        for filename, secret_key in json_files.items():
            if not os.path.exists(filename) and secret_key in st.secrets["google_files"]:
                with open(filename, "w") as f:
                    f.write(st.secrets["google_files"][secret_key])

    # 2. Reconstruction du fichier .env (pour Grok API)
    # On crée un fichier .env physique contenant la clé
    if not os.path.exists(".env") and "env_vars" in st.secrets:
        with open(".env", "w") as f:
            # On écrit chaque variable ligne par ligne
            for key, value in st.secrets["env_vars"].items():
                f.write(f"{key}={value}\n")

# Exécution immédiate au lancement
setup_remote_files()

# -------------------
# Config UI
# -------------------
st.set_page_config(page_title="FortalIA ", page_icon="📅", layout="wide")
st.markdown(
    """
    <style>
    .header {
        background: linear-gradient(90deg,#1f6feb 0%,#7ddbff 100%);
        padding: 18px;
        border-radius: 10px;
        color: white;
    }
    .card {
        background: linear-gradient(180deg,#ffffff 0%, #f7fbff 100%);
        padding: 12px;
        border-radius: 10px;
        box-shadow: 0 2px 10px rgba(31,111,235,0.08);
    }
    .stat { padding:12px; border-radius:8px; color:white; text-align:center; }
    .green{background: linear-gradient(90deg,#2ecc71,#27ae60);}
    .red{background: linear-gradient(90deg,#ff6b6b,#e74c3c);}
    .blue{background: linear-gradient(90deg,#4cc3ff,#1f6feb);}
    a.file-link { color: #054A91; font-weight:600; text-decoration: none; }
    </style>
    """, unsafe_allow_html=True
)
st.write("")

# -------------------
# Paths
# -------------------
EXCEL_FILE = "prospects.xlsx"
AGENT_SCRIPT = "agent2.py"
DATE_SCRIPT = "date.py"
CREDENTIALS = "credentials.json"
CALENDAR_TOKEN = "token_calendar.json"

# -------------------
# Helpers
# -------------------
def read_excel_df(path=EXCEL_FILE):
    if not os.path.exists(path):
        return pd.DataFrame()
    try:
        df = pd.read_excel(path)
        return df
    except Exception as e:
        st.error(f"Erreur lecture Excel: {e}")
        return pd.DataFrame()

def run_script(script_path):
    """Exécute un script via subprocess et renvoie (ok, output)."""
    if not os.path.exists(script_path):
        return False, f"Script introuvable : {script_path}"
    try:
        proc = subprocess.run([sys.executable, script_path], capture_output=True, text=True, check=False)
        out = ""
        if proc.stdout:
            out += proc.stdout
        if proc.stderr:
            out += ("\n=== ERREURS ===\n" + proc.stderr)
        return proc.returncode == 0, out
    except Exception as e:
        return False, str(e)

def get_calendar_events(max_results=8):
    if not os.path.exists(CALENDAR_TOKEN) or not os.path.exists(CREDENTIALS):
        return None, "Fichier credentials/token manquant."
    try:
        creds = Credentials.from_authorized_user_file(CALENDAR_TOKEN, ['https://www.googleapis.com/auth/calendar.readonly'])
        service = build('calendar', 'v3', credentials=creds)
        now = datetime.utcnow().isoformat() + 'Z'
        events_result = service.events().list(calendarId='primary', timeMin=now, maxResults=max_results, singleEvents=True, orderBy='startTime').execute()
        return events_result.get('items', []), None
    except Exception as e:
        return None, str(e)

def make_file_download(path):
    if os.path.exists(path):
        with open(path, "rb") as f:
            st.download_button("📥 Télécharger le fichier Excel (.xlsx)", f, file_name=os.path.basename(path), mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.warning("Fichier introuvable.")

def make_csv_download(df, default_name="prospects_export.csv"):
    if df is None or df.empty:
        st.info("Aucun contenu à exporter.")
        return
    csv = df.to_csv(index=False).encode('utf-8')
    st.download_button("⬇️ Télécharger CSV", data=csv, file_name=default_name, mime="text/csv")

# -------------------
# Layout
# -------------------
left_col, right_col = st.columns([1,2])

with left_col:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("Actions")
    st.write("")

    if st.button("📩 Récupérer les emails "):
        with st.spinner("Récupération..."):
            ok, out = run_script(AGENT_SCRIPT)
            if ok:
                st.success("✅ agent2.py exécuté")
            else:
                st.error("⚠️ Erreur lors de l'exécution")
            st.code(out, language="bash")

    if st.button("📅 Créer RDV et envoyer confirmations"):
        with st.spinner("Création RDV..."):
            ok, out = run_script(DATE_SCRIPT)
            if ok:
                st.success("✅ date.py exécuté")
            else:
                st.error("⚠️ Erreur lors de l'exécution")
            st.code(out, language="bash")

    st.write("---")
    st.markdown("**Fichier Excel (local)**")
    make_file_download(EXCEL_FILE)
    st.write("")
    st.markdown("**Télécharger CSV**")
    df_tmp = read_excel_df()
    make_csv_download(df_tmp)

    st.write("---")
    st.markdown("**Envoyer demande de documents**")
    st.caption("Sélectionne un ou plusieurs prospects ci-dessous, choisis le collaborateur puis clique sur le bouton.")

    # Load prospects for selection
    df = read_excel_df()
    if df.empty:
        st.info("Aucun prospect détecté (lance d'abord la récupération des emails).")
        prospects_options = []
    else:
        # ensure required columns exist
        for col in ["Prénom","Nom","Email","Téléphone","Traité","Jours disponibles","Plages horaires"]:
            if col not in df.columns:
                df[col] = ""

        # build option labels
        def label_from_row(idx, row):
            p = (row.get("Prénom") or "").strip()
            n = (row.get("Nom") or "").strip()
            e = (row.get("Email") or "").strip()
            return f"{idx} — {p} {n} <{e}>"

        prospects_list = []
        for idx, row in df.iterrows():
            prospects_list.append((idx, label_from_row(idx+2, row)))  # idx+2 for Excel row number (header=1)

        options = [lab for (_, lab) in prospects_list]

        selected = st.multiselect("Choisir prospects (1+)", options=options, default=None)

        # collaborator selection
        # try import collaborators list from date.py if present
        try:
            date_mod = importlib.import_module("date")
            collaborators_choices = date_mod.collaborateurs if hasattr(date_mod, "collaborateurs") else []
        except Exception:
            collaborators_choices = []

        if collaborators_choices:
            collab_labels = [f"{c['prenom']} {c['nom']} — {c['email']}" for c in collaborators_choices]
            collab_choice = st.selectbox("Collaborateur (expéditeur)", collab_labels)
            collab = collaborators_choices[collab_labels.index(collab_choice)]
        else:
            st.info("Aucun collaborateur trouvé dans date.py — utilisation d'un par défaut.")
            collab = {"prenom":"Automatique","nom":"","email":"u9373548876@gmail.com"}

        if st.button("✉️ Envoyer demande de documents"):
            if not selected:
                st.error("Sélectionne au moins un prospect.")
            else:
                # ensure date module gmail service is initialized
                try:
                    date_mod = importlib.import_module("date")
                    if not hasattr(date_mod, "gmail_service") or date_mod.gmail_service is None:
                        date_mod.gmail_service = date_mod.get_gmail_service()
                except Exception as e:
                    st.error(f"Impossible d'initialiser date.py : {e}")
                    date_mod = None

                successes = []
                failures = []
                for sel in selected:
                    # sel example: "3 — Prénom Nom <email>"
                    try:
                        # extract the email from between <>
                        email = sel.split("<")[-1].split(">")[0].strip()
                        # find row by Excel row number (we included idx+2)
                        excel_row_num = int(sel.split("—")[0].strip())
                        df_row = df.iloc[excel_row_num-2]  # convert back to 0-based
                        prenom = df_row.get("Prénom", "")
                        nom = df_row.get("Nom", "")
                        if date_mod:
                            try:
                                date_mod.send_documents_request_email(collab, prenom, nom, email)
                                successes.append(email)
                            except Exception as e:
                                failures.append((email, str(e)))
                        else:
                            failures.append((email, "Module date.py non disponible"))
                    except Exception as e:
                        failures.append((sel, str(e)))

                if successes:
                    st.success(f"Email(s) envoyés à : {', '.join(successes)}")
                if failures:
                    st.error("Erreurs :")
                    for f in failures:
                        st.write(f"- {f}")
    st.markdown('</div>', unsafe_allow_html=True)

with right_col:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("📊 Statistiques rapides")
    df_display = read_excel_df()
    if df_display.empty:
        st.warning("Aucun prospect trouvé.")
        total = treated = not_treated = 0
    else:
        total = len(df_display)
        treated = df_display['Traité'].fillna("").apply(lambda x: str(x).strip() == "✔️").sum()
        not_treated = total - treated

    c1, c2, c3 = st.columns(3)
    c1.markdown(f"<div class='stat blue'><h2 style='margin:0'>{total}</h2><div>Totaux</div></div>", unsafe_allow_html=True)
    c2.markdown(f"<div class='stat green'><h2 style='margin:0'>{treated}</h2><div>Traités</div></div>", unsafe_allow_html=True)
    c3.markdown(f"<div class='stat red'><h2 style='margin:0'>{not_treated}</h2><div>Non traités</div></div>", unsafe_allow_html=True)

    st.write("")
    st.subheader("Prochaines actions")
    st.markdown("- Utilise la colonne **Traité** dans l'Excel pour suivre l'état.\n- Tu peux sélectionner plusieurs prospects à gauche et envoyer la demande de documents.")

    st.write("---")
    st.subheader("🗓️ Aperçu Google Calendar")
    events, err = get_calendar_events(max_results=8)
    if err:
        st.info("Impossible de lister les événements (vérifie credentials/token).")
        if isinstance(err, str):
            st.write(err)
    else:
        if events:
            for ev in events:
                start = ev['start'].get('dateTime', ev['start'].get('date'))
                st.markdown(f"- **{ev.get('summary','(sans titre)')}** — {start} — [Voir]({ev.get('htmlLink')})")
        else:
            st.info("Aucun événement à venir trouvé.")

    st.markdown('</div>', unsafe_allow_html=True)

st.write("---")
if st.button("🔁 Recharger l'interface"):
    st.rerun()



