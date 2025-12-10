import datetime
import random
import time
import base64
import pandas as pd
from email.mime.text import MIMEText
from calendar_setup import get_calendar_service
from gmail_quickstart import get_service  

# Connexion aux services
calendar_service = get_calendar_service()
gmail_service = get_service()

# --- Paramètres ---
DUREE_APPEL_MIN = 30  # minutes
DELAI_REMERCIMENT_MIN = 1  # minutes après l'appel
TIMEZONE = 'Europe/Paris'
JOURS_DISPO = ["9:00-11:30", "11:30-14:30", "14:30-18:00", "18:00-20:00"]

# --- Fonction pour créer un rendez-vous Calendar ---
def create_client_meeting(client_name, client_email, start_datetime, collab_email):
    event = {
        'summary': f"Appel avec {client_name}",
        'description': f"Appel téléphonique avec {client_name} ({client_email}). Durée : {DUREE_APPEL_MIN} minutes.",
        'start': {'dateTime': start_datetime.isoformat(), 'timeZone': TIMEZONE},
        'end': {'dateTime': (start_datetime + datetime.timedelta(minutes=DUREE_APPEL_MIN)).isoformat(), 'timeZone': TIMEZONE},
        'attendees': [
            {'email': client_email},
            {'email': collab_email}
        ],
        'reminders': {
            'useDefault': False,
            'overrides': [
                {'method': 'email', 'minutes': 24*60},  # 24h avant
                {'method': 'popup', 'minutes': 30},
            ],
        },
    }
    created_event = calendar_service.events().insert(calendarId='primary', body=event, sendUpdates='all').execute()
    print(f"Rendez-vous créé pour {client_name} : {created_event.get('htmlLink')}")
    return created_event

# --- Fonction pour envoyer un mail via Gmail ---
def send_email(sender, to, subject, body_text):
    msg = MIMEText(body_text, "plain", "utf-8")
    msg['To'] = to
    msg['From'] = sender
    msg['Subject'] = subject
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    gmail_service.users().messages().send(userId="me", body={'raw': raw}).execute()

# --- Exemple de collaborateurs ---
collaborateurs = [
    {"nom": "Dupont", "prenom":"Marie", "email": "u9373548876@gmail.com"},
    {"nom": "Martin", "prenom":"Alex", "email": "alex.pro@example.com"} 
]

# --- Lecture des clients depuis un fichier Excel ---
# Assure-toi que le fichier a les colonnes : prenom, nom, email
df_clients = pd.read_excel("prospects.xlsx")  

# --- Simulation des rendez-vous ---
today = datetime.datetime.now().replace(hour=9, minute=0, second=0, microsecond=0)
for i, row in df_clients.iterrows():
    collab = random.choice(collaborateurs)
    start_time = today + datetime.timedelta(minutes=DUREE_APPEL_MIN*i)
    client_name = f"{row['prenom']} {row['nom']}"
    client_email = row['email']

    # Crée le rendez-vous Calendar
    create_client_meeting(client_name, client_email, start_time, collab['email'])

# --- Mail de notification pour le rendez-vous ---

subject = f"Confirmation de votre rendez-vous téléphonique avec {client_name}"

# Corps pour le collaborateur
body_collab = f"""Bonjour {collab['prenom']} {collab['nom']},

Vous avez un appel programmé avec {client_name} ({client_email}) le {start_time.strftime('%d/%m/%Y à %H:%M')}, d'une durée de {DUREE_APPEL_MIN} minutes.

Merci de préparer cet échange afin d'assurer un suivi optimal.

Cordialement,
L’équipe FortalIA
"""

# Corps pour le client
body_client = f"""Bonjour {row['prenom']} {row['nom']},

Nous vous confirmons votre rendez-vous téléphonique avec {collab['prenom']} {collab['nom']} le {start_time.strftime('%d/%m/%Y à %H:%M')}, d'une durée de {DUREE_APPEL_MIN} minutes.

Nous restons à votre disposition pour toute information complémentaire.

Cordialement,
L’équipe FortalIA
"""

# Envoi des mails de notification
send_email(collab['email'], collab['email'], subject, body_collab)
send_email(collab['email'], client_email, subject, body_client)

# --- Mail de remerciement envoyé au client (1 min après l'appel) ---
time.sleep(0.1)  # pour test, en prod: time.sleep((DUREE_APPEL_MIN + DELAI_REMERCIMENT_MIN) * 60)

thank_you_subject = f"Merci pour votre appel, {row['prenom']}"
thank_you_body = f"""Bonjour {row['prenom']} {row['nom']},

Nous vous remercions pour l’échange téléphonique que vous avez eu avec {collab['prenom']} {collab['nom']}.  
Nous restons à votre disposition pour toute question ou information supplémentaire concernant votre dossier.

Cordialement,
L’équipe FortalIA
"""

send_email(collab['email'], client_email, thank_you_subject, thank_you_body)

# --- Notification interne au collaborateur que le mail de remerciement a été envoyé ---
notif_subject = f"Mail de remerciement envoyé à {row['prenom']} {row['nom']}"
notif_body = f"Le mail de remerciement a été envoyé avec succès à {row['prenom']} {row['nom']} ({client_email})."
send_email(collab['email'], collab['email'], notif_subject, notif_body)

print(f"Mail de remerciement envoyé à {row['prenom']} {row['nom']}")

