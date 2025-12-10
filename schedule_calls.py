import datetime
import random
import time
import base64
from email.mime.text import MIMEText
from calendar_setup import get_calendar_service
from gmail_quickstart import get_service  # <- utilise ton gmail_quickstart

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
        'description': f"Appel téléphonique avec {client_name} ({client_email}). Durée : {DUREE_APPEL_MIN} minutes.\nTéléphone: {client_email}",
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

# --- Exemple de collaborateurs et prospects ---
collaborateurs = [
    {"nom": "Celine", "email": "celine.pro@exemple.com"},
    {"nom": "Ines", "email": "ines.pro@exemple.com"}
]

prospects = [
    {"prenom": f"Prenom{i}", "nom": f"Nom{i}", "email": f"prospect{i}@example.com"}
    for i in range(10)  # 10 pour le test
]

# --- Simulation des rendez-vous ---
today = datetime.datetime.now().replace(hour=9, minute=0, second=0, microsecond=0)
for i, prospect in enumerate(prospects):
    collab = random.choice(collaborateurs)
    start_time = today + datetime.timedelta(minutes=DUREE_APPEL_MIN*i)
    client_name = f"{prospect['prenom']} {prospect['nom']}"
    client_email = prospect['email']

    # Crée le rendez-vous Calendar
    create_client_meeting(client_name, client_email, start_time, collab['email'])

    # Mail de notification aux deux
    subject = f"Rendez-vous téléphonique avec {client_name}"
    body = f"Bonjour {collab['nom']},\n\nVous avez un appel programmé avec {client_name} ({client_email}) le {start_time.strftime('%d/%m/%Y à %H:%M')}.\nDurée : {DUREE_APPEL_MIN} minutes."
    send_email(collab['email'], collab['email'], subject, body)  # au collaborateur
    send_email(collab['email'], client_email, subject, body)      # au client

    # Mail de remerciement 1 min après l'appel
    thank_you_subject = f"Merci pour votre appel, {prospect['prenom']}"
    thank_you_body = f"Bonjour {prospect['prenom']},\n\nMerci pour votre échange avec {collab['nom']}.\nNous restons à votre disposition."
    
    # On simule le délai d'une minute après la fin de l'appel
    time_to_wait = (DUREE_APPEL_MIN + DELAI_REMERCIMENT_MIN) * 60
    time.sleep(0.1)  # <-- pour test on ne met pas toute la durée, juste un petit sleep
    send_email(collab['email'], client_email, thank_you_subject, thank_you_body)
    print(f"Mail de remerciement envoyé à {prospect['prenom']}")
