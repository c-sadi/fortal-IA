import base64
import random
import time
from email.mime.text import MIMEText
from gmail_quickstart import get_service
from googleapiclient.errors import HttpError

# TEMPLATE complet du mail
TEMPLATE = """Un prospect est intéressé pour des programmes neufs.
Chère, cher partenaire,

Un nouveau prospect-acquéreur a été qualifié par Gabby sur votre agence : {agence}.

💻  RAPPEL DE L'ANNONCE

Titre du bien : {titre}
Référence : {reference}
Ville du bien : {ville_bien}
Prix du bien : {prix}
Référent : {referent}

🪪  COORDONNÉES VÉRIFIÉES

Prénom : {prenom}
Nom : {nom}
Email : {email}
Téléphone : {telephone}
Adresse : {adresse}
Ville : {ville}
Code postal : {cp}
Département : {departement}

📝  PROFIL DU PROSPECT

Est propriétaire : {proprietaire}
Achète pour : {achat}

🏡  PROJET DU PROSPECT

Bien recherché : {bien_recherche}
Budget d'achat : {budget}
A un dossier de financement : {financement}
Délai d'achat : {delai}
Secteurs de recherche : {secteurs}

💸  APPORT D'AFFAIRES

Est intéressé par du programme neuf : {programme_neuf}

☎️  DISPONIBILITÉS DE RAPPEL

Jours disponibles : {jours}
Plages horaires : {horaires}

Gabby vous invite et vous remercie par avance à prioriser cette demande.
"""

def create_message(sender, to, subject, body_text):
    msg = MIMEText(body_text, "plain", "utf-8")
    msg['To'] = to
    msg['From'] = sender
    msg['Subject'] = subject
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    return {'raw': raw}

def random_phone():
    return "06" + "".join(str(random.randint(0, 9)) for _ in range(8))

def main():
    service = get_service()
    sender = "test@gmail.com"  # <-- mets ton email
    to = sender

    for i in range(200):
        data = {
            "agence": "Getkey Transaction",
            "titre": f"RARE – 3P SURÉLEVÉ {i} SUR JARDIN, TERRASSE 68 m² SUD-OUEST",
            "reference": f"GETKEY_{8779+i}",
            "ville_bien": random.choice(["Châtillon", "Paris", "Nanterre"]),
            "prix": random.randint(250000, 800000),
            "referent": random.choice(["Florian Lherbette", "Marie Dupont"]),
            "prenom": random.choice(["Léa", "Paul", "Lucas", "Sarah", "Emma"]),
            "nom": random.choice(["Martin", "Durand", "Petit", "Morel"]),
            "email": f"prospect{i}@example.com",
            "telephone": random_phone(),
            "adresse": f"{random.randint(1, 200)} rue Exemple",
            "ville": random.choice(["Châtillon", "Paris", "Boulogne"]),
            "cp": random.choice(["92320", "75015", "92100"]),
            "departement": random.choice(["92", "75"]),
            "proprietaire": random.choice(["Oui", "Non"]),
            "achat": random.choice(["Investir", "Résidence principale"]),
            "bien_recherche": random.choice(["Une maison", "Un appartement"]),
            "budget": random.randint(200000, 800000),
            "financement": random.choice(["Oui", "Non"]),
            "delai": random.choice(["Dès que possible", "3 mois", "6 mois"]),
            "secteurs": random.choice(["92", "75", "93"]),
            "programme_neuf": random.choice(["Oui", "Non"]),
            "jours": "Lundi, Mardi, Mercredi, Jeudi, Vendredi, Samedi",
            "horaires": "Entre 9h et 11h30, Entre 11h30 et 14h30, Entre 14h30 et 18h, Après 18h"
        }

        body = TEMPLATE.format(**data)
        subject = f"Nouveau prospect – {data['prenom']} {data['nom']}"

        msg = create_message(sender, to, subject, body)

        try:
            service.users().messages().send(userId="me", body=msg).execute()
            print(f"{i+1}/200 envoyé")
        except HttpError as e:
            print("Erreur:", e)
            if e.status_code in (429, 500, 503):
                for attempt in range(1, 6):
                    wait = 2 ** attempt
                    print(f"Retry dans {wait}s...")
                    time.sleep(wait)
                    try:
                        service.users().messages().send(userId="me", body=msg).execute()
                        print("Retry succès")
                        break
                    except HttpError as e2:
                        print("Retry échoué:", e2)
                else:
                    print(f"Abandon du message {i}")
        time.sleep(0.2)  # petit délai pour réduire les risques de quota

if __name__ == "__main__":
    main()
