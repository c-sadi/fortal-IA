import subprocess

print(" Lancement du script de récupération d'emails...")
subprocess.run(["python", "agent2.py"])

print(" Lancement du script de date...")
subprocess.run(["python", "date.py"])

print(" Tous les scripts ont été exécutés !")
