import webbrowser
import requests
import msal
from msal import PublicClientApplication
import datetime
import tkinter as tk
from tkinter import messagebox
import threading

# Remplacez les valeurs réelles par des placeholders
APPLICATION_ID = 'YOUR_APPLICATION_ID' # Remplacez 'YOUR_APPLICATION_ID' par votre ID d'application Microsoft Azure
CLIENT_SECRET = 'YOUR_CLIENT_SECRET' # Remplacez 'YOUR_CLIENT_SECRET' par votre secret client Microsoft Azure
authority_url = 'AUTHORITY_URL' # L'URL d'autorité pour l'authentification, généralement https://login.microsoftonline.com/consumers/ pour les comptes personnels Microsoft
base_url = 'BASE_URL' # URL de base de l'API Microsoft Graph, habituellement https://graph.microsoft.com/v1.0/

SCOPES = ['User.Read', 'Calendars.Read', 'Calendars.ReadBasic', 'Calendars.ReadWrite', 'Contacts.Read']

app = PublicClientApplication(APPLICATION_ID, authority=authority_url)
flow = app.initiate_device_flow(scopes=SCOPES)

global evenements_globaux
# Déclaration de la variable globale
evenements_globaux = []

global interface1
global interface2
interface1 = 1
interface2 = 0

def ouvrir_url():
    webbrowser.open(flow['verification_uri'])
    
############################

def copier_code(user_code):
    root.clipboard_clear()
    root.clipboard_append(user_code)
    messagebox.showinfo("Copié", "Code copié dans le presse-papiers")

######################

def afficher_details_evenements(reponse_evenements):
    if 'value' in reponse_evenements:
        for evenement in reponse_evenements['value']:
            print(f"Événement : {evenement.get('subject', 'Non spécifié')}")
            print(f"Date de début : {evenement.get('start', {}).get('dateTime', 'Non spécifié')}")
            print(f"Date de fin : {evenement.get('end', {}).get('dateTime', 'Non spécifié')}")
            print("Participants :")
            for participant in evenement.get('attendees', []):
                print(f"- {participant['emailAddress']['name']}")
            print("---------")
    else:
        print("Aucun événement trouvé.")
    print("")

###############

def afficher_details_contacts(reponse_contacts):
    if 'value' in reponse_contacts:
        for contact in reponse_contacts['value']:
            print(f"Nom : {contact.get('displayName', 'Non spécifié')}")
            telephones = contact.get('businessPhones', [])
            telephones = contact.get('mobilePhone', [])
            print(f"Numéro de téléphone : {telephones if telephones else 'Non spécifié'}")
			#print(f"Numéro de téléphone : {contact.get('businessPhones', ['Non spécifié'])[0]}")
            print("--------- reponse au complet -----")
            print(reponse_contacts)
            print("--------------")
            print("")
            
    else:
        print("Aucun contact trouvé.")
        print("")



def recuperer_info_participants(access_token_id, evenements_15_jours, evenements_2_jours):
    headers = {'Authorization': 'Bearer ' + access_token_id}

    # Récupérer les contacts
    endpoint_contacts = f"{base_url}me/contacts"
    reponse_contacts = requests.get(endpoint_contacts, headers=headers).json()
    contacts_par_nom = {contact['displayName'].lower(): contact for contact in reponse_contacts.get('value', [])}

    print("Contacts récupérés :")
    for nom, contact in contacts_par_nom.items():
        print(f"- {nom}: {contact.get('mobilePhone', 'Aucun numéro de téléphone')}")

    for label, evenements in [("Événements dans 15 jours", evenements_15_jours), ("Événements dans 2 jours", evenements_2_jours)]:
        print(f"\nTraitement des {label}:")
        for evenement in evenements.get('value', []):
            print(f"Traitement de l'événement : {evenement.get('subject', 'Sujet non spécifié')}")
            info_evenement = {
                'nom': evenement.get('subject', 'Sujet non spécifié'),
                'debut': evenement.get('start', {}).get('dateTime', 'Heure non spécifiée'),
                'participants': []
            }

            if 'attendees' in evenement:
                for participant in evenement['attendees']:
                    nom_participant = participant['emailAddress']['name'].lower()
                    if nom_participant in contacts_par_nom:
                        contact = contacts_par_nom[nom_participant]
                        telephone = contact.get('mobilePhone', 'Numéro non spécifié')
                        info_evenement['participants'].append({'nom': participant['emailAddress']['name'], 'telephone': telephone})
                        print(f"    -> Participant trouvé : {participant['emailAddress']['name']}, Téléphone : {telephone}")
                    else:
                        print(f"    -> Participant non trouvé dans les contacts : {participant['emailAddress']['name']}")

            else:
                print("    -> Aucun participant trouvé pour cet événement")

            # Ajout à la liste globale si non présent
            if not any(evenement_global['nom'] == info_evenement['nom'] for evenement_global in evenements_globaux):
                evenements_globaux.append(info_evenement)


def afficher_details_participants(evenements):#affiche les info des participants attendees 
    for evenement in evenements.get('value', []):
        print(f"Événement: {evenement.get('subject', 'Non spécifié')}")
        if 'attendees' in evenement:
            print("Participants:")
            for participant in evenement['attendees']:
                nom_participant = participant['emailAddress']['name']
                email_participant = participant['emailAddress']['address']
                print(f" - Nom: {nom_participant}, Email: {email_participant}")
        else:
            print(" Aucun participant.")



def afficher_evenements(evenements):
    for evenement in evenements['value']:
        print(f"Nom de l'événement: {evenement['subject']}")
        print(f"Date de début: {evenement['start']['dateTime']} (TimeZone: {evenement['start']['timeZone']})")
        print(f"Date de fin: {evenement['end']['dateTime']} (TimeZone: {evenement['end']['timeZone']})")
        print(f"Lieu: {evenement['location'].get('displayName', 'Non spécifié')}")
        print(f"Organisateur: {evenement['organizer']['emailAddress']['name']}")
        if 'attendees' in evenement and evenement['attendees']:
            print("Participants:")
            for participant in evenement['attendees']:
                print(f"- [{participant['emailAddress']['name']}]")
        else:
            print("Pas de participants")
        print(f"Description: {evenement.get('bodyPreview', 'Pas de description disponible')}")
        print("")

#############################

def envoyer_rappel(evenement):
    print("Bouton envoi rappel cliqué pour l'événement:", evenement.get('nom', 'Inconnu'))

def afficher_evenements_interface(evenements_globaux):
    print("Affichage des événements dans l'interface graphique")
    for evenement in evenements_globaux:
        frame = tk.Frame(root)
        frame.pack()

        label = tk.Label(frame, text=f"Événement: {evenement['nom']}, Date: {evenement['debut']}")
        label.pack()

        if evenement['participants']:
            for participant in evenement['participants']:
                var = tk.IntVar()
                checkbox = tk.Checkbutton(frame, text=f"{participant['nom']} - {participant['telephone']}", variable=var)
                checkbox.pack()
        else:
            label = tk.Label(frame, text="Aucun participant")
            label.pack()

        button = tk.Button(frame, text="Envoyer Rappel", command=lambda e=evenement: envoyer_rappel(e))
        button.pack()

def afficher_evenements_apres_authentification():
    print("Mise à jour de l'interface graphique après l'authentification")
    # Supprimez les widgets existants
    for widget in root.winfo_children():
        widget.destroy()

    # Affichez les événements
    afficher_evenements_interface(evenements_globaux)

def lancer_interface_graphique():
    global root, interface1, interface2
    root = tk.Tk()
    root.title("Connexion Microsoft Graph")
    root.geometry('1200x600')

    if 'user_code' in flow and interface1 == 1:
        print("interface 1")
        user_code_from_flow = flow['user_code']
        label = tk.Label(root, text=f"Votre code d'utilisateur est: {user_code_from_flow}")
        label.pack()
        copy_button = tk.Button(root, text="Copier le code", command=lambda: copier_code(user_code_from_flow))
        copy_button.pack()
        open_browser_button = tk.Button(root, text="Ouvrir le navigateur pour authentification", command=ouvrir_url)
        open_browser_button.pack()

    root.mainloop()


########################################

def gestion_api():
    # Authentification et récupération du token
    global interface1, interface2

    result = app.acquire_token_by_device_flow(flow)
    if "access_token" in result:
        access_token_id = result['access_token']
        headers = {'Authorization': 'Bearer ' + access_token_id}

        endpoint = base_url + 'me'
        response = requests.get(endpoint, headers=headers)
        print("Réponse de l'API (Utilisateur):", response.json())

        date_debut_15_jours = (datetime.datetime.now() + datetime.timedelta(days=15)).strftime("%Y-%m-%dT00:00:01")
        date_fin_15_jours = (datetime.datetime.now() + datetime.timedelta(days=15)).strftime("%Y-%m-%dT23:59:59")
        date_debut_2_jours = (datetime.datetime.now() + datetime.timedelta(days=2)).strftime("%Y-%m-%dT00:00:00")
        date_fin_2_jours = (datetime.datetime.now() + datetime.timedelta(days=2)).strftime("%Y-%m-%dT23:59:59")
        endpoint_15_jours = f"{base_url}me/calendar/calendarView?startDateTime={date_debut_15_jours}&endDateTime={date_fin_15_jours}"
        response_15_jours = requests.get(endpoint_15_jours, headers=headers)
        endpoint_2_jours = f"{base_url}me/calendar/calendarView?startDateTime={date_debut_2_jours}&endDateTime={date_fin_2_jours}"
        response_2_jours = requests.get(endpoint_2_jours, headers=headers)
        print("Événements dans 15 jours:")
        afficher_evenements(response_15_jours.json())
        print("Événements dans 2 jours:")
        afficher_evenements(response_2_jours.json())
		#print("Événements dans les prochains jours:", response_calendar.json())
        recuperer_info_participants(access_token_id, response_15_jours.json(), response_2_jours.json())
        interface1 = 0
        interface2 = 1
        afficher_evenements_apres_authentification()
        
    else:
        print("Erreur lors de l'obtention du token d'accès.")
    


# Lancement du thread pour la gestion de l'API
thread_api = threading.Thread(target=gestion_api)
thread_api.start()

# Lancement de l'interface graphique dans le thread principal
lancer_interface_graphique()
