import scribus
import tkinter as tk
from tkinter import filedialog
import os
import json

# Obtenez le répertoire actuel du script
script_dir = os.path.dirname(os.path.abspath(__file__))

# Définir la fonction pour charger la configuration utilisateur
def load_config():
    config = {}
    config_file = os.path.join(script_dir, "config.json")
    if os.path.exists(config_file):
        with open(config_file, "r", encoding="utf-8") as conf_file:
            # Ajoutez une vérification pour vous assurer que le fichier n'est pas vide
            file_content = conf_file.read()
            if file_content.strip():  # Vérifiez si le contenu n'est pas vide
                config = json.loads(file_content)
    return config

# Définir la fonction pour sauvegarder la configuration utilisateur
def save_config(config):
    config_file = os.path.join(script_dir, "config.json")
    with open(config_file, "w", encoding="utf-8") as conf_file:
        json.dump(config, conf_file, indent=4)

# Chemin vers le fichier CSV sur le bureau
csv_file = ""

# Coordonnées initiales pour le cadre texte
x, y = 20, 20 

# Dimensions du cadre de texte (35mm x 35mm)
largeur, hauteur = 35, 35

# Obtenez le répertoire actuel où se trouve ce script
script_dir = os.path.dirname(os.path.abspath(__file__))

# Liste des langues disponibles
languages = ["fr", "en", "es", "de", "tr", "ru"]

# Chemin complet vers le répertoire contenant les fichiers JSON de traduction
translations_dir = os.path.join(script_dir, "translations")

# Charger les fichiers de traduction pour toutes les langues
translations = {}
for lang in languages:
    json_path = os.path.join(translations_dir, f"{lang}.json")
    with open(json_path, "r", encoding="utf-8") as lang_file:
        translations[lang] = json.load(lang_file) 

# Chargez les préférences utilisateur au démarrage du script
config = load_config()

# Utilisez la langue préférée de l'utilisateur comme langue par défaut (par exemple, "fr" par défaut)
current_language = config.get("language", "fr")        

# Fonction pour obtenir la traduction en fonction de la langue actuelle
def get_translation(key):
    return translations[current_language].get(key, key)

# Fonction pour importer une ligne à partir du fichier CSV
def importer_ligne():
    global y
    ligne_num = entree_ligne.get()
    try:
        ligne_num = int(ligne_num)
        with open(csv_file, "r", encoding="utf-8") as file:
            lines = file.readlines()
            if 1 <= ligne_num <= len(lines):
                ligne = lines[ligne_num - 1]

                # Créez un cadre de texte dans Scribus avec les dimensions spécifiées
                text_frame = scribus.createText(x, y, largeur, hauteur)

                # Insérez la ligne nettoyée dans le cadre de texte
                ligne_actuelle = ligne.strip()
                scribus.setText(ligne_actuelle, text_frame)

                # Mettez à jour la position verticale pour la prochaine ligne
                y += hauteur + 10

                # Vérifiez si le cadre texte dépasse la hauteur de la page, et passez à la page suivante si nécessaire
                hauteur_page = scribus.getPageHeight()
                if y + hauteur > hauteur_page:
                    y = 20
                    
            else:
                scribus.messageBox(get_translation("error_label"), get_translation("invalid_line_number"), icon=scribus.ICON_ERROR)
    except ValueError:
        scribus.messageBox(get_translation("error_label"), get_translation("enter_valid_line_number"), icon=scribus.ICON_ERROR)

# Fonction pour supprimer le contenu du champ de saisie
def supprimer_champ():
    entree_ligne.delete(0, tk.END)
    actualiser_boutons()
    etiquette_mots.config(text="")

# Fonction pour ouvrir le fichier CSV à partir d'une boîte de dialogue
def ouvrir_fichier_csv():
    global csv_file
    csv_file = filedialog.askopenfilename(filetypes=[("Fichiers texte", "*.txt")])
    if csv_file:
        entree_ligne.config(state=tk.NORMAL)
        bouton_plus.config(state=tk.NORMAL)
        afficher_nombre_lignes()
        bouton_importer_tout.config(state=tk.NORMAL)

        # Mettez à jour l'état du bouton "Importer Tout" en fonction du nombre de lignes
        afficher_nombre_lignes()

# Fonction pour afficher le nombre de lignes dans le fichier ouvert
def afficher_nombre_lignes():
    if csv_file:
        with open(csv_file, "r", encoding="utf-8") as file:
            lines = file.readlines()
        nombre_lignes = len(lines)
        texte_lignes_detectees = get_translation("lines_detected")
        etiquette_nombre_lignes.config(text=f"({nombre_lignes}) {texte_lignes_detectees}", fg="blue")
        # Mettez à jour l'état du bouton "Importer Tout" en fonction du nombre de lignes
        if nombre_lignes > 0:
            bouton_importer_tout.config(state=tk.NORMAL)
        else:
            bouton_importer_tout.config(state=tk.DISABLED)
    else:
        # Si aucun fichier n'est ouvert, laissez le texte vide
        etiquette_nombre_lignes.config(text="")

# Fonction pour mettre à jour l'affichage
def mettre_a_jour_affichage():
    ligne_num = entree_ligne.get()
    try:
        ligne_num = int(ligne_num)
        with open(csv_file, "r", encoding="utf-8") as file:
            lines = file.readlines()
            if 1 <= ligne_num <= len(lines):
                ligne = lines[ligne_num - 1].strip()
                mots = ligne.split()
                if len(mots) >= 3:
                    # Affichez les trois premiers mots sans mise en forme spéciale
                    mots_affichage = " ".join(mots[:3])
                    etiquette_mots.config(text=f" {mots_affichage} ...", fg="green", font=None)  # Texte noir, police par défaut
                else:
                    # Affichez tous les mots s'il y en a moins de trois
                    etiquette_mots.config(text=f"Tous les mots: {' '.join(mots)}", fg="black", font=None)  # Texte noir, police par défaut
            else:
                # Affichez "ERREUR" en rouge sans mise en forme spéciale (police par défaut)
                etiquette_mots.config(text="ERREUR", fg="red", font=None)
    except ValueError:
        # Affichez "ERREUR" en rouge sans mise en forme spéciale (police par défaut)
        etiquette_mots.config(text="ERREUR", fg="red", font=None)

# Fonction pour afficher l'aide avec traduction
def afficher_aide():
    help_text = get_translation("help")  # Utilisez la langue actuelle
    # Ouvrez une boîte de dialogue avec le texte d'aide traduit
    scribus.messageBox(get_translation("help_button"), help_text)

# Fonction pour incrémenter le numéro de ligne
def incrementer_ligne():
    ligne_num = entree_ligne.get()
    try:
        ligne_num = int(ligne_num)
        with open(csv_file, "r", encoding="utf-8") as file:
            lines = file.readlines()
            if ligne_num < len(lines):
                ligne_num += 1
                entree_ligne.delete(0, tk.END)  # Effacez l'ancien numéro
                entree_ligne.insert(0, str(ligne_num))  # Mettez à jour le champ de saisie

                # Appel pour mettre à jour les trois premiers mots
                mettre_a_jour_affichage()
    except ValueError:
        pass

# Fonction pour décrémenter le numéro de ligne
def decrementer_ligne():
    ligne_num = entree_ligne.get()
    try:
        ligne_num = int(ligne_num)
        if ligne_num > 1:
            ligne_num -= 1
            entree_ligne.delete(0, tk.END) 
            entree_ligne.insert(0, str(ligne_num)) 
            mettre_a_jour_affichage()
    except ValueError:
        pass

# Fonction pour importer tout le fichier CSV dans un cadre texte
def importer_tout():
    global y
    try:
        with open(csv_file, "r", encoding="utf-8") as file:
            contenu = file.read()

            # Créez un cadre de texte dans Scribus avec les dimensions spécifiées
            text_frame = scribus.createText(x, y, largeur, hauteur)

            # Insérez le contenu entier du fichier dans le cadre de texte
            scribus.setText(contenu, text_frame)

            # Mettez à jour la position verticale pour la prochaine ligne
            y += hauteur + 10  # Décalez verticalement pour laisser de l'espace entre les cadres

            # Vérifiez si le cadre texte dépasse la hauteur de la page, et passez à la page suivante si nécessaire
            hauteur_page = scribus.getPageHeight()
            if y + hauteur > hauteur_page:
                y = 20  # Réinitialisez la position à l'intérieur de la page
                # Si nécessaire, ajoutez du code pour passer à la page suivante ici

        scribus.messageBox("Succès", "Le contenu du fichier a été importé avec succès.", icon=scribus.ICON_INFORMATION)
    except Exception as e:
        scribus.messageBox("Erreur", f"Une erreur s'est produite : {str(e)}", icon=scribus.ICON_ERROR)

# Fonction pour activer ou désactiver les boutons en fonction du contenu du champ et du nombre de lignes
def actualiser_boutons(event=None):
    contenu_champ = entree_ligne.get()
    if contenu_champ == "":
        bouton_importer.config(state=tk.DISABLED)
        bouton_sppr.config(state=tk.DISABLED)
        bouton_importer_tout.config(state=tk.NORMAL)
        bouton_plus.config(state=tk.NORMAL)
    elif contenu_champ.isdigit():
        bouton_importer.config(state=tk.NORMAL)
        bouton_sppr.config(state=tk.NORMAL)
        bouton_importer_tout.config(state=tk.DISABLED)
        bouton_plus.config(state=tk.DISABLED)
    else:
        bouton_importer.config(state=tk.DISABLED)
        bouton_sppr.config(state=tk.DISABLED)
        bouton_importer_tout.config(state=tk.DISABLED)
        bouton_plus.config(state=tk.DISABLED)
        
# Mettez à jour les textes de l'interface lors du changement de langue
def update_interface_texts():
    fenetre.title(get_translation("title"))
    etiquette_entree.config(text=get_translation("input_label"))
    bouton_importer.config(text=get_translation("import_button"))
    bouton_importer_tout.config(text=get_translation("import_all_button"))
    bouton_sppr.config(text=get_translation("clear_button"))
    bouton_ouvrir.config(text=get_translation("open_file_button"))
    bouton_aide.config(text=get_translation("help_button"))
    etiquette_developpe_par.config(text=get_translation("developer_label"))
    etiquette_nombre_lignes.config(text=get_translation("line_count_label"))
    afficher_nombre_lignes()

# Lorsque l'utilisateur change la langue, enregistrez-la dans les préférences
def change_language(event):
    global current_language
    current_language = langue_var.get()
    config["language"] = current_language
    save_config(config)
    update_interface_texts()
    if entree_ligne.get():
        bouton_importer_tout.config(state=tk.DISABLED)

# Fonction pour ajouter 1 au champ et réactiver le bouton "Sppr"
def ajouter_un():
    contenu_champ = entree_ligne.get()
    if not contenu_champ.isdigit():
        contenu_champ = "1"
    else:
        contenu_champ = str(int(contenu_champ) + 1)
    entree_ligne.delete(0, tk.END)
    entree_ligne.insert(0, contenu_champ)
    mettre_a_jour_affichage()
    actualiser_boutons()

# Créez une fenêtre Tkinter
fenetre = tk.Tk()
fenetre.title(get_translation("title"))

# Définissez la taille de la fenêtre en pixels
fenetre.geometry("230x350")

# Cadre pour le champ de saisie et les boutons + et Sppr
cadre_champ_et_boutons = tk.Frame(fenetre)
cadre_champ_et_boutons.pack(side=tk.TOP, pady=10)

# Étiquette pour le texte "Numéro de ligne à importer" (au-dessus)
etiquette_entree = tk.Label(cadre_champ_et_boutons, text=get_translation("input_label"))
etiquette_entree.grid(row=0, column=0, columnspan=3)

# Champ de saisie pour le numéro de ligne
entree_ligne = tk.Entry(cadre_champ_et_boutons, width=5, state="disabled", justify="center")
entree_ligne.grid(row=1, column=1, padx=4)

# Bouton "+" pour ajouter 1 au champ et réactiver le bouton "Sppr"
bouton_plus = tk.Button(cadre_champ_et_boutons, text="+", command=ajouter_un)
bouton_plus.grid(row=1, column=0, padx=4)

# Bouton "Sppr" pour supprimer le contenu du champ
bouton_sppr = tk.Button(cadre_champ_et_boutons, text=get_translation("clear_button"), command=supprimer_champ, state=tk.DISABLED)
bouton_sppr.grid(row=1, column=2, padx=4)

# Lorsque le champ de saisie du numéro de ligne change, mettez à jour les boutons et l'affichage des trois premiers mots
entree_ligne.bind("<KeyRelease>", lambda event: (mettre_a_jour_affichage(), actualiser_boutons(event)))

# Bouton pour ouvrir le fichier CSV
bouton_ouvrir = tk.Button(fenetre, text=get_translation("open_file_button"), command=ouvrir_fichier_csv)
bouton_ouvrir.pack(pady=4)

# Étiquette pour afficher le nombre de lignes (initial)
etiquette_nombre_lignes = tk.Label(fenetre, text="")
etiquette_nombre_lignes.pack(pady=4)
afficher_nombre_lignes()

# Étiquette pour afficher les trois premiers mots
etiquette_mots = tk.Label(fenetre, text="")
etiquette_mots.pack(pady=2)

# Cadre pour les boutons de déplacement
cadre_boutons_deplacement = tk.Frame(fenetre)
cadre_boutons_deplacement.pack()

# Bouton pour décrémenter le numéro de ligne
bouton_decrementer = tk.Button(cadre_boutons_deplacement, text="<------", command=decrementer_ligne, height=1)
bouton_decrementer.pack(side=tk.LEFT, padx=2)

# Bouton pour incrémenter le numéro de ligne
bouton_incrementer = tk.Button(cadre_boutons_deplacement, text="------>", command=incrementer_ligne, height=1)
bouton_incrementer.pack(side=tk.LEFT, padx=2)

# Bouton pour importer une ligne
bouton_importer = tk.Button(fenetre, text=get_translation("import_button"), command=importer_ligne, state=tk.DISABLED)
bouton_importer.pack(pady=4)

# Bouton pour importer tout le fichier
bouton_importer_tout = tk.Button(fenetre, text=get_translation("import_all_button"), command=importer_tout, state=tk.DISABLED)
bouton_importer_tout.pack(pady=4)

# Bouton pour afficher l'aide
bouton_aide = tk.Button(fenetre, text=get_translation("help_button"), command=afficher_aide)
bouton_aide.pack(pady=2)

# Chargez le texte d'aide à partir du fichier JSON
help_text = get_translation("help")

# Étiquette pour la mention "Développé par Tijo"
etiquette_developpe_par = tk.Label(fenetre, text=get_translation("developer_label"), font=("Helvetica", 8))
etiquette_developpe_par.pack(side=tk.RIGHT, anchor=tk.SE, padx=5, pady=5)

# Lorsque le champ de saisie du numéro de ligne change, mettez à jour les boutons et l'affichage des trois premiers mots
entree_ligne.bind("<KeyRelease>", lambda event: (mettre_a_jour_affichage(), actualiser_boutons(event)))

# Désactiver le bouton "Importer Tout" au lancement
bouton_importer_tout.config(state=tk.DISABLED)

# Désactiver le bouton "+" au lancement
bouton_plus.config(state=tk.DISABLED)

langue_var = tk.StringVar()
langue_var.set(current_language)
langue_menu = tk.OptionMenu(fenetre, langue_var, "fr", "en", "es", "de", "tr", "ru", command=change_language)
langue_menu.pack(side=tk.BOTTOM, padx=10, pady=10) 

# Lancez la boucle principale de l'interface Tkinter
fenetre.mainloop()
