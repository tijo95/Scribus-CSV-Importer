import scribus
import tkinter as tk
from tkinter import filedialog

# Chemin vers le fichier CSV sur le bureau
csv_file = ""

# Coordonnées initiales pour le cadre texte
x, y = 20, 20  # Ajustez ces valeurs pour définir la position de départ à l'intérieur de la page

# Dimensions du cadre de texte (35mm x 35mm)
largeur, hauteur = 35, 35

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
                y += hauteur + 10  # Décalez verticalement pour laisser de l'espace entre les cadres

                # Vérifiez si le cadre texte dépasse la hauteur de la page, et passez à la page suivante si nécessaire
                hauteur_page = scribus.getPageHeight()
                if y + hauteur > hauteur_page:
                    y = 20  # Réinitialisez la position à l'intérieur de la page
                    # Si nécessaire, ajoutez du code pour passer à la page suivante ici
            else:
                scribus.messageBox("Erreur", "Numéro de ligne invalide.", icon=scribus.ICON_ERROR)
    except ValueError:
        scribus.messageBox("Erreur", "Veuillez entrer un numéro de ligne valide.", icon=scribus.ICON_ERROR)

# Fonction pour supprimer le contenu du champ de saisie
def supprimer_champ():
    entree_ligne.delete(0, tk.END)
    actualiser_boutons()  # Mettez à jour l'état des boutons
    etiquette_mots.config(text="")

# Fonction pour ouvrir le fichier CSV à partir d'une boîte de dialogue
def ouvrir_fichier_csv():
    global csv_file
    csv_file = filedialog.askopenfilename(filetypes=[("Fichiers texte", "*.txt")])
    if csv_file:
        entree_ligne.config(state=tk.NORMAL)  # Activer la saisie du numéro de ligne
        bouton_plus.config(state=tk.NORMAL)  # Activer le bouton "+" lorsque le fichier est ouvert
        afficher_nombre_lignes()
        bouton_importer_tout.config(state=tk.NORMAL)  # Activer le bouton "Importer Tout"

        # Mettez à jour l'état du bouton "Importer Tout" en fonction du nombre de lignes
        afficher_nombre_lignes()

# Fonction pour afficher le nombre de lignes dans le fichier ouvert
def afficher_nombre_lignes():
    with open(csv_file, "r", encoding="utf-8") as file:
        lines = file.readlines()
    nombre_lignes = len(lines)
    etiquette_nombre_lignes.config(text=f"({nombre_lignes}) lignes détectées", fg="blue")
    # Mettez à jour l'état du bouton "Importer Tout" en fonction du nombre de lignes
    if nombre_lignes > 0:
        bouton_importer_tout.config(state=tk.NORMAL)
    else:
        bouton_importer_tout.config(state=tk.DISABLED)

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

# Fonction pour afficher l'aide
def afficher_aide():
    message_aide = """
    Ce script permet d'importer des lignes à partir d'un fichier CSV (.txt) dans Scribus.
    
    Assurez-vous de nettoyer et de formater votre fichier Excel avant d'importer les données dans le fichier (.txt).
    
    Utilisation :
    1. Cliquez sur le bouton "Ouvrir fichier .txt" pour sélectionner un fichier (.txt).
    2. Utilisez le bouton "+" pour démarrer la navigation sans avoir besoin de taper au clavier. Utilisez les boutons flèches gauche (<---) et droite (--->) pour ajuster le numéro de ligne.
    3. Les trois premiers mots de la ligne seront affichés comme repère.
    4. Cliquez sur le bouton "Importer" pour ajouter la ligne complète à Scribus dans son cadre texte.
    5. Utilisez le bouton "Importer Tout" pour importer tout le fichier en une seule fois.
    6. Utilisez le bouton "Sppr" pour supprimer le contenu du champ de saisie.
    
    Assurez-vous que Scribus est ouvert et que vous avez un document Scribus actif avant d'importer des lignes.
    Note : Les boutons flèches gauche et droite permettent d'ajuster le numéro de ligne sans avoir à le saisir manuellement, ce qui peut être utile pour naviguer rapidement dans le fichier (.txt).
    """
    scribus.messageBox("Aide", message_aide)

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
            entree_ligne.delete(0, tk.END)  # Effacez l'ancien numéro
            entree_ligne.insert(0, str(ligne_num))  # Mettez à jour le champ de saisie
            mettre_a_jour_affichage()  # Mettez à jour l'affichage
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
        bouton_importer_tout.config(state=tk.NORMAL)  # Activer le bouton "Importer Tout" lorsque le champ est vide
        bouton_plus.config(state=tk.NORMAL)  # Activer le bouton "+" lorsque le champ est vide
    elif contenu_champ.isdigit():
        bouton_importer.config(state=tk.NORMAL)
        bouton_sppr.config(state=tk.NORMAL)
        bouton_importer_tout.config(state=tk.DISABLED)  # Désactiver le bouton "Importer Tout" lorsque le champ est renseigné
        bouton_plus.config(state=tk.DISABLED)  # Désactiver le bouton "+" lorsque le champ est renseigné
    else:
        bouton_importer.config(state=tk.DISABLED)
        bouton_sppr.config(state=tk.DISABLED)
        bouton_importer_tout.config(state=tk.DISABLED)
        bouton_plus.config(state=tk.DISABLED)

# Fonction pour ajouter 1 au champ et réactiver le bouton "Sppr"
def ajouter_un():
    contenu_champ = entree_ligne.get()
    if not contenu_champ.isdigit():
        contenu_champ = "1"
    else:
        contenu_champ = str(int(contenu_champ) + 1)
    entree_ligne.delete(0, tk.END)
    entree_ligne.insert(0, contenu_champ)
    mettre_a_jour_affichage()  # Mettez à jour l'affichage après avoir modifié le contenu du champ
    actualiser_boutons()  # Réactive le bouton "Sppr"

# Créez une fenêtre Tkinter
fenetre = tk.Tk()
fenetre.title("Importation Scribus")

# Définissez la taille de la fenêtre en pixels
fenetre.geometry("230x305")

# Cadre pour le champ de saisie et les boutons + et Sppr
cadre_champ_et_boutons = tk.Frame(fenetre)
cadre_champ_et_boutons.pack(side=tk.TOP, pady=10)

# Étiquette pour le texte "Numéro de ligne à importer" (au-dessus)
etiquette_entree = tk.Label(cadre_champ_et_boutons, text="Numéro de ligne à importer:")
etiquette_entree.grid(row=0, column=0, columnspan=3)  # Utilisez columnspan pour étendre sur trois colonnes

# Champ de saisie pour le numéro de ligne
entree_ligne = tk.Entry(cadre_champ_et_boutons, width=5, state="disabled", justify="center")  # Utilisez justify="center"
entree_ligne.grid(row=1, column=1, padx=4)

# Bouton "+" pour ajouter 1 au champ et réactiver le bouton "Sppr"
bouton_plus = tk.Button(cadre_champ_et_boutons, text="+", command=ajouter_un)
bouton_plus.grid(row=1, column=0, padx=4)

# Bouton "Sppr" pour supprimer le contenu du champ
bouton_sppr = tk.Button(cadre_champ_et_boutons, text="Sppr", command=supprimer_champ, state=tk.DISABLED)
bouton_sppr.grid(row=1, column=2, padx=4)

# Lorsque le champ de saisie du numéro de ligne change, mettez à jour les boutons et l'affichage des trois premiers mots
entree_ligne.bind("<KeyRelease>", lambda event: (mettre_a_jour_affichage(), actualiser_boutons(event)))

# Bouton pour ouvrir le fichier CSV
bouton_ouvrir = tk.Button(fenetre, text="Ouvrir fichier .txt", command=ouvrir_fichier_csv)
bouton_ouvrir.pack(pady=4)  # Ajoutez un espacement vertical de 2 pixels

# Étiquette pour afficher le nombre de lignes
etiquette_nombre_lignes = tk.Label(fenetre, text="")
etiquette_nombre_lignes.pack(pady=4)  # Ajoutez un espacement vertical de 2 pixels

# Étiquette pour afficher les trois premiers mots
etiquette_mots = tk.Label(fenetre, text="")
etiquette_mots.pack(pady=2)  # Ajoutez un espacement vertical de 2 pixels

# Cadre pour les boutons de déplacement
cadre_boutons_deplacement = tk.Frame(fenetre)
cadre_boutons_deplacement.pack()

# Bouton pour décrémenter le numéro de ligne
bouton_decrementer = tk.Button(cadre_boutons_deplacement, text="<------", command=decrementer_ligne, height=1)
bouton_decrementer.pack(side=tk.LEFT, padx=2)  # Ajoutez un espace horizontal de 2 pixels

# Bouton pour incrémenter le numéro de ligne
bouton_incrementer = tk.Button(cadre_boutons_deplacement, text="------>", command=incrementer_ligne, height=1)
bouton_incrementer.pack(side=tk.LEFT, padx=2)  # Ajoutez un espace horizontal de 2 pixels

# Bouton pour importer une ligne
bouton_importer = tk.Button(fenetre, text="Importer", command=importer_ligne, state=tk.DISABLED)
bouton_importer.pack(pady=4)  # Ajoutez un espacement vertical de 2 pixels

# Bouton pour importer tout le fichier
bouton_importer_tout = tk.Button(fenetre, text="Importer Tout", command=importer_tout, state=tk.DISABLED)
bouton_importer_tout.pack(pady=4)  # Ajoutez un espacement vertical de 2 pixels

# Bouton pour afficher l'aide
bouton_aide = tk.Button(fenetre, text="Aide", command=afficher_aide)
bouton_aide.pack(pady=2)  # Ajoutez un espacement vertical de 2 pixels

# Étiquette pour la mention "Développé par Tijo"
etiquette_developpe_par = tk.Label(fenetre, text="By Tijo", font=("Helvetica", 8))
etiquette_developpe_par.pack(side=tk.RIGHT, anchor=tk.SE, padx=5, pady=5)

# Lorsque le champ de saisie du numéro de ligne change, mettez à jour les boutons et l'affichage des trois premiers mots
entree_ligne.bind("<KeyRelease>", lambda event: (mettre_a_jour_affichage(), actualiser_boutons(event)))

# Désactiver le bouton "Importer Tout" au lancement
bouton_importer_tout.config(state=tk.DISABLED)

# Désactiver le bouton "+" au lancement
bouton_plus.config(state=tk.DISABLED)

# Lancez la boucle principale de l'interface Tkinter
fenetre.mainloop()
