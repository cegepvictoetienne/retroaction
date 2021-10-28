#!/Library/Frameworks/Python.framework/Versions/3.9/bin/python3
"""
 Générateur de rétroaction pour les élèves.

 Crée un fichier PDF pour chaque élève avec son numéro de demande d'admission (DA)
 à partir d'un chiffrier Excel qui contient les critères de correction (un critère
 par ligne dans la colonne A) et chaque élève par colonne (à partir de la colonne B)

"""
import getopt
import os
import sys

from zipfile import BadZipFile
from zipfile import ZipFile

import openpyxl
from fpdf import FPDF

# Constantes

# Quel est le caractère qui remplace le X pour indiquer que le critère est atteint
CROCHET = chr(214)
CROCHET_POLICE = "Symbol"
CROCHET_TAILLE = 14

LIBELLE_DA = "DA"
LIBELLE_NOTES = "Notes"
LIBELLE_NOM = "Nom"
LIBELLE_PRENOM = "Prénom"
LIBELLE_COMMENTAIRES = "Commentaires"

def affiche_aide():
    """
        Affiche l'aide pour la commande.
    """

    print("")
    print("""
    retroaction.py -i <fichier_retro> -o <dossier_sortie> -s <nom_feuille> -d <denominateur>

    -i : Le chiffrier Excel contenant les rétroactions aux élèves. Chaque élément de la grille d'évaluation est en ligne et chaque élève est une colonne. Relatif au répertoire courant.
    -o : Le dossier dans lequel seront créés les pdf et l'archive zip.  Relatif au répertoire courant.
    -s : Le nom de la feuille contenant les rétroactions aux élèves.
    -d : Le dénominateur de la note de l'évaluation.
    """)

def traiter_eleve(dossier_sortie, numero_da, feuille_a_traiter, denominateur, colonne_a_traiter):
    """
        Créer le PDF pour un élève.

        Paramètres
        ----------
        dossier_sortie : str
            Chemin sur disque du dossier qui recevra le PDF
        numéro_da : str
            Le numéro de DA à utiliser dans le nom du PDF
        feuille_a_traiter : objet data_sheet
            La feuille Excel qui contient les rétroactions à traiter pour l'élève.
        denominateur : int
            Le dénominateur de la note totale
        colonne_a_traiter : int
            L'index de la colonne de la feuille Excel qui contient le détail pour l'élève

        Retour
        ------
        nom_pdf : str
            Le nom du pdf créé
    """

    # Générer le nom du PDF
    nom_pdf = f"{dossier_sortie}/{numero_da}.pdf"

    # Créer le PDF
    pdf = FPDF('P', 'in', 'Letter')
    pdf.add_page()
    pdf.set_font('Arial', 'B', 12)

    # Traiter tous les critères de correction pour l'élève
    for ligne in range(1, feuille_a_traiter.max_row + 1):

        # Générer le titre du critère, si pas de titre, pas de bordure
        bordure = 1
        titre_critere = feuille_a_traiter.cell(column=1, row=ligne).value
        if titre_critere is None:
            titre_critere = " "
            # Si le champ est vide, ne pas afficher la bordure
            bordure = 0

        # Générer la valeur du critère
        valeur_critere_brut = feuille_a_traiter.cell(
            column=colonne_a_traiter,
            row=ligne).value

        # Afficher la note avec son dénominateur
        if titre_critere == LIBELLE_NOTES:
            # Calculer sur 100
            sur_100 = round(int(valeur_critere_brut) / denominateur * 100)
            valeur_critere = f"{valeur_critere_brut} / {denominateur} ({sur_100} %)"
        else:
            if valeur_critere_brut is None:
                valeur_critere = " "
            else:
                valeur_critere = valeur_critere_brut

        # Faire le traitement de texte large pour les commentaires seulement
        if titre_critere == LIBELLE_COMMENTAIRES:
            pdf.cell(4, 0.3, titre_critere, 1, 0, 'L')
            pdf.multi_cell(4, 0.3, valeur_critere, 1)
        else:
            pdf.set_font('Arial', 'B', 12)
            pdf.cell(4, 0.3, titre_critere, bordure, 0, 'L')

            if valeur_critere == "X":
                valeur_critere = CROCHET
                pdf.set_font(CROCHET_POLICE, 'B', CROCHET_TAILLE)

            pdf.cell(4, 0.3, valeur_critere, bordure, 1, 'C')

    # Écrire le PDF sur disque
    pdf.output(nom_pdf, 'F')

    return nom_pdf

def trouver_lignes_criteres(feuille_a_traiter):
    """
    Paramètres
    ----------
    feuille_a_traiter : objet data_sheet
        La feuille Excel qui contient les rétroactions à traiter pour l'élève.

    Retour
    ------
    La liste des critères et leur ligne dans la feuille.
    """

    # Définir les critères à transférer
    criteres = {
        LIBELLE_NOM : 0,
        LIBELLE_PRENOM : 0,
        LIBELLE_DA : 0,
        LIBELLE_NOTES : 0
        }

    # Trouver la ligne correspondante aux critères
    for cle, _ in criteres.items():
        for ligne in range(1, feuille_a_traiter.max_row + 1):
            if feuille_a_traiter.cell(column=1, row=ligne).value == cle:
                criteres[cle] = ligne
    return criteres

def sommaire_notes(feuille_a_traiter, dossier_sortie, nom_feuille_a_traiter, denominateur):
    """
        Écrire un chiffrier Excel avec la liste des DA et des notes

        Paramètres
        ----------
        feuille_a_traiter : objet data_sheet
            La feuille Excel qui contient les rétroactions à traiter pour l'élève.
        dossier_sortie : str
            Chemin sur disque du dossier qui recevra le PDF
        nom_feuille_a_traiter : str
            Le nom de la feuille Excel qui contient les rétroactions à traiter pour l'élève.
        denominateur : int
            Le dénominateur de la note totale
    """
    # Définir les critères à transférer
    criteres = trouver_lignes_criteres(feuille_a_traiter)

    # Créer le chiffrier
    chiffrier = openpyxl.Workbook()
    feuille = chiffrier[chiffrier.sheetnames[0]]

    # Écrire les entêtes
    feuille.cell(row=1, column=1).value = LIBELLE_NOM
    feuille.cell(row=1, column=2).value = LIBELLE_PRENOM
    feuille.cell(row=1, column=3).value = LIBELLE_DA
    feuille.cell(row=1, column=4).value = f'Note sur {denominateur}'
    feuille.cell(row=1, column=5).value = 'Note sur 100'
    feuille.cell(row=1, column=6).value = 'Échec'


    # Traiter chaque étudiant
    for etudiant in range(2, feuille_a_traiter.max_column + 1):
        colonne = 0
        for _, valeur in criteres.items():
            colonne = colonne + 1
            feuille.cell(row=etudiant, column=colonne).value = (
              feuille_a_traiter.cell(column=etudiant, row=valeur).value
              )

        # Calculer la note sur 100
        colonne = colonne + 1
        sur_100 = (
          int(feuille_a_traiter.cell(column=etudiant, row=criteres[LIBELLE_NOTES]).value) /
          denominateur * 100
          )
        feuille.cell(row=etudiant, column=colonne).value = sur_100

        # Indiquer si l'élève est en situation d'échec
        colonne = colonne + 1
        if sur_100 < 60:
            feuille.cell(row=etudiant, column=colonne).value = "Échec"

    chiffrier.save(filename=f"{dossier_sortie}/{nom_feuille_a_traiter}.xlsx")

def traiter_feuille(fichier_retroaction, dossier_sortie, nom_feuille_a_traiter, denominateur):
    """
        Traiter tous les élèves d'une feuille Excel.

        Paramètres
        ----------
        fichier_retroaction : str
            Nom et chemin du chiffrier Excel contenant les rétroactions
        dossier_sortie : str
            Chemin sur disque du dossier qui recevra le PDF
        nom_feuille_a_traiter : str
            Le nom de la feuille Excel qui contient les rétroactions à traiter pour l'élève.
        denominateur : int
            Le dénominateur de la note totale
    """

    # data_only=True pour avoir le résultat des formules...
    chiffrier = openpyxl.load_workbook(fichier_retroaction, data_only=True)

    feuille_a_traiter = chiffrier[nom_feuille_a_traiter]

    # Trouver la ligne correspondante au DA
    ligne_da = 0
    for ligne in range(1, feuille_a_traiter.max_row + 1):
        if feuille_a_traiter.cell(column=1, row=ligne).value == LIBELLE_DA:
            ligne_da = ligne

    # Créer le fichier ZIP
    nom_zip = os.path.join(dossier_sortie, "travaux.zip")
    with ZipFile(nom_zip, "w") as fichier_zip:
        # Traiter chaque étudiant
        print(f"Création des fiches de rétroaction pour {feuille_a_traiter.max_column - 1} élèves")
        for colonne in range(2, feuille_a_traiter.max_column + 1):
            numero_da = feuille_a_traiter.cell(column=colonne, row=ligne_da).value
            fichier_zip.write(
                traiter_eleve(dossier_sortie, numero_da, feuille_a_traiter, denominateur, colonne),
                f"{numero_da}.pdf"
                )

        fichier_zip.close()

    print("Création du chiffrier de sommaire des notes")
    sommaire_notes(feuille_a_traiter, dossier_sortie, nom_feuille_a_traiter, denominateur)

def valider_parametres(fichier_retroaction, dossier_sortie, nom_feuille_a_traiter, denominateur):
    """
        Valide l'ensemble des paramètres reçus en ligne de commande.
        Vérifie que le chiffrier contient bien les critères nécessaires.

        Paramètres
        ----------
        fichier_retroaction : str
            Nom et chemin du chiffrier Excel contenant les rétroactions
        dossier_sortie : str
            Chemin sur disque du dossier qui recevra le PDF
        nom_feuille_a_traiter : str
            Le nom de la feuille Excel qui contient les rétroactions à traiter pour l'élève.
        denominateur : int
            Le dénominateur de la note totale

        Retour
        ------
        True si tout est valide.
    """

    parametres_valides = True

    # Validation des paramètres
    if not os.path.isfile(fichier_retroaction):
        print(f"Le fichier d'entrée {fichier_retroaction} n'existe pas.")
        parametres_valides = False

    # Vérifier si le fichier d'entrée est un chiffrier Excel
    try:
        chiffrier = openpyxl.load_workbook(fichier_retroaction, data_only=True)

        # Vérifier si la feuille existe
        if nom_feuille_a_traiter not in chiffrier:
            print(f"La feuille {nom_feuille_a_traiter} n'existe pas.")
            parametres_valides = False
        else:

            # Valider si les critères de base sont présents
            criteres = trouver_lignes_criteres(chiffrier[nom_feuille_a_traiter])

            for cle, valeur in criteres.items():
                if valeur == 0:
                    print(f"Le critère {cle} n'existe pas dans le chiffrier.")
                    parametres_valides = False

        chiffrier.close()
    except BadZipFile:
        print(f"Le fichier d'entrée {fichier_retroaction} n'est pas un chiffrier Excel valide.")
        parametres_valides = False

    if not os.path.isdir(dossier_sortie):
        print(f"Le dossier de sortie {dossier_sortie} n'existe pas.")
        parametres_valides = False

    if denominateur < 1:
        print("Le dénominateur doit être plus grand que zéro.")
        parametres_valides = False

    return parametres_valides

def main(argv):
    """
        Procédure principale
    """

    fichier_retroaction = ''
    dossier_sortie = ''
    nom_feuille_a_traiter = ''
    denominateur = 0

    currentdir = os.getcwd()

    try:
        opts, _ = getopt.getopt(argv,"hi:o:s:d:")
    except getopt.GetoptError:
        affiche_aide()
        sys.exit(2)
    for opt, arg in opts:
        if opt == '-h':
            affiche_aide()
            sys.exit()
        elif opt == '-i':
            fichier_retroaction = os.path.join(currentdir, arg)
        elif opt == '-o':
            dossier_sortie = os.path.join(currentdir, arg)
        elif opt == '-s':
            nom_feuille_a_traiter = arg
        elif opt == '-d':
            denominateur = int(arg)

    if valider_parametres(fichier_retroaction, dossier_sortie, nom_feuille_a_traiter, denominateur):
        print(f'Fichier d\'entrée est : "{fichier_retroaction}"')
        print(f'Dossier de sortie est : "{dossier_sortie}"')
        print(f'Nom de la feuille est "{nom_feuille_a_traiter}"')
        print(f'La note est sur : {denominateur}')
        traiter_feuille(fichier_retroaction, dossier_sortie, nom_feuille_a_traiter, denominateur)

if __name__ == "__main__":
    main(sys.argv[1:])
